// Manual workflow refactoring excerpt for Apache POI XSSFWorkbook.cloneSheet(...)
// Intended to be inserted into XSSFWorkbook alongside the existing addRelation(...) helper.

/**
 * Clone an existing sheet while preserving known workbook relationships and sheet content.
 *
 * <p>This refactoring keeps the original behavior but splits the operation into smaller,
 * purpose-specific helpers so the control flow is easier to understand, review, and test.</p>
 *
 * @param sheetIndex the source sheet index to clone
 * @param requestedName the new sheet name, or {@code null} to auto-generate one
 * @return the cloned sheet
 * @throws IllegalArgumentException if the sheet index or provided sheet name is invalid
 * @throws POIXMLException if POI cannot copy the sheet XML or its relationships
 */
public XSSFSheet cloneSheet(int sheetIndex, String requestedName) {
    validateSheetIndex(sheetIndex);

    XSSFSheet sourceSheet = sheets.get(sheetIndex);
    String cloneName = resolveCloneSheetName(sourceSheet, requestedName);
    XSSFSheet clonedSheet = createSheet(cloneName);

    XSSFDrawing sourceDrawing = copyNonDrawingRelationships(sourceSheet, clonedSheet);
    copyExternalRelationships(sourceSheet, clonedSheet);
    copyWorksheetXml(sourceSheet, clonedSheet);
    removeUnsupportedCloneFeatures(clonedSheet);
    clonedSheet.setSelected(false);
    cloneDrawingIfPresent(sourceSheet, clonedSheet, sourceDrawing);

    return clonedSheet;
}

private String resolveCloneSheetName(XSSFSheet sourceSheet, String requestedName) {
    if (requestedName == null) {
        return getUniqueSheetName(sourceSheet.getSheetName());
    }

    validateSheetName(requestedName);
    return requestedName;
}

/**
 * Copy all known internal relationships except the drawing, which must be recreated
 * against the cloned sheet to keep package references consistent.
 */
private XSSFDrawing copyNonDrawingRelationships(XSSFSheet sourceSheet, XSSFSheet clonedSheet) {
    XSSFDrawing sourceDrawing = null;

    for (RelationPart relationPart : sourceSheet.getRelationParts()) {
        POIXMLDocumentPart relatedPart = relationPart.getDocumentPart();
        if (relatedPart instanceof XSSFDrawing) {
            sourceDrawing = (XSSFDrawing) relatedPart;
            continue;
        }

        addRelation(relationPart, clonedSheet);
    }

    return sourceDrawing;
}

/**
 * Copy package-level external relationships such as hyperlinks or external references.
 */
private void copyExternalRelationships(XSSFSheet sourceSheet, XSSFSheet clonedSheet) {
    try {
        for (PackageRelationship relationship : sourceSheet.getPackagePart().getRelationships()) {
            if (relationship.getTargetMode() != TargetMode.EXTERNAL) {
                continue;
            }

            clonedSheet.getPackagePart().addExternalRelationship(
                relationship.getTargetURI().toASCIIString(),
                relationship.getRelationshipType(),
                relationship.getId()
            );
        }
    } catch (InvalidFormatException e) {
        throw new POIXMLException("Failed to copy external relationships while cloning sheet '"
            + sourceSheet.getSheetName() + "'.", e);
    }
}

/**
 * Copy the underlying worksheet XML by writing the source sheet and reading it back
 * into the newly created sheet part.
 */
private void copyWorksheetXml(XSSFSheet sourceSheet, XSSFSheet clonedSheet) {
    try (ByteArrayOutputStream output = new ByteArrayOutputStream()) {
        sourceSheet.write(output);
        try (ByteArrayInputStream input = new ByteArrayInputStream(output.toByteArray())) {
            clonedSheet.read(input);
        }
    } catch (IOException e) {
        throw new POIXMLException("Failed to copy worksheet XML while cloning sheet '"
            + sourceSheet.getSheetName() + "'.", e);
    }
}

/**
 * Remove sheet features that the current implementation still cannot clone safely.
 */
private void removeUnsupportedCloneFeatures(XSSFSheet clonedSheet) {
    CTWorksheet worksheet = clonedSheet.getCTWorksheet();

    if (worksheet.isSetLegacyDrawing()) {
        logger.log(POILogger.WARN, "Cloning sheets with comments is not yet supported.");
        worksheet.unsetLegacyDrawing();
    }

    if (worksheet.isSetPageSetup()) {
        logger.log(POILogger.WARN, "Cloning sheets with page setup is not yet supported.");
        worksheet.unsetPageSetup();
    }
}

private void cloneDrawingIfPresent(XSSFSheet sourceSheet, XSSFSheet clonedSheet, XSSFDrawing sourceDrawing) {
    if (sourceDrawing == null) {
        return;
    }

    clearExistingDrawingReference(clonedSheet.getCTWorksheet());

    XSSFDrawing clonedDrawing = clonedSheet.createDrawingPatriarch();
    clonedDrawing.getCTDrawing().set(sourceDrawing.getCTDrawing());

    for (RelationPart relationPart : sourceSheet.createDrawingPatriarch().getRelationParts()) {
        addRelation(relationPart, clonedDrawing);
    }
}

private void clearExistingDrawingReference(CTWorksheet worksheet) {
    if (worksheet.isSetDrawing()) {
        worksheet.unsetDrawing();
    }
}
