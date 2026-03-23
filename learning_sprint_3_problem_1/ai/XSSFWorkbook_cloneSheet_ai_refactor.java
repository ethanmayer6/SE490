// AI-assisted workflow refactoring excerpt for Apache POI XSSFWorkbook.cloneSheet(...)
// Intended to be inserted into XSSFWorkbook alongside the existing addRelation(...) helper.

/**
 * Clone a sheet using a staged workflow that keeps validation, relationship copying,
 * worksheet content copying, and drawing reconstruction isolated from one another.
 *
 * @param sheetIndex the source sheet index
 * @param requestedName the desired name for the clone, or {@code null} to auto-generate one
 * @return a cloned XSSFSheet with copied worksheet XML and supported relationships
 * @throws IllegalArgumentException if the sheet index or explicit name is invalid
 * @throws POIXMLException if any clone step fails
 */
public XSSFSheet cloneSheet(int sheetIndex, String requestedName) {
    validateSheetIndex(sheetIndex);

    XSSFSheet sourceSheet = sheets.get(sheetIndex);
    XSSFSheet clonedSheet = createSheet(resolveRequestedCloneName(sourceSheet, requestedName));

    XSSFDrawing sourceDrawing = copySheetPartRelationships(sourceSheet, clonedSheet);
    copyExternalSheetRelationships(sourceSheet, clonedSheet);
    cloneWorksheetBody(sourceSheet, clonedSheet);
    prepareClonedWorksheet(clonedSheet);
    recreateDrawingAndRelationships(sourceSheet, clonedSheet, sourceDrawing);

    clonedSheet.setSelected(false);
    return clonedSheet;
}

private String resolveRequestedCloneName(XSSFSheet sourceSheet, String requestedName) {
    if (requestedName == null) {
        return getUniqueSheetName(sourceSheet.getSheetName());
    }

    validateSheetName(requestedName);
    return requestedName;
}

/**
 * Copy all supported POIXML relations from the source sheet except the drawing,
 * which must be reconstructed against the clone after the sheet XML is copied.
 */
private XSSFDrawing copySheetPartRelationships(XSSFSheet sourceSheet, XSSFSheet clonedSheet) {
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

private void copyExternalSheetRelationships(XSSFSheet sourceSheet, XSSFSheet clonedSheet) {
    try {
        for (PackageRelationship relationship : sourceSheet.getPackagePart().getRelationships()) {
            if (relationship.getTargetMode() == TargetMode.EXTERNAL) {
                clonedSheet.getPackagePart().addExternalRelationship(
                    relationship.getTargetURI().toASCIIString(),
                    relationship.getRelationshipType(),
                    relationship.getId()
                );
            }
        }
    } catch (InvalidFormatException e) {
        throw new POIXMLException("Failed to clone external relationships for sheet '"
            + sourceSheet.getSheetName() + "'.", e);
    }
}

private void cloneWorksheetBody(XSSFSheet sourceSheet, XSSFSheet clonedSheet) {
    try (ByteArrayOutputStream output = new ByteArrayOutputStream()) {
        sourceSheet.write(output);
        try (ByteArrayInputStream input = new ByteArrayInputStream(output.toByteArray())) {
            clonedSheet.read(input);
        }
    } catch (IOException e) {
        throw new POIXMLException("Failed to clone worksheet XML for sheet '"
            + sourceSheet.getSheetName() + "'.", e);
    }
}

/**
 * Normalize unsupported worksheet structures after the XML copy so the clone
 * does not retain references POI cannot safely reproduce.
 */
private void prepareClonedWorksheet(XSSFSheet clonedSheet) {
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

private void recreateDrawingAndRelationships(XSSFSheet sourceSheet, XSSFSheet clonedSheet, XSSFDrawing sourceDrawing) {
    if (sourceDrawing == null) {
        return;
    }

    CTWorksheet clonedWorksheet = clonedSheet.getCTWorksheet();
    if (clonedWorksheet.isSetDrawing()) {
        clonedWorksheet.unsetDrawing();
    }

    XSSFDrawing clonedDrawing = clonedSheet.createDrawingPatriarch();
    clonedDrawing.getCTDrawing().set(sourceDrawing.getCTDrawing());

    for (RelationPart drawingRelation : sourceSheet.createDrawingPatriarch().getRelationParts()) {
        addRelation(drawingRelation, clonedDrawing);
    }
}
