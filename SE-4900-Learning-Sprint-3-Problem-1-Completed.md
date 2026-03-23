SE 4900 Learning Sprint #3
Problem 1 Draft

Submission Note
This draft completes Problem 1 from `SE-4900-Learning-Sprint-3.docx` using the tagged Apache POI source at:
`https://svn.apache.org/repos/asf/poi/tags/REL_4_1_1/src/ooxml/java/org/apache/poi/xssf/usermodel/XSSFWorkbook.java`

The assignment asks for public code links such as GitHub Gists. For now, this draft points to local code artifacts:
- Manual refactoring code: `C:\Codex\learning_sprint_3_problem_1\manual\XSSFWorkbook_cloneSheet_manual_refactor.java`
- AI-assisted refactoring code: `C:\Codex\learning_sprint_3_problem_1\ai\XSSFWorkbook_cloneSheet_ai_refactor.java`

If you want public links before submission, upload those two files to GitHub Gists or your repo and replace the local paths below.

Part 1

Workflow 1: Manual Refactoring

Manual Refactoring Plan / Notes
The original `cloneSheet(int sheetNum, String newName)` method in `XSSFWorkbook` does too much in one place:
- validates the source sheet and target name
- creates the new sheet
- copies non-drawing relationships
- copies external package relationships
- serializes and re-reads the sheet XML
- strips unsupported cloned features
- rebuilds drawing state and drawing relationships

That makes the function hard to scan and increases the chance of accidental behavior changes during maintenance. My manual plan was:
1. Preserve the current behavior and operation order before changing names or extracting helpers.
2. Separate validation/name resolution from the clone mechanics.
3. Isolate relationship copying from worksheet body copying because they fail for different reasons and deserve different exception messages.
4. Treat drawing handling as a dedicated stage because the original implementation already has special-case logic there.
5. Keep unsupported-feature cleanup explicit and close to the worksheet XML copy so that it is obvious which cloned constructs POI still cannot preserve.
6. Improve variable names and add focused comments only where the workflow is non-obvious.

Manual Refactored Code Link
`C:\Codex\learning_sprint_3_problem_1\manual\XSSFWorkbook_cloneSheet_manual_refactor.java`

Why the Manual Refactor is Higher Quality
- The top-level method now reads as a small orchestration flow instead of a large procedural block.
- Each helper has a single responsibility and a narrow reason to change.
- Exception messages identify the failing clone stage instead of collapsing every failure into one generic message.
- Drawing handling is no longer buried inside unrelated relationship logic.
- Comments explain intent rather than narrating trivial syntax.

Workflow 2: AI-Assisted Refactoring

Prompt Chain Used for Workflow 2

Prompt 1
Refactor the Apache POI `XSSFWorkbook.cloneSheet(int sheetNum, String newName)` method from REL_4_1_1 into a higher-quality version. Keep behavior equivalent, split the logic into smaller highly cohesive helpers, improve naming, add concise comments, and preserve robust error handling.

Prompt 2
The refactor is too generic. Keep Apache POI domain details explicit. Separate these responsibilities into distinct helpers:
- name resolution and validation
- non-drawing relation copying
- external relationship copying
- worksheet XML cloning
- unsupported clone cleanup
- drawing reconstruction

Prompt 3
Do not swallow exceptions or silently change behavior. Keep POI-style runtime behavior by wrapping checked failures in `POIXMLException`, and keep warnings for unsupported comment/page-setup cloning.

Prompt 4
Tighten the method names and comments so the top-level `cloneSheet` reads like a workflow. Avoid adding new classes. Keep the result suitable for insertion into the existing `XSSFWorkbook` class.

AI-Assisted Refactored Code Link
`C:\Codex\learning_sprint_3_problem_1\ai\XSSFWorkbook_cloneSheet_ai_refactor.java`

Comparison Report

Code Quality
Both outputs are better than the original because both break the monolithic method into smaller helpers with clearer names and narrower responsibilities. The manual version is slightly stronger on cohesion because I was more conservative about preserving the original control-flow stages and explicitly named helpers around POI-specific concerns such as unsupported clone features and drawing recreation. The AI-assisted version also improved readability, but it initially tended to over-generalize the relationship-copy logic until the prompt forced it back toward POI-specific structure. That is a recurring pattern with AI on legacy code: it can quickly propose cleaner shapes, but it often needs guidance to avoid abstracting away domain-specific invariants.

The biggest correctness risk in this method is not syntax, it is preserving subtle behavior: special handling of drawings, reattachment of external relationships, and post-clone cleanup of unsupported legacy drawing/page-setup sections. The manual workflow was better at keeping those invariants front and center from the start. The AI-assisted workflow still reached a good result, but only after prompt refinement that explicitly protected those edge cases.

Efficiency and Cognitive Load
The manual workflow required more sustained reasoning because I had to identify responsibilities, preserve ordering constraints, and decide where the helper boundaries should be before writing any code. That makes manual refactoring slower and more mentally demanding, especially in legacy infrastructure code with XML/package relationships. The AI-assisted workflow was faster at producing a structured first draft, but some of that speed advantage was spent on review and correction. In particular, the AI needed explicit prompting to avoid flattening POI-specific behavior into overly generic helper methods. So the AI workflow reduced drafting effort, but not verification effort.

Prompt Engineering Effectiveness
The most effective prompts were the ones that constrained behavior rather than asking for “cleaner code” in the abstract. Telling the AI exactly which responsibilities must be split out, insisting on POI-specific naming, and explicitly prohibiting silent behavior changes made the output significantly better. The weakest prompt was the initial generic refactor request; it encouraged style improvements but not enough respect for workbook-specific invariants.

Pros and Cons
Manual refactoring is safer when behavior is subtle, poorly documented, and tightly coupled to framework/package mechanics. AI-assisted refactoring is superior when the main goal is to generate a first-pass structure quickly, rename variables, and propose decomposition candidates. The best practical workflow is to let AI accelerate draft generation, but keep humans responsible for protecting invariants, reviewing edge cases, and deciding whether the abstraction boundaries actually fit the codebase.

Actionable Takeaways
In future refactoring work, I would use AI early for decomposition ideas and naming alternatives, but I would not trust it to preserve behavior in stateful legacy code unless the prompts explicitly encode the invariants I care about. For code like `cloneSheet`, human review remains essential because the hardest part is not producing smaller methods; it is preserving the package, XML, and relationship semantics while doing so.

Part 2

Case A
Lines 368-385 in `onDocumentRead`

My Manual Code Review Recommendations
That block uses a long `if / else if` chain over `instanceof` checks to classify workbook relations. The immediate problem is readability: the variable `p` is too vague, and the logic mixes direct field assignment with map population for later processing. I would extract the loop body into a helper such as `indexRelationPart(...)` and rename the local variables to `relationPart` and `relatedPart`. That alone would reduce cognitive load and make the method easier to scan.

I would also make the classification behavior more explicit. Right now, unknown relation types are silently ignored. That may be intentional, but the code does not say so. I would either add a comment that unrecognized relation parts are intentionally skipped or log them at debug level for diagnostics. A second improvement would be to isolate the two map-population cases (`XSSFSheet` and `ExternalLinksTable`) from the simple direct assignments, because those relations are used later for ordered reconstruction and represent a different responsibility than just caching singleton workbook parts.

AI-Generated Recommendations
An AI reviewer would likely recommend replacing the `instanceof` chain with a dispatch table or helper method, renaming `p` to something descriptive, and consolidating repeated assignments into smaller methods for readability. It would probably also suggest logging or handling unknown relation types explicitly and adding documentation that explains why some relations are stored directly while others are first mapped by relationship id.

Evaluation of the AI Output
I mostly agree with those suggestions because the readability and cohesion issues are real. However, a generic AI review can miss an important nuance here: not every unrecognized relation should necessarily trigger a warning because POI document parts can be extensible, and noisy logging at workbook-load time could become a maintenance problem. The human reviewer has to judge whether “unknown” truly means suspicious in this code path.

Another likely AI blind spot is that the two map cases are not just another storage style; they exist because sheets and external links are reconstructed later in workbook-defined order. If the AI only comments on style and not on that deferred-ordering invariant, then the recommendation is incomplete. So the AI suggestions are directionally good, but they need human review to preserve intent and avoid overcorrecting with unnecessary abstraction or logging.

Case B
Exceptions and error handling in `setSheetName`

My Manual Code Review Recommendations
The current method covers the obvious caller errors: null name, invalid sheet index, invalid Excel sheet name, and duplicate names. That is a solid baseline, so I would not call the method unsafe. However, the exception messages are too generic for a public API. I would make them more specific by including the sheet index and conflicting sheet name. That would make debugging easier without changing the method contract.

I would also review the failure behavior around `utils.updateSheetName(...)`. If that internal update throws after validation but before the workbook XML is updated, the rename fails atomically, which is good. Still, the method does not document that formula/name updates are part of the operation, and it does not wrap unexpected internal failures with a message that tells the caller the rename failed during reference updates rather than basic validation. I would keep `IllegalArgumentException` for caller input issues, but I would consider wrapping unexpected internal rename/update failures in a richer runtime exception with context.

AI-Generated Recommendations
An AI reviewer would likely say the method should use clearer messages, centralize validation, possibly use `Objects.requireNonNull(sheetname, ...)`, and consider wrapping failures from formula updates in a workbook-specific exception. It might also suggest documenting the silent Excel-compatible truncation to 31 characters more clearly, or changing that behavior to fail fast instead of truncating.

Evaluation of the AI Output
The AI suggestions are mostly reasonable, especially on clearer diagnostics and better documentation. I also agree that `Objects.requireNonNull(...)` would be cleaner than the current manual null check. However, I would be careful with any AI suggestion to remove silent truncation. The comment explicitly states that POI is mimicking Excel behavior there, so changing it to fail fast would be a behavior change, not just an error-handling improvement.

The other common AI weakness here is overengineering the exception model. Introducing custom exceptions may look cleaner in isolation, but for a mature public API like POI it can be disruptive unless the broader API already follows that pattern. So I agree with the AI on message clarity and documentation, but I would reject any recommendation that changes Excel-compatibility semantics or expands the public exception surface without a stronger library-wide reason.
