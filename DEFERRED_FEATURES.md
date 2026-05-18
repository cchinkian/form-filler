# Deferred / Legacy Features

Last updated: 2026-05-19

This file records features that are intentionally not part of the main product workflow now. They are kept as implementation reference only, not as daily-use UI concepts.

## Product Workflow Kept

The active product workflow is:

1. Search customer
2. Choose form
3. Fill missing fields
4. Generate PDF

This covers the normal sales use case without requiring the user to open Excel.

## Deferred

### Excel-first bulk workflow

Previous flow:

- Choose application
- Open Excel
- Fill a product sheet
- Select rows
- Execute all

Reason deferred:

- It makes Excel the daily operating surface.
- It adds extra steps when the user only needs to generate one customer's form.
- It conflicts with the new sales flow where Master DB autocomplete should provide basic customer fields.

Current decision:

- Do not expose this as the main UI.
- Keep old code temporarily as reference until the sales flow has been tested in real use.

### Dynamic PDF module combinations

Previous idea:

- Treat forms as separate modules such as A, B, C, D.
- Dynamically combine modules into ABC, ABD, CDF, etc.

Reason deferred:

- It increases coordinate and PDF ordering complexity.
- The actual use case is better served by a fixed folder per complete merged form package.

Current decision:

- Each form package folder contains exactly one merged PDF.
- Folder identity controls the form; PDF filename is ignored.

### PDF filename surrender / hash blocking

Previous flow:

- Store `template_filename` and `template_hash`.
- Block normal filling when the filename or hash changes.

Reason removed:

- The user wants to rename PDFs freely.
- The folder should be the stable identity, not the PDF filename.

Current decision:

- If a form folder has exactly one top-level PDF, use it.
- If there are multiple PDFs, show an ambiguity warning.

### Preview-first workflow

Previous idea:

- Refresh a preview page before generating final output.

Reason deferred:

- It is not needed for the first usable version.
- The current priority is to reduce data-entry friction.

Current decision:

- Generate directly.
- Add preview later only if real use shows enough mistakes that preview would save time.
