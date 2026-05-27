# Procedure Automation Release Checklist

Use this checklist before releasing the package/procedure automation flow.

| User brief | Release check |
| --- | --- |
| Excel-like customer data fills packages | Run `python3 tests/procedure_package_smoke.py` and confirm a customer-row dict populates source form fields. |
| One procedure can include multiple source forms | Confirm the smoke fixture fills two distinct temporary source PDFs. |
| Procedure step order is respected | Confirm the combined PDF pages follow `StepNo`: first form, blank page, second form, even when fixture items are unsorted. |
| Blank pages can be inserted | Confirm the combined PDF has a blank middle page with no extracted text. |
| One combined output PDF is generated | Confirm only one PDF is written to the output folder for the procedure run. |
| CIS is not exposed in output filenames | Confirm generated filenames contain neither `CIS` nor the customer's CIS value. |
| Output remains searchable/extractable | Confirm `pypdf` can extract the overlaid customer text from the filled pages. |
| No repo catalog/config dependency in smoke coverage | Confirm the smoke test uses temporary PDFs and temporary JSON fixtures, not `config/*.json`. |

Manual pre-release checks:

- Run the existing single-form smoke test: `python3 tests/sales_flow_smoke.py`.
- Run the procedure package smoke test: `python3 tests/procedure_package_smoke.py`.
- Generate one real procedure package from the app UI using a non-production/sample customer row.
- Inspect the output filename and PDF page order before distribution.
