# TEST_REPORT

Template: `Travelling claim form - updated version.xlsx`
Template SHA256 before: `a3d78782754588db9223c0e751d8e892f9fc8cc282fa1cc8027c9d2f022e23e9`
Template SHA256 after: `a3d78782754588db9223c0e751d8e892f9fc8cc282fa1cc8027c9d2f022e23e9`
Template hash unchanged: **PASS**

| Test | Data | Trips | Result | Actual Result |
|---|---:|---:|---|---|
| A | 1 trip MTR | 1 | PASS | Exported and verified case_A.xlsx. |
| B | 5 trips different transport types | 5 | PASS | Exported and verified case_B.xlsx. |
| C | Decimal amount HK$12.50 | 1 | PASS | Exported and verified case_C.xlsx. |
| D | With deduction | 1 | PASS | Exported and verified case_D.xlsx. |
| E | Fill all trip rows | 44 | PASS | Exported and verified case_E.xlsx. |
| F | Exceed max rows | 45 | PASS | Correctly blocked: Template supports 44 trips only. |

## Verified
- Workbook opens with openpyxl after export.
- Sheet count and sheet names match the template.
- Merged-cell counts match the template.
- Mapped cells contain expected values.
- Numeric amount and deduction cells are numbers.
- Formula count is unchanged. The supplied template contains no formula cells.
- No `#REF!`, `#VALUE!` or `#NAME?` strings were found.
- Template source file hash is unchanged.

## Known Limitations
- The supplied template has no separate cells for From and To because C:D is merged on each trip row; the app writes one combined location value.
- The supplied template has no Other transport column; Other is annotated in the location text and included in Sub Total.
- The supplied template has no row-level Remarks export cell, so trip remarks remain in app draft/history/preview.
- Full visual fidelity in Apple Numbers and Microsoft 365 is expected because the package parts are preserved, but automated tests here verify workbook structure and cell data rather than opening those apps.