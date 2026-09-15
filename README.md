# Profitly — Centro de Control de Rentabilidad

Premium Excel/Google Sheets profitability-management product for small business owners. All user-facing content is in Spanish; internal docs are in English for the dev team.

## Status

Following the phased build process defined in the product brief:

- [x] **Phase 1 — Technical architecture**: [`docs/01-architecture.md`](docs/01-architecture.md)
- [x] **Phase 2 — Data model & relationships (exact layout)**: [`docs/02-data-layout.md`](docs/02-data-layout.md)
- [x] **Phase 3 — Calculations & financial logic (final formulas)**: [`docs/03-calculation-formulas.md`](docs/03-calculation-formulas.md)
- [x] **Phase 4 — UX/UI structure**: [`docs/04-ux-ui.md`](docs/04-ux-ui.md)
- [x] **Phase 5 — Build the workbook**: [`build/Profitly.xlsx`](build/Profitly.xlsx) / [`build/build_workbook.py`](build/build_workbook.py)
- [x] **Phase 6 — Populate the demo**: built into Phase 5 (`_Demo_Datos`)
- [x] **Phase 7 — Protect formulas & validate inputs**: built into Phase 5
- [x] **Phase 8 — QA**: [`docs/05-qa-results.md`](docs/05-qa-results.md) — all 25 edge cases pass
- [x] **Phase 9 — Beginner-perspective review**: [`docs/06-beginner-review-and-polish.md`](docs/06-beginner-review-and-polish.md)
- [x] **Phase 10 — Visual polish**: same doc, §3 — final QA rerun against the actual shipped file, all green
- [x] **Excel compatibility fix**: [`docs/07-excel-compatibility-fix.md`](docs/07-excel-compatibility-fix.md) — the shipped file triggered Excel's repair prompt. Three rounds of root-causing against the actual package: (1) invalid OPC relationships behind internal navigation links, (2) malformed data-validation formulas + misaligned chart data ranges, (3) a `CT_Font` element-ordering schema violation on every font, caught only by Microsoft's own official `OpenXmlValidator`. All three fixed and re-verified; the file now validates with 0 schema errors across all 5 Office format versions (Office2007–Microsoft365)

**Profitly is complete.**
