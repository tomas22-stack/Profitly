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
- [⚠️] **Phase 8 — QA**: [`docs/05-qa-results.md`](docs/05-qa-results.md) — 24/25 edge cases verified correct; **1 open finding blocks sign-off** (§5 of that doc: Ventas price/commission auto-fill can retroactively rewrite historical sales — needs a decision)
- [ ] Phase 9 — Beginner-perspective review
- [ ] Phase 10 — Visual polish
