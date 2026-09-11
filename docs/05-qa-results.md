# Profitly — Build & QA Results (Phase 5)

> Scope: **Phase 5 — build the workbook**, per the brief's development process (§26), implemented strictly against `01-architecture.md`, `02-data-layout.md`, `03-calculation-formulas.md`, and `04-ux-ui.md`. This document records what was built, how it was verified, every deviation from the approved specs (all flagged, none silent), and the QA results against Phase 3's 25-case edge matrix. **All 25 cases pass; the one finding QA surfaced (§5) was resolved with an approved design change, applied here and in the affected spec docs.**

Build artifacts: `build/build_workbook.py` (the generator, source of truth — re-running it reproduces the workbook deterministically) and `build/Profitly.xlsx` (the workbook itself). `build/qa_tests.py` is the QA harness used for §4.

---

## 1. What was built

- **19 sheets**, exactly the inventory from Phase 2 §2.1 (9 App Layer: `🧭 Menú`, `🏠 Inicio`, `⚙️ Configuración`, `📦 Productos`, `🛒 Ventas`, `💸 Gastos`, `🎯 Objetivos`, `🔬 Simulador`, `🎬 Demo`; 10 Engine Layer, all hidden: `_Config`, `_Listas`, `_GastosCalc`, `_DashboardData`, `_Objetivos_Calc`, `_Insights_Rules`, `_Insights_Engine`, `_Simulador_Engine`, `_Demo_Datos`, `_Textos`). No sheet was added or removed beyond Phase 2 §0's two already-approved amendments.
- **4 core tables** as native Excel Tables (`tbl_Productos` B–P, `tbl_Ventas` B–L, `tbl_Gastos` B–F, `tbl_Objetivos` B–C), each shipped with exactly one blank starter row (the standard empty-Table state) so the real workbook opens empty and ready for the user's own data.
- **141 named ranges** covering every `Cfg_/Lst_/Txt_/KPI_/Obj_/Ins_/Sim_/Demo` name Phase 2/3 specified.
- **Full formula layer**: every calculation from Phase 3 §2–§8, including the resolved `KPI_CostosVentaPeriodo`, the `Obj_Factibilidad` 90-day feasibility engine, the Insights rule/placeholder engine, and the 3-scenario Simulador engine with its baseline-snapshot isolation.
- **`_Demo_Datos`**: 8 fictional products, 111 sales spread over a rolling 90-day window, 11 expenses, and one goal — plus a Demo-scoped aggregation block (KPIs, product ranking, trend) built as extra rows within this same engine sheet rather than a new sheet, since Phase 1 didn't allocate one for it (documented in Phase 3 §0, not a new deviation).
- **UX layer**: the shared nav/header shell, the 3-state input/autofill/calculated cell styling, KPI/ranking/insight cards, empty-state banners and dashed-row prompts, the 5-step onboarding checklist (shared between Configuración and Inicio), and the Simulador's native column-grouping for Scenarios B/C (collapsed by default).
- **3 dashboard charts** (line: evolución; horizontal bar: ganancia por producto; stacked bar: progreso) on both `🏠 Inicio` and `🎬 Demo`, each sourced from a small fixed engine-sheet range, never a live range on a growing table.
- **Data validation**: dropdowns (Producto, Categoría, Canal, Tipo de gasto, Moneda, período rápido), numeric/decimal/whole-number/date bounds, the product-name uniqueness rule, and the Objetivos first-of-month + uniqueness rule.
- **Protection**: every calculated/locked cell is `Protection(locked=False)` only where Phase 2 §7's table says it should be unlocked; every other cell keeps Excel's locked-by-default state; every sheet has `sheet.protection.sheet = True` with sort/autofilter/insert-rows left allowed (per Phase 2 §7) and structural edits (insert/delete columns, edit objects) blocked; all 10 engine sheets are hidden (`sheet_state = "hidden"`) and fully locked; workbook structure protection (`lockStructure`) discourages unhiding/deleting/reordering sheets. This protection is intentionally **passwordless** — its purpose is to stop *accidental* edits (per the brief's own framing, §17), not to secure the file against a determined user, so a lost password was never an acceptable risk to introduce.

---

## 2. Verification method

Formula correctness cannot be taken on faith — every formula in this workbook was written once, then verified by actually recalculating it, not by re-reading the spec. This environment's LibreOffice was initially missing the Calc component entirely (only `libreoffice-core`/`-common` were installed, so *no* document — not even a `.txt` — could be loaded); it was installed (`apt-get install libreoffice-calc`) before any verification could begin, and is now available for this and future sessions.

Every build stage was checked with `scripts/recalc.py` (from the `xlsx` skill), which recalculates every formula via LibreOffice and reports any cell that resolves to a raw Excel error (`#VALUE!`, `#DIV/0!`, `#N/A`, etc.). **The final workbook recalculates all 1,073 formulas with zero errors.** Verification was always run against a disposable copy, never against the shipped `Profitly.xlsx` directly — recalculation round-trips the file through LibreOffice, which was empirically confirmed to silently drop at least one Excel-native feature on save (a `DataValidation` entry vanished on round-trip: 8 validations in the pristine file, 7 after a LibreOffice save). The shipped file is always the pristine, never-round-tripped `openpyxl` output.

---

## 3. Deviations from the approved specs (all flagged)

None of these change what the workbook does for the user — every one is an internal formula-authoring or wiring fix required to make the approved design actually compute correctly or actually function as specified.

### 3.1 `[@Column]` structured self-references replaced with plain relative references

Phase 3 specified Excel formulas using the `[@Producto]`-style "current row" structured reference for same-row self-references. Empirically, this token — valid Excel syntax — is not reliably evaluated when a workbook is authored programmatically (via `openpyxl`) rather than typed inside Excel itself: it silently resolved to `#N/A` under LibreOffice's recalculation, the only tool available here to verify formulas. Cross-table structured references (`tbl_X[Column]`) and `_xlfn.MAXIFS`/`_xlfn.MINIFS` were confirmed to work correctly and were kept as specified. Every same-row self-reference was rewritten as a plain relative cell reference (e.g. `B13` instead of `[@Producto]`). This changes no formula logic or dependency — Excel's native Table calculated-column auto-fill propagates a formula down a growing table identically regardless of which reference style it uses, so Phase 2's row-insertion resilience requirement is fully preserved.

### 3.2 Two formula-correctness bugs found and fixed during build

- **`Ventas.Venta total`/`Costo total de la venta`/`Ganancia` on a blank starter row**: `E` (auto-filled Precio de venta) resolves to `""` when no product is chosen yet, and `Venta total = Cantidad*Precio - Descuento` then multiplies a number by text, throwing `#VALUE!` on the very first empty row a user sees. Fixed with an explicit `IF(Producto="","",...)` guard on `J`/`K`/`L`, so an unfilled row renders blank rather than erroring or (worse) showing "este producto ya no está en tu catálogo" on a row that was simply never filled in.
- **`AND()` does not short-circuit in Excel/Calc**: three formulas (`_Insights_Rules` R1–R3, `_Objetivos_Calc`'s facturación/ventas-necesarias guards) had an `ISNUMBER(x)` check and an arithmetic expression on `x` as sibling arguments inside one `AND()`. Excel evaluates every `AND()` argument regardless of the others, so the arithmetic still ran — and errored — on a non-numeric `x` even though the sibling guard was `FALSE`. Rewritten as nested `IF(AND(guards...), arithmetic, FALSE)` throughout, which does short-circuit correctly.

### 3.3 Configuración's business-setting fields were wired backwards — fixed

While building, `⚙️ Configuración`'s Nombre del negocio / Tipo de negocio / Moneda / Objetivo mensual cells were drafted as formulas mirroring `_Config` (`=_Config!B2`, read-only display) rather than being the actual input cells the `Cfg_*` named ranges point to. Typing into them would have done nothing — a real functional gap, not just a formula error, and one `recalc.py` cannot catch since a display-only cell isn't a formula *error*. Fixed by making `Cfg_NombreNegocio`/`Cfg_TipoNegocio`/`Cfg_Moneda`/`Cfg_ObjetivoMensualDefault` point directly at Configuración's cells (`C11`/`C12`/`C13`/`C24`), which are now the true, plain input cells. `_Config` keeps only what's genuinely computed and never typed by the user: the active date-range pair (now correctly driven by Inicio's period-selector dropdown, per Phase 1/2's own description of that wiring, which the first draft had also missed and also fixed here) and the low-margin alert threshold.

### 3.4 One unexplained tool-specific artifact, routed around

A block of formulas placed at `_Demo_Datos` rows 100–110 consistently resolved to `#VALUE!` under LibreOffice's recalculation — including a formula reduced to a bare `=B101-B102` cell reference, and even a literal constant (`=655900-358249`) re-tested at that exact address — while the *identical* formula placed one row away, or at row 300, computed correctly every time, across repeated rebuilds and repeated recalculation passes. No merged cell, conditional formatting rule, data validation, named-range collision, or circular reference was found at that location despite a direct search. This reproduced deterministically but could not be explained; it does not correspond to anything this build's design places there. Rather than ship a workbook with an unexplained failure at that address, the block was moved to rows 300–310 (confirmed clean) — a relocation, not a logic change. This is flagged here as an unresolved tooling oddity, not asserted to be understood.

None of §3.1–§3.4 change any sheet, table, column, KPI, chart, or user-facing behavior beyond what was already approved.

---

## 4. QA results against the Phase 3 §10 edge-case matrix

Verified with real injected data and actual recalculated output (`build/qa_tests.py`), not by re-reading the formulas. Grouped into 7 test workbooks; each ran with **zero formula errors**.

| Test | Cases covered | What was injected | Result |
|---|---|---|---|
| A — vacío | 1, 4, 8, 14, 15 | Pristine, untouched workbook (0 productos/ventas/gastos, no objetivo) | Productos rankings show `Txt_SinVentas`; Inicio shows the onboarding banner ("Te falta 1 paso..."); Objetivos' `Ganancia diaria necesaria` = 0, not an error. All correct. |
| B — un producto, una venta | 2, 5, 7 | 1 product, all cost fields = 0, 1 sale | `Costo real de venta` = 0, `Margen` = 100%, dashboard Ventas/Ganancia real = 1000, Margen = 100%. All correct. |
| C — casos de margen | 3, 9, 10, 16, 17, 18 | 4 products (healthy margin, never sold, margin below threshold, priced below cost) + a discounted sale | Healthy-margin product: 55% (hand-verified). Unsold product: `Unidades vendidas` = 0, margin still computed. Below-threshold product: margin 13% < 15%, `Alerta margen bajo` flag = 1. Negative-margin product: margin **-40%**, shown as a real number (not blocked), flag = 1. Discount: `Venta total` correctly nets the discount; `Ganancia` correctly nets both the discount and the real cost. All hand-verified exact. |
| D — objetivo imposible | 14 | `tbl_Objetivos` row set to 0 | `Ganancia diaria necesaria` = 0 (not `#DIV/0!`); guarded chain confirmed via 0 formula errors. |
| E — cambio de precio | 19 | 1 product/sale, then the product's price edited afterward | Failed on the first pass — see §5 for the finding and the fix. **Retested after the fix**: sale recorded at $1,000 stayed at $1,000 (`Ventas!E10`, `J10`) after the product's price was edited to $2,000 in Productos. Passes. |
| F — números grandes y decimales | 23, 24 | Price ~$15M with cents, 8-decimal commission, quantity 9,999 | `Costo real de venta`, `Ganancia por unidad`, `Margen`, `Venta total`, `Ganancia` all hand-verified exact to the cent — no precision loss, no overflow. |
| G — simulador | 21 | (G1) 0 real sales; (G2) 1 product/sale + Δ Precio +10% | G1: simulator shows *"Necesitás cargar ventas reales antes de simular escenarios"*, not a projection off $0. G2: baseline Ganancia real $6,000 → projected $7,000 → **Diferencia +$1,000**, hand-verified exact against the Phase 3 §8.3 formula chain. |

Cases not run as a separate isolated test because they're already exercised by the demo dataset or are structural, not numerical, and were verified by direct inspection: **6** (111 demo sales), **11** (all demo products carry a commission %), **12** (Publicidad atribuida vs. Δ Publicidad kept as separate fields, per Phase 3 §3.3/§8.3), **13** (demo goal + progress panel), **20** (quantity is a plain multiplicand, no special-cased path exists to fail), **22** (Ventas' optional fields — Canal, Descuento, Otros costos — default to blank/0 without any formula requiring them, confirmed by the same blank-row guard in §3.2), **25** (every percentage in the workbook — Margen, Comisión, Progreso, all Simulador Δ-percent inputs — is stored and formatted as a 0–1 decimal throughout; the demo Margen value read back as `0.2695`, not `26.95` or `2695`, confirming the convention held).

**All 25 of 25 cases pass** (after the §5 fix).

---

## 5. Finding — resolved (option 1 approved)

**`Ventas.Precio de venta` and `Comisión %` could silently rewrite already-recorded sales when a product's price changed later.**

Phase 3 §3.2 originally designed these two columns as "auto-filled, editable": a lookup formula pre-fills the value the moment a product is chosen, and the user *may* overwrite it. QA Test E found that if the user does **not** overwrite it — the normal path, since nothing prompts them to — the cell keeps the live lookup formula forever. Recording a sale at $1,000, then later editing that product's price to $2,000 in `📦 Productos`, silently changed the historical sale's recorded price *and* revenue to $2,000 too, retroactively. This directly contradicted Phase 3 §10's own case-19 requirement ("a mid-catalog price change never rewrites historical sales") and the brief's own priority order (§27: mathematical accuracy ranks above automation/convenience). There is no macro-free way to "snapshot" a formula's result into a cell at the moment a row is completed.

Three options were presented; **option 1 was chosen**: `Precio de venta` and `Comisión %` are now plain required inputs, exactly like `Cantidad`/`Fecha` — no formula, no auto-fill, nothing to silently drift. The user types the price actually charged at the time of sale. This trades away the "one less thing to type" convenience for a guarantee that a recorded sale can never change after the fact, matching the brief's stated priority order.

Applied changes:
- `build/build_workbook.py`: `Ventas` columns E/G moved from category `"autofill"` to `"input"`; the lookup formulas were removed; new tooltips point the user to Productos as a manual reference instead of auto-filling; numeric validation (decimal ≥ 0 for E, 0–100% for G) was added, matching every other plain numeric input column; the shared legend was simplified to the 2 states actually used anywhere in the workbook now (no column uses "auto-filled, editable" any more).
- `docs/02-data-layout.md`, `docs/03-calculation-formulas.md`, `docs/04-ux-ui.md`: updated in place with amendment notes at each affected section (table definitions, formula spec, 3-state design system, Ventas UX spec, tooltip list) rather than silently rewritten — each amendment names what changed, why, and points back here.
- Retested (QA Test E, rerun after the fix): confirmed the historical sale's price no longer changes when the product's price is edited afterward (§4). Full workbook recalculated clean afterward: 0 errors across 1,073 formulas.

No other Phase 1–4 decision changed as a result of this fix.

---

## 6. What's confirmed against the pre-finalization checklist

| Requirement | Status |
|---|---|
| No formula errors exposed to users | ✅ 0 errors / 1,073 formulas, verified by recalculation, not inspection |
| Real data cannot be modified by the simulator | ✅ `_Simulador_Engine` reads only its own frozen baseline block; no formula path exists from Simulador back into `tbl_Productos`/`tbl_Ventas`/`tbl_Gastos`; confirmed by QA Test G1/G2 |
| Demo data cannot contaminate real data | ✅ `_Demo_Datos` and the real tables share no cell, formula, or named range; `🎬 Demo` is fully locked/read-only |
| Protected/calculated cells cannot be accidentally edited | ✅ every calculated cell is locked, every sheet has protection enabled; verified in the pristine (unrecalculated) file's saved state |
| Dashboard KPIs reconcile with underlying data | ✅ hand-verified in QA Tests B, C, F against the raw table data |
| Productos/Ventas/Gastos/Objetivos/Simulador stay correctly connected | ✅ confirmed for aggregation (Ventas→Productos rollups, Gastos→Ganancia real, Objetivos→Simulador baseline) and, after the §5 fix, for isolation (a later Productos price edit no longer reaches into historical Ventas rows) |

**Phase 5 (build), Phase 6 (demo), Phase 7 (protection/validation), and Phase 8 (QA) are complete.** All 25 edge cases pass; all pre-finalization checklist items are green.
