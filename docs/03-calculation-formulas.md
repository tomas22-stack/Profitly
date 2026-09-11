# Profitly — Final Calculation Formulas (Phase 3)

> Scope: **Phase 3 — final calculation formulas**, per the brief's development process (§26). This document supersedes the "draft syntax" formulas in `02-data-layout.md` with final, verified Excel **and** Google Sheets syntax, resolves every item Phase 2 §8 deferred, and maps every formula to the edge cases the workbook must survive (Phase 1 §21). No workbook, UI, or styling is built here. Phase 1 and Phase 2 stand except for the two amendments in §0, which are additive column insertions — nothing already specified is removed or renamed beyond what §0 states.

---

## 0. Amendments to Phase 2 (flagged, with justification)

Both amendments were already anticipated as open items in Phase 2 §8; this section resolves them concretely rather than introducing new scope.

### 0.1 `tbl_Ventas` gains a new calculated column: `Costo total de la venta`

**Problem:** Phase 2 flagged that `KPI_CostosVentaPeriodo` (Ventas del período's cost-of-sale total) can't be a plain `SUMIFS` because the per-row cost figure (`Costo real de venta unitario × Cantidad`) doesn't exist as a column anywhere — it was only ever computed inline inside `tbl_Ventas[Ganancia]`'s formula. Without a standalone column, aggregating it requires a `SUMPRODUCT` over the whole table on every dashboard recalculation, which is exactly the "scattered full-column array formula" pattern Phase 1 §1 rules out on performance grounds.

**Resolution:** Add column **K** to `tbl_Ventas`: **`Costo total de la venta`** (locked, calculated). This shifts the existing `Ganancia` column from K to **L**, and `Ganancia`'s formula simplifies to reference the new column instead of recomputing the lookup. `KPI_CostosVentaPeriodo` becomes a plain `SUMIFS` over this column (§5.1). Net effect on Phase 2's table: `tbl_Ventas` now has 11 columns (B–L) instead of 10 (B–K).

### 0.2 `tbl_Productos` gains a new calculated column: `Alerta margen bajo`

**Problem:** Phase 2's draft formula for the "⚠️ Producto con margen bajo" ranking card was `MATCH(TRUE, (Margen<umbral)*(Unidades>0), 0)` — a multi-condition array match. Modern Excel (365/2021+) and Google Sheets evaluate this natively, but older Excel (2019 and earlier, still common among small-business users) requires it to be entered as a legacy Ctrl+Shift+Enter array formula, which a non-technical user could accidentally break by re-entering the cell normally — directly contradicting the brief's "user must never touch a formula, and if they do, it must not silently corrupt" posture (§2, §17).

**Resolution:** Add column **P** to `tbl_Productos`: **`Alerta margen bajo`** (locked, calculated, boolean 1/0) = `IF(AND(Margen < Cfg_MargenAlertaUmbral, Unidades vendidas > 0), 1, 0)`. The ranking card then does a plain, non-array `MATCH(1, tbl_Productos[Alerta margen bajo], 0)`, which every supported Excel version and Sheets handle identically with zero array-entry risk. `tbl_Productos` now has 15 columns (B–P) instead of 14 (B–O).

No sheet, KPI, chart, insight rule, or Simulator variable from Phase 1/2 is added, removed, or renamed beyond these two additive columns.

---

## 1. Cross-Platform Formula Conventions

Stated once here, applied consistently below instead of being re-explained per formula.

| Concern | Excel | Google Sheets |
|---|---|---|
| Table/column references | Native structured references: `tbl_Productos[Margen]`, `[@Producto]` | No true structured references — plain bounded ranges are used instead (see row caps below), referenced as `Productos!$B$12:$B$511` etc. |
| Row caps | Not needed — Tables auto-expand natively with no pre-declared limit. | Every table gets a **pre-declared, generous row cap** so range-based formulas are finite and non-volatile. Caps (tunable, not a hard technical ceiling): `tbl_Productos` → 500 rows (12–511), `tbl_Ventas` → 5,000 rows (8–5007), `tbl_Gastos` → 5,000 rows (8–5007), `tbl_Objetivos` → 120 rows (10 years monthly). Sized for "realistic small-business dataset" per Phase 1 §23; a build note for Phase 5 is to make these easy to extend without touching every formula (i.e., named ranges for the caps themselves, not literal row numbers scattered through formulas). |
| Calculated columns that are **never** manually overridden (e.g. `Costo real de venta`, `Margen`, `Ganancia`, `Costo total de la venta`) | Excel Table calculated column: enter the formula once in the first data row; Excel auto-propagates it to every row the table expands into. | One `ARRAYFORMULA` per column, entered once in the first data-row cell, spanning the full row cap, wrapped in a blank-row guard so unused rows render empty instead of `$0`/`0%`: `=ARRAYFORMULA(IF(<key column range>="", "", <calculation>))`. |
| Calculated columns that are **auto-filled defaults but user-editable** (`Ventas.Precio de venta`, `Ventas.Comisión %`) | Same Excel Table calculated-column behavior; typing over one cell simply replaces that row's formula with a static value, per Excel's normal Table behavior — no special handling needed. | **Cannot** use `ARRAYFORMULA` here — Sheets refuses to let a user edit a single cell inside an array formula's output range, which would make the "editable override" requirement impossible. Instead, the same per-row formula is **pre-filled individually into every row of the row cap at workbook build time** (Phase 5/6 task), exactly like Excel's behavior, so every row — whether the 8th or the 4,999th — already carries the lookup formula the moment the user starts typing into it, and overriding one cell only affects that cell. |
| Lookups | `INDEX/MATCH` (not `VLOOKUP`, to stay resilient to column reordering; not `XLOOKUP`, for older-Excel compatibility) | Same `INDEX/MATCH` pattern; when the match value is itself a range (i.e. computing a whole column at once), the whole expression is wrapped in `ARRAYFORMULA` — Sheets iterates `MATCH`/`SUMIFS`/`COUNTIFS`/`AVERAGEIFS` correctly per row when a range is passed as the "criteria value" argument inside `ARRAYFORMULA`. This is the exact mechanism that resolves Phase 2's open item on Sheets calculated-column equivalence. |
| Aggregation | `SUMIFS`/`COUNTIFS`/`AVERAGEIFS`/`MAXIFS` | Identical syntax and argument order — this is one of the reasons Phase 1 chose these functions over alternatives. |
| Error suppression | `IFERROR(expr, Txt_*)` | Identical. |
| Volatile functions | `TODAY()` used only for period defaults and date-bound checks (Phase 1/2 already accepted this); `OFFSET`/`INDIRECT` not used anywhere in this spec. | Same. |

---

## 2. `tbl_Productos` — Final Formulas

Header row 11, data rows 12–511 (Sheets cap). Columns B–J are input (unchanged from Phase 2). Columns K–P below are final.

| Col | Field | Category | Excel | Google Sheets (entered once, row 12) |
|---|---|---|---|---|
| K | Costo real de venta | Calc, never overridden | `=[@Costo del producto] + ([@Precio de venta]*[@Comisión %]) + [@Packaging] + [@Envío / costo asociado] + [@Publicidad atribuida] + [@Otros costos]` | `=ARRAYFORMULA(IF($B$12:$B$511="", "", $E$12:$E$511+($D$12:$D$511*$F$12:$F$511)+$G$12:$G$511+$H$12:$H$511+$I$12:$I$511+$J$12:$J$511))` |
| L | Ganancia por unidad | Calc | `=[@Precio de venta] - [@Costo real de venta]` | `=ARRAYFORMULA(IF($B$12:$B$511="", "", $D$12:$D$511-$K$12:$K$511))` |
| M | Margen | Calc, guarded | `=IF([@Precio de venta]=0, Txt_SinDatos, [@Ganancia por unidad]/[@Precio de venta])` | `=ARRAYFORMULA(IF($B$12:$B$511="", "", IF($D$12:$D$511=0, Txt_SinDatos, $L$12:$L$511/$D$12:$D$511)))` |
| N | Unidades vendidas | Calc | `=SUMIFS(tbl_Ventas[Cantidad], tbl_Ventas[Producto], [@Producto])` | `=ARRAYFORMULA(IF($B$12:$B$511="", "", SUMIF(Ventas!$C$8:$C$5007, $B$12:$B$511, Ventas!$D$8:$D$5007)))` |
| O | Ganancia total | Calc | `=SUMIFS(tbl_Ventas[Ganancia], tbl_Ventas[Producto], [@Producto])` | `=ARRAYFORMULA(IF($B$12:$B$511="", "", SUMIF(Ventas!$C$8:$C$5007, $B$12:$B$511, Ventas!$L$8:$L$5007)))` |
| P | Alerta margen bajo | Calc, new (§0.2) | `=IF(AND(ISNUMBER([@Margen]), [@Margen]<Cfg_MargenAlertaUmbral, [@Unidades vendidas]>0), 1, 0)` | `=ARRAYFORMULA(IF($B$12:$B$511="", "", IF((ISNUMBER($M$12:$M$511))*($M$12:$M$511<Cfg_MargenAlertaUmbral)*($N$12:$N$511>0), 1, 0)))` |

`N`/`O` use `SUMIF` rather than `SUMIFS` in the Sheets version specifically because they have only one criteria pair — `SUMIF` accepts a range as its criteria argument under `ARRAYFORMULA` slightly more predictably than `SUMIFS` across Sheets versions; both are non-volatile and equivalent here. Where a formula needs 2+ criteria (period-bounded aggregates, §5), `SUMIFS` is used in both platforms per §1.

**Why no `IFERROR` on K/L:** these are pure arithmetic over numeric inputs that default to 0 when blank — there is no division, lookup, or match that can fail, so wrapping them would hide a real error rather than a business-logic edge case. This is a general rule applied throughout: **only wrap a formula in `IFERROR`/`IF`-guards when it contains a division, `MATCH`, or `INDEX` that can genuinely fail** — not reflexively on every cell. Over-guarding a pure-arithmetic cell would make a genuine bug (e.g. a broken named range) silently disappear into `Txt_SinDatos` instead of surfacing during QA.

### 2.1 Rankings block (rows 6–8, above the table — unchanged position from Phase 2)

| Card | Excel | Sheets |
|---|---|---|
| 🏆 Producto más rentable | `=IF(COUNTIFS(tbl_Productos[Unidades vendidas],">0")=0, Txt_SinVentas, INDEX(tbl_Productos[Producto], MATCH(MAXIFS(tbl_Productos[Margen], tbl_Productos[Unidades vendidas], ">0"), tbl_Productos[Margen], 0)))` | `=IF(COUNTIFS($N$12:$N$511,">0")=0, Txt_SinVentas, INDEX($B$12:$B$511, MATCH(MAXIFS($M$12:$M$511,$N$12:$N$511,">0"), $M$12:$M$511, 0)))` |
| 🔥 Producto más vendido | `=IF(COUNTA(tbl_Productos[Producto])=0, Txt_SinProductos, IF(MAX(tbl_Productos[Unidades vendidas])=0, Txt_SinVentas, INDEX(tbl_Productos[Producto], MATCH(MAX(tbl_Productos[Unidades vendidas]), tbl_Productos[Unidades vendidas], 0))))` | same pattern with bounded ranges |
| 💰 Producto que más dinero genera | `=IF(COUNTA(tbl_Productos[Producto])=0, Txt_SinProductos, IF(MAX(tbl_Productos[Ganancia total])<=0, Txt_SinVentas, INDEX(tbl_Productos[Producto], MATCH(MAX(tbl_Productos[Ganancia total]), tbl_Productos[Ganancia total], 0))))` | same pattern |
| ⚠️ Producto con margen bajo | `=IFERROR(INDEX(tbl_Productos[Producto], MATCH(1, tbl_Productos[Alerta margen bajo], 0)), Txt_SinAlertas)` | `=IFERROR(INDEX($B$12:$B$511, MATCH(1, $P$12:$P$511, 0)), Txt_SinAlertas)` |

`Txt_SinAlertas` (new message, §10): *"Ninguno de tus productos tiene el margen bajo. ¡Buen trabajo!"* — this is a positive-case message, not an error; it's included in `_Textos` because it's the natural "nothing to show" state for that specific card, distinct from `Txt_SinProductos`/`Txt_SinVentas`.

---

## 3. `tbl_Ventas` — Final Formulas

Header row 7, data rows 8–5007 (Sheets cap). Columns B–J input. Columns K–L below are final, reflecting §0.1.

> **Amendment (post-build, approved):** `E` (Precio de venta) and `G` (Comisión %) were originally "auto-filled, overridable" via the lookup formulas struck through below. QA found this unsafe: a user who never retypes the auto-filled value leaves a live formula in the cell, so a later price edit in `tbl_Productos` silently rewrites that historical sale's recorded price and revenue — violating the case-19 requirement in §10 below. Approved fix: `E`/`G` are now **plain required inputs**, identical in kind to `Cantidad`/`Fecha` — no formula, no default, nothing to silently drift. Full detail: `docs/05-qa-results.md` §5.

| Col | Field | Category | Excel | Google Sheets |
|---|---|---|---|---|
| E | Precio de venta | **Input** (plain, required — not auto-filled; see amendment above) | — | — |
| G | Comisión % | **Input** (plain, required — not auto-filled; see amendment above) | — | — |
| J | Venta total | Calc, never overridden | `=([@Cantidad]*[@Precio de venta]) - [@Descuento]` | `=ARRAYFORMULA(IF($C$8:$C$5007="", "", ($D$8:$D$5007*$E$8:$E$5007)-$H$8:$H$5007))` |
| K | **Costo total de la venta** (new, §0.1) | Calc, guarded | `=IFERROR(INDEX(tbl_Productos[Costo real de venta], MATCH([@Producto], tbl_Productos[Producto], 0)) * [@Cantidad], Txt_ProductoNoEncontrado)` | `=ARRAYFORMULA(IF($C$8:$C$5007="", "", IFERROR(INDEX(Productos!$K$12:$K$511, MATCH($C$8:$C$5007, Productos!$B$12:$B$511, 0)) * $D$8:$D$5007, Txt_ProductoNoEncontrado)))` |
| L | Ganancia (was K in Phase 2) | Calc, guarded | `=IFERROR([@Venta total] - [@Costo total de la venta] - [@Otros costos], Txt_ProductoNoEncontrado)` | `=ARRAYFORMULA(IF($C$8:$C$5007="", "", IFERROR($J$8:$J$5007-$K$8:$K$5007-$I$8:$I$5007, Txt_ProductoNoEncontrado)))` |

**Edge case — product deleted after it has recorded sales:** if a user removes a product row from `tbl_Productos` that already has sales in `tbl_Ventas` (the dropdown validation only constrains *new* entries, not existing rows if the source list changes), `K`'s `MATCH` fails and both `K` and `L` for that sale resolve to `Txt_ProductoNoEncontrado` ("Este producto ya no está en tu catálogo.") instead of silently defaulting the cost to 0 — which would have overstated profit. Because `SUMIFS`/`SUMIF` skip non-numeric cells automatically, that sale's `Venta total` still counts toward `Ventas del período`, but its cost/profit is correctly excluded rather than wrongly zeroed — the dashboard shows a smaller, accurate `Ganancia real` rather than an inflated wrong one. This is a deliberate "fail visibly, not silently" choice per Phase 1 §18, and is why `tbl_Productos` should never truly delete a product with sales history — Configuración's help text (Phase 4 content task) should say so.

---

## 4. `tbl_Gastos` / `_GastosCalc` — Final Formulas

`tbl_Gastos` has no calculated columns (unchanged from Phase 2). `_GastosCalc`:

| Cell/Name | Excel | Google Sheets |
|---|---|---|
| Gastos del período | `=SUMIFS(tbl_Gastos[Importe], tbl_Gastos[Fecha], ">="&Cfg_PeriodoInicio, tbl_Gastos[Fecha], "<="&Cfg_PeriodoFin)` | `=SUMIFS(Gastos!$E$8:$E$5007, Gastos!$B$8:$B$5007, ">="&Cfg_PeriodoInicio, Gastos!$B$8:$B$5007, "<="&Cfg_PeriodoFin)` |
| Gastos del período anterior | Same, with `Cfg_PeriodoInicio`/`Fin` shifted back one period (§5.2) | Same |
| Gastos por categoría (período actual) | One `SUMIFS` row per `Lst_CategoriasGasto` entry: `=SUMIFS(tbl_Gastos[Importe], tbl_Gastos[Categoría], <categoría>, tbl_Gastos[Fecha], ">="&Cfg_PeriodoInicio, tbl_Gastos[Fecha], "<="&Cfg_PeriodoFin)` | Same pattern with bounded ranges |

No edge case here beyond the standard "0 expenses this period" → `SUMIFS` returns `0` cleanly (not an error — `Gastos del negocio = $0` is a valid, correct state, not a data-insufficiency message).

---

## 5. `_DashboardData` — Final Formulas

### 5.1 Block A/B — KPIs, current + previous period

| Name | Excel/Sheets (identical syntax; ranges differ per §1) | Guard |
|---|---|---|
| `KPI_VentasPeriodo` | `=SUMIFS(<Ventas.Venta total>, <Ventas.Fecha>, ">="&Cfg_PeriodoInicio, <Ventas.Fecha>, "<="&Cfg_PeriodoFin)` | none (0 is valid) |
| `KPI_CostosVentaPeriodo` *(resolved, §0.1)* | `=SUMIFS(<Ventas.Costo total de la venta>, <Ventas.Fecha>, ">="&Cfg_PeriodoInicio, <Ventas.Fecha>, "<="&Cfg_PeriodoFin)` | none |
| `KPI_GananciaProductos` | `=KPI_VentasPeriodo - KPI_CostosVentaPeriodo` | none |
| `KPI_GastosPeriodo` | `=_GastosCalc!<Gastos del período>` | none |
| `KPI_GananciaReal` | `=KPI_GananciaProductos - KPI_GastosPeriodo` | none (can be negative — a real, valid business state, shown in red by the design system, not blocked) |
| `KPI_MargenReal` | `=IF(KPI_VentasPeriodo=0, Txt_SinVentas, KPI_GananciaReal/KPI_VentasPeriodo)` | guarded (division) |
| `KPI_CantidadVentas` | `=COUNTIFS(<Ventas.Fecha>, ">="&Cfg_PeriodoInicio, <Ventas.Fecha>, "<="&Cfg_PeriodoFin)` | none |
| `KPI_TicketPromedio` | `=IF(KPI_CantidadVentas=0, Txt_SinVentas, KPI_VentasPeriodo/KPI_CantidadVentas)` | guarded |

Block B (`_Ant` suffix) repeats the same 8 formulas with `Cfg_PeriodoInicio`/`Cfg_PeriodoFin` replaced by a shifted pair:

```
Cfg_PeriodoInicio_Ant = EDATE(Cfg_PeriodoInicio, -1)
Cfg_PeriodoFin_Ant    = EDATE(Cfg_PeriodoInicio, 0) - 1     ' i.e. the day before this period's start
```
(`EDATE` is supported identically by Excel and Sheets.) KPI card ▲/▼ badges compute `=IFERROR((KPI_X - KPI_X_Ant)/ABS(KPI_X_Ant), Txt_SinDatos)` — guarded against a zero or non-numeric prior period (e.g. the business's first month, where there is no "previous period" to compare against — QA case "empty months").

### 5.2 Block C — Trend chart source (resolves the weekly/monthly threshold deferred in Phase 2 §8)

```
DiasDeHistorial = IF(COUNTA(<Ventas.Fecha>)=0, 0, TODAY() - MIN(<Ventas.Fecha>))
Granularidad    = IF(DiasDeHistorial < 60, "Semanal", "Mensual")
```

If `Granularidad = "Mensual"`: 6 rows, `Periodo = EDATE(TODAY(), -5) … TODAY()` bucketed by calendar month, each `Ventas`/`Ganancia` cell a `SUMIFS` over that month's date range.
If `Granularidad = "Semanal"`: 6 rows, `Periodo` = the last 6 rolling 7-day windows ending today, same `SUMIFS` pattern with 7-day boundaries.
If `COUNTA(<Ventas.Fecha>)=0` (zero sales ever — QA case #4): the block resolves to a single flag `Txt_SinVentas` and the chart is suppressed in favor of an inline message on Inicio (a chart with no data is worse than no chart, per Phase 1 §4's "do not overload the dashboard" principle applied to the empty state).

**Threshold rationale:** 60 days (~2 months) was chosen over the previously-sketched "8 weeks" because it aligns with the 6-bucket, monthly-default view switching to weekly only when there isn't yet enough history for 2 full monthly buckets — avoiding a chart that shows "Mes 1: $0 | Mes 2: (partial)".

### 5.3 Block D — Ganancia por producto (chart 2 source)

```
Top 8: for k = 1 to 8, ValorK = LARGE(tbl_Productos[Ganancia total], k) [Excel] / LARGE($O$12:$O$511, k) [Sheets]
       ProductoK = INDEX(Producto range, MATCH(ValorK, Ganancia-total range, 0))
Otros = MAX(SUM(tbl_Productos[Ganancia total]) - SUM(Top 8 values), 0)
```
Guarded: if `COUNTA(<Productos.Producto>) <= 8`, the "Otros" row is suppressed (0 or omitted) rather than shown as a meaningless $0 bar — handled by the chart's own "hide zero series" setting (Phase 4 concern) plus the `MAX(...,0)` floor here preventing a negative "Otros" bar if `LARGE` picks up fewer than 8 real products (its behavior when `k` exceeds the count of numeric entries — guarded via `IFERROR(LARGE(...),"")` per k so unused ranks render blank, not `#NUM!`).

### 5.4 Block E — Objetivo progress (chart 3 source)

Pass-through: `= '_Objetivos_Calc'!Obj_ProgresoPct` (single source of truth stays in `_Objetivos_Calc`, unchanged from Phase 2).

---

## 6. `_Objetivos_Calc` — Final Formulas (resolves `Obj_Factible`, Phase 2 §8)

| Name | Formula | Guard |
|---|---|---|
| `Obj_GananciaObjetivo` | `=IFERROR(INDEX(tbl_Objetivos[Objetivo de ganancia], MATCH(EOMONTH(TODAY(),-1)+1, tbl_Objetivos[Periodo (Mes-Año)], 0)), Cfg_ObjetivoMensualDefault)` | falls back to default when no explicit row exists for the current month |
| `Obj_GananciaActual` | `=KPI_GananciaReal` | none |
| `Obj_Restante` | `=MAX(Obj_GananciaObjetivo - Obj_GananciaActual, 0)` | floored at 0 |
| `Obj_ProgresoPct` | `=IF(Obj_GananciaObjetivo<=0, Txt_ObjetivoInvalido, MIN(Obj_GananciaActual/Obj_GananciaObjetivo, 1))` | guarded |
| `Obj_DiasRestantes` | `=DAY(EOMONTH(TODAY(),0)) - DAY(TODAY())` | none (can be 0 on the last day of the month) |
| `Obj_GananciaDiariaNecesaria` | `=IF(Obj_DiasRestantes=0, Txt_UltimoDia, Obj_Restante/Obj_DiasRestantes)` | guarded (new message `Txt_UltimoDia`, §10 — the div/0 here is not a data problem but a calendar edge, so it gets its own message rather than reusing `Txt_SinDatos`) |
| `Obj_FacturacionNecesaria` | `=IF(KPI_MargenReal<=0, Txt_ObjetivoInvalido, Obj_Restante/KPI_MargenReal)` | guarded |
| `Obj_VentasNecesarias` | `=IF(KPI_TicketPromedio<=0, Txt_FaltaHistorial, Obj_FacturacionNecesaria/KPI_TicketPromedio)` | guarded |

### 6.1 Feasibility — final design (renamed `Obj_Factible` → `Obj_Factibilidad`, flagged rename)

**Why the rename:** Phase 2 modeled this as a boolean (`Obj_Factible`). A boolean can't represent "not enough history to say" — which the brief explicitly requires communicating (§10: *"If there is insufficient historical data, clearly communicate that"*) — without a second cell. A single 3-state **text** output does both jobs in one named cell and is exactly the string Inicio/Insights display inline, so it's also simpler downstream. This is a data-type refinement of an item Phase 2 explicitly deferred to Phase 3, not a change to anything Phase 2 committed to.

New helper block in `_Objetivos_Calc`, rows 20–110 (90 rolling calendar days — bounded, non-volatile per row):

```
Fecha (row i)        = TODAY() - 89 + (i-20)                                    ' 90 consecutive calendar days ending today
GananciaDelDía (i)   = SUMIFS(<Ventas.Ganancia>, <Ventas.Fecha>, Fecha(i))       ' product-level profit for that day
```

`SUMIFS` per day, not `SUMIFS` on `Ganancia real` — business operating expenses (`Gastos`) are lumpy/monthly by nature and have no meaningful "daily" figure, so the historical benchmark deliberately compares like with like: daily figures needed vs. daily figures achieved are both **product profit**, not full net profit. This simplification is stated explicitly here because it's a modeling choice, not an oversight.

```
GananciaDiariaMaximaHistorica = MAX(GananciaDelDía range)
DiasDesdePrimeraVenta          = IF(COUNTA(<Ventas.Fecha>)=0, 0, TODAY() - MIN(<Ventas.Fecha>))

Obj_Factibilidad =
  IF(COUNTA(<Ventas.Fecha>)=0, Txt_SinVentas,
    IF(DiasDesdePrimeraVenta<14, Txt_FaltaHistorial,
      IF(Obj_GananciaDiariaNecesaria <= GananciaDiariaMaximaHistorica,
         "Factible según tu historial", "Ambicioso según tu historial")))
```

14-day minimum chosen as the smallest window that includes at least one full weekly sales cycle (most small businesses have day-of-week sales patterns), avoiding a feasibility verdict based on e.g. a single unusually good or bad day.

---

## 7. `_Insights_Rules` / `_Insights_Engine` — Final Formulas

### 7.1 Placeholder resolution block (`_Insights_Engine`, rows 2–6)

| Token | Formula |
|---|---|
| `{delta}` | `=IFERROR(TEXT(ABS(KPI_MargenReal - KPI_MargenReal_Ant), "0,0%"), "")` |
| `{producto}` | `=IF(COUNTA(<Productos.Producto>)=0, "", INDEX(<Productos.Producto>, MATCH(MAX(<Productos.Ganancia total>), <Productos.Ganancia total>, 0)))` |
| `{ventasDelta}` | `=IFERROR(TEXT((KPI_VentasPeriodo-KPI_VentasPeriodo_Ant)/KPI_VentasPeriodo_Ant, "0,0%"), "")` |
| `{gananciaDelta}` | `=IFERROR(TEXT((KPI_GananciaReal-KPI_GananciaReal_Ant)/ABS(KPI_GananciaReal_Ant), "0,0%"), "")` |
| `{restante}` | `=Cfg_Moneda & TEXT(Obj_Restante, "#,##0")` |

### 7.2 Rule conditions (`_Insights_Rules`, final)

| RuleID | Condición (final) | Prioridad |
|---|---|---|
| R0 (fallback) | `=COUNTA(<Ventas.Fecha>) < 5` | 0 — overrides all others when TRUE |
| R1 | `=AND(ISNUMBER(KPI_MargenReal), ISNUMBER(KPI_MargenReal_Ant), (KPI_MargenReal_Ant - KPI_MargenReal) > 0.02)` | 1 |
| R2 | `=AND(SUM(<Productos.Ganancia total>)>0, MAX(<Productos.Ganancia total>)/SUM(<Productos.Ganancia total>) > 0.5)` | 2 |
| R3 | `=AND(ISNUMBER(VALUE_ventasDelta), ISNUMBER(VALUE_gananciaDelta), (VALUE_ventasDelta - VALUE_gananciaDelta) > 0.10)` | 2 |
| R4 | `=AND(Obj_GananciaObjetivo>0, Obj_Restante>0)` | 3 |

`VALUE_ventasDelta`/`VALUE_gananciaDelta` are the *numeric* (pre-`TEXT()`) versions of §7.1's tokens, kept as separate helper cells so the rule conditions can do numeric comparison while the placeholder cells hold display-ready strings — avoids parsing a formatted percentage string back into a number.

### 7.3 Engine output

```
Ins_Texto1..4 = the Plantilla of up to 4 rows where (Activa=TRUE AND Condición=TRUE), sorted by Prioridad ascending,
                with every {token} in Plantilla replaced via nested SUBSTITUTE() using the resolved values above.
                If R0 is TRUE, Ins_Texto1 = R0's Plantilla ("Agregá más datos para obtener insights.") and
                Ins_Texto2..4 are blank — R0 short-circuits everything else per Phase 1 §4.6.
```

Ranking rows-that-are-TRUE by priority without dynamic arrays (for older-Excel compatibility, consistent with §0.2's reasoning): a small fixed helper column in `_Insights_Rules` computes each active-and-true row's effective sort key (`Prioridad` if true, blank if false or inactive), and `Ins_Texto1..4` each use `SMALL`/`MATCH`/`INDEX` against that key column to pick the 1st, 2nd, 3rd, 4th lowest — a non-array, fully compatible pattern.

---

## 8. `_Simulador_Engine` — Final Formulas (resolves Δ validation bounds, Phase 2 §8)

### 8.1 Data-sufficiency gate

```
Sim_DatosSuficientes = KPI_CantidadVentas >= 1
```
Every output formula below is wrapped `=IF(NOT(Sim_DatosSuficientes), Txt_FaltaHistorial, <calc>)` — the Simulator with zero real sales recorded (QA case: simulator with incomplete data) shows one clear message instead of projecting off a $0 baseline, which would make every "diferencia" trivially equal to the delta inputs and mislead the user into thinking it means something.

### 8.2 Final Δ input validation bounds (per scenario column)

| Input | Bound | Rationale |
|---|---|---|
| `Δ Precio %` | −50% to +200% | Covers realistic repricing (half-off to triple) without permitting a −100%+ input that breaks downstream math (§8.3) |
| `Δ Cantidad de ventas %` | −90% to +300% | Same reasoning; −100% would zero out `Ventas proyectadas` and produce a `MargenPorPrecio` divide-by-zero, guarded regardless (§8.3) but bounded to stay in a meaningful business range |
| `Δ Margen (pp)` | −50 to +50 points | A swing larger than 50 points is not a realistic single scenario and risks the clamp in §8.3 dominating the result |
| `Δ Gastos` ($) | No fixed bound (numeric only) | Valid range depends on the user's own baseline, which varies — bounding it with a static number would be arbitrary. Instead the *formula* clamps the result (§8.3), and the input validation only rejects non-numeric entries. |
| `Δ Publicidad` ($) | No fixed bound (numeric only) | Same reasoning as `Δ Gastos` |

### 8.3 Projected scenario — final formulas, with clamps

```
Ventas proyectadas           = Sim_Base_Ventas * (1 + Sim_A_DeltaVentas) * (1 + Sim_A_DeltaPrecio)
CostosVenta proyectados      = Sim_Base_CostosVenta * (1 + Sim_A_DeltaVentas)
MargenPorPrecio               = IF(Ventas proyectadas <= 0, 0, (Ventas proyectadas - CostosVenta proyectados) / Ventas proyectadas)
MargenProyectado (final)      = MIN(MAX(MargenPorPrecio + Sim_A_DeltaMargen, -2), 0.95)      ' clamped to a sane [-200%, 95%] band
GananciaProductos proyectada  = Ventas proyectadas * MargenProyectado (final)
Gastos proyectados            = MAX(Sim_Base_Gastos + Sim_A_DeltaGastos + Sim_A_DeltaPublicidad, 0)   ' clamped ≥ 0
Sim_A_GastosAjustados (flag)  = (Sim_Base_Gastos + Sim_A_DeltaGastos + Sim_A_DeltaPublicidad) < 0      ' drives a small inline note if TRUE
GananciaReal proyectada       = GananciaProductos proyectada - Gastos proyectados
Diferencia                    = GananciaReal proyectada - Sim_Base_GananciaReal
```

`Sim_A_GastosAjustados = TRUE` displays the note *"Ajustamos tus gastos proyectados a $0 porque el valor ingresado los volvía negativos."* next to the scenario's output — surfacing the clamp instead of silently absorbing it, consistent with §18's error-handling philosophy.

All three scenario columns (A/B/C) use the identical formula set against their own delta inputs; none reference each other.

---

## 9. `_Textos` — Final Message List

Consolidated (supersedes the partial list in Phase 2 §2; additions from this phase marked **new**):

| Name | Spanish text |
|---|---|
| `Txt_SinDatos` | No hay suficientes datos para calcular este indicador. |
| `Txt_SinVentas` | Todavía no cargaste ventas en este período. |
| `Txt_SinProductos` | Agregá tu primer producto para empezar. |
| `Txt_ProductoNoEncontrado` | Este producto ya no está en tu catálogo. |
| `Txt_ObjetivoInvalido` | Definí un objetivo mayor a tu ganancia actual. |
| `Txt_ObjetivoAlcanzado` | ¡Ya alcanzaste tu objetivo! 🎉 |
| `Txt_FaltaHistorial` | Necesitás más historial de ventas para calcular esto con precisión. |
| `Txt_NombreDuplicado` | Ya existe un producto con este nombre. Elegí un nombre distinto para diferenciarlo. |
| `Txt_SinAlertas` **(new)** | Ninguno de tus productos tiene el margen bajo. ¡Buen trabajo! |
| `Txt_UltimoDia` **(new)** | Hoy es el último día del período — no queda tiempo para repartir lo que falta. |

Every guard formula in §2–§8 resolves to one of these names — no formula in the workbook contains a literal Spanish string, so wording stays centralized and consistent (Phase 1's original rationale for `_Textos`, unchanged).

---

## 10. Edge-Case / QA Matrix

Maps every scenario from Phase 1 §21 to the exact formula behavior defined above. This is a design check, not the QA execution itself (Phase 8 runs and records real results against a built workbook).

| # | Case | Behavior |
|---|---|---|
| 1 | No products | Rankings → `Txt_SinProductos`; Dashboard KPIs unaffected (all sales-based, return $0/guarded messages independently); Insights → R0 fires (`Txt_SinVentas`-adjacent fallback via `COUNTA(Ventas.Fecha)<5`) |
| 2 | One product | All rankings resolve to that single product trivially (MAX/INDEX/MATCH over a 1-row set) |
| 3 | Multiple products | Standard path, no special handling needed |
| 4 | No sales | `KPI_*` guarded (§5.1); trend chart suppressed (§5.2); rankings → `Txt_SinVentas`; Simulator gated (§8.1) |
| 5 | One sale | `KPI_CantidadVentas=1`; `Obj_Factibilidad` → `Txt_FaltaHistorial` (< 14 days of history by construction, §6.1) |
| 6 | Multiple sales | Standard path |
| 7 | Zero cost (all cost fields 0 for a product) | `Costo real de venta=0`, `Margen=100%` — valid, not an error |
| 8 | Zero expenses | `Gastos del período=$0` — valid, `Ganancia real = Ganancia de productos` exactly |
| 9 | Negative margin | Shown as a negative percentage (red, per design system), never suppressed — §2's `M` formula has no floor |
| 10 | Discounts | `Ventas.Descuento` flows straight into `Venta total` (§3); a discount larger than the sale value can drive `Venta total` negative — accepted as a valid (if unusual) real state, not blocked, since a same-row `Ganancia` will correctly reflect a loss |
| 11 | Commissions | Baked into `Costo real de venta` (Productos K) and thus every downstream figure |
| 12 | Advertising | `Publicidad atribuida` (per-unit, Productos) vs. `Δ Publicidad`/Gastos "Publicidad" category (business-level) — kept distinct per Phase 1 §3.3 |
| 13 | Goals | Full `_Objetivos_Calc` chain, §6 |
| 14 | Impossible goals (objetivo ≤ 0) | `Obj_ProgresoPct`/`Obj_FacturacionNecesaria` → `Txt_ObjetivoInvalido` |
| 15 | Empty months (no activity in a period) | Current-vs-previous delta badges → guarded via `IFERROR` (§5.1); trend chart shows a real $0 bucket for that period rather than erroring |
| 16 | Products with no sales | `Unidades vendidas=0`, excluded from "más rentable"/"margen bajo" candidacy (both require `Unidades vendidas>0`) but still visible in the catalog and in chart 2 (at $0) |
| 17 | Very low margins | Flows into `Alerta margen bajo` (§2.1) once below `Cfg_MargenAlertaUmbral` |
| 18 | Negative margins | Same as #9 — no floor, shown as-is |
| 19 | Price changes | `tbl_Ventas.Precio de venta` is per-sale (§3), so a mid-catalog price change on `tbl_Productos` never rewrites historical sales — each `Ventas` row keeps the price it was sold at |
| 20 | Quantity changes | Pure input; every downstream formula recalculates from `Cantidad` directly, no caching |
| 21 | Simulator | §8, including the data-sufficiency gate and clamps |
| 22 | Incomplete data | Every aggregate/ratio formula in §5–§8 is individually guarded; no formula assumes another has already validated its inputs |
| 23 | Large numbers | No formula imposes an artificial ceiling; currency columns are unbounded decimals — only the Sheets *row* caps in §1 are fixed, not value magnitudes |
| 24 | Decimal values | All currency/percent fields are plain decimals throughout — no rounding is applied mid-calculation (only `TEXT()` formatting for display in Insights, §7.1), so precision isn't lost between chained formulas |
| 25 | Percentage calculations | `Margen`, `Comisión %`, `Obj_ProgresoPct`, all Δ-percent Simulator inputs consistently stored as decimals (0–1 range) with percent number formatting applied for display, never as raw "0–100" numbers mixed with 0–1 ones — a common source of silent 100x errors, avoided by this consistent convention stated once here |

---

## 11. Summary — Deferred Items From Phase 2, Resolved

| Phase 2 open item | Resolution |
|---|---|
| Exact non-volatile formula for `KPI_CostosVentaPeriodo` | §0.1 + §5.1 — new `tbl_Ventas.Costo total de la venta` column, plain `SUMIFS` |
| Google Sheets equivalent of Excel calculated-column auto-fill | §1 — two patterns depending on column category: `ARRAYFORMULA` for never-overridden columns, pre-filled per-row formulas (bounded to the row cap) for auto-filled-but-overridable columns |
| Exact historical-max expression behind `Obj_Factible` | §6.1 — renamed `Obj_Factibilidad` (3-state text), 90-day daily `SUMIFS` helper block, 14-day minimum-history gate |
| Final Δ input validation bounds for the Simulator | §8.2 |
| Weekly-vs-monthly bucket threshold for the trend chart | §5.2 — 60 days |

No items remain open from Phase 2. No changes were made to Phase 1's sheet inventory, dashboard KPI set, chart set, or Simulator variable set — only the two additive columns in §0, both of which were already anticipated as Phase 3 work in Phase 2 §8.

Phase 3 is complete and ready for review. Phase 4 (UX/UI structure) has not been started.
