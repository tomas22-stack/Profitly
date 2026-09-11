# Profitly — Data Layout Specification (Phase 2)

> Scope: **Phase 2 — exact data layout**, per the brief's development process (§26). This is an implementation specification for review — no workbook, formulas-in-final-syntax, or UI styling are built yet (formulas here are shown in draft/representative syntax to make dependencies reviewable; Phase 3 finalizes and verifies exact Excel/Sheets syntax; Phase 4 finalizes visual styling). Phase 1 (`01-architecture.md`) is carried forward unchanged **except for two amendments below**, which are flagged per instruction rather than applied silently.

---

## 0. Amendments to Phase 1 (flagged, with justification)

Two concrete technical problems surfaced while specifying exact layout. Both are narrow, structural fixes — no sheet is added or removed from the app-facing navigation, no calculation formula from Phase 1 changes.

### 0.1 `_ProductosCalc` and `_VentasCalc` are removed; their columns move into `tbl_Productos` / `tbl_Ventas` directly

**Problem found:** Phase 1 described these two engine sheets as mirroring `tbl_Productos`/`tbl_Ventas` **"1:1 by row."** Phase 1's own protection policy (§7) also says sheet protection must leave **"sort"** enabled on App Layer sheets, because users need to sort their product or sales list. But if a user sorts `tbl_Productos` and `_ProductosCalc` is a separate table on a separate sheet, the two tables' rows fall out of alignment the moment one is sorted and the other isn't — every calculated cell would silently show the wrong product's numbers. This is a data-integrity bug, not a cosmetic one, and it's undetectable to the user (no error is thrown; the numbers are just wrong).

**Resolution:** Fold the calculated columns directly into `tbl_Productos` and `tbl_Ventas` as additional **locked columns of the same table** (same row = same record, always, by construction — there is no second table to fall out of sync). The input/output visual distinction Phase 1 required (§17) is preserved at the *cell* level (unlocked input columns vs. locked grey calculated columns within one table), which was always the real requirement — the separate hidden sheet was one possible implementation of it, not a hard constraint. This also removes 2 of the 21 sheets from Phase 1's inventory (19 total now) and is a net simplification.

`_GastosCalc`, `_DashboardData`, `_Objetivos_Calc`, `_Insights_Engine`, `_Simulador_Engine` are unaffected — none of them relies on row-position mirroring; all aggregate via `SUMIFS`/lookups keyed by a value (product name, period), which is exactly the pattern that avoids this bug. See §3 for why that pattern is safe.

### 0.2 `Producto` (product name) must be enforced unique — it is the de facto foreign key

**Problem found:** `tbl_Ventas[Producto]` links a sale to a product by **name** (a dropdown of `tbl_Productos[Producto]`), and every aggregation (`Unidades vendidas`, `Ganancia total`, the rankings, `_DashboardData`) is a `SUMIFS`/`MATCH` keyed on that name. Phase 1 never stated a uniqueness constraint on product names. If two products share a name (e.g., a copy-pasted row renamed incompletely), every `SUMIFS` keyed on `Producto` silently merges their sales — profitability numbers become wrong with no error shown, which directly violates the brief's error-handling requirement (§18) and its "never invent/misreport information" requirement for insights (§5).

**Resolution:** Add a custom data-validation rule on `tbl_Productos[Producto]` (§6 below) rejecting duplicate names at entry time, with the message *"Ya existe un producto con este nombre. Elegí un nombre distinto para diferenciarlo."* No hidden ID column is introduced — the brief's simplicity principle (§13/§24) argues against adding an invisible key the user never sees when a validation rule closes the gap just as reliably, and product name was already meant to be the natural, human-legible identifier used throughout Productos, Ventas, and the rankings.

Everything else below reflects Phase 1's architecture unchanged.

---

## 1. Global Layout Conventions

These conventions apply to every **App Layer** sheet (`🏠 Inicio`, `⚙️ Configuración`, `📦 Productos`, `🛒 Ventas`, `💸 Gastos`, `🎯 Objetivos`, `🔬 Simulador`, `🧭 Menú`, `🎬 Demo`) so the workbook behaves predictably as the user moves between them.

| Rule | Detail |
|---|---|
| Column A | Always a blank margin column (never content). Content starts at column B. |
| Row 1 | Navigation bar: one hyperlinked cell per section (`🏠 Inicio`, `⚙️ Configuración`, `📦 Productos`, `🛒 Ventas`, `💸 Gastos`, `🎯 Objetivos`, `🔬 Simulador`), starting at B1, one per column-pair. The current sheet's cell is styled distinctly (filled) rather than hyperlinked to itself. |
| Row 2 | Spacer (row height reduced, no content). |
| Row 3 | Page title (e.g. `📦 PRODUCTOS`) + one-line instruction/subtitle directly under it — merged cells, large/medium type respectively. |
| Row 4 | Spacer. |
| Row 5+ | Page content begins here. Every sheet's content region is anchored to row 5 so cross-sheet chart/shape placement stays predictable. |
| Growth rule | An auto-expanding Table is always the **last** content in its column band on the sheet — nothing is ever placed directly below a Table in the same columns, since new rows would overwrite it. Anything that must coexist with a growing Table (a panel, a KPI strip) is placed **above** it or **beside it in a separate column band**, never below. |
| Chart/shape anchoring | Charts are anchored to a fixed top-left cell and given a fixed row/column span; their *data source* is always a small, fixed-size range on an Engine sheet (never a live range on a growing Table), so a chart never needs to be manually resized/reanchored as data grows. |

---

## 2. Named Range Master List

Every cross-sheet formula reference uses a name, never a raw address, so inserting rows never silently breaks a downstream formula. Prefix indicates origin/purpose.

| Prefix | Origin sheet | Example names | Used by |
|---|---|---|---|
| `Cfg_` | `_Config` | `Cfg_NombreNegocio`, `Cfg_Moneda`, `Cfg_TipoNegocio`, `Cfg_ObjetivoMensualDefault`, `Cfg_PeriodoInicio`, `Cfg_PeriodoFin`, `Cfg_ModoActivo`, `Cfg_MargenAlertaUmbral` | Every sheet |
| `Lst_` | `_Listas` (or direct table refs) | `Lst_Categorias`, `Lst_Canales`, `Lst_TiposGasto`, `Lst_CategoriasGasto`, `Lst_Monedas`, `Lst_PeriodosRapidos` | Data validation dropdowns across Productos, Ventas, Gastos, Configuración |
| `Txt_` | `_Textos` | `Txt_SinDatos`, `Txt_SinVentas`, `Txt_SinProductos`, `Txt_ProductoNoEncontrado`, `Txt_ObjetivoInvalido`, `Txt_ObjetivoAlcanzado`, `Txt_FaltaHistorial`, `Txt_NombreDuplicado` | Every guarded formula |
| `KPI_` | `_DashboardData` | `KPI_VentasPeriodo`, `KPI_CostosVentaPeriodo`, `KPI_GananciaProductos`, `KPI_GastosPeriodo`, `KPI_GananciaReal`, `KPI_MargenReal`, `KPI_TicketPromedio`, `KPI_CantidadVentas`, plus a `_Ant` suffix set for the prior-period block (e.g. `KPI_GananciaReal_Ant`) | `🏠 Inicio`, `_Simulador_Engine`, `_Insights_Engine` |
| `Obj_` | `_Objetivos_Calc` | `Obj_GananciaObjetivo`, `Obj_GananciaActual`, `Obj_Restante`, `Obj_ProgresoPct`, `Obj_DiasRestantes`, `Obj_GananciaDiariaNecesaria`, `Obj_FacturacionNecesaria`, `Obj_VentasNecesarias`, `Obj_Factible` (boolean flag, see §5) | `🎯 Objetivos`, `🏠 Inicio`, `_Insights_Engine` |
| `Ins_` | `_Insights_Engine` | `Ins_Texto1` … `Ins_Texto4` | `🏠 Inicio` |
| `Sim_` | `🔬 Simulador` / `_Simulador_Engine` | Inputs: `Sim_A_DeltaPrecio`, `Sim_A_DeltaVentas`, `Sim_A_DeltaMargen`, `Sim_A_DeltaGastos`, `Sim_A_DeltaPublicidad` (×3 for scenarios A/B/C). Outputs: `Sim_A_GananciaProyectada`, `Sim_A_Diferencia`, etc. | `🔬 Simulador` |
| `tbl_` | native Table objects | `tbl_Productos`, `tbl_Ventas`, `tbl_Gastos`, `tbl_Objetivos` | Everywhere (structured references) |

Table column references use native structured references (`tbl_Productos[Margen]`) rather than named ranges — Excel and Sheets both support this, and it keeps the name list above from having to enumerate every column.

---

## 3. Core Tables — Exact Columns

All four are native Tables (`ListObject`/Sheets equivalent) with header row + auto-expanding body. Column letters below assume the table starts at column B (per §1); they are relative anchors for review, not final pixel placement (Phase 4).

### 3.1 `tbl_Productos` — sheet `📦 Productos`, header row 11, data from row 12

| Col | Header (ES) | Kind | Type | Formula (draft) |
|---|---|---|---|---|
| B | Producto | **Input** | text, unique | — |
| C | Categoría | **Input** | dropdown (`Lst_Categorias`) | — |
| D | Precio de venta | **Input** | currency ≥ 0 | — |
| E | Costo del producto | **Input** | currency ≥ 0 | — |
| F | Comisión % | **Input** | percent 0–100 | — |
| G | Packaging | **Input** | currency ≥ 0 | — |
| H | Envío / costo asociado | **Input** | currency ≥ 0 | — |
| I | Publicidad atribuida | **Input** | currency ≥ 0 | — |
| J | Otros costos | **Input** | currency ≥ 0 | — |
| K | Costo real de venta | **Calc** (locked) | currency | `=E + (D*F) + G + H + I + J` |
| L | Ganancia por unidad | **Calc** (locked) | currency | `=D - K` |
| M | Margen | **Calc** (locked) | percent | `=IF(D=0, Txt_SinDatos, L/D)` |
| N | Unidades vendidas | **Calc** (locked) | number | `=SUMIFS(tbl_Ventas[Cantidad], tbl_Ventas[Producto], [@Producto])` |
| O | Ganancia total | **Calc** (locked) | currency | `=SUMIFS(tbl_Ventas[Ganancia], tbl_Ventas[Producto], [@Producto])` |

Because K–O are now Table calculated columns (§0.1), Excel auto-propagates the formula to every new row the instant a product is added — this is a native Excel Table behavior, not something we build. Google Sheets has no exact equivalent auto-fill; the Phase 3/5 build notes must specify a `ARRAYFORMULA`-based substitute for Sheets so parity holds (flagged here as a build-time task, not a layout change).

**Rankings block** (rows 6–8, above the table, per the growth rule in §1) — four cards, each an `INDEX/MATCH` against columns M/N/O:

| Card | Logic (draft) |
|---|---|
| 🏆 Producto más rentable | `INDEX(tbl_Productos[Producto], MATCH(MAX(IF(tbl_Productos[Unidades vendidas]>0, tbl_Productos[Margen])), tbl_Productos[Margen], 0))` (array-entered), guarded for "no sales yet" |
| 🔥 Producto más vendido | `INDEX(...MATCH(MAX(tbl_Productos[Unidades vendidas])...)` |
| 💰 Producto que más dinero genera | `INDEX(...MATCH(MAX(tbl_Productos[Ganancia total])...)` |
| ⚠️ Producto con margen bajo | List of `tbl_Productos[Producto]` where `Margen < Cfg_MargenAlertaUmbral` AND `Unidades vendidas > 0`, first match shown |

All four resolve to `Txt_SinProductos` / `Txt_SinDatos` when `tbl_Productos` is empty or has no sales yet (QA case: 0 products, products with 0 sales — see Phase 8).

### 3.2 `tbl_Ventas` — sheet `🛒 Ventas`, header row 7, data from row 8

| Col | Header (ES) | Kind | Type | Formula (draft) |
|---|---|---|---|---|
| B | Fecha | **Input** | date, ≤ today | — |
| C | Producto | **Input** | dropdown (`tbl_Productos[Producto]`) | — |
| D | Cantidad | **Input** | whole number ≥ 1 | — |
| E | Precio de venta | **Input** | currency ≥ 0 | — (plain input, not auto-filled — see amendment below) |
| F | Canal | **Input** | dropdown (`Lst_Canales`) | — |
| G | Comisión % | **Input** | percent 0–100 | — (plain input, not auto-filled — see amendment below) |
| H | Descuento | **Input** | currency ≥ 0, default 0 | — |
| I | Otros costos | **Input** | currency ≥ 0, default 0 | — |
| J | Venta total | **Calc** (locked) | currency | `=(D*E) - H` |
| K | Ganancia | **Calc** (locked) | currency | `=J - (IFERROR(INDEX(tbl_Productos[Costo real de venta], MATCH([@Producto], tbl_Productos[Producto], 0)), 0) * D) - I` |

> **Amendment (post-build, approved):** `E`/`G` were originally specified as "auto-filled, editable" — the default formula above populates the cell the moment `Producto` is chosen, and typing over it replaces just that cell's formula with a static value. QA (build Phase 5/8) found the flaw in that design: a user who accepts the auto-filled value without retyping it leaves a **live formula** in place, so editing the product's price later silently rewrites every past sale's recorded price and revenue too — contradicting this very document's own §5 relationship rule and Phase 3's case-19 requirement. Approved fix: `E`/`G` are now **plain required inputs**, exactly like `Cantidad`/`Fecha` — no default formula, no auto-fill. This guarantees a recorded sale can never change after the fact, at the cost of the user typing (or copying) the price each time. Full detail: `docs/05-qa-results.md` §5.

### 3.3 `tbl_Gastos` — sheet `💸 Gastos`, header row 7, data from row 8

| Col | Header (ES) | Kind | Type |
|---|---|---|---|
| B | Fecha | **Input** | date, ≤ today |
| C | Categoría | **Input** | dropdown (`Lst_CategoriasGasto`: Publicidad, Software/Suscripciones, Internet, Packaging general, Herramientas, Otros) |
| D | Descripción | **Input** | text |
| E | Importe | **Input** | currency ≥ 0 |
| F | Tipo de gasto | **Input** | dropdown: Fijo / Variable / Único |

No calculated columns at the row level — this table is read only by aggregate (`_GastosCalc`), never joined to another table, per Phase 1 §3.3's split between per-unit "costos de venta" (live in Productos/Ventas) and "gastos del negocio" (live only here).

### 3.4 `tbl_Objetivos` — sheet `🎯 Objetivos`, header row 5, data from row 6

| Col | Header (ES) | Kind | Type |
|---|---|---|---|
| B | Periodo (Mes-Año) | **Input** | date, forced to first-of-month (validation, see §6); unique per month |
| C | Objetivo de ganancia | **Input** | currency, defaults to `Cfg_ObjetivoMensualDefault`, warns if ≤ 0 |

Small table (typically ≤ 12–24 rows/year), placed in columns B:C so the read-only progress panel (§4.5) can sit beside it in columns F onward without collision (growth rule, §1).

---

## 4. Engine Layer — Exact Structure

Engine sheets use a uniform two-column convention: column A = label (for developer readability only, not shown to user since the sheet is hidden), column B (or C where a comparison column is needed) = value/formula.

### 4.1 `_Config`

| Cell | Name | Set by |
|---|---|---|
| B2 | `Cfg_NombreNegocio` | Configuración onboarding field |
| B3 | `Cfg_Moneda` | Configuración onboarding field (symbol string, e.g. `"$"`) |
| B4 | `Cfg_TipoNegocio` | Configuración onboarding field |
| B5 | `Cfg_ObjetivoMensualDefault` | Configuración onboarding field |
| B6 | `Cfg_PeriodoInicio` | Formula, default = first day of current month; overridden by the Inicio period-selector dropdown |
| B7 | `Cfg_PeriodoFin` | Formula, default = last day of current month; overridden by the Inicio period-selector dropdown |
| B8 | `Cfg_ModoActivo` | Static per sheet group (`"Mi Negocio"` on Inicio-side links, `"Demo"` on Demo-side links) — display only, never branches a calculation |
| B9 | `Cfg_MargenAlertaUmbral` | Fixed default (15%), not user-editable in v1 |

### 4.2 `_Listas`

One column per list, headers in row 1, values below — each column is the literal source range for its dropdown (`Lst_Categorias` = `_Listas!$B$2:$B$50`, etc.). Configuración's onboarding writes new rows here when the user adds a category/canal; §0.2's uniqueness logic does not apply here since list values are curated, not a join key.

### 4.3 `_GastosCalc`

| Row | Label | Formula (draft) |
|---|---|---|
| 2 | Gastos del período | `=SUMIFS(tbl_Gastos[Importe], tbl_Gastos[Fecha], ">="&Cfg_PeriodoInicio, tbl_Gastos[Fecha], "<="&Cfg_PeriodoFin)` |
| 3 | Gastos del período anterior | same, with a shifted period pair |
| 5–11 | Gastos por categoría (período actual) | one `SUMIFS` row per `Lst_CategoriasGasto` entry, feeds an optional Pro-only breakdown (deferred per Phase 1 §8) |
| 14–19 | Gastos por mes (últimos 6 meses) | `Periodo | Importe` pairs, feeds `_DashboardData` trend block |

### 4.4 `_DashboardData`

| Block | Rows | Content |
|---|---|---|
| A — KPIs, período actual | 2–9 | `KPI_VentasPeriodo`, `KPI_CostosVentaPeriodo`, `KPI_GananciaProductos`, `KPI_GastosPeriodo` (= `_GastosCalc!B2`), `KPI_GananciaReal`, `KPI_MargenReal`, `KPI_TicketPromedio`, `KPI_CantidadVentas` |
| B — KPIs, período anterior | 12–19 | same 8 metrics, `_Ant` suffix, period pair shifted back one |
| C — Tendencia (chart 1 source) | 22–28 | `Periodo | Ventas | Ganancia`, last 6 periods, monthly buckets (auto-switches to weekly if `<8` weeks of history exist — see Phase 1 §8.3) |
| D — Ganancia por producto (chart 2 source) | 31–40 | `Producto | Ganancia total`, top 8 by `LARGE(tbl_Productos[Ganancia total],k)` + one `"Otros"` row summing the remainder |
| E — Progreso objetivo (chart 3 source) | 43 | pass-through reference to `Obj_ProgresoPct` (single source of truth stays in `_Objetivos_Calc`, §4.5) |

Representative formulas:

```
KPI_VentasPeriodo        = SUMIFS(tbl_Ventas[Venta total], tbl_Ventas[Fecha], ">="&Cfg_PeriodoInicio, tbl_Ventas[Fecha], "<="&Cfg_PeriodoFin)
KPI_CostosVentaPeriodo   = SUMPRODUCT((tbl_Ventas[Fecha]>=Cfg_PeriodoInicio)*(tbl_Ventas[Fecha]<=Cfg_PeriodoFin)*(tbl_Ventas[Cantidad]*<costo unitario lookup>))
KPI_GananciaProductos    = KPI_VentasPeriodo - KPI_CostosVentaPeriodo
KPI_GananciaReal         = KPI_GananciaProductos - KPI_GastosPeriodo
KPI_MargenReal           = IF(KPI_VentasPeriodo=0, Txt_SinDatos, KPI_GananciaReal/KPI_VentasPeriodo)
KPI_TicketPromedio       = IF(KPI_CantidadVentas=0, Txt_SinVentas, KPI_VentasPeriodo/KPI_CantidadVentas)
KPI_CantidadVentas       = COUNTIFS(tbl_Ventas[Fecha], ">="&Cfg_PeriodoInicio, tbl_Ventas[Fecha], "<="&Cfg_PeriodoFin)
```

`KPI_CostosVentaPeriodo` is the one aggregate that can't be a plain `SUMIFS` (it needs a per-row unit-cost lookup multiplied by quantity); Phase 3 will verify the exact non-volatile way to express this in both Excel and Sheets (candidates: a `SUMPRODUCT`, or exposing "Costo total de la venta" as an extra locked column on `tbl_Ventas` itself — which, per the §0.1 reasoning, is likely the safer choice and is flagged here for Phase 3 to decide rather than assumed).

### 4.5 `_Objetivos_Calc`

| Row | Name | Formula (draft) |
|---|---|---|
| 2 | `Obj_GananciaObjetivo` | `=IFERROR(INDEX(tbl_Objetivos[Objetivo de ganancia], MATCH(EOMONTH(TODAY(),0)-DAY(TODAY())+1, tbl_Objetivos[Periodo (Mes-Año)], 0)), Cfg_ObjetivoMensualDefault)` |
| 3 | `Obj_GananciaActual` | `=KPI_GananciaReal` |
| 4 | `Obj_Restante` | `=MAX(Obj_GananciaObjetivo - Obj_GananciaActual, 0)` |
| 5 | `Obj_ProgresoPct` | `=IF(Obj_GananciaObjetivo<=0, Txt_ObjetivoInvalido, MIN(Obj_GananciaActual/Obj_GananciaObjetivo, 1))` |
| 6 | `Obj_DiasRestantes` | `=DAY(EOMONTH(TODAY(),0)) - DAY(TODAY())` |
| 7 | `Obj_GananciaDiariaNecesaria` | `=IF(Obj_DiasRestantes=0, Txt_SinDatos, Obj_Restante/Obj_DiasRestantes)` |
| 8 | `Obj_FacturacionNecesaria` | `=IF(KPI_MargenReal<=0, Txt_ObjetivoInvalido, Obj_Restante/KPI_MargenReal)` |
| 9 | `Obj_VentasNecesarias` | `=IF(KPI_TicketPromedio<=0, Txt_FaltaHistorial, Obj_FacturacionNecesaria/KPI_TicketPromedio)` |
| 10 | `Obj_Factible` (boolean) | `=Obj_GananciaDiariaNecesaria <= MAXIFS(<ganancia diaria histórica>)` — feasibility flag referenced by `_Insights_Engine`; exact historical-max expression to be finalized Phase 3 once the daily-bucketing approach for `tbl_Ventas` is settled |

### 4.6 `_Insights_Rules` + `_Insights_Engine`

`_Insights_Rules` (authored content, not user-editable) — Table with columns:

| Col | Header | Content |
|---|---|---|
| A | RuleID | R1, R2, R3, R4, R0 (fallback) |
| B | Condición | boolean formula referencing `KPI_*`/`Obj_*`/`tbl_Productos` names |
| C | Plantilla | Spanish text with `{token}` placeholders |
| D | Prioridad | 0 (fallback, overrides all) – 3 |
| E | Activa | TRUE/FALSE — lets a rule be disabled without deleting the row |

`_Insights_Engine` — a placeholder-resolution block (rows 2–6: `{delta}`, `{producto}`, `{ventasDelta}`, `{gananciaDelta}`, `{restante}`, each a formula pulling the live value) feeding a ranking block (rows 10–13) that picks the top 4 `Activa=TRUE` rows by `Prioridad` where `Condición=TRUE`, substitutes tokens into `Plantilla` via nested `SUBSTITUTE`, and exposes the results as `Ins_Texto1`…`Ins_Texto4`. If `R0` (data-insufficiency) is TRUE, it is the only rule allowed to fire (per Phase 1 §4.6) — enforced by evaluating R0 first and short-circuiting the rest.

### 4.7 `_Simulador_Engine` — baseline snapshot (exact structure)

This is the piece Phase 1 left conceptual; here is the literal cell layout.

**Block A — Baseline (frozen, one-way copy of current real KPIs), rows 2–8, column B:**

| Row | Name | Source |
|---|---|---|
| 2 | `Sim_Base_Ventas` | `=KPI_VentasPeriodo` |
| 3 | `Sim_Base_CostosVenta` | `=KPI_CostosVentaPeriodo` |
| 4 | `Sim_Base_GananciaProductos` | `=KPI_GananciaProductos` |
| 5 | `Sim_Base_Gastos` | `=KPI_GastosPeriodo` |
| 6 | `Sim_Base_GananciaReal` | `=KPI_GananciaReal` |
| 7 | `Sim_Base_Margen` | `=KPI_MargenReal` |
| 8 | `Sim_Base_CantidadVentas` | `=KPI_CantidadVentas` |

These are the **only** cells in `_Simulador_Engine` that reference `_DashboardData`/`KPI_*` — every projected-scenario formula below reads only from this frozen block, never from `tbl_Productos`/`tbl_Ventas`/`tbl_Gastos` directly. This is what makes the "simulator can never touch real data" guarantee a property of the sheet's wiring (no formula path exists), not a convention someone could break by editing a formula elsewhere.

**Block B — Scenario inputs**, entered by the user on the visible `🔬 Simulador` sheet, read here by name. Three parallel scenario columns (A/B/C) per Phase 1 §4.7:

`Sim_A_DeltaPrecio` (%), `Sim_A_DeltaVentas` (%), `Sim_A_DeltaMargen` (pp), `Sim_A_DeltaGastos` ($), `Sim_A_DeltaPublicidad` ($) — and the same five for `_B_`/`_C_`.

**Block C — Projected scenario (per scenario column), rows 12–19:**

```
Ventas proyectadas          = Sim_Base_Ventas * (1 + Sim_A_DeltaVentas) * (1 + Sim_A_DeltaPrecio)
CostosVenta proyectados     = Sim_Base_CostosVenta * (1 + Sim_A_DeltaVentas)
MargenPorPrecio             = IF(Ventas proyectadas=0, 0, (Ventas proyectadas - CostosVenta proyectados) / Ventas proyectadas)
MargenProyectado (final)    = MargenPorPrecio + Sim_A_DeltaMargen
GananciaProductos proyectada = Ventas proyectadas * MargenProyectado (final)
Gastos proyectados          = Sim_Base_Gastos + Sim_A_DeltaGastos + Sim_A_DeltaPublicidad
GananciaReal proyectada     = GananciaProductos proyectada - Gastos proyectados
Diferencia                  = GananciaReal proyectada - Sim_Base_GananciaReal
```

**Modeling decision flagged for review:** `Δ Margen` is modeled as an *additive* adjustment on top of whatever margin results mechanically from `Δ Precio` (raising price alone already raises margin, since unit cost is unchanged) — not a replacement for it. This avoids double-counting the two levers but means the two are not fully independent in the user's mental model. The ⓘ tooltip on the Simulador's "Margen" input must say this explicitly (draft: *"Este ajuste se suma al efecto que ya tiene un cambio de precio sobre tu margen."*) — content task for Phase 4, flagged here because it originates from this formula decision.

`Δ Publicidad` is modeled as incremental **business-level** ad spend added to `Gastos proyectados`, kept separate from `Δ Gastos` (general) only for the sake of matching the brief's explicit example ("¿Qué pasa si aumento publicidad $100.000?", §11) with its own dedicated input — mathematically the two deltas are equivalent (both just add to Gastos proyectados) but keeping them as separate named inputs preserves that direct answer without extra derivation.

---

## 5. IDs & Relationships Between Tables

| Relationship | Key | Enforcement |
|---|---|---|
| `tbl_Ventas.Producto` → `tbl_Productos.Producto` | Product name (natural key) | Dropdown restricts entry to existing names (referential integrity by construction); uniqueness validation on `tbl_Productos.Producto` prevents key collisions (§0.2) |
| `tbl_Objetivos.Periodo` → implicit calendar month | First-of-month date (natural key) | Validation forces first-of-month value + uniqueness per month (§6) so `_Objetivos_Calc`'s `MATCH` never returns an ambiguous row |
| `tbl_Gastos` | standalone, no FK | Aggregated only by date range/category — never joined to Productos or Ventas, by design (§3.3) |
| `_Demo_Datos` | mirrors `tbl_Productos`/`tbl_Ventas`/`tbl_Gastos`/`tbl_Objetivos` schema exactly, same keys/uniqueness rules | Fully separate table instances — no shared row, no shared name range, no formula crosses from real tables into Demo or vice versa (Phase 1 §2.2) |

No numeric surrogate IDs are introduced anywhere in v1 (see §0.2's rationale) — every relationship in Profitly is small enough (products: dozens, not millions; one goal row per month) that a validated natural key is simpler for a non-technical user to reason about than a hidden ID column, and performs identically at this data scale.

---

## 6. Data Validation — Master Table

| Sheet.Column | Rule | Error message shown |
|---|---|---|
| Productos.Producto | Text required; **custom formula** `=COUNTIF(tbl_Productos[Producto],[@Producto])=1` | `Ya existe un producto con este nombre. Elegí un nombre distinto para diferenciarlo.` |
| Productos.Categoría | List from `Lst_Categorias` | `Elegí una categoría de la lista.` |
| Productos.Precio de venta / Costo del producto / Packaging / Envío / Publicidad atribuida / Otros costos | Decimal ≥ 0 | `Ingresá un número mayor o igual a 0.` |
| Productos.Comisión % | Decimal, 0–100% | `Ingresá un porcentaje entre 0% y 100%.` |
| Ventas.Fecha / Gastos.Fecha | Date, ≤ TODAY() (+ small future tolerance, exact days TBD Phase 3) | `La fecha no puede ser futura.` |
| Ventas.Producto | List from `tbl_Productos[Producto]` | `Elegí un producto de tu catálogo.` |
| Ventas.Cantidad | Whole number ≥ 1 | `Ingresá una cantidad de al menos 1.` |
| Ventas.Canal | List from `Lst_Canales` | `Elegí un canal de la lista.` |
| Ventas.Descuento / Otros costos | Decimal ≥ 0 | `Ingresá un número mayor o igual a 0.` |
| Gastos.Categoría | List from `Lst_CategoriasGasto` | `Elegí una categoría de la lista.` |
| Gastos.Importe | Decimal ≥ 0 | `Ingresá un número mayor o igual a 0.` |
| Gastos.Tipo de gasto | List: Fijo / Variable / Único | `Elegí un tipo de la lista.` |
| Objetivos.Periodo | Date; **custom formula** forcing first-of-month (`=DAY([@Periodo])=1`) + `COUNTIF(tbl_Objetivos[Periodo (Mes-Año)],[@Periodo])=1` | `Elegí el primer día del mes que querés planificar.` / `Ya definiste un objetivo para este mes — editá la fila existente.` |
| Objetivos.Objetivo de ganancia | Decimal; soft warning (not blocked) if ≤ 0 | `Definí un objetivo mayor a $0 para que Profitly pueda calcular tu progreso.` |
| Simulador.Δ Precio %, Δ Cantidad de ventas % | Decimal, reasonable bound e.g. −90% to +500% (guards against nonsensical projections, exact bound TBD Phase 3) | `Ingresá un porcentaje razonable.` |
| Simulador.Δ Margen (pp) | Decimal, −100 to +100 | `Ingresá un valor entre -100 y 100 puntos.` |

All list-based validations reference `_Listas` columns or `tbl_Productos[Producto]` directly (never a hardcoded literal list), so adding a category in Configuración immediately updates every dropdown that uses it.

---

## 7. Protection Boundaries — Master Table

| Sheet | Unlocked (editable) | Locked (protected) |
|---|---|---|
| `🏠 Inicio` | Period-selector dropdown only | Everything else (KPI cards, insights, charts) |
| `⚙️ Configuración` | All onboarding fields (writes to `_Config`/`_Listas`) | Layout labels, instructions |
| `📦 Productos` | `tbl_Productos` columns B–J | `tbl_Productos` columns K–O (calc), rankings block |
| `🛒 Ventas` | `tbl_Ventas` columns B–I (all plain inputs, including E/G — see §3.2 amendment) | `tbl_Ventas` columns J, K |
| `💸 Gastos` | `tbl_Gastos` columns B–F | — (no calc columns on this sheet) |
| `🎯 Objetivos` | `tbl_Objetivos` columns B–C | Progress panel (all of `_Objetivos_Calc`'s exposed values) |
| `🔬 Simulador` | The 5×3 scenario input cells | Baseline display + all projected/difference output cells |
| `🧭 Menú`, `🎬 Demo` | Nothing | Fully locked (Demo is read-only by design, §2.2 of Phase 1) |
| All `_`-prefixed Engine sheets | Nothing | Fully locked, sheet hidden. Onboarding fields on Configuración write into specific `_Config`/`_Listas` cells via **direct cell reference from an unlocked front-end cell**, not by unlocking the engine sheet itself — i.e. the engine sheet stays 100% locked; only the App Layer input cells that happen to feed it are unlocked. |

Sheet-level protection settings (all App Layer sheets): allow *sort*, *use autofilter*, *insert table rows*; disallow *insert/delete columns*, *edit objects* (charts/shapes), *edit scenarios*. This matches Phase 1 §7 and is now safe under the §0.1 amendment (sorting `tbl_Productos`/`tbl_Ventas` can no longer desynchronize a calculated column, since the calculated columns live in the same table).

---

## 8. Summary of Open Items Carried to Phase 3

- Exact non-volatile Excel **and** Sheets syntax for `KPI_CostosVentaPeriodo` (candidate: add a locked "Costo total de la venta" column to `tbl_Ventas` — leaning yes, per §4.4).
- Final bound values for the Simulador's Δ input validation ranges.
- Exact historical-max expression behind `Obj_Factible` (§4.5), once daily-bucketing of `tbl_Ventas` is finalized.
- Google Sheets substitute for Excel Table calculated-column auto-fill (§3.1) on `tbl_Productos`/`tbl_Ventas`.
- Weekly-vs-monthly bucket threshold for the trend chart (Phase 1 §8.3), to be tuned against QA datasets in Phase 8.

No changes to Phase 1's sheet inventory, dashboard KPI set, chart set, insight rules, or simulator variable set are proposed — only the two amendments in §0.
