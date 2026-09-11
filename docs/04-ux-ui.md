# Profitly — UX/UI Structure Specification (Phase 4)

> Scope: **Phase 4 — UX/UI structure**, per the brief's development process (§26). This defines the complete visible design of the 7 App Layer sheets — layout, hierarchy, navigation, components, empty states, tooltips, alerts, and the input/calculated visual system — on top of Phase 1's architecture, Phase 2's exact cell layout, and Phase 3's final formulas, all taken as given and unchanged. No workbook, styling implementation, or pixel-level asset is built here (that is Phase 5 build / Phase 10 polish). **No metric, KPI, chart, card, or calculation appears in this document that Phase 1–3 didn't already define** — Phase 4 only decides how the approved things look and are organized, never what exists.

---

## 0. Design Mandate (from the brief, restated as design rules)

| Competitor-review pitfall the brief flags | Design rule applied throughout this document |
|---|---|
| Too many tabs / confusing navigation | Exactly one navigation mechanism (the row-1 tab strip, §2) repeated identically on all 7 sheets — no sheet introduces its own separate nav pattern. `🧭 Menú` is scoped to a single job (Demo vs. Mi Negocio fork), not a second navigation hub. |
| Overwhelming metrics / dense tables | Inicio shows exactly the 8 KPIs + 3 charts + ≤4 insights Phase 1 defined — nothing more. Every table leads with input columns; calculated columns are visually pushed to a distinct "Profitly calculó esto" zone so the eye isn't asked to parse 15 columns of undifferentiated numbers at once. |
| Confusing/opaque calculations | Every non-obvious metric gets a one-line ⓘ tooltip in plain Spanish (§5) explaining *how* it's derived, not just what it's called. `Obj_Factibilidad`'s 3-state result is always shown as a labeled badge, never a bare number. |
| No guidance for first-time users | A single, consistent onboarding mechanism: the 5-step checklist (§3.2) appears both in Configuración (where it's completed) and, conditionally, as a banner on Inicio (where it's a reminder) — one mechanism, two surfaces, not two different onboarding systems. |
| Inputs not obviously inputs | The 3-state cell system (§1.3) is taught once via a small legend component, then applied identically everywhere — never re-invented per sheet. |
| Dashboard overloaded with charts/data | Fixed at 3 charts, arranged in a strict hierarchy (trend → detail → detail), per Phase 1 §4's own cap. |

---

## 1. Design System Foundations

### 1.1 Color roles

Roles, not final hex values (hex given as a reference starting point; Phase 10 owns final brand polish) — what matters for Phase 4 is that each role is visually distinct and used consistently:

| Role | Reference color | Used for |
|---|---|---|
| Marca / Primario | Deep indigo-teal, `#1E3A5F` | Nav strip active tab, section titles, primary buttons, chart 1 primary series |
| Éxito / Positivo | `#1E8E5A` | Positive deltas, "Factible" badge, positive Simulador diferencia |
| Alerta / Atención | `#B7791F` (amber) | Warning insights (⚠️), "Ambicioso" badge, low-margin flag |
| Error suave / Negativo | `#C0392B` | Negative deltas, negative margin/ganancia values, negative Simulador diferencia |
| Info | `#3B6E9E` | ⓘ tooltip icon color, informational banners |
| Fondo neutro | `#F6F7F9` | Sheet background outside cards/tables |
| Superficie de tarjeta | `#FFFFFF` with soft shadow-equivalent (thin `#E3E6EA` border) | Every card component (§1.4) |
| Entrada requerida | `#FFFFFF` fill, `#B8C4D9` border | Input cells, state 1 (§1.3) |
| Entrada automática editable | `#F3F0FC` fill, `#C9BEEA` dashed border | Auto-filled overridable cells, state 2 |
| Calculado (bloqueado) | `#EEF2F0` fill, no border | Calculated/locked cells, state 3 |

Delta-color rule (important nuance): color follows *business favorability*, not the arithmetic sign. For Ventas, Ganancia real, Margen, Ticket promedio: ▲ green / ▼ red. For **Gastos** specifically, the polarity inverts: a ▲ (spending more) is shown in amber/red and a ▼ (spending less) in green, since a rising expense is not a "win" the way rising sales is. This is called out explicitly in the KPI Card component (§1.4) so it is never implemented backwards.

### 1.2 Typography & iconography

- Headings: bold, brand-color, larger scale (page title > section title > card label).
- Body/values: regular weight; KPI values set larger and bold to read at a glance; supporting labels smaller and muted gray.
- Numbers are right-aligned and use consistent decimal/thousands formatting per `Cfg_Moneda` (established Phase 1 §8.2) — never mixed formats in the same column.
- Iconography is the existing emoji set already fixed by Phase 1's navigation (🏠⚙️📦🛒💸🎯🔬) plus a small closed set used consistently: ⓘ (tooltip), ▲▼ (deltas), 🏆🔥💰⚠️ (rankings, unchanged from Phase 1 §7), 💡 (insights), ✏️/✎ (editable-auto-filled hint), 🔒 (locked, used sparingly — the fill-color system in §1.3 is the primary signal, the lock icon is a reinforcement only on the legend, not on every cell, to avoid visual noise).

### 1.3 The Input/Calculated Visual System (applies to every sheet with a table)

Three states, taught once via a legend, applied everywhere without exception:

| State | Meaning | Visual treatment | Where it appears |
|---|---|---|---|
| 1 — **Entrada requerida** | The user must type this | White fill, solid light-blue border, normal text | Producto, Fecha, Cantidad, Categoría, Importe, etc. — the bulk of every input column |
| 2 — **Entrada automática (editable)** | Pre-filled for convenience, safe to overwrite | Light lavender fill, dashed border, small ✎ note | `Ventas.Precio de venta`, `Ventas.Comisión %` only (Phase 3 §1/§3) |
| 3 — **Calculado por Profitly** | Never edit; Profitly computed it | Soft gray-green fill, no border, locked | `Costo real de venta`, `Margen`, `Ganancia`, all ranking cards, all dashboard KPIs/charts, Objetivos progress panel, Simulador outputs |

**Legend component**: a small, once-per-sheet horizontal strip placed directly above any table that mixes states (Productos, Ventas), reading:

> `⬜ Vos cargás esto   ·   🔸 Profitly lo completa (lo podés cambiar)   ·   🔒 Profitly lo calculó`

with each label swatched in its actual cell color, so the visual vocabulary is learned once and recognized everywhere after.

### 1.4 Reusable Components

| Component | Structure | Used on |
|---|---|---|
| **KPI Card** | 3 stacked lines: (1) label + optional ⓘ, small caps, muted; (2) value, large bold; (3) delta badge `▲/▼ X% vs. período anterior`, colored per §1.1's favorability rule, or blank if no prior-period data (guarded). Fixed card footprint so all cards in a row align. | Inicio (7 of the 8 KPIs — Progreso gets a variant, below) |
| **Progress KPI Card** | Same header/value structure as KPI Card, but the value row is replaced by a horizontal progress bar (conditional-formatting data bar) showing `Obj_ProgresoPct`, with the percentage overlaid as text | Inicio ("Progreso"), Objetivos (bigger version, §3.6) |
| **Ranking Card** | Icon + short label ("🏆 Más rentable") + product name (bold) + supporting value (margin/units/profit) + guarded empty text if none qualifies | Productos (4 cards) |
| **Insight Row** | Icon (per rule, §7.2 of Phase 3) + one sentence, left-aligned, light background tint matching the icon's semantic color (amber for ⚠️, blue for 📈/🎯, teal for 🔥) | Inicio (up to 4 rows) |
| **Alert / Banner** | Full-width or card-width strip, icon + short message + optional link, 4 tone variants: info (blue), success (green), warning (amber), guidance/empty-state (neutral gray with a CTA) | Empty states, onboarding reminders, Gastos vs. Productos cost clarification |
| **Primary Button (link-styled)** | Solid brand-color fill, white bold text, short verb-first label + `→`, implemented as a styled hyperlinked cell/shape (no macro) | "Ir a Productos →", "Ir a Ventas →" in empty states |
| **Secondary Link** | Plain text, brand-color, underlined-on-nothing (no hover state exists in Excel/Sheets, so it's simply styled as a link at rest) | Inline references, "Ver ⓘ" |
| **Tooltip (ⓘ)** | Native Data Validation "input message" attached to the ⓘ-marked cell — appears on selection without a click, no macro | Every metric listed in §5 |

---

## 2. Global Navigation Shell

Applies identically to all 7 App sheets, per Phase 2 §1 (row 1 = nav, row 2 = spacer, row 3 = title, row 4 = spacer, row 5+ = content). Nothing here changes that row plan — Phase 4 only specifies what fills it.

**Row 1 — Tab strip.** Seven merged-cell "tabs," one per section, left to right in the brief's own order (🏠 Inicio, ⚙️ Configuración, 📦 Productos, 🛒 Ventas, 💸 Gastos, 🎯 Objetivos, 🔬 Simulador). The current sheet's tab: solid brand-color fill, white bold text, no hyperlink (you're already here). The other six: light-gray fill, brand-color text, hyperlinked to that sheet. At the far right of the same row, a small persistent **mode badge**: `Mi Negocio · [Cfg_NombreNegocio]` in a neutral pill, or `🎬 DEMO` in a distinct amber-tinted pill on the Demo sheet group — so the user always knows, on every single screen, whether they're looking at their real numbers or the walkthrough. This directly satisfies the brief's demo/real separation requirement (§15) at the point of use, not only at the `🧭 Menú` fork.

**`🧭 Menú`'s scope (for context, not part of the 7 sheets in this phase's deep-dive):** a single-purpose landing screen shown before the user has chosen Demo or Mi Negocio — two large cards ("🎬 Ver una demo" / "🚀 Empezar con mi negocio"), nothing else. It is not a second navigation surface; once inside either sheet group, the row-1 tab strip is the only way to move around, which is what keeps navigation to "one mechanism" per §0.

**Row 3 — Page header.** Page title (icon + Spanish name, e.g. `📦 PRODUCTOS`) in the largest type on the sheet, with a one-line subtitle underneath answering "what is this screen for" in plain language (e.g. *"Tu catálogo de productos y cuánto te deja ganar cada uno."*). This single line is the sheet's entire "manual" — consistent with the brief's ban on long instruction manuals (§13).

---

## 3. Sheet-by-Sheet UX/UI Specification

### 3.1 🏠 Inicio

**Job it does for the user:** answer *"¿cómo va mi negocio ahora mismo?"* in under 5 seconds, then let them drill into why.

**Content, top to bottom (rows per Phase 2 §1's row-5-start convention):**

1. **Onboarding banner** (conditional, guidance-tone Alert Banner, §1.4) — shown only while any of the 5 checklist steps (§3.2) is incomplete: *"🚀 Te falta [N] paso(s) para ver tu panorama completo: [next incomplete step] → [Primary Button to that sheet]"*. Disappears entirely once all 5 are done — never shown again after, to avoid a permanently-visible nag once the user is past onboarding.
2. **Period selector** — a single dropdown, right-aligned, `Lst_PeriodosRapidos` (Este mes / Mes pasado / Últimos 3 meses / Personalizado), driving `Cfg_PeriodoInicio`/`Fin`. This is the only control on Inicio that isn't read-only.
3. **KPI row 1** — 4 KPI Cards, equal width, left to right: **VENTAS**, **GANANCIA REAL** (ⓘ: *"Lo que realmente te queda después de descontar los costos de cada venta y los gastos generales de tu negocio."*), **MARGEN** (ⓘ: *"Porcentaje de tus ventas que queda después de descontar los costos asociados a vender."* — this exact line is the brief's own §13 example, reused verbatim since it's already the right level of plain-language explanation), **GASTOS**.
4. **KPI row 2** — 3 cards: **TICKET PROMEDIO**, **OBJETIVO MENSUAL**, **PROGRESO** (Progress KPI Card variant).
5. **Insights panel** — heading *"💡 Lo que Profitly encontró en tus números"*, up to 4 Insight Rows. Empty state (data-insufficient, `R0`): a single centered neutral line, *"Agregá más datos para obtener insights."* — no icon row shown at all in that state, since 4 empty rows would look broken.
6. **Chart 1 — full width**: *"Evolución de ventas y ganancia"* (the trend story leads, since it's the headline narrative per §0's hierarchy rule).
7. **Charts 2 and 3 — side by side, half width each**: *"Ganancia por producto"* (horizontal bar, so long product names stay legible) and *"Progreso hacia el objetivo"* (the same progress visual as the KPI card, but larger, with `Obj_GananciaObjetivo`/`Obj_GananciaActual` amounts labeled directly on it).

**Dashboard empty states (3 possible states, not just one blanket "no data" message):**

| State | Condition | What's shown instead |
|---|---|---|
| A — Completamente vacío | 0 productos | KPI rows, insights, and charts are all replaced by a single centered onboarding card: *"Todavía no cargaste información. Empezá agregando tus productos."* + Primary Button "Ir a Productos →". Showing 7 `$0` cards here would read as broken, not empty — so the structure itself is swapped, not just the numbers inside it. |
| B — Productos sin ventas | ≥1 producto, 0 ventas | KPI cards render normally but show their guarded text (`Txt_SinVentas`) instead of `$0`; charts are replaced with a lighter placeholder card: *"Todavía no hay ventas para graficar. Registrá tu primera venta para ver esta evolución."* + link to Ventas; insights show the `R0` fallback line. |
| C — Con datos | ≥1 venta | Full dashboard as designed above. |

---

### 3.2 ⚙️ Configuración

**Job it does for the user:** get from "empty workbook" to "ready to use" in one pass, without needing a manual.

**Content:**

1. **"EMPEZÁ ACÁ" roadmap** (brief §6, verbatim structure) — 5 small numbered step-cards in a row, each icon + 1-line label + link, always visible at the very top: `1. Configurá tu negocio` · `2. Agregá tus productos →📦` · `3. Registrá tus ventas →🛒` · `4. Registrá tus gastos →💸` · `5. Mirá tu rentabilidad →🏠`. Each step's card shows ✅ once done (see progress logic below), ⬜ while pending — this is the same checklist referenced from Inicio's banner (§3.1), one underlying state, two display surfaces.
2. **Tu negocio** — input section: `Nombre del negocio` (text), `Tipo de negocio` (dropdown, cosmetic per Phase 1 §3.1), `Moneda` (dropdown).
3. **Categorías de productos** — a short editable list (feeds `Lst_Categorias`); presented as a simple one-column input list with a "+ Agregar categoría" affordance row at the bottom (same extendable-row pattern as tables, §3.3).
4. **Canales de venta** — same pattern, feeds `Lst_Canales`.
5. **💡 Gastos recurrentes** (info Alert Banner, not a new input area — Phase 2/3 defined no separate recurring-expense mechanism, so this step is guidance, not a new feature): *"Cargá tus gastos fijos (alquiler, herramientas, suscripciones) en 💸 Gastos para que tu Ganancia real sea precisa."* + link to Gastos.
6. **Objetivo mensual** — single currency input, writes `Cfg_ObjetivoMensualDefault`.
7. A small closing note, addressing a common SaaS-habit mismatch explicitly: *"No hay botón 'Guardar' — todo se guarda automáticamente a medida que escribís, como en cualquier planilla."* This exists because users arriving with app-not-spreadsheet expectations (which is the whole premise of Profitly's positioning, §2) may otherwise look for a save action that doesn't exist here.

**Progress checklist logic** (drives both this sheet's step cards and Inicio's banner): Step 1 done when `Cfg_NombreNegocio`, `Cfg_Moneda` are non-blank; Step 2 when `tbl_Productos` has ≥1 row; Step 3 when `tbl_Ventas` has ≥1 row; Step 4 when `tbl_Gastos` has ≥1 row; Step 5 always shown as "done" once 1–4 are (it's not a data condition, it's "go look at Inicio").

---

### 3.3 📦 Productos

**Job it does for the user:** know, at a glance, which products make money and which don't — then let them manage the full catalog.

**Content:**

1. **Rankings row** — 4 Ranking Cards (🏆🔥💰⚠️, Phase 1 §7 / Phase 3 §2.1), with the "más vendido ≠ más rentable" distinction made explicit via tooltip on both cards: 🔥's ⓘ reads *"El producto que más unidades vendiste. No es necesariamente el que más ganancia te deja — mirá 'Más rentable' para eso."*; 🏆's ⓘ reads the inverse. This is the brief's own explicit §7 requirement, resolved as a tooltip pair rather than extra on-sheet text, keeping the card row visually light.
2. **Legend** (§1.3) directly above the table.
3. **Catalog table** — columns B–P per Phase 3 §2, visually grouped by the 3-state system: B–J (state 1, white), K–O (state 3, gray-green), P (state 3, gray-green, and additionally not given its own visible header emphasis — it's a flag column feeding the ⚠️ card, not something the user is meant to read directly, so it's styled identically to the other calculated columns rather than called out). A dashed-border **"+ Agregar producto"** affordance sits as the visual cue at the table's first empty row, making the extendable tail obvious rather than relying on the user noticing a blank grid.

**Empty state:** 0 productos → the rankings row shows 4 cards all reading the same guarded text (`Txt_SinProductos`) rather than 4 different broken-looking states, and the table area shows one centered line above the empty first row: *"Todavía no tenés productos cargados. Agregá el primero para empezar 👇"*.

---

### 3.4 🛒 Ventas

**Job it does for the user:** log a sale in under 10 seconds, without re-typing anything already known about the product.

**Content:**

1. **Mini stat strip** — two small inline stats, not full KPI Cards (this sheet is for data entry, not analysis — keeping it light per §0's "avoid overwhelming" rule): `Ventas este mes: [N registros]` and `Última venta: [fecha]`.
2. **Legend** (§1.3), now demonstrating all 3 states since this is the one sheet where all 3 coexist.
3. **Entry table** — columns B–L per Phase 3 §3: B, C, D, F, H, I in state 1 (white); **E and G in state 2** (light lavender, dashed, with a small ✎ note reading *"Se completa solo cuando elegís un producto — lo podés cambiar."*); J–L in state 3 (gray-green). Same "+ Agregar venta" dashed-row affordance as Productos.

**Empty states:**
- 0 productos in catalog yet → the `Producto` dropdown has nothing to offer; instead of a silently-empty dropdown, a guidance Alert Banner sits above the table: *"Necesitás cargar al menos un producto antes de registrar ventas."* + Primary Button "Ir a Productos →".
- ≥1 producto, 0 ventas yet → same dashed-row affordance as Productos: *"Registrá tu primera venta para empezar a ver tu rentabilidad 👇"*.

---

### 3.5 💸 Gastos

**Job it does for the user:** track the business's general operating costs, kept unmistakably separate from per-sale costs.

**Content:**

1. **Clarification banner** (info Alert Banner, always visible — this is a permanent orientation aid, not a dismissible onboarding tip, because the productos/gastos cost split is a recurring point of confusion the brief calls out by name, §9): *"ⓘ Acá cargás los gastos generales de tu negocio (alquiler, software, herramientas). Los costos de cada venta — comisión, envío, packaging, publicidad — se cargan en 📦 Productos."*
2. **Mini stat strip**: `Gastos del negocio este mes: [$X]`.
3. **Entry table** — columns B–F, all state 1 (no calculated columns at row level, per Phase 2/3). Same dashed-row "+ Agregar gasto" affordance.

**Empty state:** 0 gastos → *"Todavía no cargaste gastos del negocio. Si por ahora no tenés, tu Ganancia real usa $0 — podés agregar gastos en cualquier momento."* (reassuring, not alarming — $0 gastos is a valid state, not an error, per Phase 3 §4).

---

### 3.6 🎯 Objetivos

**Job it does for the user:** know exactly how close they are to their goal and precisely what's needed to close the gap.

**Content:**

1. **Goal table** (columns B–C, small) on the left.
2. **Progress panel** beside it (columns F onward, per Phase 2's collision-avoidance placement) — a larger Progress KPI Card at the top (big radial/bar showing `Obj_ProgresoPct`, with `Obj_GananciaActual` / `Obj_GananciaObjetivo` labeled directly beneath), followed by a clean stat list, each row with an ⓘ:
   - `Restante` — ⓘ *"Lo que te falta ganar este mes para llegar a tu objetivo."*
   - `Ganancia diaria necesaria` — ⓘ *"Cuánto necesitás ganar, en promedio, cada día que queda del mes."*
   - `Facturación necesaria` — ⓘ *"Cuánto necesitás vender (no ganar) para llegar a tu objetivo, según tu margen actual."* — this pairing (with the next one) is the exact spot the brief's own competitor-review concerns flagged as easy to confuse, so both tooltips exist specifically to be read side by side.
   - `Ventas necesarias` — ⓘ *"Cuántas ventas necesitás, según tu ticket promedio, para llegar a esa facturación."*
3. **Factibilidad badge** — a single colored pill under the stat list, driven by `Obj_Factibilidad` (Phase 3 §6.1): green *"✅ Factible según tu historial"*, amber *"🚀 Ambicioso según tu historial"*, or neutral gray *"📊 Necesitás más historial de ventas"* — always one of these three, never a bare number, so the feasibility read is immediate.

**Empty state:** no goal defined anywhere (`Cfg_ObjetivoMensualDefault` blank and no `tbl_Objetivos` row) → the progress panel is replaced by a single prompt: *"Todavía no definiste un objetivo. ¿Cuánto querés ganar este mes?"* with the goal input cell itself highlighted (state 1 styling, slightly emphasized border) as the obvious next action.

---

### 3.7 🔬 Simulador

**Job it does for the user:** try a pricing/volume/spend decision safely, see the money impact immediately, and be completely sure their real data was never touched.

**Title, exactly per brief §11:** *"¿QUÉ PASARÍA SI...?"*

**Content:**

1. **ESCENARIO ACTUAL** — a read-only stat block (state 3 styling throughout, explicitly labeled so its "locked" look is understood as intentional, not broken): Ventas, Costos de venta, Ganancia de productos, Gastos, Ganancia real, Margen — the frozen baseline from `_Simulador_Engine` (Phase 3 §8).
2. **Scenario A — inputs**, shown expanded by default (the primary, always-visible scenario): 5 inputs (Δ Precio %, Δ Cantidad de ventas %, Δ Margen pp, Δ Gastos $, Δ Publicidad $), each state-2-styled (auto default 0%, editable) with a one-line plain-language helper directly under it, most notably under Δ Margen: *"Este ajuste se suma al efecto que ya tiene un cambio de precio sobre tu margen."* (the exact modeling-clarification line Phase 3 §8.3 flagged as needed here).
3. **Scenarios B and C** — placed to the right of Scenario A as two more compact columns, wrapped in a **native column group** (Excel/Sheets outline grouping, no macro) collapsed by default behind a `+` control. First-time users see only Scenario A; anyone who wants to compare up to three what-ifs side by side expands the group. This is how "multiple hypothetical scenarios" (brief §11) is supported without showing three full input panels to a user who only wants one — progressive disclosure via a native spreadsheet feature, not a second UI mode (keeping intact the Phase 1 §8 decision to skip a separate Simple/Pro mode).
4. **ESCENARIO PROYECTADO** — same 6-stat layout as the baseline, per active scenario, state 3 styled.
5. **DIFERENCIA** — the single most prominent element on this sheet: one large number, `+$X` in green or `-$X` in red per §1.1, directly under/beside the projected Ganancia real, matching the brief's own example format exactly (`Ganancia: $510.000 → $687.000` → `DIFERENCIA +$177.000`).
6. **Clamp notice** (conditional Alert Banner, amber, only shown when `Sim_A_GastosAjustados`/B/C is TRUE): *"Ajustamos tus gastos proyectados a $0 porque el valor ingresado los volvía negativos."*

**Empty state:** `Sim_DatosSuficientes = FALSE` (no real sales recorded yet) → the entire input/output area is replaced with: *"Necesitás cargar ventas reales antes de simular escenarios."* + Primary Button "Ir a Ventas →" — inputs are hidden entirely rather than shown-but-meaningless, since a simulator projecting off a $0 baseline would otherwise produce numbers that look real but mean nothing.

---

## 4. Demo Presentation (`🎬 Demo`)

Per Phase 1 §2.2, Demo is a fully separate, read-only sheet sourced from `_Demo_Datos`, never sharing a cell with real data. UX-wise, it is styled **identically** to Inicio (same KPI cards, same chart layout, same insight panel) so that what the user learns from Demo transfers directly — the only differences are: (1) the mode badge in the nav row reads `🎬 DEMO` in its distinct amber-tinted pill (§2), and (2) a persistent banner directly under the page header: *"Estás viendo datos de ejemplo. Cuando quieras, pasá a 🚀 Mi Negocio para cargar los tuyos."* + link. No input cells exist anywhere on this sheet (fully state-3/locked), reinforcing that it's a walkthrough, not a second workspace.

---

## 5. Tooltip / Help Text Master List

Consolidated (every ⓘ referenced above, in one place for review):

| Sheet | Metric | Tooltip text |
|---|---|---|
| Inicio | Ganancia real | Lo que realmente te queda después de descontar los costos de cada venta y los gastos generales de tu negocio. |
| Inicio | Margen | Porcentaje de tus ventas que queda después de descontar los costos asociados a vender. |
| Productos | 🏆 Más rentable | El producto con mejor margen entre los que ya vendiste. |
| Productos | 🔥 Más vendido | El producto que más unidades vendiste. No es necesariamente el que más ganancia te deja — mirá "Más rentable" para eso. |
| Ventas | Precio de venta / Comisión % | Se completa solo cuando elegís un producto — lo podés cambiar. |
| Objetivos | Restante | Lo que te falta ganar este mes para llegar a tu objetivo. |
| Objetivos | Ganancia diaria necesaria | Cuánto necesitás ganar, en promedio, cada día que queda del mes. |
| Objetivos | Facturación necesaria | Cuánto necesitás vender (no ganar) para llegar a tu objetivo, según tu margen actual. |
| Objetivos | Ventas necesarias | Cuántas ventas necesitás, según tu ticket promedio, para llegar a esa facturación. |
| Simulador | Δ Margen | Este ajuste se suma al efecto que ya tiene un cambio de precio sobre tu margen. |

---

## 6. Empty-State Matrix (consolidated)

| Sheet | Trigger | Treatment |
|---|---|---|
| Inicio | 0 productos | Full-dashboard swap to onboarding card (§3.1, State A) |
| Inicio | Productos sin ventas | Guarded KPI text + chart placeholder + insight fallback (§3.1, State B) |
| Productos | 0 productos | Rankings show guarded text; dashed-row prompt on table |
| Ventas | 0 productos in catalog | Guidance banner blocking entry, link to Productos |
| Ventas | 0 ventas | Dashed-row prompt on table |
| Gastos | 0 gastos | Reassuring (not alarming) inline note — $0 is valid |
| Objetivos | No goal defined | Panel swapped for a single prompt + highlighted input |
| Simulador | No real sales yet | Entire input/output area swapped for guidance + link to Ventas |

---

## 7. Buttons / Links Inventory

| Label | Style | Sheet | Destination |
|---|---|---|---|
| Ir a Productos → | Primary | Inicio (State A), Ventas (0-productos guard), Simulador (empty state) | 📦 Productos |
| Ir a Ventas → | Primary | Inicio (State B chart placeholder), Simulador (empty state) | 🛒 Ventas |
| Ir a Gastos → | Primary | Onboarding banner (when step 4 is next) | 💸 Gastos |
| (per-step links) | Secondary | Configuración roadmap, Inicio onboarding banner | respective sheet |
| + Agregar producto / venta / gasto | Dashed-row affordance (not a true button — a styled empty row) | Productos, Ventas, Gastos | n/a (same sheet, next table row) |
| 🎬 Ver una demo / 🚀 Empezar con mi negocio | Primary, large card | 🧭 Menú | 🎬 Demo / 🏠 Inicio |
| Pasá a 🚀 Mi Negocio | Secondary | 🎬 Demo banner | 🏠 Inicio |

No "Guardar" button exists anywhere (§3.2) — flagged once, applies workbook-wide.

---

## 8. Explicit Scope Confirmation

- Every KPI, chart, card, ranking, insight rule, goal metric, and Simulator variable referenced above is the same one Phase 1–3 already defined — none renamed, none added, none dropped.
- No new sheet was introduced; `🧭 Menú` and `🎬 Demo` are used exactly as Phase 1 scoped them (§2.2), with their UX role clarified here (§2, §4) rather than expanded.
- Column groups for Simulador Scenarios B/C (§3.7) use a native, macro-free Excel/Sheets feature (outline grouping) and change no formula or data structure from Phase 3 — a presentation choice, not an architecture change, so it is not flagged as an amendment.
- Deferred to Phase 10 (visual polish): final hex values, exact fonts, pixel-level card spacing/shadows, and any animation-equivalent micro-interactions Excel/Sheets can support (e.g. exact conditional-formatting icon sets). Deferred to Phase 5 (build): actually applying any of this to a live workbook.

Phase 4 is complete and ready for review. Phase 5 (build) has not been started.
