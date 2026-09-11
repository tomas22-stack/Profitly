# Profitly — Beginner-Perspective Review & Visual Polish (Phases 9–10)

> Scope: **Phase 9 — beginner-perspective review** and **Phase 10 — final visual polish**, per the brief's development process (§26), performed against the actual shipped `build/Profitly.xlsx` (not a spec re-read). All approved functionality, calculations, architecture, and data relationships are unchanged. No new feature or metric was added; every change below is either a genuine usability defect found by reading the shipped file as a first-time user would, or a presentational refinement.

---

## 1. Method

Phase 9 was performed by systematically dumping every visible cell (value, formula, style, hyperlink, validation, protection state) from the actual shipped workbook, in the order a first-time user would encounter it — nav bar, page title, top-to-bottom content — on all 7 App sheets plus Menú, and asking of each element: *would a first-time user with basic Excel knowledge know what to do here, where to type, what this number means, and how to get to the next screen, without outside help?* Findings were verified by checking the underlying cell/formula, not assumed from the spec. Every fix below was rebuilt and reverified with a full LibreOffice recalculation (0 errors) and the QA suite before moving on.

---

## 2. Phase 9 — Findings and fixes

### 2.1 Configuración's "editable" category/channel lists pointed at a sheet the user cannot reach or edit

**Finding:** The Categorías/Canales sections on `⚙️ Configuración` read *"editá la lista base en la hoja `_Listas`"* — but `_Listas` is a hidden, protected engine sheet with no nav entry. A first-time user cannot unhide a sheet they don't know exists, and even if they somehow did, it's locked. The instruction as written was impossible to follow, and directly contradicted Phase 4's own approved spec (§3.2), which called for a genuinely editable list *on Configuración itself*.

**Fix:** Configuración now has real editable list cells (8 rows for Categorías, 6 for Canales), pre-seeded with the same defaults that used to live on `_Listas`. `Lst_Categorias`/`Lst_Canales` — the named ranges every dropdown in the workbook reads from — now point directly at these Configuración cells (the same pattern already used for `Cfg_NombreNegocio` etc.), so editing a category there updates every dropdown immediately, no hidden sheet involved.

### 2.2 Productos' Categoría column had no dropdown

**Finding:** `docs/02-data-layout.md` and `docs/03-calculation-formulas.md` both document `Productos.Categoría` as a dropdown sourced from `Lst_Categorias`. Re-reading the shipped file's validation list found it was never actually wired during the Phase 5 build — the cell accepted any free text. **Fixed**: added the missing `DataValidation`.

### 2.3 Inicio's onboarding banner said "→ Ir a X" but wasn't clickable

**Finding:** Every other "→ Ir a X" guidance banner in the workbook (Ventas, Simulador, Configuración's own roadmap cards) is a real hyperlink. Inicio's onboarding banner — the single most prominent thing a brand-new user sees — used the identical visual phrasing but had no `.hyperlink` at all, because its message text is dynamic (one of 4 possible next steps) and a plain cell can only carry one static link target. A user who naturally tries to click it would find nothing happens.

**Fix:** Rewritten using Excel's `HYPERLINK()` function with a *computed* target — `HYPERLINK("#'"&IF(...)&"'!A1", IF(...))` — so one cell correctly links to whichever step is actually next. Verified error-free under recalculation and confirmed the sheet-name literals used match the workbook's actual tab names exactly.

### 2.4 The "objetivo missing" condition used OR where it needed AND

**Finding, found while fixing 2.3:** Inicio's banner considered the goal-setting step "missing" via `OR(Cfg_ObjetivoMensualDefault=0, COUNTA(tbl_Objetivos[...])<=1)`. The second clause is true by construction for any workbook that has never had an extra goal row added (the starter row's formula always counts as non-blank), which makes the `OR` **always true** — so a user who sets a real default goal via Configuración, without also adding an explicit dated row on `🎯 Objetivos`, would see the "define your goal" prompt forever, never able to complete onboarding. Objetivos' own equivalent prompt (`B9`) already used `AND` correctly for the same check. **Fixed** by changing Inicio's clause to `AND`, matching the logic that was already correct elsewhere in the same workbook.

### 2.5 Demo's chart-3 source data sat unlabeled on the visible sheet

**Finding:** `🎬 Demo`'s "Alcanzado/Restante" values (feeding the progress chart) were written directly onto the visible Demo sheet at `R115:S116` — far below the visible dashboard, but still reachable by scrolling, and completely unlabeled to a curious user. The real `🏠 Inicio` does the equivalent thing correctly, on the hidden `_DashboardData` engine sheet. **Fixed** by moving Demo's version to the hidden `_Demo_Datos` sheet, making the two dashboards consistent and keeping the architecture's "engine layer is invisible" rule intact on Demo too.

### 2.6 Simulador's Scenario B/C column grouping only collapsed one column

**Finding:** Phase 4 specified Scenarios B and C as collapsed-by-default via native Excel column grouping. The build code set this up via `ColumnDimensionHolder.group()`, which — checked directly in the shipped file — only actually applied `outlineLevel`/`hidden` to the first column of the range (`F`); columns `G`, `H`, `J`, `K`, `L` were left expanded and visible, so the collapse control existed but didn't do anything useful. **Fixed** by setting `outline_level`/`hidden` explicitly on every column in the block (including the gutter column between Scenarios B and C, so the collapse leaves no visible sliver).

### 2.7 "Δ Margen (pp)" didn't explain what "pp" means

**Finding:** A first-time user unfamiliar with "puntos porcentuales" (percentage points) as a finance term could misread this field. **Fixed**: the tooltip now opens with a one-line definition and example (*"pp = puntos porcentuales (ej.: de 20% a 25% es +5 pp)."*) before the existing modeling note.

---

## 3. Phase 10 — Visual polish applied

All presentational; none change a formula, layout structure, or data relationship.

- **Chart colors**: the 3 dashboard charts used Excel's generic numbered chart styles, which don't draw from the brand palette Phase 4 defined. Now: Ganancia (chart 1) in the brand success green, Ventas as a muted supporting line; Ganancia por producto (chart 2) in brand primary; Alcanzado/Restante (chart 3) in success green / neutral gray.
- **Productos' "Alerta margen bajo" column** showed a raw `0`/`1` — correct, but not polished, and easy to mistake for a stray value. Now formatted (`[=1]"⚠️";;`) to show blank for 0 and a small warning icon for 1.
- **Freeze panes** on Productos, Ventas, and Gastos: the nav bar, page header, and table column headers now stay visible while scrolling a long table, instead of scrolling out of view — a real usability win as much as a polish one.
- **KPI card row heights** on Inicio/Demo: the bold, size-15 value row now gets explicit breathing room (24pt) instead of Excel's cramped ~15pt default, with the label/delta rows sized to match.
- **Sheet tab colors**: every App sheet tab now carries the brand primary color (Demo gets the amber warning tint instead, matching its mode badge), so the tab strip itself reads as one product rather than a stack of default-gray Excel tabs.

**Considered and deliberately not done:** a dashed-border "+ Agregar" visual cue on each table's starter row (Phase 4 §3.3/§3.4 mentioned this). Investigated during Phase 10, but Excel Tables are prone to carrying a row's formatting forward when a new row is added by typing into the row below it, and there was no reliable way to confirm the dashed styling wouldn't propagate onto the user's *first real data row* and read as a mistake rather than a hint. The existing, already-verified-working empty-state banner text (*"Agregá el primero para empezar 👇"*) achieves the same "here's where to start" signal without that risk, so it was kept as the sole affordance rather than risk a visible glitch on real data.

---

## 4. Final QA — rerun against the actual shipped `build/Profitly.xlsx`

Re-run in full after every Phase 9/10 change, not against an intermediate copy. (One process issue surfaced and was fixed along the way: the build script used to save to an intermediate `build/_stage.xlsx`, and two verification passes earlier in this project were mistakenly run against a stale `Profitly.xlsx` that hadn't been refreshed from it. The script now saves directly to `build/Profitly.xlsx`, so "rebuild" and "the shipped file" can no longer drift apart.)

| Check | Result |
|---|---|
| Formula errors | **0** across 1,071 formulas, verified by LibreOffice recalculation of the actual shipped file |
| 25-case edge matrix (`build/qa_tests.py`) | **25/25 pass**, rerun against the current shipped file with identical results to the prior verified run — no regression from any Phase 9/10 change |
| Simulator isolation | Confirmed: `_Simulador_Engine` reads only its own frozen baseline; no formula anywhere leads from Simulador back into `tbl_Productos`/`tbl_Ventas`/`tbl_Gastos` |
| Demo isolation | Confirmed: `🎬 Demo` is 100% locked (scanned every cell — zero unlocked cells found) and shares no cell, formula, or named range with the real tables |
| Protection boundaries | Confirmed directly: every input cell (including the newly-added Configuración list cells) is unlocked, every calculated cell (e.g. `Productos!K13`) is locked, every sheet has protection enabled |
| Dashboard reconciliation | Confirmed via QA Tests B/C/F: KPI cards match hand-computed totals from the underlying table data exactly |
| Navigation | Confirmed: every static hyperlink in the workbook resolves to a real sheet name (checked programmatically against `wb.sheetnames`); the one dynamic `HYPERLINK()` formula's computed sheet-name literals match the actual tab names exactly |
| Spanish user-facing text | Confirmed: scanned all 9 App sheets for common English artifacts — none found |

**All items pass. Phases 9 and 10 are complete.**
