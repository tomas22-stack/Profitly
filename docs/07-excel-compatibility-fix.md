# Profitly — Excel Compatibility Fix: Repair/Recovery Prompt

> Critical post-ship fix. The shipped `build/Profitly.xlsx` opened in Microsoft Excel with *"We found a problem with some content in 'Profitly.xlsx'. Do you want us to recover as much as we can?"* This document records the root cause, the fix, and how it was verified — found and fixed by inspecting the actual XLSX package, not assumed.
>
> **Update:** the round 1 fix (§2–§4 below) was real and necessary, but the prompt persisted after it — meaning there was more than one structural defect. §7 documents a second round of investigation and fixes. Both rounds are kept here rather than rewritten into one, since the round 1 finding is still accurate and still a prerequisite fix.

---

## 1. Why earlier QA didn't catch this

Every QA pass up to this point (Phases 5, 8, 9, 10) verified the workbook by recalculating it with LibreOffice and checking for `#VALUE!`/`#DIV/0!`/etc. — that tests whether *formulas* are correct. It says nothing about whether the *OOXML package itself* is well-formed enough for Microsoft Excel's stricter parser, because LibreOffice's reader is materially more tolerant of malformed package structure than Excel's. A file can recalculate perfectly and still be structurally invalid. This class of bug requires inspecting the actual ZIP/XML package, which no prior phase did.

## 2. Root cause

**Every internal ("go to this other sheet") hyperlink in the workbook was written as an invalid OPC relationship.**

The build script set navigation links like this throughout (`add_nav`'s tab strip, guidance banners, the Menú cards, Configuración's roadmap):

```python
ws["B7"].hyperlink = f"#'{SH_PRODUCTOS}'!A1"
```

Assigning a plain **string** to `cell.hyperlink` is the problem. openpyxl's `Cell.hyperlink` setter does not distinguish "this looks like an internal reference" from "this is an external URL" — any string is wrapped as `Hyperlink(ref=cell.coordinate, target=<string>)`, and a `target=` is *always* serialized as an OPC relationship with `TargetMode="External"`. Unzipping the shipped file confirmed exactly this, e.g. in `xl/worksheets/_rels/sheet1.xml.rels`:

```xml
<Relationship Type=".../hyperlink" Target="#'🎬 Demo'!A1" TargetMode="External" Id="rId1" />
```

This is invalid per the OPC spec: an `External` relationship's `Target` is supposed to be an independently resolvable URI (an `http://` address, a file path, `mailto:`, etc.) — not a bare `#`-fragment naming a location *inside the same package*, and not one containing an embedded literal single-quote and non-ASCII emoji character the way a real URI would need to be percent-encoded. Excel's package parser validates this strictly and rejects it; LibreOffice's reader is far more forgiving and silently accepted it, which is why every prior recalculation-based QA pass came back clean. Every one of the 8 places in the codebase that set `.hyperlink` to an internal reference string had this defect — 46 individual hyperlink cells across the workbook were affected in total.

The correct OOXML construction for a same-workbook hyperlink doesn't use a relationship at all — it's a plain `location` attribute directly on the `<hyperlink>` element:

```xml
<hyperlink ref="D1" location="'⚙️ Configuración'!A1" />
```

## 3. The fix

Added one helper (`internal_link(cell, sheet_name)` in `build/build_workbook.py`) that constructs the `Hyperlink` object correctly — `location=` instead of `target=`, so openpyxl serializes it as the plain `location` attribute above, with no relationship and no `TargetMode` at all:

```python
def internal_link(cell, sheet_name):
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sheet_name}'!A1", target=None)
```

Replaced all 8 call sites (`cell.hyperlink = f"#'{X}'!A1"`) with `internal_link(cell, X)`. The one *formula-based* link — Inicio's dynamic onboarding banner, which uses the worksheet function `=HYPERLINK("#'sheet'!A1", message)` — was already correct and untouched: the `HYPERLINK()` **function's** own argument syntax legitimately uses the `#` prefix for an internal reference (this is Microsoft's own documented convention for that function); the bug was specific to the `cell.hyperlink` **attribute** mechanism, which is a completely different code path with different serialization rules.

## 4. Package-level verification performed

Beyond the specific fix, the entire OOXML package was inspected systematically (unzipped and checked with `xml.etree.ElementTree` and direct XML analysis, not just spot-checked):

| Check | Result |
|---|---|
| Every XML/rels part in the package is well-formed | ✅ all parts parse cleanly |
| `[Content_Types].xml` declares a content type for every physical file in the package | ✅ complete |
| `styles.xml` collection `count` attributes (`numFmts`, `fonts`, `fills`, `borders`, `cellStyleXfs`, `cellXfs`, `dxfs`, `cellStyles`) match actual child-element counts | ✅ all match |
| Every `cellXfs` entry's `fontId`/`fillId`/`borderId`/`numFmtId` reference resolves to a real, in-bounds entry | ✅ none dangling |
| Every worksheet's `s="N"` style index is within `cellXfs` bounds | ✅ none out of bounds |
| Every `dxfId` referenced by conditional formatting resolves to a real `dxfs` entry | ✅ resolves (2 defined, 2 used) |
| All 4 table definitions: unique `id`, `ref` matches declared column count, no duplicate column names, `autoFilter` ref matches `ref` | ✅ all 4 clean |
| Every `r:id`/`r:embed` referenced inside a part resolves to a real entry in that part's `.rels` file | ✅ confirmed after the fix (this is exactly what the broken hyperlinks were failing) |
| No `TargetMode="External"` relationship remains anywhere in the package | ✅ confirmed zero, post-fix |
| Pristine (never LibreOffice-touched) shipped file loads via `openpyxl.load_workbook()` with `warnings-as-errors` | ✅ zero warnings |

The two `UserWarning`s seen during earlier QA sessions ("Unknown extension is not supported", "Conditional Formatting extension... will be removed") only ever appeared when re-reading a copy **after LibreOffice had already resaved it** for recalculation — LibreOffice adds its own extension metadata on save that openpyxl doesn't recognize when reading it back. That is LibreOffice's own round-trip artifact, confirmed absent from the actual pristine generated file, and unrelated to the Excel repair prompt.

## 5. Functional re-verification (unchanged from before the fix)

Re-run in full against the regenerated file:

| Check | Result |
|---|---|
| Formula errors | 0 across 1,071 formulas |
| 25-case edge matrix | 25/25 pass, identical values to the pre-fix run — no regression |
| Sheets | 19/19 present, same names, same order |
| Named ranges | 141 (unchanged) |
| Tables | 4 (unchanged) |
| Data validations | 63 (unchanged) |
| Input cells editable | confirmed directly (e.g. `Productos!B13.protection.locked == False`) |
| Calculated cells protected | confirmed directly (e.g. `Productos!K13.protection.locked == True`) |
| Sheet/workbook protection | confirmed enabled on every sheet |

No formula, named range, validation rule, table, chart, or protection setting was touched by this fix — only the serialization mechanism for internal navigation links.

## 6. What could not be verified

This environment has no Microsoft Excel installation, so "opens cleanly in Excel with no repair prompt" was verified by fixing the specific, identified structural defect (confirmed via direct package inspection to be the exact pattern OOXML's spec disallows and Excel's validator is known to reject) and by exhaustively re-validating every other structural aspect of the package by hand, rather than by observing Excel itself. If any further Excel-specific issue exists beyond what a from-first-principles OPC/XML audit can surface, it was not caught here — that residual risk is disclosed rather than papered over.

---

## 7. Round 2 — the prompt persisted; a second real defect found and fixed

The round-1 fix was confirmed correct (verified by unzipping and re-checking the package) but the user reported the repair prompt still appeared on the regenerated file. That means round 1 removed one real defect but not the only one. Went back into the package with a wider net rather than assuming round 1 was complete.

### 7.1 Data validation `formula1` values carried a spurious leading `=`

Every `list`-type data validation that referenced a named range, and every `custom`-type validation (the product-name/period uniqueness checks), was built with `formula1="=Lst_Categorias"` etc. — a literal `=` character stored as the first character of the formula text. Confirmed directly in the XML:

```xml
<dataValidation type="list" ...><formula1>=Lst_Categorias</formula1></dataValidation>
```

The `=` is a UI-only convention — when a person types `=Lst_Categorias` into Excel's Data Validation dialog, Excel stores the reference in `<formula1>` *without* the leading `=`, since the element's content is already implicitly a formula/reference expression. A doubled `=` here is non-standard output that a genuine Excel-authored file never produces, present on 10 validations across the workbook: every dropdown (Producto, Canal, Categoría ×2, Categoría de gasto, Tipo de gasto, Moneda, Período rápido) and every custom uniqueness/date check.

**Fix:** stripped the leading `=` from all 10 `formula1` values in `build/build_workbook.py`. `formula1="Lst_Categorias"` now, matching genuine Excel output.

### 7.2 Chart series-title references pointed at the wrong (often blank) cell

While investigating, found — and fixed, as a correctness bug independent of the repair prompt — that the "Evolución de ventas y ganancia" and "Ganancia por producto" charts (both on `🏠 Inicio` and `🎬 Demo`, 4 of the 6 chart parts) had their data range built one row too high. `add_dashboard_charts()` assumed the row *above* each data block's header was itself a header row (a pattern that happened to be right for the 3rd chart but wrong for these two), so `titles_from_data=True` picked up the series name from an unrelated or blank cell instead of the real "Ventas"/"Ganancia"/"Ganancia total" header, and the category axis included the header row as if it were a data point. Confirmed directly: before the fix, chart 1's series title referenced `_DashboardData!B21` (a granularity flag cell, not a header); after, it correctly references `_DashboardData!B22` ("Ventas"). Fixed by correcting the row arithmetic for both charts.

### 7.3 Re-audit performed after both fixes (not just the two specific issues)

Every check from §4 was re-run against the newly regenerated file, plus additional checks this round specifically motivated:

| Check | Result |
|---|---|
| All §4 checks (well-formedness, Content-Types, styles counts, r:id consistency, `TargetMode="External"`, defined names) | ✅ all still clean |
| Merged-cell ranges on every sheet, checked pairwise for overlap | ✅ none found |
| `<pane>` (freeze pane) XML on Productos/Ventas/Gastos | ✅ well-formed, consistent `ySplit`/`topLeftCell`/`activePane` |
| Chart XML element ordering (`c:ser` child sequence) on all 6 chart parts, against the CT_LineSer/CT_BarSer schema sequence | ✅ correct order on every series |
| Drawing anchors (`oneCellAnchor`/`from`/`ext`/`graphicFrame`) on both drawing parts | ✅ well-formed |
| No `tabSelected="1"` on any hidden sheet | ✅ none |
| Every `<formula1>` in every worksheet's data validations, rescanned for a leading `=` | ✅ zero remaining |
| `DataBarRule` calls that explicitly passed `minLength=None`/`maxLength=None` | ✅ confirmed these serialize as *omitted* attributes, not literal `"None"` strings |

### 7.4 Functional re-verification, round 2

Identical method to §5, re-run against the round-2 file: 0 formula errors (1,071 formulas), 25/25 edge cases pass with values identical to both prior runs, all 19 sheets/141 named ranges/4 tables/63 validations unchanged, input/calculated cell protection unchanged.

### 7.5 Where this leaves things, honestly

Two independent, concrete, verifiable defects have now been found and fixed by direct package inspection — not by guessing. Both were real and each is individually confirmed gone. Whether a third exists cannot be ruled out with certainty from static analysis alone in an environment with no Microsoft Excel to observe directly: this round's audit was substantially broader than round 1 (adding merges, freeze panes, chart element order, and drawing anchors to what's checked), but "broader" is not "exhaustive against Excel's exact validator," which is not public. If the prompt appears again, the next most useful thing to capture is *which specific content Excel says it recovered* — Excel's repair log (shown after clicking "Yes") usually names the affected part or feature by type, which would point straight at whatever remains rather than requiring another blind sweep.

---

**Regenerated file:** `build/Profitly.xlsx` (rebuilt from `build/build_workbook.py` after both rounds of fixes).
