# Profitly — Excel Compatibility Fix: Repair/Recovery Prompt

> Critical post-ship fix. The shipped `build/Profitly.xlsx` opened in Microsoft Excel with *"We found a problem with some content in 'Profitly.xlsx'. Do you want us to recover as much as we can?"* This document records the root cause, the fix, and how it was verified — found and fixed by inspecting the actual XLSX package, not assumed.

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

**Regenerated file:** `build/Profitly.xlsx` (rebuilt from `build/build_workbook.py` after the fix).
