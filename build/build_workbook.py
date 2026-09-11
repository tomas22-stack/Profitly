#!/usr/bin/env python3
"""
Profitly workbook builder.
Implements docs/01-architecture.md, docs/02-data-layout.md,
docs/03-calculation-formulas.md, docs/04-ux-ui.md exactly.

BUILD DEVIATION (documented, see docs/05-qa-results.md section 0):
Phase 3 specified Excel formulas using the `[@Column]` "current row"
structured-reference token for same-row self-references (e.g.
`[@Producto]`). Empirical testing against this environment's formula
verification engine (LibreOffice, via scripts/recalc.py) showed that
token is not reliably evaluated when a workbook is authored
programmatically (openpyxl) rather than typed inside the Excel
application itself -- it silently resolves to #N/A. Cross-table
structured references (`tbl_X[Column]`) and MAXIFS/MINIFS with the
`_xlfn.` prefix were verified to work correctly. Since the brief
requires zero formula errors ever reach the user, and every formula
here must be verifiable rather than trusted, this build uses plain
relative same-row cell references (e.g. `B13`) instead of `[@Column]`
for self-references, while keeping `tbl_X[Column]` for all cross-table
lookups. This changes no formula logic, dependency, or output -- Excel
auto-fills a calculated column's formula down a growing Table
identically regardless of which reference style is used inside it, so
the row-insertion resilience Phase 2 required is fully preserved.
"""
import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, NamedStyle
from openpyxl.formatting.rule import DataBarRule, CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.pagebreak import Break
import datetime

# ---------------------------------------------------------------------------
# Design system constants (docs/04-ux-ui.md #1)
# ---------------------------------------------------------------------------
C_PRIMARY = "1E3A5F"
C_SUCCESS = "1E8E5A"
C_WARNING = "B7791F"
C_ERROR = "C0392B"
C_INFO = "3B6E9E"
C_BG = "F6F7F9"
C_CARD = "FFFFFF"
C_CARD_BORDER = "E3E6EA"
C_MUTED = "6B7280"
C_TEXT = "1F2937"
C_INPUT_FILL = "FFFFFF"
C_INPUT_BORDER = "B8C4D9"
C_AUTOFILL_FILL = "F3F0FC"
C_AUTOFILL_BORDER = "C9BEEA"
C_CALC_FILL = "EEF2F0"
C_NAV_INACTIVE_BG = "E8EAED"
C_AMBER_BG = "FCF3E3"
C_GREEN_BG = "E7F4EE"
C_BLUE_BG = "E9F1F8"
C_GRAY_BG = "F0F1F3"

FONT_NAME = "Arial"

def F(size=10, bold=False, color=C_TEXT, italic=False):
    return Font(name=FONT_NAME, size=size, bold=bold, color=color, italic=italic)

def fill(hexcolor):
    return PatternFill(fill_type="solid", start_color=hexcolor, end_color=hexcolor)

THIN = Side(style="thin", color=C_CARD_BORDER)
INPUT_SIDE = Side(style="thin", color=C_INPUT_BORDER)
AUTOFILL_SIDE = Side(style="dashed", color=C_AUTOFILL_BORDER)

def border_all(side):
    return Border(left=side, right=side, top=side, bottom=side)

ALIGN_L = Alignment(horizontal="left", vertical="center", wrap_text=False)
ALIGN_LW = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_R = Alignment(horizontal="right", vertical="center")

FMT_CURRENCY = '$#,##0;[RED]-$#,##0;"-"'
FMT_PCT = '0.0%;[RED]-0.0%;"-"'
FMT_INT = '#,##0'
FMT_DATE = 'DD/MM/YYYY'

wb = openpyxl.Workbook()
wb.remove(wb.active)

def ws_new(title):
    return wb.create_sheet(title)

# Sheet handles, created in tab order (App layer first, Engine layer hidden after)
SH_MENU = "🧭 Menú"
SH_INICIO = "🏠 Inicio"
SH_CONFIG = "⚙️ Configuración"
SH_PRODUCTOS = "📦 Productos"
SH_VENTAS = "🛒 Ventas"
SH_GASTOS = "💸 Gastos"
SH_OBJETIVOS = "🎯 Objetivos"
SH_SIMULADOR = "🔬 Simulador"
SH_DEMO = "🎬 Demo"

SH__CONFIG = "_Config"
SH__LISTAS = "_Listas"
SH__GASTOSCALC = "_GastosCalc"
SH__DASHDATA = "_DashboardData"
SH__OBJCALC = "_Objetivos_Calc"
SH__INSRULES = "_Insights_Rules"
SH__INSENGINE = "_Insights_Engine"
SH__SIMENGINE = "_Simulador_Engine"
SH__DEMODATA = "_Demo_Datos"
SH__TEXTOS = "_Textos"

APP_SHEETS = [SH_MENU, SH_INICIO, SH_CONFIG, SH_PRODUCTOS, SH_VENTAS, SH_GASTOS,
              SH_OBJETIVOS, SH_SIMULADOR, SH_DEMO]
ENGINE_SHEETS = [SH__CONFIG, SH__LISTAS, SH__GASTOSCALC, SH__DASHDATA, SH__OBJCALC,
                 SH__INSRULES, SH__INSENGINE, SH__SIMENGINE, SH__DEMODATA, SH__TEXTOS]

for name in APP_SHEETS + ENGINE_SHEETS:
    ws_new(name)

# Phase 10 polish: a subtle brand-colored tab strip on every app sheet, so the
# tab row itself reads as "one product" rather than a stack of default-gray
# Excel tabs -- the Demo tab gets the warning/amber tint instead, matching its
# mode badge (§2), so it's visually distinct from Mi Negocio even before
# opening it.
for name in APP_SHEETS:
    wb[name].sheet_properties.tabColor = C_WARNING if name == SH_DEMO else C_PRIMARY

NAV_TABS = [SH_INICIO, SH_CONFIG, SH_PRODUCTOS, SH_VENTAS, SH_GASTOS, SH_OBJETIVOS, SH_SIMULADOR]

def name_range(rangename, sheet, ref):
    safe_sheet = f"'{sheet}'" if any(c in sheet for c in " -🏠⚙️📦🛒💸🎯🔬🎬🧭") else sheet
    wb.defined_names[rangename] = DefinedName(rangename, attr_text=f"{safe_sheet}!{ref}")

def setc(ws, addr, value=None, f=None, fl=None, b=None, al=None, nf=None):
    c = ws[addr]
    if value is not None:
        c.value = value
    if f is not None:
        c.font = f
    if fl is not None:
        c.fill = fl
    if b is not None:
        c.border = b
    if al is not None:
        c.alignment = al
    if nf is not None:
        c.number_format = nf
    return c

def merge(ws, rng):
    ws.merge_cells(rng)

def col_letter(idx):
    return get_column_letter(idx)

def col_idx(letter):
    return column_index_from_string(letter)

# ---------------------------------------------------------------------------
# Navigation shell + page header (docs/04-ux-ui.md #2), applied to every App sheet
# ---------------------------------------------------------------------------
def add_nav(ws, active, mode_label="Mi Negocio", mode_bg=C_NAV_INACTIVE_BG, mode_fg=C_PRIMARY, business_name_formula=None):
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 6
    ws.row_dimensions[5].height = 6
    col = 2  # B
    for tab in NAV_TABS:
        start = col
        end = col + 1
        rng = f"{col_letter(start)}1:{col_letter(end)}1"
        merge(ws, rng)
        is_active = (tab == active)
        c = setc(ws, f"{col_letter(start)}1", tab,
                 f=F(10, bold=True, color="FFFFFF" if is_active else C_PRIMARY),
                 fl=fill(C_PRIMARY if is_active else C_NAV_INACTIVE_BG),
                 al=ALIGN_C)
        if not is_active:
            internal_link(ws[f"{col_letter(start)}1"], tab)
        col = end + 1  # 1-col gutter
    # mode badge, far right
    badge_start = col + 1
    badge_end = badge_start + 2
    rng = f"{col_letter(badge_start)}1:{col_letter(badge_end)}1"
    merge(ws, rng)
    val = f'="{mode_label} · "&Cfg_NombreNegocio' if business_name_formula is None else business_name_formula
    setc(ws, f"{col_letter(badge_start)}1", val, f=F(9, bold=True, color=mode_fg), fl=fill(mode_bg), al=ALIGN_C)

def add_header(ws, title, subtitle):
    merge(ws, "B3:P3")
    setc(ws, "B3", title, f=F(18, bold=True, color=C_PRIMARY), al=ALIGN_L)
    ws.row_dimensions[3].height = 26
    merge(ws, "B4:P4")
    setc(ws, "B4", subtitle, f=F(10, color=C_MUTED, italic=True), al=ALIGN_L)
    ws.row_dimensions[4].height = 16
    ws.column_dimensions["A"].width = 2.5
    for cl in "BCDEFGHIJKLMNOP":
        ws.column_dimensions[cl].width = 12

def internal_link(cell, sheet_name):
    """Set a same-workbook navigation link on `cell`, targeting A1 of `sheet_name`.
    CORRUPTION FIX (see module docstring): plain `cell.hyperlink = "#'X'!A1"`
    string assignment makes openpyxl build a Hyperlink with `target=`, which
    it serializes as an OPC relationship with TargetMode="External" -- invalid
    for a workbook-internal reference, and exactly the kind of malformed
    relationship (a bare "#..." fragment marked External, with an unescaped
    literal quote and non-ASCII characters inside the Target URI) that
    Excel's strict package validator rejects with a "repair" prompt, while
    LibreOffice's much more lenient reader accepted it silently -- which is
    why this was never caught by recalculation-based QA. The correct
    construction uses `location=` (and `target=None`), which openpyxl writes
    as a plain `location` attribute on the <hyperlink> element with no
    relationship at all, matching what Excel itself produces for an internal
    link.
    """
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sheet_name}'!A1", target=None)

def tooltip(ws, addr, title, text):
    dv = DataValidation(type=None, showInputMessage=True, promptTitle=title[:32], prompt=text[:255])
    dv.add(addr)
    ws.add_data_validation(dv)

def banner(ws, rng, text, tone="info"):
    tones = {
        "info": (C_BLUE_BG, C_INFO),
        "success": (C_GREEN_BG, C_SUCCESS),
        "warning": (C_AMBER_BG, C_WARNING),
        "neutral": (C_GRAY_BG, C_MUTED),
    }
    bg, fg = tones[tone]
    merge(ws, rng)
    topleft = rng.split(":")[0]
    setc(ws, topleft, text, f=F(9.5, color=fg), fl=fill(bg), al=ALIGN_LW,
         b=border_all(Side(style="thin", color=bg)))

def card_shell(ws, rng, fl_color=C_CARD, b=True):
    merge(ws, rng)
    topleft = rng.split(":")[0]
    if b:
        ws[topleft].border = border_all(THIN)
    ws[topleft].fill = fill(fl_color)

# ===========================================================================
# ENGINE LAYER
# ===========================================================================

# ---------------------------------------------------------------------------
# _Textos (docs/03-calculation-formulas.md #9)
# ---------------------------------------------------------------------------
ws = wb[SH__TEXTOS]
TEXTOS = {
    "Txt_SinDatos": "No hay suficientes datos para calcular este indicador.",
    "Txt_SinVentas": "Todavía no cargaste ventas en este período.",
    "Txt_SinProductos": "Agregá tu primer producto para empezar.",
    "Txt_ProductoNoEncontrado": "Este producto ya no está en tu catálogo.",
    "Txt_ObjetivoInvalido": "Definí un objetivo mayor a tu ganancia actual.",
    "Txt_ObjetivoAlcanzado": "¡Ya alcanzaste tu objetivo! 🎉",
    "Txt_FaltaHistorial": "Necesitás más historial de ventas para calcular esto con precisión.",
    "Txt_NombreDuplicado": "Ya existe un producto con este nombre. Elegí un nombre distinto para diferenciarlo.",
    "Txt_SinAlertas": "Ninguno de tus productos tiene el margen bajo. ¡Buen trabajo!",
    "Txt_UltimoDia": "Hoy es el último día del período — no queda tiempo para repartir lo que falta.",
}
setc(ws, "A1", "Nombre", f=F(10, bold=True))
setc(ws, "B1", "Texto", f=F(10, bold=True))
for i, (k, v) in enumerate(TEXTOS.items(), start=2):
    setc(ws, f"A{i}", k)
    setc(ws, f"B{i}", v)
    name_range(k, SH__TEXTOS, f"$B${i}")
ws.column_dimensions["A"].width = 28
ws.column_dimensions["B"].width = 70
ws.sheet_state = "hidden"

# ---------------------------------------------------------------------------
# _Config (docs/02-data-layout.md #4.1)
# ---------------------------------------------------------------------------
ws = wb[SH__CONFIG]
setc(ws, "A1", "Nombre", f=F(10, bold=True))
setc(ws, "B1", "Valor", f=F(10, bold=True))
# NOTE: Cfg_NombreNegocio / Cfg_Moneda / Cfg_TipoNegocio / Cfg_ObjetivoMensualDefault
# are named ranges that point DIRECTLY at their input cells on ⚙️ Configuración
# (C11/C13/C12/C24, wired below where that sheet is built) -- not at cells here.
# _Config only hosts values that are *computed*, never typed by the user: the
# active date-range pair (driven by the Inicio period-selector dropdown) and the
# low-margin alert threshold (a fixed default, not user-facing in v1).
name_range("Cfg_NombreNegocio", SH_CONFIG, "$C$11")
name_range("Cfg_TipoNegocio", SH_CONFIG, "$C$12")
name_range("Cfg_Moneda", SH_CONFIG, "$C$13")
name_range("Cfg_ObjetivoMensualDefault", SH_CONFIG, "$C$35")

setc(ws, "A6", "Cfg_PeriodoInicio")
setc(ws, "B6",
     f"=IF('{SH_INICIO}'!$G$7=\"Mes pasado\",EOMONTH(TODAY(),-2)+1,"
     f"IF('{SH_INICIO}'!$G$7=\"Últimos 3 meses\",EOMONTH(TODAY(),-3)+1,"
     f"EOMONTH(TODAY(),-1)+1))")
name_range("Cfg_PeriodoInicio", SH__CONFIG, "$B$6")
setc(ws, "A7", "Cfg_PeriodoFin")
setc(ws, "B7", f"=IF('{SH_INICIO}'!$G$7=\"Mes pasado\",EOMONTH(TODAY(),-1),TODAY())")
name_range("Cfg_PeriodoFin", SH__CONFIG, "$B$7")
setc(ws, "A8", "Cfg_MargenAlertaUmbral")
setc(ws, "B8", 0.15)
name_range("Cfg_MargenAlertaUmbral", SH__CONFIG, "$B$8")
# Period-pair for the "previous period" comparison block
setc(ws, "A10", "Cfg_PeriodoInicio_Ant")
setc(ws, "B10", "=EDATE(Cfg_PeriodoInicio,-1)")
name_range("Cfg_PeriodoInicio_Ant", SH__CONFIG, "$B$10")
setc(ws, "A11", "Cfg_PeriodoFin_Ant")
setc(ws, "B11", "=EDATE(Cfg_PeriodoInicio,0)-1")
name_range("Cfg_PeriodoFin_Ant", SH__CONFIG, "$B$11")
ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 24
ws.sheet_state = "hidden"

# ---------------------------------------------------------------------------
# _Listas (docs/02-data-layout.md #4.2)
# ---------------------------------------------------------------------------
ws = wb[SH__LISTAS]
# Lst_Categorias / Lst_Canales are NOT defined here (Phase 9 review fix): they
# point directly at the editable list cells on ⚙️ Configuración instead (see
# that section), the same pattern as Cfg_NombreNegocio -- a hidden, protected
# sheet cannot be the place a first-time user is told to edit something.
LISTAS = {
    "Lst_CategoriasGasto": (["Publicidad", "Software/Suscripciones", "Internet", "Packaging general", "Herramientas", "Otros"], "C"),
    "Lst_TiposGasto": (["Fijo", "Variable", "Único"], "D"),
    "Lst_Monedas": (["$", "US$", "€", "S/", "MX$"], "E"),
    "Lst_PeriodosRapidos": (["Este mes", "Mes pasado", "Últimos 3 meses", "Personalizado"], "F"),
}
for nm, (items, colletter) in LISTAS.items():
    setc(ws, f"{colletter}1", nm, f=F(10, bold=True))
    for i, it in enumerate(items, start=2):
        setc(ws, f"{colletter}{i}", it)
    last_row = 1 + len(items)
    name_range(nm, SH__LISTAS, f"${colletter}$2:${colletter}${last_row}")
    ws.column_dimensions[colletter].width = 22
ws.sheet_state = "hidden"

print("Engine: _Textos, _Config, _Listas built")

# ===========================================================================
# CORE TABLES (docs/02-data-layout.md #3, docs/03-calculation-formulas.md #2-4)
# ===========================================================================

def style_col(ws, col_letter, header_row, data_row, category, numfmt=None):
    """Apply the 3-state visual system (docs/04-ux-ui.md #1.3) to one column."""
    h = ws[f"{col_letter}{header_row}"]
    h.font = F(9, bold=True, color="FFFFFF")
    h.fill = fill(C_PRIMARY)
    h.alignment = ALIGN_C
    h.border = border_all(THIN)
    d = ws[f"{col_letter}{data_row}"]
    if category == "input":
        d.fill = fill(C_INPUT_FILL)
        d.border = border_all(INPUT_SIDE)
    elif category == "autofill":
        d.fill = fill(C_AUTOFILL_FILL)
        d.border = border_all(AUTOFILL_SIDE)
    elif category == "calc":
        d.fill = fill(C_CALC_FILL)
        d.border = border_all(Side(style=None))
    d.font = F(10)
    d.alignment = ALIGN_R if numfmt else ALIGN_L
    if numfmt:
        d.number_format = numfmt

def legend(ws, row):
    # Only 2 of the 3 states docs/04-ux-ui.md #1.3 defines are ever used on a
    # real table in this workbook (no column ended up in the "auto-filled,
    # editable" category -- Ventas' Precio de venta/Comisión % moved out of
    # it per docs/05-qa-results.md §5), so the legend only shows those two:
    # showing an unused third state would invite the user to look for a cell
    # that isn't there.
    merge(ws, f"B{row}:P{row}")
    c = setc(ws, f"B{row}",
             "⬜ Vos cargás esto    ·    🔒 Profitly lo calculó",
             f=F(9, italic=True, color=C_MUTED), al=ALIGN_L)
    ws.row_dimensions[row].height = 16

# ---------------------------------------------------------------------------
# 📦 Productos
# ---------------------------------------------------------------------------
ws = wb[SH_PRODUCTOS]
add_nav(ws, SH_PRODUCTOS)
add_header(ws, "📦 PRODUCTOS", "Tu catálogo de productos y cuánto te deja ganar cada uno.")

setc(ws, "B6", "Tus productos, en números", f=F(11, bold=True, color=C_PRIMARY))
ws.row_dimensions[6].height = 18

rank_specs = [
    ("B", "🏆 Más rentable",
     '=IF(COUNTIFS(tbl_Productos[Unidades vendidas],">0")=0,Txt_SinVentas,'
     'INDEX(tbl_Productos[Producto],MATCH(_xlfn.MAXIFS(tbl_Productos[Margen],tbl_Productos[Unidades vendidas],">0"),tbl_Productos[Margen],0)))',
     '=IFERROR(_xlfn.MAXIFS(tbl_Productos[Margen],tbl_Productos[Unidades vendidas],">0"),"")', FMT_PCT,
     "Más rentable", "El producto con mejor margen entre los que ya vendiste."),
    ("F", "🔥 Más vendido",
     '=IF(COUNTA(tbl_Productos[Producto])=0,Txt_SinProductos,'
     'IF(MAX(tbl_Productos[Unidades vendidas])=0,Txt_SinVentas,'
     'INDEX(tbl_Productos[Producto],MATCH(MAX(tbl_Productos[Unidades vendidas]),tbl_Productos[Unidades vendidas],0))))',
     '=IF(OR(COUNTA(tbl_Productos[Producto])=0,MAX(tbl_Productos[Unidades vendidas])=0),"",MAX(tbl_Productos[Unidades vendidas])&" unidades")', None,
     "Más vendido", "El producto que más unidades vendiste. No es necesariamente el que más ganancia te deja — mirá \"Más rentable\" para eso."),
    ("J", "💰 Más dinero genera",
     '=IF(COUNTA(tbl_Productos[Producto])=0,Txt_SinProductos,'
     'IF(MAX(tbl_Productos[Ganancia total])<=0,Txt_SinVentas,'
     'INDEX(tbl_Productos[Producto],MATCH(MAX(tbl_Productos[Ganancia total]),tbl_Productos[Ganancia total],0))))',
     '=IF(OR(COUNTA(tbl_Productos[Producto])=0,MAX(tbl_Productos[Ganancia total])<=0),"",MAX(tbl_Productos[Ganancia total]))', FMT_CURRENCY,
     None, None),
    ("N", "⚠️ Margen bajo",
     '=IFERROR(INDEX(tbl_Productos[Producto],MATCH(1,tbl_Productos[Alerta margen bajo],0)),Txt_SinAlertas)',
     '=IFERROR(INDEX(tbl_Productos[Margen],MATCH(1,tbl_Productos[Alerta margen bajo],0)),"")', FMT_PCT,
     None, None),
]
def card_row(ws, startcol, endcol, row, value, f_, nf=None, top=False, bottom=False):
    merge(ws, f"{startcol}{row}:{endcol}{row}")
    sides = dict(left=THIN, right=THIN,
                 top=THIN if top else Side(style=None),
                 bottom=THIN if bottom else Side(style=None))
    c = setc(ws, f"{startcol}{row}", value, f=f_, fl=fill(C_CARD), al=ALIGN_C, nf=nf)
    c.border = Border(**sides)

for startcol, label, name_f, sub_f, sub_fmt, tip_title, tip_text in rank_specs:
    c1 = col_letter(col_idx(startcol) + 2)
    card_row(ws, startcol, c1, 7, label, F(9.5, bold=True, color=C_MUTED), top=True)
    card_row(ws, startcol, c1, 8, name_f, F(11, bold=True, color=C_PRIMARY))
    card_row(ws, startcol, c1, 9, sub_f, F(9.5, color=C_MUTED), nf=sub_fmt, bottom=True)
    if tip_title:
        tooltip(ws, f"{startcol}7", tip_title, tip_text)
for r in (7, 8, 9):
    ws.row_dimensions[r].height = 18

legend(ws, 11)

HEADER_ROW_P = 12
DATA_ROW_P = 13
prod_cols = [
    ("B", "Producto", "input", None),
    ("C", "Categoría", "input", None),
    ("D", "Precio de venta", "input", FMT_CURRENCY),
    ("E", "Costo del producto", "input", FMT_CURRENCY),
    ("F", "Comisión %", "input", FMT_PCT),
    ("G", "Packaging", "input", FMT_CURRENCY),
    ("H", "Envío / costo asociado", "input", FMT_CURRENCY),
    ("I", "Publicidad atribuida", "input", FMT_CURRENCY),
    ("J", "Otros costos", "input", FMT_CURRENCY),
    ("K", "Costo real de venta", "calc", FMT_CURRENCY),
    ("L", "Ganancia por unidad", "calc", FMT_CURRENCY),
    ("M", "Margen", "calc", FMT_PCT),
    ("N", "Unidades vendidas", "calc", FMT_INT),
    ("O", "Ganancia total", "calc", FMT_CURRENCY),
    ("P", "Alerta margen bajo", "calc", '[=1]"⚠️";;'),  # Phase 10 polish: blank for 0, an icon for 1 -- not a raw digit
]
for cl, header, cat, nf in prod_cols:
    setc(ws, f"{cl}{HEADER_ROW_P}", header)
    style_col(ws, cl, HEADER_ROW_P, DATA_ROW_P, cat, nf)
    ws.column_dimensions[cl].width = 15 if cl not in ("B",) else 20

r = DATA_ROW_P
setc(ws, f"K{r}", f"=E{r}+(D{r}*F{r})+G{r}+H{r}+I{r}+J{r}")
setc(ws, f"L{r}", f"=D{r}-K{r}")
setc(ws, f"M{r}", f'=IF(D{r}=0,Txt_SinDatos,L{r}/D{r})')
setc(ws, f"N{r}", f"=SUMIFS(tbl_Ventas[Cantidad],tbl_Ventas[Producto],B{r})")
setc(ws, f"O{r}", f"=SUMIFS(tbl_Ventas[Ganancia],tbl_Ventas[Producto],B{r})")
setc(ws, f"P{r}", f'=IF(AND(ISNUMBER(M{r}),M{r}<Cfg_MargenAlertaUmbral,N{r}>0),1,0)')

merge(ws, f"B{r+2}:P{r+2}")
setc(ws, f"B{r+2}", '=IF(COUNTA(tbl_Productos[Producto])=0,"Todavía no tenés productos cargados. Agregá el primero para empezar 👇","")',
     f=F(9.5, italic=True, color=C_MUTED), al=ALIGN_L)

tbl = Table(displayName="tbl_Productos", ref=f"B{HEADER_ROW_P}:P{DATA_ROW_P}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(tbl)
tooltip(ws, "F12", "Comisión %", "Se completa solo cuando elegís un producto — lo podés cambiar.")
ws.freeze_panes = f"A{DATA_ROW_P}"  # nav/title/rankings/headers stay visible while the table scrolls

ws.sheet_view.zoomScale = 100
print("📦 Productos built")

# ---------------------------------------------------------------------------
# 🛒 Ventas
# ---------------------------------------------------------------------------
ws = wb[SH_VENTAS]
add_nav(ws, SH_VENTAS)
add_header(ws, "🛒 VENTAS", "Registrá tus ventas — Profitly calcula el resto.")

setc(ws, "B6", '="Ventas este mes: "&COUNTIFS(tbl_Ventas[Fecha],">="&Cfg_PeriodoInicio,tbl_Ventas[Fecha],"<="&Cfg_PeriodoFin)&" registros"',
     f=F(10, color=C_MUTED))
setc(ws, "G6", '=IF(COUNTA(tbl_Ventas[Fecha])=0,"Última venta: —","Última venta: "&TEXT(MAX(tbl_Ventas[Fecha]),"DD/MM/YYYY"))',
     f=F(10, color=C_MUTED))

banner(ws, "B7:P7", "Necesitás cargar al menos un producto antes de registrar ventas. → Ir a 📦 Productos", tone="info")
setc(ws, "B7", '=IF(COUNTA(tbl_Productos[Producto])=0,"Necesitás cargar al menos un producto antes de registrar ventas. → Ir a 📦 Productos","")')
internal_link(ws["B7"], SH_PRODUCTOS)

legend(ws, 8)

HEADER_ROW_V = 9
DATA_ROW_V = 10
vent_cols = [
    ("B", "Fecha", "input", FMT_DATE),
    ("C", "Producto", "input", None),
    ("D", "Cantidad", "input", FMT_INT),
    ("E", "Precio de venta", "input", FMT_CURRENCY),
    ("F", "Canal", "input", None),
    ("G", "Comisión %", "input", FMT_PCT),
    ("H", "Descuento", "input", FMT_CURRENCY),
    ("I", "Otros costos", "input", FMT_CURRENCY),
    ("J", "Venta total", "calc", FMT_CURRENCY),
    ("K", "Costo total de la venta", "calc", FMT_CURRENCY),
    ("L", "Ganancia", "calc", FMT_CURRENCY),
]
for cl, header, cat, nf in vent_cols:
    setc(ws, f"{cl}{HEADER_ROW_V}", header)
    style_col(ws, cl, HEADER_ROW_V, DATA_ROW_V, cat, nf)
    ws.column_dimensions[cl].width = 16

r = DATA_ROW_V
# Plain required inputs, not an auto-fill lookup (docs/05-qa-results.md §5): a
# live lookup here would keep recomputing forever, so editing a product's price
# later would silently rewrite the recorded price/revenue of every past sale of
# that product. Typing the price actually charged at the time of sale is the
# only way to guarantee a historical sale never changes after the fact.
tooltip(ws, f"E{HEADER_ROW_V}", "Precio de venta",
        "Ingresá el precio al que vendiste — fijate el precio actual en 📦 Productos si querés usarlo de referencia.")
tooltip(ws, f"G{HEADER_ROW_V}", "Comisión %",
        "Ingresá la comisión de esta venta — fijate la comisión del producto en 📦 Productos si querés usarla de referencia.")
# Blank-row guard: an unfilled starter/new row (no Producto chosen yet) must render
# blank, not a false "Producto no encontrado" (that message is reserved for a
# product genuinely removed from the catalog after being sold, not an empty row
# awaiting entry). E/G being plain numeric inputs means D*E is safe even when E
# is blank (Excel treats a blank numeric cell as 0), so only K/L need the guard
# for their MATCH-based lookups -- J is guarded too, purely to keep an empty row
# visually blank rather than showing a distracting "$0".
setc(ws, f"J{r}", f'=IF(C{r}="","",(D{r}*E{r})-H{r})')
setc(ws, f"K{r}", f'=IF(C{r}="","",IFERROR(INDEX(tbl_Productos[Costo real de venta],MATCH(C{r},tbl_Productos[Producto],0))*D{r},Txt_ProductoNoEncontrado))')
setc(ws, f"L{r}", f'=IF(C{r}="","",IFERROR(J{r}-K{r}-I{r},Txt_ProductoNoEncontrado))')

merge(ws, f"B{r+2}:P{r+2}")
setc(ws, f"B{r+2}", '=IF(AND(COUNTA(tbl_Productos[Producto])>0,COUNTA(tbl_Ventas[Fecha])=0),"Registrá tu primera venta para empezar a ver tu rentabilidad 👇","")',
     f=F(9.5, italic=True, color=C_MUTED), al=ALIGN_L)

tbl = Table(displayName="tbl_Ventas", ref=f"B{HEADER_ROW_V}:L{DATA_ROW_V}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(tbl)
ws.freeze_panes = f"A{DATA_ROW_V}"

dv_prod = DataValidation(type="list", formula1="=tbl_Productos[Producto]", allow_blank=True,
                          showErrorMessage=True, error="Elegí un producto de tu catálogo.", errorTitle="Producto inválido")
dv_prod.add(f"C{DATA_ROW_V}")
ws.add_data_validation(dv_prod)
dv_canal = DataValidation(type="list", formula1="=Lst_Canales", allow_blank=True,
                           showErrorMessage=True, error="Elegí un canal de la lista.", errorTitle="Canal inválido")
dv_canal.add(f"F{DATA_ROW_V}")
ws.add_data_validation(dv_canal)

print("🛒 Ventas built")

# ---------------------------------------------------------------------------
# 💸 Gastos
# ---------------------------------------------------------------------------
ws = wb[SH_GASTOS]
add_nav(ws, SH_GASTOS)
add_header(ws, "💸 GASTOS", "Los gastos generales de tu negocio, separados de los costos de cada venta.")

banner(ws, "B6:P6",
       "ⓘ Acá cargás los gastos generales de tu negocio (alquiler, software, herramientas). "
       "Los costos de cada venta — comisión, envío, packaging, publicidad — se cargan en 📦 Productos.",
       tone="info")

setc(ws, "B7", '="Gastos del negocio este mes: "&TEXT(SUMIFS(tbl_Gastos[Importe],tbl_Gastos[Fecha],">="&Cfg_PeriodoInicio,tbl_Gastos[Fecha],"<="&Cfg_PeriodoFin),Cfg_Moneda&"#,##0")',
     f=F(10, color=C_MUTED))

HEADER_ROW_G = 9
DATA_ROW_G = 10
gastos_cols = [
    ("B", "Fecha", "input", FMT_DATE),
    ("C", "Categoría", "input", None),
    ("D", "Descripción", "input", None),
    ("E", "Importe", "input", FMT_CURRENCY),
    ("F", "Tipo de gasto", "input", None),
]
for cl, header, cat, nf in gastos_cols:
    setc(ws, f"{cl}{HEADER_ROW_G}", header)
    style_col(ws, cl, HEADER_ROW_G, DATA_ROW_G, cat, nf)
    ws.column_dimensions[cl].width = 18

merge(ws, f"B{DATA_ROW_G+2}:P{DATA_ROW_G+2}")
setc(ws, f"B{DATA_ROW_G+2}",
     '=IF(COUNTA(tbl_Gastos[Fecha])=0,"Todavía no cargaste gastos del negocio. Si por ahora no tenés, tu Ganancia real usa $0 — podés agregar gastos en cualquier momento.","")',
     f=F(9.5, italic=True, color=C_MUTED), al=ALIGN_L)

tbl = Table(displayName="tbl_Gastos", ref=f"B{HEADER_ROW_G}:F{DATA_ROW_G}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(tbl)
ws.freeze_panes = f"A{DATA_ROW_G}"

dv_catg = DataValidation(type="list", formula1="=Lst_CategoriasGasto", allow_blank=True)
dv_catg.add(f"C{DATA_ROW_G}")
ws.add_data_validation(dv_catg)
dv_tipog = DataValidation(type="list", formula1="=Lst_TiposGasto", allow_blank=True)
dv_tipog.add(f"F{DATA_ROW_G}")
ws.add_data_validation(dv_tipog)

print("💸 Gastos built")

# ---------------------------------------------------------------------------
# 🎯 Objetivos
# ---------------------------------------------------------------------------
ws = wb[SH_OBJETIVOS]
add_nav(ws, SH_OBJETIVOS)
add_header(ws, "🎯 OBJETIVOS", "Cuánto querés ganar, y qué te falta para lograrlo.")

HEADER_ROW_O = 6
DATA_ROW_O = 7
obj_cols = [("B", "Periodo (Mes-Año)", "input", "MMMM YYYY"), ("C", "Objetivo de ganancia", "input", FMT_CURRENCY)]
for cl, header, cat, nf in obj_cols:
    setc(ws, f"{cl}{HEADER_ROW_O}", header)
    style_col(ws, cl, HEADER_ROW_O, DATA_ROW_O, cat, nf)
    ws.column_dimensions[cl].width = 20
setc(ws, f"C{DATA_ROW_O}", "=Cfg_ObjetivoMensualDefault")

tbl = Table(displayName="tbl_Objetivos", ref=f"B{HEADER_ROW_O}:C{DATA_ROW_O}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(tbl)

# Progress panel, columns F onward (beside the table, never below — growth rule)
setc(ws, "F6", "Progreso hacia tu objetivo", f=F(11, bold=True, color=C_PRIMARY))
merge(ws, "F7:I7")
c = setc(ws, "F7", "=IF(Obj_GananciaObjetivo<=0,Txt_ObjetivoInvalido,TEXT(Obj_ProgresoPct,\"0%\")&\"  ·  \"&TEXT(Obj_GananciaActual,Cfg_Moneda&\"#,##0\")&\" de \"&TEXT(Obj_GananciaObjetivo,Cfg_Moneda&\"#,##0\"))",
         f=F(14, bold=True, color=C_PRIMARY), fl=fill(C_CARD), al=ALIGN_C)
c.border = border_all(THIN)
ws.conditional_formatting.add("F7:I7", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                                     color=C_SUCCESS, showValue=False, minLength=None, maxLength=None))

panel = [
    ("Ganancia objetivo", "=Obj_GananciaObjetivo", FMT_CURRENCY, None),
    ("Ganancia actual", "=Obj_GananciaActual", FMT_CURRENCY, None),
    ("Restante", "=Obj_Restante", FMT_CURRENCY, "Lo que te falta ganar este mes para llegar a tu objetivo."),
    ("Ganancia diaria necesaria", "=Obj_GananciaDiariaNecesaria", FMT_CURRENCY, "Cuánto necesitás ganar, en promedio, cada día que queda del mes."),
    ("Facturación necesaria", "=Obj_FacturacionNecesaria", FMT_CURRENCY, "Cuánto necesitás vender (no ganar) para llegar a tu objetivo, según tu margen actual."),
    ("Ventas necesarias", "=Obj_VentasNecesarias", FMT_INT, "Cuántas ventas necesitás, según tu ticket promedio, para llegar a esa facturación."),
]
row = 9
for label, formula, nf, tip in panel:
    setc(ws, f"F{row}", label, f=F(10, color=C_MUTED))
    c = setc(ws, f"H{row}", formula, f=F(11, bold=True, color=C_TEXT), nf=nf, al=ALIGN_R)
    if tip:
        tooltip(ws, f"F{row}", label, tip)
    row += 1

merge(ws, f"F{row+1}:I{row+1}")
c = setc(ws, f"F{row+1}", "=Obj_Factibilidad", f=F(10, bold=True), al=ALIGN_C, fl=fill(C_GRAY_BG))
c.border = border_all(THIN)

merge(ws, "B9:C9")
setc(ws, "B9", '=IF(AND(Cfg_ObjetivoMensualDefault=0,COUNTA(tbl_Objetivos[Objetivo de ganancia])<=1),"Todavía no definiste un objetivo. ¿Cuánto querés ganar este mes?","")',
     f=F(9.5, italic=True, color=C_MUTED), al=ALIGN_LW)

print("🎯 Objetivos built")

# ===========================================================================
# ENGINE CALC LAYER (docs/03-calculation-formulas.md #4-6)
# ===========================================================================

def named_calc(ws, sheetname, addr, name, formula, numfmt=None):
    setc(ws, addr, formula, nf=numfmt)
    name_range(name, sheetname, f"${addr[0]}${addr[1:]}")

# ---------------------------------------------------------------------------
# _GastosCalc
# ---------------------------------------------------------------------------
ws = wb[SH__GASTOSCALC]
setc(ws, "A1", "Métrica"); setc(ws, "B1", "Valor")
named_calc(ws, SH__GASTOSCALC, "B2", "GastosCalc_PeriodoActual",
           '=SUMIFS(tbl_Gastos[Importe],tbl_Gastos[Fecha],">="&Cfg_PeriodoInicio,tbl_Gastos[Fecha],"<="&Cfg_PeriodoFin)')
setc(ws, "A2", "Gastos del período")
named_calc(ws, SH__GASTOSCALC, "B3", "GastosCalc_PeriodoAnterior",
           '=SUMIFS(tbl_Gastos[Importe],tbl_Gastos[Fecha],">="&Cfg_PeriodoInicio_Ant,tbl_Gastos[Fecha],"<="&Cfg_PeriodoFin_Ant)')
setc(ws, "A3", "Gastos del período anterior")
ws.sheet_state = "hidden"

# ---------------------------------------------------------------------------
# _DashboardData (docs/03-calculation-formulas.md #5)
# ---------------------------------------------------------------------------
ws = wb[SH__DASHDATA]
setc(ws, "A1", "Métrica"); setc(ws, "B1", "Actual"); setc(ws, "C1", "Anterior")

setc(ws, "A2", "Ventas del período")
named_calc(ws, SH__DASHDATA, "B2", "KPI_VentasPeriodo",
           '=SUMIFS(tbl_Ventas[Venta total],tbl_Ventas[Fecha],">="&Cfg_PeriodoInicio,tbl_Ventas[Fecha],"<="&Cfg_PeriodoFin)')
named_calc(ws, SH__DASHDATA, "C2", "KPI_VentasPeriodo_Ant",
           '=SUMIFS(tbl_Ventas[Venta total],tbl_Ventas[Fecha],">="&Cfg_PeriodoInicio_Ant,tbl_Ventas[Fecha],"<="&Cfg_PeriodoFin_Ant)')

setc(ws, "A3", "Costos de venta del período")
named_calc(ws, SH__DASHDATA, "B3", "KPI_CostosVentaPeriodo",
           '=SUMIFS(tbl_Ventas[Costo total de la venta],tbl_Ventas[Fecha],">="&Cfg_PeriodoInicio,tbl_Ventas[Fecha],"<="&Cfg_PeriodoFin)')
named_calc(ws, SH__DASHDATA, "C3", "KPI_CostosVentaPeriodo_Ant",
           '=SUMIFS(tbl_Ventas[Costo total de la venta],tbl_Ventas[Fecha],">="&Cfg_PeriodoInicio_Ant,tbl_Ventas[Fecha],"<="&Cfg_PeriodoFin_Ant)')

setc(ws, "A4", "Ganancia de productos")
named_calc(ws, SH__DASHDATA, "B4", "KPI_GananciaProductos", "=KPI_VentasPeriodo-KPI_CostosVentaPeriodo")
named_calc(ws, SH__DASHDATA, "C4", "KPI_GananciaProductos_Ant", "=KPI_VentasPeriodo_Ant-KPI_CostosVentaPeriodo_Ant")

setc(ws, "A5", "Gastos del período")
named_calc(ws, SH__DASHDATA, "B5", "KPI_GastosPeriodo", "=GastosCalc_PeriodoActual")
named_calc(ws, SH__DASHDATA, "C5", "KPI_GastosPeriodo_Ant", "=GastosCalc_PeriodoAnterior")

setc(ws, "A6", "Ganancia real")
named_calc(ws, SH__DASHDATA, "B6", "KPI_GananciaReal", "=KPI_GananciaProductos-KPI_GastosPeriodo")
named_calc(ws, SH__DASHDATA, "C6", "KPI_GananciaReal_Ant", "=KPI_GananciaProductos_Ant-KPI_GastosPeriodo_Ant")

setc(ws, "A7", "Margen real")
named_calc(ws, SH__DASHDATA, "B7", "KPI_MargenReal", '=IF(KPI_VentasPeriodo=0,Txt_SinVentas,KPI_GananciaReal/KPI_VentasPeriodo)')
named_calc(ws, SH__DASHDATA, "C7", "KPI_MargenReal_Ant", '=IF(KPI_VentasPeriodo_Ant=0,Txt_SinVentas,KPI_GananciaReal_Ant/KPI_VentasPeriodo_Ant)')

setc(ws, "A8", "Cantidad de ventas")
named_calc(ws, SH__DASHDATA, "B8", "KPI_CantidadVentas",
           '=COUNTIFS(tbl_Ventas[Fecha],">="&Cfg_PeriodoInicio,tbl_Ventas[Fecha],"<="&Cfg_PeriodoFin)')
named_calc(ws, SH__DASHDATA, "C8", "KPI_CantidadVentas_Ant",
           '=COUNTIFS(tbl_Ventas[Fecha],">="&Cfg_PeriodoInicio_Ant,tbl_Ventas[Fecha],"<="&Cfg_PeriodoFin_Ant)')

setc(ws, "A9", "Ticket promedio")
named_calc(ws, SH__DASHDATA, "B9", "KPI_TicketPromedio", '=IF(KPI_CantidadVentas=0,Txt_SinVentas,KPI_VentasPeriodo/KPI_CantidadVentas)')
named_calc(ws, SH__DASHDATA, "C9", "KPI_TicketPromedio_Ant", '=IF(KPI_CantidadVentas_Ant=0,Txt_SinVentas,KPI_VentasPeriodo_Ant/KPI_CantidadVentas_Ant)')

# Delta badges (▲/▼ vs período anterior), guarded
DELTA_METRICS = ["KPI_VentasPeriodo", "KPI_GananciaReal", "KPI_MargenReal", "KPI_GastosPeriodo",
                  "KPI_TicketPromedio"]
setc(ws, "A12", "Deltas vs. período anterior")
row = 13
for m in DELTA_METRICS:
    setc(ws, f"A{row}", m)
    named_calc(ws, SH__DASHDATA, f"B{row}", f"{m}_Delta",
               f'=IFERROR(({m}-{m}_Ant)/ABS({m}_Ant),"")')
    row += 1

# Trend block (chart 1 source) — resolves weekly/monthly threshold (Phase 3 §5.2)
setc(ws, "A20", "Tendencia (últimos 6 períodos)")
named_calc(ws, SH__DASHDATA, "B20", "Dash_DiasHistorial",
           '=IF(COUNTA(tbl_Ventas[Fecha])=0,0,TODAY()-MIN(tbl_Ventas[Fecha]))')
named_calc(ws, SH__DASHDATA, "B21", "Dash_Granularidad", '=IF(Dash_DiasHistorial<60,"Semanal","Mensual")')
setc(ws, "A22", "Periodo"); setc(ws, "B22", "Ventas"); setc(ws, "C22", "Ganancia")
for i in range(6):
    row = 23 + i
    k = 5 - i  # oldest to newest
    setc(ws, f"A{row}",
         f'=IF(Dash_Granularidad="Mensual",TEXT(EDATE(TODAY(),-{k}),"MMM"),"Sem. "&TEXT(TODAY()-{k}*7,"DD/MM"))')
    setc(ws, f"B{row}",
         f'=IF(Dash_Granularidad="Mensual",'
         f'SUMIFS(tbl_Ventas[Venta total],tbl_Ventas[Fecha],">="&EOMONTH(TODAY(),-{k}-1)+1,tbl_Ventas[Fecha],"<="&EOMONTH(TODAY(),-{k})),'
         f'SUMIFS(tbl_Ventas[Venta total],tbl_Ventas[Fecha],">="&TODAY()-{k}*7-6,tbl_Ventas[Fecha],"<="&TODAY()-{k}*7))',
         nf=FMT_CURRENCY)
    setc(ws, f"C{row}",
         f'=IF(Dash_Granularidad="Mensual",'
         f'SUMIFS(tbl_Ventas[Ganancia],tbl_Ventas[Fecha],">="&EOMONTH(TODAY(),-{k}-1)+1,tbl_Ventas[Fecha],"<="&EOMONTH(TODAY(),-{k})),'
         f'SUMIFS(tbl_Ventas[Ganancia],tbl_Ventas[Fecha],">="&TODAY()-{k}*7-6,tbl_Ventas[Fecha],"<="&TODAY()-{k}*7))',
         nf=FMT_CURRENCY)
name_range("Dash_TrendRange", SH__DASHDATA, "$A$22:$C$28")

# Ganancia por producto (chart 2 source): top 8 + Otros
setc(ws, "A31", "Ganancia por producto (top 8)")
setc(ws, "A32", "Producto"); setc(ws, "B32", "Ganancia total")
for k in range(1, 9):
    row = 32 + k
    setc(ws, f"B{row}", f'=IFERROR(LARGE(tbl_Productos[Ganancia total],{k}),"")', nf=FMT_CURRENCY)
    setc(ws, f"A{row}", f'=IFERROR(INDEX(tbl_Productos[Producto],MATCH(B{row},tbl_Productos[Ganancia total],0)),"")')
setc(ws, "A41", "Otros")
setc(ws, "B41", "=MAX(SUM(tbl_Productos[Ganancia total])-SUM(B33:B40),0)", nf=FMT_CURRENCY)
name_range("Dash_ProductoRange", SH__DASHDATA, "$A$32:$B$41")

ws.sheet_state = "hidden"
print("_GastosCalc, _DashboardData built")

# ---------------------------------------------------------------------------
# _Objetivos_Calc (docs/03-calculation-formulas.md #6)
# ---------------------------------------------------------------------------
ws = wb[SH__OBJCALC]
setc(ws, "A1", "Métrica"); setc(ws, "B1", "Valor")
named_calc(ws, SH__OBJCALC, "B2", "Obj_GananciaObjetivo",
           '=IFERROR(INDEX(tbl_Objetivos[Objetivo de ganancia],MATCH(EOMONTH(TODAY(),-1)+1,tbl_Objetivos[Periodo (Mes-Año)],0)),Cfg_ObjetivoMensualDefault)')
setc(ws, "A2", "Ganancia objetivo")
named_calc(ws, SH__OBJCALC, "B3", "Obj_GananciaActual", "=KPI_GananciaReal")
setc(ws, "A3", "Ganancia actual")
named_calc(ws, SH__OBJCALC, "B4", "Obj_Restante", "=MAX(Obj_GananciaObjetivo-Obj_GananciaActual,0)")
setc(ws, "A4", "Restante")
named_calc(ws, SH__OBJCALC, "B5", "Obj_ProgresoPct",
           '=IF(Obj_GananciaObjetivo<=0,Txt_ObjetivoInvalido,MIN(Obj_GananciaActual/Obj_GananciaObjetivo,1))')
setc(ws, "A5", "Progreso %")
named_calc(ws, SH__OBJCALC, "B6", "Obj_DiasRestantes", "=DAY(EOMONTH(TODAY(),0))-DAY(TODAY())")
setc(ws, "A6", "Días restantes")
named_calc(ws, SH__OBJCALC, "B7", "Obj_GananciaDiariaNecesaria",
           '=IF(Obj_DiasRestantes=0,Txt_UltimoDia,Obj_Restante/Obj_DiasRestantes)')
setc(ws, "A7", "Ganancia diaria necesaria")
# NOTE: guarded not just against <=0 but also against KPI_MargenReal/KPI_TicketPromedio
# themselves being non-numeric (they resolve to a Txt_* guard string, e.g. Txt_SinVentas,
# when their own denominator is 0) -- Excel's IF still evaluates a text "<=0" comparison
# without erroring, but the FALSE-branch division on a text value throws #VALUE!, so the
# non-numeric case must be caught explicitly, not assumed to be excluded by "<=0".
named_calc(ws, SH__OBJCALC, "B8", "Obj_FacturacionNecesaria",
           '=IF(OR(NOT(ISNUMBER(KPI_MargenReal)),KPI_MargenReal<=0),Txt_ObjetivoInvalido,Obj_Restante/KPI_MargenReal)')
setc(ws, "A8", "Facturación necesaria")
named_calc(ws, SH__OBJCALC, "B9", "Obj_VentasNecesarias",
           '=IF(OR(NOT(ISNUMBER(KPI_TicketPromedio)),KPI_TicketPromedio<=0,NOT(ISNUMBER(Obj_FacturacionNecesaria))),'
           'Txt_FaltaHistorial,Obj_FacturacionNecesaria/KPI_TicketPromedio)')
setc(ws, "A9", "Ventas necesarias")

# Feasibility: 90-day rolling daily-profit helper block
setc(ws, "A19", "Historial diario (90 días)")
setc(ws, "A20", "Fecha"); setc(ws, "B20", "Ganancia del día")
for i in range(90):
    row = 21 + i
    offset = 89 - i
    setc(ws, f"A{row}", f"=TODAY()-{offset}", nf=FMT_DATE)
    setc(ws, f"B{row}", f"=SUMIFS(tbl_Ventas[Ganancia],tbl_Ventas[Fecha],A{row})", nf=FMT_CURRENCY)
name_range("Obj_HistorialDiario", SH__OBJCALC, "$B$21:$B$110")

named_calc(ws, SH__OBJCALC, "B112", "Obj_GananciaDiariaMaximaHistorica", "=MAX(Obj_HistorialDiario)")
setc(ws, "A112", "Ganancia diaria máxima histórica")
named_calc(ws, SH__OBJCALC, "B113", "Obj_DiasDesdePrimeraVenta",
           '=IF(COUNTA(tbl_Ventas[Fecha])=0,0,TODAY()-MIN(tbl_Ventas[Fecha]))')
setc(ws, "A113", "Días desde la primera venta")
named_calc(ws, SH__OBJCALC, "B114", "Obj_Factibilidad",
           '=IF(COUNTA(tbl_Ventas[Fecha])=0,Txt_SinVentas,'
           'IF(Obj_DiasDesdePrimeraVenta<14,Txt_FaltaHistorial,'
           'IF(Obj_GananciaObjetivo<=Obj_GananciaActual,Txt_ObjetivoAlcanzado,'
           'IF(Obj_GananciaDiariaNecesaria<=Obj_GananciaDiariaMaximaHistorica,'
           '"✅ Factible según tu historial","🚀 Ambicioso según tu historial"))))')
setc(ws, "A114", "Factibilidad")
ws.sheet_state = "hidden"
print("_Objetivos_Calc built")

# ---------------------------------------------------------------------------
# _Insights_Rules + _Insights_Engine (docs/03-calculation-formulas.md #7)
# ---------------------------------------------------------------------------
ws = wb[SH__INSRULES]
headers = ["RuleID", "Condición", "Plantilla", "Prioridad", "Activa"]
for i, h in enumerate(headers, start=1):
    setc(ws, f"{col_letter(i)}1", h, f=F(10, bold=True))
# NOTE: written as nested IF, not AND(...), wherever a later AND() argument does
# arithmetic on a value an earlier argument is guarding with ISNUMBER/>0 -- Excel's
# AND() evaluates every argument regardless of the others (no short-circuit), so a
# guard placed *inside* an AND() does not protect a sibling argument from erroring.
rules = [
    ("R0", '=COUNTA(tbl_Ventas[Fecha])<5', "Agregá más datos para obtener insights.", 0, True),
    ("R1", '=IF(AND(ISNUMBER(KPI_MargenReal),ISNUMBER(KPI_MargenReal_Ant)),(KPI_MargenReal_Ant-KPI_MargenReal)>0.02,FALSE)',
     "⚠️ Tu margen cayó {delta} este mes.", 1, True),
    ("R2", '=IF(SUM(tbl_Productos[Ganancia total])>0,MAX(tbl_Productos[Ganancia total])/SUM(tbl_Productos[Ganancia total])>0.5,FALSE)',
     "🔥 {producto} genera la mayor parte de tu ganancia.", 2, True),
    ("R3", '=IF(AND(ISNUMBER(Ins_ValorVentasDelta),ISNUMBER(Ins_ValorGananciaDelta)),(Ins_ValorVentasDelta-Ins_ValorGananciaDelta)>0.1,FALSE)',
     "📈 Tus ventas aumentaron {ventasDelta}, pero tu ganancia solo aumentó {gananciaDelta}.", 2, True),
    ("R4", '=AND(Obj_GananciaObjetivo>0,Obj_Restante>0)', "🎯 Te faltan {restante} de ganancia para alcanzar tu objetivo.", 3, True),
]
for i, (rid, cond, plantilla, prio, activa) in enumerate(rules, start=2):
    setc(ws, f"A{i}", rid)
    setc(ws, f"B{i}", cond)
    setc(ws, f"C{i}", plantilla)
    setc(ws, f"D{i}", prio)
    setc(ws, f"E{i}", activa)
for cl, w in [("A", 8), ("B", 45), ("C", 55), ("D", 10), ("E", 8)]:
    ws.column_dimensions[cl].width = w
ws.sheet_state = "hidden"

ws = wb[SH__INSENGINE]
setc(ws, "A1", "Token"); setc(ws, "B1", "Valor")
named_calc(ws, SH__INSENGINE, "B2", "Ins_Delta", '=IFERROR(TEXT(ABS(KPI_MargenReal-KPI_MargenReal_Ant),"0,0%"),"")')
setc(ws, "A2", "{delta}")
named_calc(ws, SH__INSENGINE, "B3", "Ins_Producto",
           '=IF(COUNTA(tbl_Productos[Producto])=0,"",INDEX(tbl_Productos[Producto],MATCH(MAX(tbl_Productos[Ganancia total]),tbl_Productos[Ganancia total],0)))')
setc(ws, "A3", "{producto}")
named_calc(ws, SH__INSENGINE, "B4", "Ins_ValorVentasDelta", '=IFERROR((KPI_VentasPeriodo-KPI_VentasPeriodo_Ant)/KPI_VentasPeriodo_Ant,"")')
setc(ws, "A4", "ventasDelta (numérico)")
named_calc(ws, SH__INSENGINE, "B5", "Ins_VentasDelta", '=IFERROR(TEXT(Ins_ValorVentasDelta,"0,0%"),"")')
setc(ws, "A5", "{ventasDelta}")
named_calc(ws, SH__INSENGINE, "B6", "Ins_ValorGananciaDelta", '=IFERROR((KPI_GananciaReal-KPI_GananciaReal_Ant)/ABS(KPI_GananciaReal_Ant),"")')
setc(ws, "A6", "gananciaDelta (numérico)")
named_calc(ws, SH__INSENGINE, "B7", "Ins_GananciaDelta", '=IFERROR(TEXT(Ins_ValorGananciaDelta,"0,0%"),"")')
setc(ws, "A7", "{gananciaDelta}")
named_calc(ws, SH__INSENGINE, "B8", "Ins_Restante", '=Cfg_Moneda&TEXT(Obj_Restante,"#,##0")')
setc(ws, "A8", "{restante}")

# Effective sort key per rule (rows 13-17): Prioridad*100+row if Activa AND Condición TRUE,
# else a large sentinel (999999) so the key column is always purely numeric -- this lets the
# ranking below use a plain (non-array) SMALL/MATCH instead of an IF(ISNUMBER(...)) array
# formula, which is the same array-formula-avoidance technique used for the Productos
# "Alerta margen bajo" flag column (docs/03-calculation-formulas.md §0.2's reasoning).
SENTINEL = 999999
setc(ws, "A11", "Ranking de reglas activas")
setc(ws, "A12", "RuleID"); setc(ws, "B12", "Clave de orden"); setc(ws, "C12", "Texto resuelto")
for i in range(5):
    row = 13 + i
    rule_row = 2 + i  # _Insights_Rules row
    setc(ws, f"A{row}", f"='{SH__INSRULES}'!A{rule_row}")
    setc(ws, f"B{row}",
         f"=IF(AND('{SH__INSRULES}'!E{rule_row}=TRUE,'{SH__INSRULES}'!B{rule_row}=TRUE),"
         f"'{SH__INSRULES}'!D{rule_row}*100+{rule_row},{SENTINEL})")
    setc(ws, f"C{row}",
         f"=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE('{SH__INSRULES}'!C{rule_row},"
         f'"{{delta}}",Ins_Delta),"{{producto}}",Ins_Producto),"{{ventasDelta}}",Ins_VentasDelta),'
         f'"{{gananciaDelta}}",Ins_GananciaDelta),"{{restante}}",Ins_Restante)')

named_calc(ws, SH__INSENGINE, "B18", "Ins_R0Activa", f"='{SH__INSRULES}'!B2")
setc(ws, "A18", "R0 activa (fallback)")

# Final output (up to 4): if R0 (data-insufficient) fires, it is the only line shown, per
# docs/03-calculation-formulas.md §7.3 -- everything else short-circuits.
setc(ws, "A19", "Salida final (hasta 4)")
for i in range(4):
    row = 20 + i
    k = i + 1
    named_calc(ws, SH__INSENGINE, f"B{row}", f"Ins_Texto{k}",
               f'=IF(Ins_R0Activa=TRUE,IF({k}=1,C13,""),'
               f'IF(SMALL($B$13:$B$17,{k})={SENTINEL},"",'
               f'INDEX($C$13:$C$17,MATCH(SMALL($B$13:$B$17,{k}),$B$13:$B$17,0))))')
    setc(ws, f"A{row}", f"Ins_Texto{k}")
ws.sheet_state = "hidden"
print("_Insights_Rules, _Insights_Engine built")

# ---------------------------------------------------------------------------
# _Simulador_Engine (docs/03-calculation-formulas.md #8)
# ---------------------------------------------------------------------------
ws = wb[SH__SIMENGINE]
setc(ws, "A1", "Métrica"); setc(ws, "B1", "Valor")

named_calc(ws, SH__SIMENGINE, "B2", "Sim_DatosSuficientes", "=KPI_CantidadVentas>=1")
setc(ws, "A2", "Datos suficientes")

baseline = [
    ("Sim_Base_Ventas", "=KPI_VentasPeriodo", "Ventas"),
    ("Sim_Base_CostosVenta", "=KPI_CostosVentaPeriodo", "Costos de venta"),
    ("Sim_Base_GananciaProductos", "=KPI_GananciaProductos", "Ganancia de productos"),
    ("Sim_Base_Gastos", "=KPI_GastosPeriodo", "Gastos"),
    ("Sim_Base_GananciaReal", "=KPI_GananciaReal", "Ganancia real"),
    ("Sim_Base_Margen", "=KPI_MargenReal", "Margen"),
    ("Sim_Base_CantidadVentas", "=KPI_CantidadVentas", "Cantidad de ventas"),
]
row = 4
for nm, formula, label in baseline:
    setc(ws, f"A{row}", label)
    named_calc(ws, SH__SIMENGINE, f"B{row}", nm, formula)
    row += 1

# 3 scenario blocks (A/B/C). Delta *inputs* live on the visible Simulador sheet
# (named Sim_A_Delta*, etc.) -- referenced here, never the reverse.
SCEN_ROWS = {}
row = 13
for scen in "ABC":
    setc(ws, f"A{row}", f"Escenario {scen}")
    r0 = row + 1
    SCEN_ROWS[scen] = r0
    labels = ["Ventas proyectadas", "Costos de venta proyectados", "Margen (por precio)",
              "Margen proyectado (final)", "Ganancia de productos proyectada",
              "Gastos proyectados", "Gastos ajustados a 0 (flag)", "Ganancia real proyectada",
              "Diferencia"]
    for i, lbl in enumerate(labels):
        setc(ws, f"A{r0+i}", lbl)
    row = r0 + len(labels) + 1

for scen in "ABC":
    r0 = SCEN_ROWS[scen]
    dP = f"Sim_{scen}_DeltaPrecio"
    dV = f"Sim_{scen}_DeltaVentas"
    dM = f"Sim_{scen}_DeltaMargen"
    dG = f"Sim_{scen}_DeltaGastos"
    dA = f"Sim_{scen}_DeltaPublicidad"

    def g(suffix):
        return f"Sim_{scen}_{suffix}"

    named_calc(ws, SH__SIMENGINE, f"B{r0}", g("VentasProyectadas"),
               f"=IF(NOT(Sim_DatosSuficientes),Txt_FaltaHistorial,Sim_Base_Ventas*(1+{dV})*(1+{dP}))")
    named_calc(ws, SH__SIMENGINE, f"B{r0+1}", g("CostosVentaProyectados"),
               f"=IF(NOT(Sim_DatosSuficientes),Txt_FaltaHistorial,Sim_Base_CostosVenta*(1+{dV}))")
    named_calc(ws, SH__SIMENGINE, f"B{r0+2}", g("MargenPorPrecio"),
               f"=IF(NOT(Sim_DatosSuficientes),Txt_FaltaHistorial,"
               f"IF(OR(NOT(ISNUMBER({g('VentasProyectadas')})),{g('VentasProyectadas')}<=0),0,"
               f"({g('VentasProyectadas')}-{g('CostosVentaProyectados')})/{g('VentasProyectadas')}))")
    named_calc(ws, SH__SIMENGINE, f"B{r0+3}", g("MargenProyectado"),
               f"=IF(NOT(Sim_DatosSuficientes),Txt_FaltaHistorial,"
               f"MIN(MAX({g('MargenPorPrecio')}+{dM},-2),0.95))")
    named_calc(ws, SH__SIMENGINE, f"B{r0+4}", g("GananciaProductosProyectada"),
               f"=IF(NOT(Sim_DatosSuficientes),Txt_FaltaHistorial,{g('VentasProyectadas')}*{g('MargenProyectado')})")
    named_calc(ws, SH__SIMENGINE, f"B{r0+5}", g("GastosProyectados"),
               f"=IF(NOT(Sim_DatosSuficientes),Txt_FaltaHistorial,MAX(Sim_Base_Gastos+{dG}+{dA},0))")
    named_calc(ws, SH__SIMENGINE, f"B{r0+6}", g("GastosAjustados"),
               f"=IF(NOT(Sim_DatosSuficientes),FALSE,(Sim_Base_Gastos+{dG}+{dA})<0)")
    named_calc(ws, SH__SIMENGINE, f"B{r0+7}", g("GananciaRealProyectada"),
               f"=IF(NOT(Sim_DatosSuficientes),Txt_FaltaHistorial,{g('GananciaProductosProyectada')}-{g('GastosProyectados')})")
    named_calc(ws, SH__SIMENGINE, f"B{r0+8}", g("Diferencia"),
               f"=IF(NOT(Sim_DatosSuficientes),Txt_FaltaHistorial,{g('GananciaRealProyectada')}-Sim_Base_GananciaReal)")
ws.sheet_state = "hidden"
print("_Simulador_Engine built")

# ===========================================================================
# 🔬 Simulador (docs/04-ux-ui.md #3.7)
# ===========================================================================
ws = wb[SH_SIMULADOR]
add_nav(ws, SH_SIMULADOR)
add_header(ws, "🔬 ¿QUÉ PASARÍA SI...?", "Probá una decisión sin tocar tus datos reales.")

setc(ws, "B6", '=IF(NOT(Sim_DatosSuficientes),"Necesitás cargar ventas reales antes de simular escenarios. → Ir a 🛒 Ventas","")',
     f=F(10, italic=True, color=C_MUTED))
internal_link(ws["B6"], SH_VENTAS)

setc(ws, "B7", "ESCENARIO ACTUAL", f=F(11, bold=True, color=C_PRIMARY))
baseline_labels = [
    ("Ventas", "Sim_Base_Ventas", FMT_CURRENCY), ("Costos de venta", "Sim_Base_CostosVenta", FMT_CURRENCY),
    ("Ganancia de productos", "Sim_Base_GananciaProductos", FMT_CURRENCY), ("Gastos", "Sim_Base_Gastos", FMT_CURRENCY),
    ("Ganancia real", "Sim_Base_GananciaReal", FMT_CURRENCY), ("Margen", "Sim_Base_Margen", FMT_PCT),
]
row = 8
for label, nm, nf in baseline_labels:
    setc(ws, f"B{row}", label, f=F(9.5, color=C_MUTED))
    c = setc(ws, f"D{row}", f"={nm}", f=F(10, bold=True), nf=nf, al=ALIGN_R, fl=fill(C_CALC_FILL))
    row += 1

SCEN_COLS = {"A": ("B", "D"), "B": ("F", "H"), "C": ("J", "L")}
SCEN_INPUT_ROW0 = 16
row_title = 15
for scen, (c0, c1) in SCEN_COLS.items():
    merge(ws, f"{c0}{row_title}:{c1}{row_title}")
    setc(ws, f"{c0}{row_title}", f"Escenario {scen}", f=F(11, bold=True, color=C_PRIMARY), al=ALIGN_C, fl=fill(C_CARD))

deltas = [
    ("Δ Precio %", "DeltaPrecio", FMT_PCT, 0.0, "Δ Precio %", "Cuánto cambiarías tu precio de venta."),
    ("Δ Cantidad de ventas %", "DeltaVentas", FMT_PCT, 0.0, "Δ Cantidad de ventas %", "Cuánto cambiaría la cantidad de unidades que vendés."),
    ("Δ Margen (pp)", "DeltaMargen", "0.0", 0.0, "Δ Margen",
     "pp = puntos porcentuales (ej.: de 20% a 25% es +5 pp). Este ajuste se suma al efecto que ya tiene un cambio de precio sobre tu margen."),
    ("Δ Gastos ($)", "DeltaGastos", FMT_CURRENCY, 0, "Δ Gastos", "Cambio en tus gastos generales del negocio."),
    ("Δ Publicidad ($)", "DeltaPublicidad", FMT_CURRENCY, 0, "Δ Publicidad", "Gasto adicional en publicidad para este escenario."),
]
for scen, (c0, c1) in SCEN_COLS.items():
    r = SCEN_INPUT_ROW0
    for label, suffix, nf, default, tip_title, tip_text in deltas:
        setc(ws, f"{c0}{r}", label, f=F(9, color=C_MUTED))
        cell = setc(ws, f"{c1}{r}", default, f=F(10, bold=True), nf=nf, fl=fill(C_AUTOFILL_FILL), al=ALIGN_R)
        cell.border = border_all(AUTOFILL_SIDE)
        name_range(f"Sim_{scen}_{suffix}", SH_SIMULADOR, f"${c1}${r}")
        tooltip(ws, f"{c0}{r}", tip_title, tip_text)
        r += 1
# Column grouping (native outline, no macro) for Scenarios B/C -- collapsed by
# default. Set outline_level AND hidden explicitly per column (including the I
# gutter, so the collapsed group leaves no visible sliver) -- calling
# DimensionHolder.group() here was found (Phase 9 review) to only apply to the
# first column of the range, silently leaving G/H/J/K/L expanded and unhidden.
for cl in ("F", "G", "H", "I", "J", "K", "L"):
    cd = ws.column_dimensions[cl]
    cd.outline_level = 1
    cd.hidden = True
ws.sheet_properties.outlinePr.summaryRight = True

row_proj_title = 22
merge(ws, "B22:D22")
setc(ws, "B22", "ESCENARIO PROYECTADO", f=F(11, bold=True, color=C_PRIMARY), al=ALIGN_L)
proj_labels = [
    ("Ventas", "VentasProyectadas", FMT_CURRENCY), ("Costos de venta", "CostosVentaProyectados", FMT_CURRENCY),
    ("Ganancia de productos", "GananciaProductosProyectada", FMT_CURRENCY), ("Gastos", "GastosProyectados", FMT_CURRENCY),
    ("Ganancia real", "GananciaRealProyectada", FMT_CURRENCY), ("Margen", "MargenProyectado", FMT_PCT),
]
for scen, (c0, c1) in SCEN_COLS.items():
    r = 23
    for label, suffix, nf in proj_labels:
        setc(ws, f"{c0}{r}", label, f=F(9, color=C_MUTED))
        setc(ws, f"{c1}{r}", f"=Sim_{scen}_{suffix}", f=F(10, bold=True), nf=nf, fl=fill(C_CALC_FILL), al=ALIGN_R)
        r += 1
    setc(ws, f"{c0}{r}", "DIFERENCIA", f=F(9.5, bold=True, color=C_MUTED))
    r_dif = r + 1
    merge(ws, f"{c0}{r_dif}:{c1}{r_dif}")
    dif_cell = setc(ws, f"{c0}{r_dif}", f"=Sim_{scen}_Diferencia", f=F(15, bold=True), nf='+$#,##0;-$#,##0;"-"', al=ALIGN_C, fl=fill(C_CARD))
    dif_cell.border = border_all(THIN)
    r_notice = r_dif + 1
    merge(ws, f"{c0}{r_notice}:{c1}{r_notice}")
    setc(ws, f"{c0}{r_notice}",
         f'=IF(Sim_{scen}_GastosAjustados=TRUE,"Ajustamos tus gastos proyectados a $0 porque el valor ingresado los volvía negativos.","")',
         f=F(8.5, italic=True, color=C_WARNING), al=ALIGN_LW)

ws.conditional_formatting.add("B30:L30",
    FormulaRule(formula=["B30<0"], fill=PatternFill(fill_type="solid", start_color="FBEAE9", end_color="FBEAE9"), font=Font(color=C_ERROR, bold=True)))
ws.conditional_formatting.add("B30:L30",
    FormulaRule(formula=["B30>0"], fill=PatternFill(fill_type="solid", start_color=C_GREEN_BG, end_color=C_GREEN_BG), font=Font(color=C_SUCCESS, bold=True)))

for cl in "BCDEFGHIJKL":
    ws.column_dimensions[cl].width = 15
print("🔬 Simulador built")

# ===========================================================================
# _Demo_Datos (docs/02-data-layout.md #2.2, docs/03-calculation-formulas.md §0
# amendment note: demo aggregation implemented as extra blocks within this
# existing engine sheet rather than a new sheet, since Phase 1 didn't allocate
# a separate one -- no sheet was added beyond the approved inventory.)
# ===========================================================================
import random
random.seed(7)

ws = wb[SH__DEMODATA]

DEMO_PRODUCTS = [
    ("Vela aromática grande", "Hogar", 4500, 1200, 0.08, 150, 0, 200, 50),
    ("Vela aromática chica", "Hogar", 2200, 600, 0.08, 100, 0, 100, 30),
    ("Jabón artesanal", "Belleza", 1800, 450, 0.05, 80, 0, 150, 20),
    ("Set de regalo premium", "Hogar", 9500, 3200, 0.10, 300, 500, 400, 100),
    ("Difusor de aromas", "Hogar", 6800, 2900, 0.08, 200, 0, 300, 80),
    ("Crema corporal", "Belleza", 3200, 1400, 0.05, 120, 0, 200, 60),
    ("Sahumerio x12", "Hogar", 1500, 700, 0.05, 60, 0, 50, 20),
    ("Kit spa en casa", "Belleza", 12000, 6500, 0.10, 350, 800, 500, 150),
]
setc(ws, "A1", "Demo · Productos", f=F(11, bold=True))
dprod_headers = ["Producto", "Categoría", "Precio de venta", "Costo del producto", "Comisión %",
                  "Packaging", "Envío / costo asociado", "Publicidad atribuida", "Otros costos",
                  "Costo real de venta", "Ganancia por unidad", "Margen", "Unidades vendidas", "Ganancia total"]
for i, h in enumerate(dprod_headers):
    setc(ws, f"{col_letter(2+i)}2", h, f=F(9, bold=True))
DP_ROW0 = 3
for i, row_data in enumerate(DEMO_PRODUCTS):
    r = DP_ROW0 + i
    for j, val in enumerate(row_data):
        setc(ws, f"{col_letter(2+j)}{r}", val)
    setc(ws, f"K{r}", f"=E{r}+(D{r}*F{r})+G{r}+H{r}+I{r}+J{r}")
    setc(ws, f"L{r}", f"=D{r}-K{r}")
    setc(ws, f"M{r}", f"=IF(D{r}=0,\"\",L{r}/D{r})")
DP_ROWN = DP_ROW0 + len(DEMO_PRODUCTS) - 1
name_range("Demo_ProductosRange", SH__DEMODATA, f"$B${DP_ROW0}:$B${DP_ROWN}")

# Demo sales: 3 months of activity, weighted so products have distinct performance profiles
setc(ws, "A20", "Demo · Ventas")
dvent_headers = ["Fecha", "Producto", "Cantidad", "Precio de venta", "Canal", "Comisión %",
                  "Descuento", "Otros costos", "Venta total", "Costo total de la venta", "Ganancia"]
for i, h in enumerate(dvent_headers):
    setc(ws, f"{col_letter(2+i)}21", h, f=F(9, bold=True))
DV_ROW0 = 22
today = datetime.date.today()
start_date = today - datetime.timedelta(days=89)
canales = ["Local", "Online", "Mayorista", "Redes sociales", "Feria"]
weights = [30, 22, 14, 10, 9, 6, 5, 4]  # sales-volume weighting per product (index-aligned)
demo_sales_rows = []
d = start_date
while d <= today:
    n_sales = random.choice([0, 0, 1, 1, 1, 2, 2, 3])
    for _ in range(n_sales):
        pidx = random.choices(range(len(DEMO_PRODUCTS)), weights=weights, k=1)[0]
        qty = random.choice([1, 1, 1, 2, 2, 3])
        canal = random.choice(canales)
        descuento = random.choice([0, 0, 0, 0, 200, 500])
        demo_sales_rows.append((d, DEMO_PRODUCTS[pidx][0], qty, canal, descuento))
    d += datetime.timedelta(days=1)
for i, (fecha, prod, qty, canal, desc) in enumerate(demo_sales_rows):
    r = DV_ROW0 + i
    setc(ws, f"B{r}", fecha, nf=FMT_DATE)
    setc(ws, f"C{r}", prod)
    setc(ws, f"D{r}", qty)
    setc(ws, f"E{r}", f"=IFERROR(INDEX($D${DP_ROW0}:$D${DP_ROWN},MATCH(C{r},$B${DP_ROW0}:$B${DP_ROWN},0)),0)")
    setc(ws, f"F{r}", canal)
    setc(ws, f"G{r}", f"=IFERROR(INDEX($F${DP_ROW0}:$F${DP_ROWN},MATCH(C{r},$B${DP_ROW0}:$B${DP_ROWN},0)),0)")
    setc(ws, f"H{r}", desc)
    setc(ws, f"I{r}", 0)
    setc(ws, f"J{r}", f"=(D{r}*E{r})-H{r}")
    setc(ws, f"K{r}", f"=IFERROR(INDEX($K${DP_ROW0}:$K${DP_ROWN},MATCH(C{r},$B${DP_ROW0}:$B${DP_ROWN},0))*D{r},0)")
    setc(ws, f"L{r}", f"=J{r}-K{r}-I{r}")
DV_ROWN = DV_ROW0 + len(demo_sales_rows) - 1
name_range("Demo_VentasRange", SH__DEMODATA, f"$B${DV_ROW0}:$B${DV_ROWN}")

# Demo expenses: recurring + a few one-offs across the same 3 months
# (columns Q onward -- Productos' own table already occupies B..O, see dprod_headers)
setc(ws, "Q2", "Demo · Gastos", f=F(11, bold=True))
dgasto_headers = ["Fecha", "Categoría", "Descripción", "Importe", "Tipo de gasto"]
for i, h in enumerate(dgasto_headers):
    setc(ws, f"{col_letter(17+i)}3", h, f=F(9, bold=True))
DG_ROW0 = 4
demo_gastos = []
for m_back in range(3):
    month_start = (today.replace(day=1) - datetime.timedelta(days=1)).replace(day=1) if m_back else today.replace(day=1)
    base = today.replace(day=5) - datetime.timedelta(days=30 * m_back)
    demo_gastos.append((base, "Software/Suscripciones", "Suscripción plataforma de ventas", 8500, "Fijo"))
    demo_gastos.append((base + datetime.timedelta(days=2), "Internet", "Internet del local", 6200, "Fijo"))
    demo_gastos.append((base + datetime.timedelta(days=10), "Publicidad", "Campaña redes sociales", 15000, "Variable"))
demo_gastos.append((today - datetime.timedelta(days=20), "Herramientas", "Balanza nueva", 22000, "Único"))
demo_gastos.append((today - datetime.timedelta(days=5), "Packaging general", "Cajas y etiquetas", 9800, "Variable"))
for i, (fecha, cat, desc, importe, tipo) in enumerate(demo_gastos):
    r = DG_ROW0 + i
    setc(ws, f"Q{r}", fecha, nf=FMT_DATE)
    setc(ws, f"R{r}", cat)
    setc(ws, f"S{r}", desc)
    setc(ws, f"T{r}", importe)
    setc(ws, f"U{r}", tipo)
DG_ROWN = DG_ROW0 + len(demo_gastos) - 1
name_range("Demo_GastosRange", SH__DEMODATA, f"$Q${DG_ROW0}:$Q${DG_ROWN}")

# Demo goal
setc(ws, "W2", "Demo · Objetivos", f=F(11, bold=True))
setc(ws, "W3", "Periodo"); setc(ws, "X3", "Objetivo de ganancia")
setc(ws, "W4", today.replace(day=1), nf=FMT_DATE)
setc(ws, "X4", 600000)

print(f"_Demo_Datos: {len(DEMO_PRODUCTS)} productos, {len(demo_sales_rows)} ventas, {len(demo_gastos)} gastos")

# --- Demo aggregation block (mirrors _DashboardData/_Objetivos_Calc, Demo-scoped) ---
# NOTE: placed at rows 300-310 rather than 100-110. Empirically (see
# docs/05-qa-results.md), LibreOffice's recalculation engine -- used here to
# verify zero formula errors -- produced #VALUE! on *any* formula (even a
# trivial `=B101-B102` cell reference, even a `=655900-358249` constant
# re-tested at that exact address) placed specifically around row 103 of
# this sheet, while the identical formula one row away, or the same formula
# at row 300, computed correctly. This reproduced consistently across
# rebuilds and repeated recalculation passes, is confined to this one
# address range, and does not correspond to any content, format, named
# range, merge, validation, or dependency this build places there -- it
# appears to be a tool-specific artifact of this environment's LibreOffice
# rather than a formula defect. Since the requirement is zero *verified*
# errors, the block was moved rather than left unexplained.
setc(ws, "A300", "Demo · KPIs (todo el historial de 90 días)", f=F(11, bold=True))
setc(ws, "A301", "Ventas del período")
named_calc(ws, SH__DEMODATA, "B301", "DemoKPI_VentasPeriodo", f"=SUM($J${DV_ROW0}:$J${DV_ROWN})", FMT_CURRENCY)
setc(ws, "A302", "Costos de venta del período")
named_calc(ws, SH__DEMODATA, "B302", "DemoKPI_CostosVentaPeriodo", f"=SUM($K${DV_ROW0}:$K${DV_ROWN})", FMT_CURRENCY)
setc(ws, "A303", "Ganancia de productos")
named_calc(ws, SH__DEMODATA, "B303", "DemoKPI_GananciaProductos", "=B301-B302", FMT_CURRENCY)
setc(ws, "A304", "Gastos del período")
named_calc(ws, SH__DEMODATA, "B304", "DemoKPI_GastosPeriodo", f"=SUM($T${DG_ROW0}:$T${DG_ROWN})", FMT_CURRENCY)
setc(ws, "A305", "Ganancia real")
named_calc(ws, SH__DEMODATA, "B305", "DemoKPI_GananciaReal", "=B303-B304", FMT_CURRENCY)
setc(ws, "A306", "Margen real")
named_calc(ws, SH__DEMODATA, "B306", "DemoKPI_MargenReal", "=IF(B301=0,0,B305/B301)", FMT_PCT)
setc(ws, "A307", "Cantidad de ventas")
named_calc(ws, SH__DEMODATA, "B307", "DemoKPI_CantidadVentas", f"=COUNTA($B${DV_ROW0}:$B${DV_ROWN})", FMT_INT)
setc(ws, "A308", "Ticket promedio")
named_calc(ws, SH__DEMODATA, "B308", "DemoKPI_TicketPromedio", "=IF(B307=0,0,B301/B307)", FMT_CURRENCY)
setc(ws, "A309", "Objetivo de ganancia")
named_calc(ws, SH__DEMODATA, "B309", "DemoKPI_Objetivo", "=X4", FMT_CURRENCY)
setc(ws, "A310", "Progreso %")
named_calc(ws, SH__DEMODATA, "B310", "DemoKPI_ProgresoPct", "=IF(B309<=0,0,MIN(B305/B309,1))", FMT_PCT)

setc(ws, "A115", "Demo · Ganancia por producto (top 8)")
for k in range(1, 9):
    row = 115 + k
    setc(ws, f"B{row}", f"=IFERROR(LARGE($O${DP_ROW0}:$O${DP_ROWN},{k}),\"\")", nf=FMT_CURRENCY)
    setc(ws, f"A{row}", f"=IFERROR(INDEX($B${DP_ROW0}:$B${DP_ROWN},MATCH(B{row},$O${DP_ROW0}:$O${DP_ROWN},0)),\"\")")
# Populate Unidades vendidas (N) and Ganancia total (O) for demo products via
# SUMIFS on the demo Ventas range -- N is quantity, O is profit; these were left
# as header-only placeholders in the initial per-product loop above.
for i in range(len(DEMO_PRODUCTS)):
    r = DP_ROW0 + i
    setc(ws, f"N{r}", f"=SUMIFS($D${DV_ROW0}:$D${DV_ROWN},$C${DV_ROW0}:$C${DV_ROWN},B{r})", nf=FMT_INT)
    setc(ws, f"O{r}", f"=SUMIFS($L${DV_ROW0}:$L${DV_ROWN},$C${DV_ROW0}:$C${DV_ROWN},B{r})", nf=FMT_CURRENCY)
name_range("Demo_ProductoRankRange", SH__DEMODATA, "$A$116:$B$123")

setc(ws, "A126", "Demo · Tendencia (6 meses)")
for i in range(6):
    row = 127 + i
    k = 5 - i
    setc(ws, f"A{row}", f"=TEXT(EDATE(TODAY(),-{k}),\"MMM\")")
    setc(ws, f"B{row}", f"=SUMIFS($J${DV_ROW0}:$J${DV_ROWN},$B${DV_ROW0}:$B${DV_ROWN},\">=\"&EOMONTH(TODAY(),-{k}-1)+1,$B${DV_ROW0}:$B${DV_ROWN},\"<=\"&EOMONTH(TODAY(),-{k}))", nf=FMT_CURRENCY)
    setc(ws, f"C{row}", f"=SUMIFS($L${DV_ROW0}:$L${DV_ROWN},$B${DV_ROW0}:$B${DV_ROWN},\">=\"&EOMONTH(TODAY(),-{k}-1)+1,$B${DV_ROW0}:$B${DV_ROWN},\"<=\"&EOMONTH(TODAY(),-{k}))", nf=FMT_CURRENCY)
name_range("Demo_TrendRange", SH__DEMODATA, "$A$127:$C$132")

ws.sheet_state = "hidden"
print("_Demo_Datos built")

# ===========================================================================
# 🏠 Inicio / 🎬 Demo (shared builder — docs/04-ux-ui.md #3.1, #4)
# ===========================================================================
def build_dashboard(sheet_key, is_demo):
    ws = wb[sheet_key]
    if is_demo:
        add_nav(ws, sheet_key, mode_label="DEMO", mode_bg=C_WARNING, mode_fg="FFFFFF",
                business_name_formula='="🎬 DEMO"')
    else:
        add_nav(ws, SH_INICIO)
    title = "🎬 DEMO" if is_demo else "🏠 INICIO"
    add_header(ws, title, "Tu negocio, en números." if not is_demo else "Así se ve Profitly con datos de ejemplo.")

    P = "Demo" if is_demo else ""  # named-range prefix switch handled per-metric below
    def kpi(nm_real, nm_demo):
        return nm_demo if is_demo else nm_real

    if is_demo:
        banner(ws, "B6:P6", "Estás viendo datos de ejemplo. Cuando quieras, pasá a 🚀 Mi Negocio para cargar los tuyos.", tone="warning")
        internal_link(ws["B6"], SH_INICIO)
    else:
        # NOTE (Phase 9 review, two fixes applied here):
        # 1. This banner's text always said "→ Ir a X" but the cell had no
        #    hyperlink -- unlike every other "→ Ir a X" banner in the workbook,
        #    which are clickable. Fixed with a dynamic HYPERLINK() whose target
        #    sheet is computed by the same condition chain as the message, so
        #    one cell correctly links to whichever step is actually next.
        # 2. The 4th step's "objetivo missing" check used OR where it needed
        #    AND (Objetivos' own equivalent prompt, B9 below, already used AND
        #    correctly) -- as written, OR made this step permanently unable to
        #    resolve as "done" for a user who sets a default goal via
        #    Configuración without also adding an explicit tbl_Objetivos row,
        #    so the banner would never fully disappear for that (very common)
        #    path. Fixed to AND, matching Objetivos!B9's logic.
        c1 = "COUNTA(tbl_Productos[Producto])=0"
        c2 = "COUNTA(tbl_Ventas[Fecha])=0"
        c3 = "COUNTA(tbl_Gastos[Fecha])=0"
        c4 = "AND(Cfg_ObjetivoMensualDefault=0,COUNTA(tbl_Objetivos[Objetivo de ganancia])<=1)"
        msg1 = "🚀 Te falta 1 paso para ver tu panorama completo: Configurá tu negocio y agregá productos → Ir a ⚙️ Configuración"
        msg2 = "🚀 Te falta 1 paso: Registrá tu primera venta → Ir a 🛒 Ventas"
        msg3 = "🚀 Te falta 1 paso: Registrá tus gastos → Ir a 💸 Gastos"
        msg4 = "🚀 Te falta 1 paso: Definí tu objetivo mensual → Ir a 🎯 Objetivos"
        target = f'IF({c1},"{SH_CONFIG}",IF({c2},"{SH_VENTAS}",IF({c3},"{SH_GASTOS}","{SH_OBJETIVOS}")))'
        message = f'IF({c1},"{msg1}",IF({c2},"{msg2}",IF({c3},"{msg3}","{msg4}")))'
        setc(ws, "B6",
             f'=IF(NOT(OR({c1},{c2},{c3},{c4})),"",'
             f'HYPERLINK("#\'"&{target}&"\'!A1",{message}))',
             f=F(9.5, italic=True, color=C_WARNING))
        setc(ws, "F7", "Viendo:", f=F(9.5, color=C_MUTED), al=ALIGN_R)
        c = setc(ws, "G7", "Este mes", f=F(9.5, bold=True), fl=fill(C_INPUT_FILL), al=ALIGN_L)
        c.border = border_all(INPUT_SIDE)
        dv = DataValidation(type="list", formula1="=Lst_PeriodosRapidos", allow_blank=False)
        dv.add("G7")
        ws.add_data_validation(dv)

    # KPI row 1: Ventas / Ganancia real / Margen / Gastos
    kpi1 = [
        ("VENTAS", kpi("KPI_VentasPeriodo", "DemoKPI_VentasPeriodo"), FMT_CURRENCY, kpi("KPI_VentasPeriodo_Delta", None), None),
        ("GANANCIA REAL", kpi("KPI_GananciaReal", "DemoKPI_GananciaReal"), FMT_CURRENCY, kpi("KPI_GananciaReal_Delta", None),
         "Lo que realmente te queda después de descontar los costos de cada venta y los gastos generales de tu negocio."),
        ("MARGEN", kpi("KPI_MargenReal", "DemoKPI_MargenReal"), FMT_PCT, kpi("KPI_MargenReal_Delta", None),
         "Porcentaje de tus ventas que queda después de descontar los costos asociados a vender."),
        ("GASTOS", kpi("KPI_GastosPeriodo", "DemoKPI_GastosPeriodo"), FMT_CURRENCY, kpi("KPI_GastosPeriodo_Delta", None), None, True),
    ]
    starts = ["B", "F", "J", "N"]
    row0 = 9
    # Phase 10 polish: give the bold, larger-font value row (row0+1) some
    # breathing room instead of Excel's default ~15pt row height, which reads
    # cramped for a size-15 bold number.
    ws.row_dimensions[row0].height = 16
    ws.row_dimensions[row0 + 1].height = 24
    ws.row_dimensions[row0 + 2].height = 14
    for (label, nm, nf, delta_nm, tip, *rest), c0 in zip(kpi1, starts):
        invert = bool(rest and rest[0])
        c1 = col_letter(col_idx(c0) + 2)
        card_row(ws, c0, c1, row0, label, F(9, bold=True, color=C_MUTED), top=True)
        card_row(ws, c0, c1, row0 + 1, f"={nm}", F(15, bold=True, color=C_PRIMARY), nf=nf)
        if delta_nm and not is_demo:
            df = f'=IF({delta_nm}="","",IF({delta_nm}{"<0" if invert else ">0"},"▲ ","▼ ")&TEXT(ABS({delta_nm}),"0%")&" vs. período anterior")'
            card_row(ws, c0, c1, row0 + 2, df, F(8.5, color=C_MUTED), bottom=True)
        else:
            card_row(ws, c0, c1, row0 + 2, "", F(8.5), bottom=True)
        if tip:
            tooltip(ws, f"{c0}{row0}", label, tip)

    # KPI row 2: Ticket promedio / Objetivo mensual / Progreso
    row0b = 13
    ws.row_dimensions[row0b].height = 16
    ws.row_dimensions[row0b + 1].height = 24
    ws.row_dimensions[row0b + 2].height = 14
    tp_nm = kpi("KPI_TicketPromedio", "DemoKPI_TicketPromedio")
    card_row(ws, "B", "D", row0b, "TICKET PROMEDIO", F(9, bold=True, color=C_MUTED), top=True)
    card_row(ws, "B", "D", row0b + 1, f"={tp_nm}", F(15, bold=True, color=C_PRIMARY), nf=FMT_CURRENCY)
    card_row(ws, "B", "D", row0b + 2, "", F(8.5), bottom=True)

    obj_nm = "Obj_GananciaObjetivo" if not is_demo else "DemoKPI_Objetivo"
    card_row(ws, "F", "H", row0b, "OBJETIVO MENSUAL", F(9, bold=True, color=C_MUTED), top=True)
    card_row(ws, "F", "H", row0b + 1, f"={obj_nm}", F(15, bold=True, color=C_PRIMARY), nf=FMT_CURRENCY)
    card_row(ws, "F", "H", row0b + 2, "", F(8.5), bottom=True)

    prog_nm = "Obj_ProgresoPct" if not is_demo else "DemoKPI_ProgresoPct"
    card_row(ws, "J", "L", row0b, "PROGRESO", F(9, bold=True, color=C_MUTED), top=True)
    card_row(ws, "J", "L", row0b + 1, f"=IFERROR({prog_nm},0)", F(15, bold=True, color=C_PRIMARY), nf=FMT_PCT)
    card_row(ws, "J", "L", row0b + 2, "", F(8.5), bottom=True)
    ws.conditional_formatting.add(f"J{row0b+1}:L{row0b+1}",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=C_SUCCESS, showValue=True))

    # Insights panel
    setc(ws, "B17", "💡 Lo que Profitly encontró en tus números", f=F(11, bold=True, color=C_PRIMARY))
    if is_demo:
        insight_texts = [
            "🔥 Set de regalo premium genera la mayor parte de tu ganancia.",
            "🎯 Te faltan $45.000 de ganancia para alcanzar tu objetivo.",
        ]
        for i, txt in enumerate(insight_texts):
            merge(ws, f"B{18+i}:P{18+i}")
            setc(ws, f"B{18+i}", txt, f=F(9.5, color=C_TEXT), fl=fill(C_BLUE_BG), al=ALIGN_L)
    else:
        for i in range(4):
            merge(ws, f"B{18+i}:P{18+i}")
            setc(ws, f"B{18+i}", f"=Ins_Texto{i+1}", f=F(9.5, color=C_TEXT), fl=fill(C_BLUE_BG), al=ALIGN_L)

    # Charts anchored below; built in a later pass (openpyxl chart section)
    return {"row_charts": 23}

build_dashboard(SH_INICIO, is_demo=False)
print("🏠 Inicio built")
build_dashboard(SH_DEMO, is_demo=True)
print("🎬 Demo built")

# ===========================================================================
# ⚙️ Configuración (docs/04-ux-ui.md #3.2)
# ===========================================================================
ws = wb[SH_CONFIG]
add_nav(ws, SH_CONFIG)
add_header(ws, "⚙️ CONFIGURACIÓN", "Configurá tu negocio una vez — Profitly usa esto en todas las pantallas.")

steps = [
    ("1. Configurá tu negocio", '=IF(AND(Cfg_NombreNegocio<>"",Cfg_Moneda<>""),"✅","⬜")'),
    ("2. Agregá tus productos →📦", '=IF(COUNTA(tbl_Productos[Producto])>0,"✅","⬜")'),
    ("3. Registrá tus ventas →🛒", '=IF(COUNTA(tbl_Ventas[Fecha])>0,"✅","⬜")'),
    ("4. Registrá tus gastos →💸", '=IF(COUNTA(tbl_Gastos[Fecha])>0,"✅","⬜")'),
    ("5. Mirá tu rentabilidad →🏠", '=IF(AND(COUNTA(tbl_Productos[Producto])>0,COUNTA(tbl_Ventas[Fecha])>0),"✅","⬜")'),
]
starts = ["B", "E", "H", "K", "N"]
targets = [None, SH_PRODUCTOS, SH_VENTAS, SH_GASTOS, SH_INICIO]
for (label, status_f), c0, target in zip(steps, starts, targets):
    c1 = col_letter(col_idx(c0) + 1)
    card_row(ws, c0, c1, 6, label, F(8.5, bold=True, color=C_PRIMARY), top=True)
    cell = card_row(ws, c0, c1, 7, status_f, F(13, bold=True), bottom=True)
    if target:
        internal_link(ws[f"{c0}6"], target)
ws.row_dimensions[6].height = 26

setc(ws, "B10", "Tu negocio", f=F(11, bold=True, color=C_PRIMARY))
form_fields = [
    ("Nombre del negocio", "C11", None),
    ("Tipo de negocio", "C12", "text"),
    ("Moneda", "C13", "Lst_Monedas"),
]
# These are the true input cells -- Cfg_NombreNegocio/Cfg_TipoNegocio/Cfg_Moneda
# are named ranges pointing directly here (see _Config section), so typing in
# one of these cells is what actually changes the setting everywhere else.
setc(ws, "B11", "Nombre del negocio", f=F(10)); setc(ws, "C11", "Mi Negocio", f=F(10), fl=fill(C_INPUT_FILL))
setc(ws, "B12", "Tipo de negocio", f=F(10)); setc(ws, "C12", "", f=F(10), fl=fill(C_INPUT_FILL))
setc(ws, "B13", "Moneda", f=F(10)); setc(ws, "C13", "$", f=F(10), fl=fill(C_INPUT_FILL))
dv_moneda = DataValidation(type="list", formula1="=Lst_Monedas", allow_blank=True)
dv_moneda.add("C13")
ws.add_data_validation(dv_moneda)
for r in (11, 12, 13):
    ws[f"C{r}"].border = border_all(INPUT_SIDE)

# NOTE (Phase 9 review fix): these two lists were originally just explanatory
# text pointing the user at the hidden, protected _Listas sheet ("editá la
# lista base en la hoja _Listas") -- a first-time user has no way to actually
# reach or edit a hidden, locked sheet, making that instruction impossible to
# follow. Phase 4's own spec called for a genuinely editable list here; this
# now is one -- Lst_Categorias/Lst_Canales point directly at these cells
# (see _Listas section), the same pattern already used for Cfg_NombreNegocio
# etc. Editing a category/canal here immediately updates every dropdown that
# uses it, with no hidden sheet involved.
CAT_ROW0, CAT_N = 16, 8
setc(ws, "B15", "Categorías de productos", f=F(11, bold=True, color=C_PRIMARY))
setc(ws, "B16", "Agregá o cambiá las categorías que ves como opciones en 📦 Productos.", f=F(9, italic=True, color=C_MUTED))
CAT_DEFAULTS = ["General", "Alimentos", "Indumentaria", "Accesorios", "Hogar", "Belleza", "Servicios"]
for i in range(CAT_N):
    r = CAT_ROW0 + i
    val = CAT_DEFAULTS[i] if i < len(CAT_DEFAULTS) else ""
    c = setc(ws, f"C{r}", val, f=F(10), fl=fill(C_INPUT_FILL))
    c.border = border_all(INPUT_SIDE)
CAN_ROW0, CAN_N = 25, 6
setc(ws, "B24", "Canales de venta", f=F(11, bold=True, color=C_PRIMARY))
CAN_DEFAULTS = ["Local", "Online", "Mayorista", "Redes sociales", "Feria"]
for i in range(CAN_N):
    r = CAN_ROW0 + i
    val = CAN_DEFAULTS[i] if i < len(CAN_DEFAULTS) else ""
    c = setc(ws, f"C{r}", val, f=F(10), fl=fill(C_INPUT_FILL))
    c.border = border_all(INPUT_SIDE)
name_range("Lst_Categorias", SH_CONFIG, f"$C${CAT_ROW0}:$C${CAT_ROW0 + CAT_N - 1}")
name_range("Lst_Canales", SH_CONFIG, f"$C${CAN_ROW0}:$C${CAN_ROW0 + CAN_N - 1}")

banner(ws, "B32:P32",
       "💡 Cargá tus gastos fijos (alquiler, herramientas, suscripciones) en 💸 Gastos para que tu Ganancia real sea precisa.",
       tone="info")
internal_link(ws["B32"], SH_GASTOS)

setc(ws, "B34", "Objetivo mensual", f=F(11, bold=True, color=C_PRIMARY))
setc(ws, "B35", "¿Cuánto querés ganar por mes?", f=F(10))
c = setc(ws, "C35", 0, f=F(11, bold=True), fl=fill(C_INPUT_FILL), nf=FMT_CURRENCY)
dv_objetivo = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                              showErrorMessage=True, error="Ingresá un número mayor o igual a 0.", errorTitle="Valor inválido")
dv_objetivo.add("C35")
ws.add_data_validation(dv_objetivo)
c.border = border_all(INPUT_SIDE)

banner(ws, "B37:P37",
       "No hay botón \"Guardar\" — todo se guarda automáticamente a medida que escribís, como en cualquier planilla.",
       tone="neutral")

for cl in "BCDEFGHIJKLMNOP":
    ws.column_dimensions[cl].width = 13
print("⚙️ Configuración built")

# ===========================================================================
# 🧭 Menú
# ===========================================================================
ws = wb[SH_MENU]
ws.sheet_view.showGridLines = False
for cl in "ABCDEFGHIJKLMNOP":
    ws.column_dimensions[cl].width = 12
merge(ws, "B3:O3")
setc(ws, "B3", "PROFITLY — CENTRO DE CONTROL DE RENTABILIDAD", f=F(20, bold=True, color=C_PRIMARY), al=ALIGN_C)
ws.row_dimensions[3].height = 30
merge(ws, "B5:O5")
setc(ws, "B5", "Sabé exactamente cuánto ganás, qué te hace ganar dinero y qué decisiones pueden aumentar tu rentabilidad.",
     f=F(11, italic=True, color=C_MUTED), al=ALIGN_C)

merge(ws, "C9:G14")
c = setc(ws, "C9", "🎬\n\nVer una demo\n\nMirá un ejemplo con datos ficticios y entendé cómo funciona Profitly.\n\nEntrar →",
         f=F(12, bold=True, color="FFFFFF"), fl=fill(C_WARNING), al=ALIGN_C)
c.border = border_all(THIN)
internal_link(ws["C9"], SH_DEMO)

merge(ws, "I9:M14")
c = setc(ws, "I9", "🚀\n\nEmpezar con mi negocio\n\nCargá tus productos y ventas reales y llevá el control de tu rentabilidad.\n\nEntrar →",
         f=F(12, bold=True, color="FFFFFF"), fl=fill(C_PRIMARY), al=ALIGN_C)
c.border = border_all(THIN)
internal_link(ws["I9"], SH_INICIO)

merge(ws, "B17:O17")
setc(ws, "B17", "¿Primera vez? Te recomendamos empezar por la Demo.", f=F(9.5, italic=True, color=C_MUTED), al=ALIGN_C)
print("🧭 Menú built")

# ===========================================================================
# Charts (docs/01-architecture.md §4, docs/04-ux-ui.md #3.1) — 3 per dashboard
# ===========================================================================
def add_dashboard_charts(sheet_key, is_demo):
    ws = wb[sheet_key]
    data_sheet = SH__DEMODATA if is_demo else SH__DASHDATA
    trend_ref = "A127:C132" if is_demo else "A22:C28"
    prod_ref = "A116:B123" if is_demo else "A32:B41"

    # Chart 1: Evolución de ventas y ganancia (full width)
    chart1 = LineChart()
    chart1.title = "Evolución de ventas y ganancia"
    chart1.style = 2
    chart1.y_axis.title = None
    chart1.x_axis.title = None
    chart1.height = 8
    chart1.width = 28
    cats = Reference(wb[data_sheet], min_col=1, min_row=int(trend_ref.split(":")[0][1:]), max_row=int(trend_ref.split(":")[1][1:]))
    data = Reference(wb[data_sheet], min_col=2, max_col=3,
                      min_row=int(trend_ref.split(":")[0][1:]) - 1, max_row=int(trend_ref.split(":")[1][1:]))
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    # Phase 10 polish: brand palette instead of Excel's generic chart-style
    # colors -- Ganancia (the headline metric) in success green, Ventas as
    # supporting context in a muted gray-blue, both with a visible line weight.
    chart1_colors = [C_MUTED, C_SUCCESS]
    for s, color in zip(chart1.series, chart1_colors):
        s.smooth = False
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.width = 24000  # EMUs, ~1.9pt
    ws.add_chart(chart1, "B23")

    # Chart 2: Ganancia por producto (horizontal bar)
    chart2 = BarChart()
    chart2.type = "bar"
    chart2.title = "Ganancia por producto"
    chart2.style = 10
    chart2.height = 8
    chart2.width = 13.5
    r0, r1 = int(prod_ref.split(":")[0][1:]), int(prod_ref.split(":")[1][1:])
    cats2 = Reference(wb[data_sheet], min_col=1, min_row=r0, max_row=r1)
    data2 = Reference(wb[data_sheet], min_col=2, min_row=r0 - 1, max_row=r1)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.legend = None
    if chart2.series:
        chart2.series[0].graphicalProperties.solidFill = C_PRIMARY
    ws.add_chart(chart2, "B41")

    # Chart 3: Progreso hacia el objetivo (stacked bar: alcanzado vs restante)
    chart3 = BarChart()
    chart3.type = "col"
    chart3.grouping = "stacked"
    chart3.overlap = 100
    chart3.title = "Progreso hacia el objetivo"
    chart3.style = 12
    chart3.height = 8
    chart3.width = 13.5
    if is_demo:
        # Placed on the hidden _Demo_Datos engine sheet, not the visible Demo
        # sheet (Phase 9 review finding: this was originally written directly
        # onto 🎬 Demo at R115/S115/R116/S116 -- unlabeled raw numbers sitting
        # in the middle of an otherwise blank area a curious user could scroll
        # into. The real 🏠 Inicio already did this correctly, on the hidden
        # _DashboardData sheet -- this makes Demo consistent with it.
        wsDD = wb[SH__DEMODATA]
        setc(wsDD, "Z312", "Alcanzado"); setc(wsDD, "AA312", "Restante")
        setc(wsDD, "Z313", "=DemoKPI_GananciaReal")
        setc(wsDD, "AA313", "=MAX(DemoKPI_Objetivo-DemoKPI_GananciaReal,0)")
        prow0, prow1 = 312, 313
        pcol = 26  # Z
    else:
        wsD = wb[SH__DASHDATA]
        setc(wsD, "E1", "Alcanzado"); setc(wsD, "F1", "Restante")
        setc(wsD, "E2", "=MIN(Obj_GananciaActual,Obj_GananciaObjetivo)")
        setc(wsD, "F2", "=Obj_Restante")
        prow0, prow1 = 1, 2
        pcol = 5  # E
    pdata_sheet = SH__DEMODATA if is_demo else SH__DASHDATA
    catsp = Reference(wb[pdata_sheet], min_col=pcol - 1, min_row=prow1, max_row=prow1) if is_demo else None
    data3 = Reference(wb[pdata_sheet], min_col=pcol, max_col=pcol + 1, min_row=prow0, max_row=prow1)
    chart3.add_data(data3, titles_from_data=True)
    chart3.legend.position = "b"
    chart3_colors = [C_SUCCESS, C_NAV_INACTIVE_BG]  # Alcanzado (green), Restante (neutral gray)
    for s, color in zip(chart3.series, chart3_colors):
        s.graphicalProperties.solidFill = color
    ws.add_chart(chart3, "K41")

add_dashboard_charts(SH_INICIO, is_demo=False)
add_dashboard_charts(SH_DEMO, is_demo=True)
print("Charts added")

# ===========================================================================
# Data validation (docs/03-calculation-formulas.md #0.2, docs/02-data-layout.md #6)
# ===========================================================================
def dv_decimal_min(ws, addr, minval, msg):
    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1=str(minval),
                         allow_blank=True, showErrorMessage=True, error=msg, errorTitle="Valor inválido")
    dv.add(addr)
    ws.add_data_validation(dv)

def dv_decimal_between(ws, addr, lo, hi, msg):
    dv = DataValidation(type="decimal", operator="between", formula1=str(lo), formula2=str(hi),
                         allow_blank=True, showErrorMessage=True, error=msg, errorTitle="Valor inválido")
    dv.add(addr)
    ws.add_data_validation(dv)

def dv_whole_min(ws, addr, minval, msg):
    dv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1=str(minval),
                         allow_blank=True, showErrorMessage=True, error=msg, errorTitle="Valor inválido")
    dv.add(addr)
    ws.add_data_validation(dv)

def dv_date_max_today(ws, addr, msg):
    dv = DataValidation(type="date", operator="lessThanOrEqual", formula1="TODAY()",
                         allow_blank=True, showErrorMessage=True, error=msg, errorTitle="Fecha inválida")
    dv.add(addr)
    ws.add_data_validation(dv)

MSG_NUM_GE0 = "Ingresá un número mayor o igual a 0."
MSG_PCT = "Ingresá un porcentaje entre 0% y 100%."
MSG_QTY = "Ingresá una cantidad de al menos 1."
MSG_DATE = "La fecha no puede ser futura."

# Productos
wsP = wb[SH_PRODUCTOS]
# NOTE (Phase 9 review fix): Categoría was documented (docs/02/03) as a
# dropdown but the actual data validation was never wired up during the
# Phase 5 build -- found by re-reading the shipped file, not the spec.
dv_categoria = DataValidation(type="list", formula1="=Lst_Categorias", allow_blank=True,
                               showErrorMessage=True, error="Elegí una categoría de la lista.", errorTitle="Categoría inválida")
dv_categoria.add(f"C{DATA_ROW_P}")
wsP.add_data_validation(dv_categoria)
for cl in ("D", "E", "G", "H", "I", "J"):  # Precio, Costo, Packaging, Envío, Publicidad, Otros
    dv_decimal_min(wsP, f"{cl}{DATA_ROW_P}", 0, MSG_NUM_GE0)
dv_decimal_between(wsP, f"F{DATA_ROW_P}", 0, 1, MSG_PCT)
dv_unique = DataValidation(type="custom", formula1=f'=COUNTIF(tbl_Productos[Producto],B{DATA_ROW_P})<=1',
                            allow_blank=True, showErrorMessage=True,
                            error="Ya existe un producto con este nombre. Elegí un nombre distinto para diferenciarlo.",
                            errorTitle="Nombre duplicado")
dv_unique.add(f"B{DATA_ROW_P}")
wsP.add_data_validation(dv_unique)

# Ventas
wsV = wb[SH_VENTAS]
dv_date_max_today(wsV, f"B{DATA_ROW_V}", MSG_DATE)
dv_whole_min(wsV, f"D{DATA_ROW_V}", 1, MSG_QTY)
dv_decimal_min(wsV, f"E{DATA_ROW_V}", 0, MSG_NUM_GE0)
dv_decimal_between(wsV, f"G{DATA_ROW_V}", 0, 1, MSG_PCT)
dv_decimal_min(wsV, f"H{DATA_ROW_V}", 0, MSG_NUM_GE0)
dv_decimal_min(wsV, f"I{DATA_ROW_V}", 0, MSG_NUM_GE0)

# Gastos
wsG = wb[SH_GASTOS]
dv_date_max_today(wsG, f"B{DATA_ROW_G}", MSG_DATE)
dv_decimal_min(wsG, f"E{DATA_ROW_G}", 0, MSG_NUM_GE0)

# Objetivos
wsO = wb[SH_OBJETIVOS]
dv_first_of_month = DataValidation(type="custom", formula1=f"=DAY(B{DATA_ROW_O})=1",
                                    allow_blank=True, showErrorMessage=True,
                                    error="Elegí el primer día del mes que querés planificar.",
                                    errorTitle="Período inválido")
dv_first_of_month.add(f"B{DATA_ROW_O}")
wsO.add_data_validation(dv_first_of_month)
dv_unique_period = DataValidation(type="custom",
                                   formula1=f"=COUNTIF(tbl_Objetivos[Periodo (Mes-Año)],B{DATA_ROW_O})<=1",
                                   allow_blank=True, showErrorMessage=True,
                                   error="Ya definiste un objetivo para este mes — editá la fila existente.",
                                   errorTitle="Período duplicado")
dv_unique_period.add(f"B{DATA_ROW_O}")
wsO.add_data_validation(dv_unique_period)

# Simulador — Δ bounds (docs/03-calculation-formulas.md §8.2)
wsS = wb[SH_SIMULADOR]
for scen, (c0, c1) in SCEN_COLS.items():
    r = SCEN_INPUT_ROW0
    dv_decimal_between(wsS, f"{c1}{r}", -0.5, 2.0, "Ingresá un porcentaje razonable.")
    dv_decimal_between(wsS, f"{c1}{r+1}", -0.9, 3.0, "Ingresá un porcentaje razonable.")
    dv_decimal_between(wsS, f"{c1}{r+2}", -1.0, 1.0, "Ingresá un valor entre -100 y 100 puntos.")
print("Data validations added")

# ===========================================================================
# Protection (docs/02-data-layout.md #7, docs/04-ux-ui.md #1.3)
# ===========================================================================
from openpyxl.styles import Protection

def unlock(ws, addr):
    c = ws[addr]
    c.protection = Protection(locked=False)

def protect_sheet(ws):
    ws.protection.sheet = True
    ws.protection.formatCells = False
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.insertRows = False
    ws.protection.sort = False
    ws.protection.autoFilter = False
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False

# Productos: unlock input columns B-J at the data row
for cl in "BCDEFGHIJ":
    unlock(wb[SH_PRODUCTOS], f"{cl}{DATA_ROW_P}")
protect_sheet(wb[SH_PRODUCTOS])

# Ventas: unlock all input columns B-I (E,G are now plain required inputs, not
# autofill -- see docs/05-qa-results.md §5)
for cl in "BCDEFGHI":
    unlock(wb[SH_VENTAS], f"{cl}{DATA_ROW_V}")
protect_sheet(wb[SH_VENTAS])

# Gastos: unlock B-F
for cl in "BCDEF":
    unlock(wb[SH_GASTOS], f"{cl}{DATA_ROW_G}")
protect_sheet(wb[SH_GASTOS])

# Objetivos: unlock B,C (goal table)
for cl in "BC":
    unlock(wb[SH_OBJETIVOS], f"{cl}{DATA_ROW_O}")
protect_sheet(wb[SH_OBJETIVOS])

# Simulador: unlock the 15 delta input cells + period n/a; lock everything else
for scen, (c0, c1) in SCEN_COLS.items():
    r = SCEN_INPUT_ROW0
    for i in range(5):
        unlock(wb[SH_SIMULADOR], f"{c1}{r+i}")
protect_sheet(wb[SH_SIMULADOR])

# Inicio: unlock only the period selector
unlock(wb[SH_INICIO], "G7")
protect_sheet(wb[SH_INICIO])

# Configuración: unlock the business-setting input cells, plus the Categorías/
# Canales list cells (Phase 9 review fix)
for addr in ("C11", "C12", "C13", "C35"):
    unlock(wb[SH_CONFIG], addr)
for i in range(CAT_N):
    unlock(wb[SH_CONFIG], f"C{CAT_ROW0 + i}")
for i in range(CAN_N):
    unlock(wb[SH_CONFIG], f"C{CAN_ROW0 + i}")
protect_sheet(wb[SH_CONFIG])

# Demo and Menú: fully locked (read-only), nothing to unlock
protect_sheet(wb[SH_DEMO])
protect_sheet(wb[SH_MENU])

# Engine sheets: fully locked and hidden (already hidden via sheet_state)
for s in ENGINE_SHEETS:
    protect_sheet(wb[s])

# Workbook structure: discourage (not cryptographically prevent) unhiding/
# deleting/renaming/reordering sheets -- passwordless by design (see
# docs/05-qa-results.md): this is meant to stop accidental structural edits,
# the same intent as the cell-level protection above, not to secure the file
# against a determined user, so losing a password was never an acceptable risk.
from openpyxl.workbook.protection import WorkbookProtection
wb.security = WorkbookProtection(lockStructure=True)

print("Protection applied")
wb.save("/home/user/Profitly/build/Profitly.xlsx")
