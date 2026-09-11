#!/usr/bin/env python3
"""QA harness: injects edge-case data into copies of the pristine workbook,
recalculates via LibreOffice, and reports actual computed values against
the 25-case matrix in docs/03-calculation-formulas.md #10."""
import openpyxl, shutil, subprocess, sys, json, datetime

SRC = "/home/user/Profitly/build/Profitly.xlsx"
RECALC = "/root/.claude/skills/synced/939c706b-0650-4382-8400-7cfa64db24ac_f1fe1cd6-63d3-4d8c-aefc-d1372735b983/xlsx/scripts/recalc.py"

def add_product_row(ws, row, name, cat, precio, costo, comision, packaging, envio, publicidad, otros):
    for col, val in zip("BCDEFGHIJ", [name, cat, precio, costo, comision, packaging, envio, publicidad, otros]):
        ws[f"{col}{row}"] = val
    ws[f"K{row}"] = f"=E{row}+(D{row}*F{row})+G{row}+H{row}+I{row}+J{row}"
    ws[f"L{row}"] = f"=D{row}-K{row}"
    ws[f"M{row}"] = f'=IF(D{row}=0,Txt_SinDatos,L{row}/D{row})'
    ws[f"N{row}"] = f"=SUMIFS(tbl_Ventas[Cantidad],tbl_Ventas[Producto],B{row})"
    ws[f"O{row}"] = f"=SUMIFS(tbl_Ventas[Ganancia],tbl_Ventas[Producto],B{row})"
    ws[f"P{row}"] = f'=IF(AND(ISNUMBER(M{row}),M{row}<Cfg_MargenAlertaUmbral,N{row}>0),1,0)'

def extend_table(ws, tablename, new_ref):
    ws.tables[tablename].ref = new_ref

def add_sale_row(ws, row, fecha, producto, cantidad, canal="Local", descuento=0, otros_costos=0):
    ws[f"B{row}"] = fecha
    ws[f"C{row}"] = producto
    ws[f"D{row}"] = cantidad
    ws[f"E{row}"] = f'=IFERROR(INDEX(tbl_Productos[Precio de venta],MATCH(C{row},tbl_Productos[Producto],0)),"")'
    ws[f"F{row}"] = canal
    ws[f"G{row}"] = f'=IFERROR(INDEX(tbl_Productos[Comisión %],MATCH(C{row},tbl_Productos[Producto],0)),"")'
    ws[f"H{row}"] = descuento
    ws[f"I{row}"] = otros_costos
    ws[f"J{row}"] = f'=IF(C{row}="","",(D{row}*E{row})-H{row})'
    ws[f"K{row}"] = f'=IF(C{row}="","",IFERROR(INDEX(tbl_Productos[Costo real de venta],MATCH(C{row},tbl_Productos[Producto],0))*D{row},Txt_ProductoNoEncontrado))'
    ws[f"L{row}"] = f'=IF(C{row}="","",IFERROR(J{row}-K{row}-I{row},Txt_ProductoNoEncontrado))'

def add_gasto_row(ws, row, fecha, categoria, desc, importe, tipo="Variable"):
    for col, val in zip("BCDEF", [fecha, categoria, desc, importe, tipo]):
        ws[f"{col}{row}"] = val

def recalc(path, timeout=90):
    r = subprocess.run(["python3", RECALC, path, str(timeout)], capture_output=True, text=True)
    try:
        return json.loads(r.stdout)
    except Exception:
        return {"error": r.stdout + r.stderr}

def read(path, addrs, sheet):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    return {a: ws[a].value for a in addrs}

results = []

def run_test(name, mutate_fn, checks_sheet_addrs, description):
    path = f"/tmp/qa_{name}.xlsx"
    shutil.copy(SRC, path)
    wb = openpyxl.load_workbook(path)
    mutate_fn(wb)
    wb.save(path)
    rc = recalc(path, 100)
    status = rc.get("status", "ERROR")
    errors = rc.get("total_errors", "?")
    values = {}
    if status == "success" or status == "errors_found":
        for sheet, addrs in checks_sheet_addrs.items():
            values.update({f"{sheet}!{a}": v for a, v in read(path, addrs, sheet).items()})
    results.append({"test": name, "description": description, "recalc_status": status,
                     "formula_errors": errors, "values": values})
    print(f"=== {name} === status={status} errors={errors}")
    for k, v in values.items():
        print(f"  {k} = {v!r}")

TODAY = datetime.date.today()

# TEST A: fully empty (0 products, 0 sales, 0 gastos, no goal) — cases 1, 4, 8, 14, 15
def mut_empty(wb):
    pass  # pristine file already ships with empty starter rows
run_test("A_empty", mut_empty,
    {"📦 Productos": ["B8"], "🏠 Inicio": ["B9", "B6"], "🎯 Objetivos": ["H12"]},
    "Cases 1,4,8,14,15: no products/sales/gastos/goal")

# TEST B: one product, one sale, zero extra costs — cases 2, 5, 7
def mut_one(wb):
    wsP = wb["📦 Productos"]
    add_product_row(wsP, 13, "Producto Unico", "General", 1000, 0, 0, 0, 0, 0, 0)
    extend_table(wsP, "tbl_Productos", "B12:P13")
    wsV = wb["🛒 Ventas"]
    add_sale_row(wsV, 10, TODAY, "Producto Unico", 1)
    extend_table(wsV, "tbl_Ventas", "B9:L10")
run_test("B_one_product_one_sale", mut_one,
    {"📦 Productos": ["K13", "L13", "M13"], "🏠 Inicio": ["B10", "F10", "J10"]},
    "Cases 2,5,7: one product, one sale, zero cost => margin 100%")

# TEST C: two products (one w/ sales, one without), discount, negative margin, low margin — cases 3,9,10,16,17,18
def mut_multi(wb):
    wsP = wb["📦 Productos"]
    wsP.unmerge_cells("B15:P15")  # the "sin productos" empty-state banner sits here; test writes over it
    add_product_row(wsP, 13, "Producto A (vendido)", "General", 1000, 400, 0.05, 0, 0, 0, 0)   # healthy margin
    add_product_row(wsP, 14, "Producto B (sin ventas)", "General", 500, 100, 0, 0, 0, 0, 0)      # never sold
    add_product_row(wsP, 15, "Producto C (margen bajo)", "General", 1000, 870, 0, 0, 0, 0, 0)    # ~13% margin, below 15% threshold
    add_product_row(wsP, 16, "Producto D (margen negativo)", "General", 500, 700, 0, 0, 0, 0, 0)  # priced below cost
    extend_table(wsP, "tbl_Productos", "B12:P16")
    wsV = wb["🛒 Ventas"]
    wsV.unmerge_cells("B12:P12")  # the "sin ventas" empty-state banner sits here; test writes over it
    add_sale_row(wsV, 10, TODAY, "Producto A (vendido)", 3, descuento=200)  # discount case
    add_sale_row(wsV, 11, TODAY, "Producto C (margen bajo)", 1)
    add_sale_row(wsV, 12, TODAY, "Producto D (margen negativo)", 1)
    extend_table(wsV, "tbl_Ventas", "B9:L12")
run_test("C_multi_margin_cases", mut_multi,
    {"📦 Productos": ["M13", "M14", "N14", "M15", "P15", "M16", "P16"],
     "🛒 Ventas": ["J10", "L10", "L12"]},
    "Cases 3,9,10,16,17,18: multiple products, discount, negative/low margin, product with no sales")

# TEST D: impossible goal (objetivo = 0) and unreachable goal — case 14
def mut_goal(wb):
    wsO = wb["🎯 Objetivos"]
    wsO["C7"] = 0
run_test("D_impossible_goal", mut_goal,
    {"🎯 Objetivos": ["H9", "H12", "F14"]},
    "Case 14: objetivo = 0 => Txt_ObjetivoInvalido, not #DIV/0!")

# TEST E: price change after a sale — case 19 (historical sale price must not retroactively change)
def mut_price_change(wb):
    wsP = wb["📦 Productos"]
    add_product_row(wsP, 13, "Producto Precio", "General", 1000, 400, 0, 0, 0, 0, 0)
    extend_table(wsP, "tbl_Productos", "B12:P13")
    wsV = wb["🛒 Ventas"]
    add_sale_row(wsV, 10, TODAY, "Producto Precio", 1)
    extend_table(wsV, "tbl_Ventas", "B9:L10")
    # Now change the product's price -- the sale's own Precio de venta cell
    # is a static value once written (E10 is a formula only until overwritten;
    # here it was never manually overridden, so per Ventas' own design (Phase 3
    # §3.2) it re-evaluates -- this test documents that behavior explicitly.
    wsP["D13"] = 2000
run_test("E_price_change_after_sale", mut_price_change,
    {"🛒 Ventas": ["E10", "J10"], "📦 Productos": ["D13"]},
    "Case 19: verifies whether a post-sale price edit retroactively changes historical sale value")

# TEST F: large numbers and decimals — cases 23, 24
def mut_large(wb):
    wsP = wb["📦 Productos"]
    add_product_row(wsP, 13, "Producto Grande", "General", 15000000.50, 6234567.25, 0.0825, 1234.10, 0, 0, 0)
    extend_table(wsP, "tbl_Productos", "B12:P13")
    wsV = wb["🛒 Ventas"]
    add_sale_row(wsV, 10, TODAY, "Producto Grande", 9999)
    extend_table(wsV, "tbl_Ventas", "B9:L10")
run_test("F_large_decimal_numbers", mut_large,
    {"📦 Productos": ["K13", "L13", "M13"], "🛒 Ventas": ["J10", "L10"]},
    "Cases 23,24: very large quantity and decimal currency values")

# TEST G: Simulador with insufficient data vs with data — case 21
def mut_sim_empty(wb):
    pass
run_test("G1_simulator_no_data", mut_sim_empty,
    {"🔬 Simulador": ["B6", "D29"]},
    "Case 21: simulator with zero real sales must show guidance, not a projection")

def mut_sim_data(wb):
    wsP = wb["📦 Productos"]
    add_product_row(wsP, 13, "Producto Sim", "General", 1000, 400, 0, 0, 0, 0, 0)
    extend_table(wsP, "tbl_Productos", "B12:P13")
    wsV = wb["🛒 Ventas"]
    add_sale_row(wsV, 10, TODAY, "Producto Sim", 10)
    extend_table(wsV, "tbl_Ventas", "B9:L10")
    wsS = wb["🔬 Simulador"]
    wsS["D16"] = 0.10  # Δ Precio +10%
run_test("G2_simulator_with_data", mut_sim_data,
    {"🔬 Simulador": ["D8", "D23", "D30"]},
    "Case 21: simulator projects correctly off real baseline with a +10% price delta")

with open("/tmp/qa_results.json", "w") as f:
    json.dump(results, f, indent=2, ensure_ascii=False, default=str)
print("\nAll QA tests complete. Results saved to /tmp/qa_results.json")
