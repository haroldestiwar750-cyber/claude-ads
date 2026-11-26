"""Compara los resultados de las fórmulas (recalculadas en LibreOffice) contra la simulación Python.
Uso: python verificar_formulas.py ../SistemaEmpresarial_base.xlsx"""
import sys, json
sys.path.insert(0, "../build")
import demo as DM
from lo_util import arrancar, abrir, val
s = DM.generar()
proc, ctx, desk = arrancar()
fallos, ok = [], 0
try:
    d = abrir(desk, sys.argv[1]); d.calculateAll()
    inv = d.Sheets.getByName("INVENTARIO")
    for i, p in enumerate(DM.PRODUCTOS):
        r = 9 + i
        st = inv.getCellRangeByName(f"O{r}").Value
        exp = s.stock(p[0])
        if abs(st - exp) > 1e-6: fallos.append(("stock", p[0], st, exp))
        else: ok += 1
        cr = inv.getCellRangeByName(f"J{r}").Value
        if abs(cr - s.costo_real(p[0])) > 0.6: fallos.append(("costoreal", p[0], cr, s.costo_real(p[0])))
        else: ok += 1
    ventas = sum(r[14] for r in s.bd if r[20] == "ACTIVA")
    uti = d.Sheets.getByName("UTILIDADES")
    v = uti.getCellRangeByName("C19").Value
    (fallos if abs(v - ventas) > 1 else [None]).append(("ventas año", v, ventas)) if abs(v-ventas)>1 else None
    ok += abs(v - ventas) <= 1
    cxc = sum(c[6] - c[7] for c in s.car if c[1] == "POR COBRAR" and c[9] == "ACTIVA") - sum(i[10] for i in s.ing if i[3] == "ABONO" and i[13] == "ACTIVO")
    car = d.Sheets.getByName("CARTERA").getCellRangeByName("K5").Value
    print("CxC formula", car, "sim", cxc)
    ok += abs(car - cxc) < 1
    rep = {}
    for sh, a in [("INICIO","C7"),("INICIO","F7"),("INICIO","I7"),("INICIO","L7"),("INICIO","O7"),
                  ("INICIO","C11"),("INICIO","F11"),("INICIO","I11"),("INICIO","L11"),("INICIO","O11"),
                  ("INICIO","C15"),("INICIO","F15"),("INICIO","I15"),("INICIO","L15"),("INICIO","O15"),
                  ("CAJA Y BANCOS","H11"),("CLIENTES","J8"),("CLIENTES","K8"),("CLIENTES","N8"),("CLIENTES","Q8"),
                  ("CALC","B71"),("CALC","C71"),("CALC","B78"),("CALC","E31"),("CALC","B85"),
                  ("CARTERA","R15"),("CARTERA","P15"),("COSTOS Y RENTABILIDAD","R35"),("COSTOS Y RENTABILIDAD","T35"),
                  ("GASTOS FIJOS","J7"),("GASTOS FIJOS","K7"),("INGRESOS","K14"),("EGRESOS","K22"),("UTILIDADES","H19")]:
        rep[f"{sh}!{a}"] = val(d, sh, a)
    print(json.dumps(rep, ensure_ascii=False, indent=0))
    d.close(True)
finally:
    desk.terminate(); proc.wait(timeout=30)
print(json.dumps({"verificaciones_ok": ok, "fallos": fallos}, ensure_ascii=False))
