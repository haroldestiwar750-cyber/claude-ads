"""
Construye el libro base del Sistema Empresarial (formato .xlsx).

Genera las 22 hojas con formato profesional, fórmulas conectadas,
listas desplegables, validaciones, formato condicional, gráficos,
áreas de impresión, protección y datos de demostración coherentes.
El instalador de Windows (instalador/instalar.ps1) inyecta luego las
macros VBA, crea los botones y guarda el archivo final .xlsm.

Uso CLI:
    python construir_libro.py --out ../SistemaEmpresarial_base.xlsx [--sin-demo]
"""
import argparse
import datetime as dt
import json
import os

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

import demo as DM
from estructura import (SH, ORDEN, TABLAS, col_index, FAC, FAC_L1, FAC_LN, GUI,
                        GUI_L1, GUI_LN, COM, COM_L1, COM_LN, ING, EGR, CAR, CAJ,
                        REP, INI, CFG_ROWS, CFG_FIRST_ROW, CFG_LISTS,
                        CFG_LIST_HDR, CFG_LIST_FIRST, CFG_LIST_N, cfg_cell)

FONT = "Segoe UI"
NAVY, BLUE, TEAL = "#1E293B", "#1D4ED8", "#0F766E"
LIGHT, LINE, INPUT, INLINE = "#F1F5F9", "#CBD5E1", "#FFF8E1", "#E5C77F"
RED, AMBER, GREEN, GREY = "#DC2626", "#D97706", "#15803D", "#64748B"
NUMF = {"cop": '"$"#,##0;[Red]-"$"#,##0', "int": "#,##0;[Red]-#,##0",
        "pct": "0.0%", "date": "dd/mm/yyyy", "dt": "dd/mm/yyyy hh:mm",
        "num": "#,##0.00", "txt": None}
CLAVE_HOJAS = "1234"


def L(i):
    """Letra de columna desde índice 1-based."""
    return xl_col_to_name(i - 1)


def q(sheet_key):
    return "'" + SH[sheet_key] + "'"


def rng(tk, key):
    """Rango acotado de una columna de tabla, p.ej. 'BASE DE DATOS'!$O$5:$O$100004."""
    t = TABLAS[tk]
    c = L(col_index(tk, key))
    return f"{q(t['sheet'])}!${c}${t['first']}:${c}${t['last']}"


def cell(tk, key, row):
    return f"{L(col_index(tk, key))}{row}"


class Fmts:
    def __init__(self, wb):
        self.wb, self.cache = wb, {}

    def __call__(self, **kw):
        base = {"font_name": FONT, "font_size": 10, "valign": "vcenter"}
        base.update({k: v for k, v in kw.items() if v is not None})
        nf = base.pop("nf", None)
        if nf and NUMF.get(nf):
            base["num_format"] = NUMF[nf]
        key = tuple(sorted((k, str(v)) for k, v in base.items()))
        if key not in self.cache:
            self.cache[key] = self.wb.add_format(base)
        return self.cache[key]

    # atajos de estilos
    def inp(self, nf=None, **kw):
        d = dict(bg_color=INPUT, border=1, border_color=INLINE, locked=False, nf=nf)
        d.update(kw)
        return self(**d)

    def out(self, nf=None, **kw):
        d = dict(border=1, border_color=LINE, nf=nf)
        d.update(kw)
        return self(**d)

    def hdr(self, **kw):
        d = dict(bold=True, font_color="white", bg_color=NAVY, border=1,
                 border_color="#334155", text_wrap=True, align="center")
        d.update(kw)
        return self(**d)

    def lbl(self, **kw):
        d = dict(bold=True, font_color="#334155")
        d.update(kw)
        return self(**d)


def todt(v):
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return dt.datetime(v.year, v.month, v.day)
    return v


class Builder:
    def __init__(self, out, con_demo=True):
        self.out = out
        self.wb = xlsxwriter.Workbook(out)
        self.wb.set_properties({"title": "Sistema Integral de Gestión Empresarial",
                                "subject": "ERP en Excel - ropa, calzado y accesorios",
                                "author": "Sistema Empresarial", "company": "",
                                "comments": "Generado por construir_libro.py"})
        self.wb.set_calc_mode("auto")
        self.F = Fmts(self.wb)
        self.ws = {}
        for k in ORDEN:
            self.ws[k] = self.wb.add_worksheet(SH[k])
        self.sim = DM.generar() if con_demo else None
        self.demo = con_demo
        self.logo = os.path.join(os.path.dirname(os.path.abspath(out)), "logo_demo.png")

    # ------------------------------------------------------------ comunes
    def banner(self, k, title, subtitle, ncols=16, link_col=0):
        ws, F = self.ws[k], self.F
        ws.set_row(0, 34, F(bg_color=NAVY))
        ws.set_row(1, 20, F(bg_color=LIGHT))
        ws.write(0, 0, title, F(bold=True, font_size=16, font_color="white", bg_color=NAVY))
        ws.write(1, 2 if link_col == 0 else 0, subtitle,
                 F(italic=True, font_color=GREY, bg_color=LIGHT))
        if link_col is not None and k != "INI":
            ws.write_url(1 if link_col == 0 else 0, link_col, f"internal:'{SH['INI']}'!A1",
                         F(bold=True, font_color="white", bg_color=BLUE, align="center",
                           border=1, border_color=BLUE), string="⌂ INICIO")
        ws.hide_gridlines(2)
        ws.set_zoom(95)

    def tabla_hdr(self, k, tk, freeze=True, filtro=True):
        ws, F, t = self.ws[k], self.F, TABLAS[tk]
        r = t["hdr"] - 1
        ws.set_row(r, 32)
        for i, c in enumerate(t["cols"]):
            ci = t["col0"] - 1 + i
            ws.write(r, ci, c[1], F.hdr(bg_color=(TEAL if c[4] == "f" else (NAVY if c[4] == "v" else BLUE))))
            colfmt = F(nf=c[3]) if (c[4] == "v" and NUMF.get(c[3])) else None
            ws.set_column(ci, ci, c[2], colfmt, {"hidden": c[1].endswith("(técnico)")})
        if filtro:
            ws.autofilter(r, t["col0"] - 1, t["hdr"] + 2000 if t["last"] > 5000 else t["last"] - 1,
                          t["col0"] - 2 + len(t["cols"]))
        if freeze:
            ws.freeze_panes(t["hdr"], 0)

    def tabla_valores(self, k, tk, rows):
        ws, F, t = self.ws[k], self.F, TABLAS[tk]
        for ri, row in enumerate(rows):
            r = t["first"] - 1 + ri
            for i, v in enumerate(row):
                c = t["cols"][i]
                fm = F.out(c[3], border_color="#E2E8F0")
                ci = t["col0"] - 1 + i
                if v is None or v == "":
                    continue
                if c[3] in ("date", "dt"):
                    ws.write_datetime(r, ci, todt(v), fm)
                elif isinstance(v, str):
                    ws.write_string(r, ci, v, fm)
                else:
                    ws.write_number(r, ci, v, fm)

    def tabla_maestra(self, k, tk, rows, formulas, extra_fmt=None):
        """Escribe filas de un maestro: capturas desbloqueadas + fórmulas en toda la capacidad."""
        ws, F, t = self.ws[k], self.F, TABLAS[tk]
        n = t["last"] - t["first"] + 1
        fcache = {}
        for ri in range(n):
            r1 = t["first"] + ri  # 1-based
            data = rows[ri] if ri < len(rows) else None
            for i, c in enumerate(t["cols"]):
                ci = t["col0"] - 1 + i
                key, kind, nf = c[0], c[4], c[3]
                if kind == "f":
                    if nf not in fcache:
                        fcache[nf] = F.out(nf, bg_color="#F8FAFC", border_color="#E2E8F0")
                    ws.write_formula(r1 - 1, ci, formulas[key](r1), fcache[nf])
                else:
                    fm = F.inp(nf) if kind == "i" else F.out(nf, border_color="#E2E8F0")
                    v = data[i] if data and i < len(data) else None
                    if v is None or v == "":
                        ws.write_blank(r1 - 1, ci, None, fm)
                    elif nf in ("date", "dt"):
                        ws.write_datetime(r1 - 1, ci, todt(v), fm)
                    elif isinstance(v, str):
                        ws.write_string(r1 - 1, ci, v, fm)
                    else:
                        ws.write_number(r1 - 1, ci, v, fm)

    def proteger(self, k, extra=None):
        opts = {"autofilter": True, "format_columns": True, "format_rows": True,
                "select_locked_cells": True, "select_unlocked_cells": True,
                "objects": False}
        if extra:
            opts.update(extra)
        self.ws[k].protect(CLAVE_HOJAS, opts)

    def dv_list(self, k, rango, source, kind="stop", msg=None):
        d = {"validate": "list", "source": source, "error_type": kind,
             "ignore_blank": True}
        if msg:
            d["error_message"] = msg
            d["error_title"] = "Valor no válido"
        self.ws[k].data_validation(rango, d)

    # ------------------------------------------------------------ nombres
    def nombres(self):
        wb = self.wb
        for i, r in enumerate(CFG_ROWS):
            wb.define_name(r[0], f"={q('CFG')}!$B${CFG_FIRST_ROW + i}")
        last = CFG_LIST_FIRST + CFG_LIST_N - 1
        for name, (c, _t, _v) in CFG_LISTS.items():
            wb.define_name(name, f"=OFFSET({q('CFG')}!${c}${CFG_LIST_FIRST},0,0,"
                                 f"MAX(1,COUNTA({q('CFG')}!${c}${CFG_LIST_FIRST}:${c}${last})),1)")
        ti = TABLAS["INV"]
        wb.define_name("lstSKU", f"=OFFSET({q('INV')}!$A${ti['first']},0,0,MAX(1,COUNTA({q('INV')}!$A${ti['first']}:$A${ti['last']})),1)")
        tc = TABLAS["CLI"]
        cq = L(col_index("CLI", "COMBO"))
        wb.define_name("lstClientes", f"=OFFSET({q('CLI')}!${cq}${tc['first']},0,0,MAX(1,COUNTA({q('CLI')}!$A${tc['first']}:$A${tc['last']})),1)")
        tp = TABLAS["PRO"]
        wb.define_name("lstProveedores", f"=OFFSET({q('PRO')}!$B${tp['first']},0,0,MAX(1,COUNTA({q('PRO')}!$B${tp['first']}:$B${tp['last']})),1)")
        tr = TABLAS["CAR"]
        li, ix = L(col_index("CAR", "LISTA")), L(col_index("CAR", "IDX"))
        wb.define_name("lstPendientes", f"=OFFSET({q('CAR')}!${li}${tr['first']},0,0,MAX(1,COUNT({q('CAR')}!${ix}${tr['first']}:${ix}${tr['last']})),1)")
        wb.define_name("lstCatFiltro", f"=OFFSET({q('CALC')}!$Y$1,0,0,1+COUNTA({q('CFG')}!$D${CFG_LIST_FIRST}:$D${last}),1)")
        wb.define_name("lstTiposReporte", "=lstReportes")

    # ------------------------------------------------------------ CONFIGURACIÓN
    def hoja_cfg(self):
        k, ws, F = "CFG", self.ws["CFG"], self.F
        self.banner(k, "CONFIGURACIÓN EMPRESARIAL",
                    "Datos de la empresa, numeración, márgenes y listas del sistema. Celdas amarillas = editables.")
        ws.set_column("A:A", 44)
        ws.set_column("B:B", 38)
        ws.set_column("C:C", 3)
        ws.set_column("D:T", 18)
        ws.write("A3", "PARÁMETRO", F.hdr(align="left"))
        ws.write("B3", "VALOR", F.hdr())
        sig = {"cfgSigFactura": 1, "cfgSigGuia": 1, "cfgSigCompra": 1}
        if self.sim:
            sig = {"cfgSigFactura": self.sim.sig["FV"], "cfgSigGuia": self.sim.sig["GE"],
                   "cfgSigCompra": self.sim.sig["OC"]}
        for i, (name, label, val, nf) in enumerate(CFG_ROWS):
            r = CFG_FIRST_ROW - 1 + i
            v = sig.get(name, val)
            ws.write(r, 0, label, F.lbl(border=1, border_color=LINE, bg_color=LIGHT))
            if nf == "secret":
                ws.write(r, 1, v, F(bg_color="#E2E8F0", font_color="#E2E8F0", border=1, border_color=LINE,
                                    locked=True, hidden=True))
            elif name.startswith("cfgSig"):
                ws.write(r, 1, v, F.out("int", bold=True, bg_color="#E2E8F0"))
            else:
                ws.write(r, 1, v, F.inp(nf if nf != "txt" else None, text_wrap=(name == "cfgPie")))
        nr = CFG_FIRST_ROW + len(CFG_ROWS)
        notas = ["• Los consecutivos (gris) solo se cambian con el botón CAMBIAR NUMERACIÓN (requiere clave).",
                 "• Las claves (gris) se cambian con el botón CAMBIAR CLAVES. Clave inicial de autorización: admin123 · hojas: 1234.",
                 "• Margen s/ventas: (Precio − Costo) / Precio.   Recargo s/costo: (Precio − Costo) / Costo.",
                 "• Las listas de la derecha alimentan todos los desplegables. Agregue elementos al final de cada columna.",
                 "• 'Cuenta asociada al método' define a qué cuenta de CAJA Y BANCOS entra o sale el dinero de cada método."]
        for i, n in enumerate(notas):
            ws.write(nr + 1 + i, 0, n, F(italic=True, font_color=GREY))
        for name, (c, title, vals) in CFG_LISTS.items():
            ci = ord(c) - 65
            ws.write(CFG_LIST_HDR - 1, ci, title, F.hdr())
            lock = name in ("lstCuentas", "lstMeses", "lstSiNo", "lstEstadosPago",
                            "lstReportes", "lstTipoGasto")
            for j in range(CFG_LIST_N):
                v = vals[j] if j < len(vals) else None
                fm = F.out(bg_color="#E2E8F0") if lock else F.inp()
                if name == "lstCuentas" and j < 5:
                    fm = F.inp(bold=True)
                if lock and name == "lstCuentas" and j >= 5:
                    fm = F.out(bg_color="#E2E8F0")
                if v is None:
                    ws.write_blank(CFG_LIST_FIRST - 1 + j, ci, None, fm)
                else:
                    ws.write_string(CFG_LIST_FIRST - 1 + j, ci, v, fm)
        ws.write("D2", "Listas del sistema (amarillo = editable; gris = fijo del sistema)",
                 F(italic=True, font_color=GREY, bg_color=LIGHT))
        ws.freeze_panes(3, 0)
        self.proteger(k)

    # ------------------------------------------------------------ INVENTARIO
    def hoja_inv(self):
        k, ws, F = "INV", self.ws["INV"], self.F
        self.banner(k, "INVENTARIO", "Existencias en tiempo real: stock = inicial + entradas − salidas (calculado desde MOVIMIENTOS INV). Nunca edite Entradas/Salidas a mano.")
        self.tabla_hdr(k, "INV")
        f = TABLAS["INV"]["first"]
        C = lambda key, r: cell("INV", key, r)
        cmp = lambda: f'{rng("COM","ESTADO")},"ACTIVA"'
        fx = {
            "TRANSP": lambda r: f'=IF({C("SKU",r)}="","",IFERROR((SUMIFS({rng("COM","TRANSPORTE")},{rng("COM","SKU")},{C("SKU",r)},{cmp()})+SUMIFS({rng("COM","OTROS")},{rng("COM","SKU")},{C("SKU",r)},{cmp()}))/SUMIFS({rng("COM","CANT")},{rng("COM","SKU")},{C("SKU",r)},{cmp()}),0))',
            "COSTOREAL": lambda r: f'=IF({C("SKU",r)}="","",N({C("COSTO",r)})+N({C("TRANSP",r)}))',
            "ENTRADAS": lambda r: f'=IF({C("SKU",r)}="","",SUMIFS({rng("MOV","ENTRADA")},{rng("MOV","SKU")},{C("SKU",r)}))',
            "SALIDAS": lambda r: f'=IF({C("SKU",r)}="","",SUMIFS({rng("MOV","SALIDA")},{rng("MOV","SKU")},{C("SKU",r)}))',
            "STOCK": lambda r: f'=IF({C("SKU",r)}="","",N({C("STOCKINI",r)})+{C("ENTRADAS",r)}-{C("SALIDAS",r)})',
            "VALOR": lambda r: f'=IF({C("SKU",r)}="","",{C("STOCK",r)}*{C("COSTOREAL",r)})',
            "VENDIDAS": lambda r: f'=IF({C("SKU",r)}="","",SUMIFS({rng("BD","CANT")},{rng("BD","SKU")},{C("SKU",r)},{rng("BD","ESTADO")},"ACTIVA"))',
            "ROTACION": lambda r: f'=IF({C("SKU",r)}="","",IF({C("VENDIDAS",r)}<=0,"SIN VENTAS",IF({C("VENDIDAS",r)}/MAX(1,N({C("STOCKINI",r)})+{C("ENTRADAS",r)})>=0.5,"ALTA",IF({C("VENDIDAS",r)}/MAX(1,N({C("STOCKINI",r)})+{C("ENTRADAS",r)})>=0.2,"MEDIA","BAJA"))))',
            "ESTADO": lambda r: f'=IF({C("SKU",r)}="","",IF({C("ACTIVO",r)}="NO","INACTIVO",IF({C("STOCK",r)}<=0,"AGOTADO",IF({C("STOCK",r)}<=N({C("STOCKMIN",r)}),"BAJO STOCK","DISPONIBLE"))))',
        }
        rows = []
        if self.sim:
            for p in DM.PRODUCTOS:
                i = self.sim.inv[p[0]]
                rows.append([p[0], p[1], p[2], p[3], p[4], p[5], DM.P[p[6]],
                             round(i["costo"], 2), None, None, p[8], i["ini"], None, None,
                             None, p[10], None, None, None, None, "SI",
                             None, None, None])
        self.tabla_maestra(k, "INV", rows, fx)
        t = TABLAS["INV"]
        a, b = t["first"], t["last"]
        ws.data_validation(f"A{a}:A{b}", {"validate": "custom", "value": f"=COUNTIF($A${a}:$A${b},A{a})=1",
                                          "error_title": "SKU duplicado",
                                          "error_message": "Ese código SKU ya existe. Cada producto debe tener un SKU único."})
        self.dv_list(k, f"C{a}:C{b}", "=lstCategorias", msg="Seleccione una categoría de la lista (CONFIGURACIÓN).")
        self.dv_list(k, f"E{a}:E{b}", "=lstTallas", "information")
        self.dv_list(k, f"F{a}:F{b}", "=lstColores", "information")
        self.dv_list(k, f"G{a}:G{b}", "=lstProveedores", "warning")
        self.dv_list(k, f"U{a}:U{b}", "=lstSiNo")
        for c in ("H", "K", "L", "P", "V", "X"):
            ws.data_validation(f"{c}{a}:{c}{b}", {"validate": "decimal", "criteria": ">=", "value": 0,
                                                  "error_message": "Ingrese un valor mayor o igual a cero."})
        ws.data_validation(f"W{a}:W{b}", {"validate": "decimal", "criteria": "between", "minimum": 0, "maximum": 0.9})
        est = f"T{a}:T{b}"
        ws.conditional_format(est, {"type": "cell", "criteria": "==", "value": '"AGOTADO"', "format": F(bg_color="#FEE2E2", font_color=RED, bold=True)})
        ws.conditional_format(est, {"type": "cell", "criteria": "==", "value": '"BAJO STOCK"', "format": F(bg_color="#FEF3C7", font_color=AMBER, bold=True)})
        ws.conditional_format(est, {"type": "cell", "criteria": "==", "value": '"DISPONIBLE"', "format": F(font_color=GREEN, bold=True)})
        ws.conditional_format(est, {"type": "cell", "criteria": "==", "value": '"INACTIVO"', "format": F(font_color=GREY, italic=True)})
        ws.conditional_format(f"O{a}:O{b}", {"type": "cell", "criteria": "<", "value": 0, "format": F(bg_color="#FEE2E2", font_color=RED, bold=True)})
        # panel de consulta y KPIs
        ws.set_row(2, 8)
        ws.write("A4", "CONSULTA DE EXISTENCIAS", F(bold=True, font_color=BLUE))
        labels = [("A5", "Categoría"), ("C5", "Talla"), ("E5", "Color"), ("G5", "SKU contiene")]
        for addr, txt in labels:
            ws.write(addr, txt, F.lbl(align="right"))
        for addr in ("B5", "D5", "F5", "H5"):
            ws.write_blank(addr, None, F.inp())
        self.dv_list(k, "B5", "=lstCategorias", "information")
        self.dv_list(k, "D5", "=lstTallas", "information")
        self.dv_list(k, "F5", "=lstColores", "information")
        crit = lambda c: f'IF({c}="","<>§",{c})'
        crs = f'{L(1)}{a}:{L(1)}{b}'
        base = (f'$C${a}:$C${b},{crit("$B$5")},$E${a}:$E${b},{crit("$D$5")},$F${a}:$F${b},{crit("$F$5")},'
                f'$A${a}:$A${b},IF($H$5="","<>§","*"&$H$5&"*")')
        ws.write("A6", "Unidades en stock", F.lbl(align="right"))
        ws.write_formula("B6", f"=SUMIFS($O${a}:$O${b},{base})", F.out("int", bold=True, bg_color="#DBEAFE"))
        ws.write("C6", "Valor", F.lbl(align="right"))
        ws.write_formula("D6", f"=SUMIFS($Q${a}:$Q${b},{base})", F.out("cop", bold=True, bg_color="#DBEAFE"))
        ws.write("E6", "Referencias", F.lbl(align="right"))
        ws.write_formula("F6", f"=COUNTIFS({base})", F.out("int", bold=True, bg_color="#DBEAFE"))
        kp = [("J4", "Referencias", f"=COUNTA($A${a}:$A${b})", "int"),
              ("L4", "Unidades totales", f"=SUM($O${a}:$O${b})", "int"),
              ("N4", "Valor inventario (costo)", f"=SUM($Q${a}:$Q${b})", "cop"),
              ("P4", "Agotados", f'=COUNTIF($T${a}:$T${b},"AGOTADO")', "int"),
              ("R4", "Bajo stock", f'=COUNTIF($T${a}:$T${b},"BAJO STOCK")', "int"),
              ("T4", "Más vendido", f'=IFERROR(INDEX($A${a}:$A${b},MATCH(MAX($R${a}:$R${b}),$R${a}:$R${b},0)),"-")', "txt")]
        for addr, lab, fm, nf in kp:
            ws.write(addr, lab, F(font_color=GREY, font_size=9, bg_color="#EEF2FF", border=1, border_color=LINE))
            r = int(addr[1:])
            ws.write_formula(f"{addr[0]}{r+1}", fm, F(bold=True, font_size=13, bg_color="#EEF2FF",
                                                          border=1, border_color=LINE, nf=nf))
        ws.write("A7", "Botones: NUEVO PRODUCTO · DEVOLUCIÓN · AJUSTE · VER MOVIMIENTOS (columna derecha). Encabezados: azul = captura, verde = fórmula.", F(italic=True, font_color=GREY, font_size=9))
        self.proteger(k)

    # ------------------------------------------------------------ MOVIMIENTOS, BD, HFA, HGU, AUD
    def hojas_historial(self):
        info = {"BD": ("BASE DE DATOS — VENTAS (DETALLE)", "Una fila por producto facturado. Registro automático desde FACTURACIÓN. No editar."),
                "HFA": ("HISTORIAL DE FACTURAS", "Encabezado de cada factura guardada. Las anuladas se conservan con su motivo y responsable."),
                "HGU": ("HISTORIAL DE GUÍAS DE ENVÍO", "Registro de cada guía guardada."),
                "MOV": ("MOVIMIENTOS DE INVENTARIO (KARDEX)", "Quién, cuándo y por qué cambió cada existencia. Base del stock actual."),
                "AUD": ("AUDITORÍA", "Bitácora de acciones sensibles: anulaciones, ajustes, cambios de configuración, copias de seguridad.")}
        data = {}
        if self.sim:
            data = {"BD": self.sim.bd, "HFA": self.sim.hfa, "HGU": self.sim.hgu,
                    "MOV": self.sim.mov, "AUD": self.sim.aud}
        for k, (tt, sub) in info.items():
            self.banner(k, tt, sub)
            self.tabla_hdr(k, k)
            self.tabla_valores(k, k, data.get(k, []))
            t = TABLAS[k]
            if "ESTADO" in [c[0] for c in t["cols"]]:
                c = L(col_index(k, "ESTADO"))
                self.ws[k].conditional_format(f"A{t['first']}:{L(len(t['cols']))}{t['first']+20000}",
                                              {"type": "formula", "criteria": f'=${c}{t["first"]}="ANULADA"',
                                               "format": self.F(font_color="#94A3B8", italic=True)})
            self.proteger(k)

    # ------------------------------------------------------------ CLIENTES
    def hoja_cli(self):
        k, ws, F = "CLI", self.ws["CLI"], self.F
        self.banner(k, "CLIENTES", "Base de datos de clientes. Totales, saldos y última compra se calculan solos desde facturación y cartera.")
        self.tabla_hdr(k, "CLI")
        C = lambda key, r: cell("CLI", key, r)
        hc, hq, hb = rng("HFA", "CODCLI"), rng("HFA", "ESTADO"), rng("HFA", "FECHA")
        fx = {
            "PRIMERA": lambda r: f'=IF({C("CODIGO",r)}="","",IF({C("NFACT",r)}=0,"",SUMPRODUCT(MIN(({hc}={C("CODIGO",r)})*({hq}="ACTIVA")*{hb}+((({hc}<>{C("CODIGO",r)})+({hq}<>"ACTIVA"))>0)*999999))))',
            "ULTIMA": lambda r: f'=IF({C("CODIGO",r)}="","",IF({C("NFACT",r)}=0,"",SUMPRODUCT(MAX(({hc}={C("CODIGO",r)})*({hq}="ACTIVA")*{hb}))))',
            "NFACT": lambda r: f'=IF({C("CODIGO",r)}="","",COUNTIFS({hc},{C("CODIGO",r)},{hq},"ACTIVA"))',
            "TOTAL": lambda r: f'=IF({C("CODIGO",r)}="","",SUMIFS({rng("BD","SUBTOTAL")},{rng("BD","CODCLI")},{C("CODIGO",r)},{rng("BD","ESTADO")},"ACTIVA"))',
            "SALDO": lambda r: f'=IF({C("CODIGO",r)}="","",SUMIFS({rng("CAR","SALDO")},{rng("CAR","CODTER")},{C("CODIGO",r)},{rng("CAR","TIPO")},"POR COBRAR"))',
            "CLASIF": lambda r: f'=IF({C("CODIGO",r)}="","",IF({C("NFACT",r)}>=cfgFrecuente,"FRECUENTE",IF({C("NFACT",r)}>0,"OCASIONAL","SIN COMPRAS")))',
            "CARTERA": lambda r: f'=IF({C("CODIGO",r)}="","",IF({C("SALDO",r)}>0,IF(COUNTIFS({rng("CAR","CODTER")},{C("CODIGO",r)},{rng("CAR","TIPO")},"POR COBRAR",{rng("CAR","ESTADO")},"VENCIDA")>0,"VENCIDA","CON SALDO"),"AL DÍA"))',
            "COMBO": lambda r: f'=IF({C("CODIGO",r)}="","",{C("CODIGO",r)}&" | "&{C("NOMBRE",r)})',
        }
        rows = [list(c[:9]) for c in DM.CLIENTES] if self.sim else [list(DM.CLIENTES[0][:9])]
        self.tabla_maestra(k, "CLI", rows, fx)
        t = TABLAS["CLI"]
        a, b = t["first"], t["last"]
        ws.data_validation(f"A{a}:A{b}", {"validate": "custom", "value": f"=COUNTIF($A${a}:$A${b},A{a})=1",
                                          "error_message": "Código de cliente duplicado."})
        self.dv_list(k, f"I{a}:I{b}", "=lstDeptos", "warning")
        for col, val, fm in (("O", "FRECUENTE", F(font_color=GREEN, bold=True)),
                             ("P", "VENCIDA", F(bg_color="#FEE2E2", font_color=RED, bold=True)),
                             ("P", "CON SALDO", F(font_color=AMBER, bold=True))):
            ws.conditional_format(f"{col}{a}:{col}{b}", {"type": "cell", "criteria": "==", "value": f'"{val}"', "format": fm})
        kp = [("A4", "Clientes", f"=COUNTA($A${a}:$A${b})", "int"),
              ("C4", "Frecuentes", f'=COUNTIF($O${a}:$O${b},"FRECUENTE")', "int"),
              ("E4", "Con cartera", f'=COUNTIF($N${a}:$N${b},">0")', "int"),
              ("G4", "Cartera total clientes", f"=SUM($N${a}:$N${b})", "cop")]
        for addr, lab, fm, nf in kp:
            ws.write(addr, lab, F(font_color=GREY, font_size=9, bg_color="#EEF2FF"))
            col = addr[0]
            ws.write_formula(f"{chr(ord(col)+1)}4", fm, F(bold=True, font_size=12, bg_color="#EEF2FF", nf=nf))
        self.proteger(k)

    # ------------------------------------------------------------ PROVEEDORES
    def hoja_pro(self):
        k, ws, F = "PRO", self.ws["PRO"], self.F
        self.banner(k, "PROVEEDORES", "Base de proveedores. Total comprado y saldo pendiente provienen de COMPRAS y CARTERA.")
        self.tabla_hdr(k, "PRO")
        C = lambda key, r: cell("PRO", key, r)
        act = f'{rng("COM","ESTADO")},"ACTIVA"'
        fx = {
            "UNIDADES": lambda r: f'=IF({C("NOMBRE",r)}="","",SUMIFS({rng("COM","CANT")},{rng("COM","PROVEEDOR")},{C("NOMBRE",r)},{act}))',
            "TOTAL": lambda r: f'=IF({C("NOMBRE",r)}="","",SUMIFS({rng("COM","COSTOREAL")},{rng("COM","PROVEEDOR")},{C("NOMBRE",r)},{act}))',
            "SALDO": lambda r: f'=IF({C("NOMBRE",r)}="","",SUMIFS({rng("CAR","SALDO")},{rng("CAR","TERCERO")},{C("NOMBRE",r)},{rng("CAR","TIPO")},"POR PAGAR"))',
            "ESTADO": lambda r: f'=IF({C("NOMBRE",r)}="","",IF({C("SALDO",r)}>0,IF(COUNTIFS({rng("CAR","TERCERO")},{C("NOMBRE",r)},{rng("CAR","TIPO")},"POR PAGAR",{rng("CAR","ESTADO")},"VENCIDA")>0,"VENCIDA","CON SALDO"),"AL DÍA"))',
        }
        rows = [list(p) for p in DM.PROVEEDORES] if self.sim else []
        self.tabla_maestra(k, "PRO", rows, fx)
        t = TABLAS["PRO"]
        a, b = t["first"], t["last"]
        ws.data_validation(f"B{a}:B{b}", {"validate": "custom", "value": f"=COUNTIF($B${a}:$B${b},B{a})=1",
                                          "error_message": "Ya existe un proveedor con ese nombre."})
        ws.conditional_format(f"M{a}:M{b}", {"type": "cell", "criteria": "==", "value": '"VENCIDA"', "format": F(bg_color="#FEE2E2", font_color=RED, bold=True)})
        ws.write("A4", "Total por pagar a proveedores", F(font_color=GREY, bg_color="#EEF2FF"))
        ws.write_formula("B4", f"=SUM($L${a}:$L${b})", F(bold=True, font_size=12, bg_color="#EEF2FF", nf="cop"))
        self.proteger(k)

    # ------------------------------------------------------------ COMPRAS
    def hoja_com(self):
        k, ws, F = "COM", self.ws["COM"], self.F
        self.banner(k, "COMPRAS DE MERCANCÍA", "Diligencie el formulario y pulse REGISTRAR COMPRA: aumenta el inventario, actualiza el costo promedio y genera el pago o la cuenta por pagar.")
        ws.write("B4", "NUEVA COMPRA", F(bold=True, font_size=12, font_color=BLUE))
        lab = F.lbl(align="right")
        for addr, txt in (("B5", "N° compra"), ("E5", "Fecha"), ("H5", "Proveedor"),
                          ("B6", "Forma de pago"), ("E6", "Estado de pago"), ("H6", "Valor pagado"),
                          ("B7", "Transporte (flete)"), ("E7", "Otros costos"), ("H7", "Vencimiento (crédito)"),
                          ("B8", "Observaciones")):
            ws.write(addr, txt, lab)
        ws.write_formula(COM["NUM"], '=cfgPrefCompra&TEXT(cfgSigCompra,"0000")', F.out(bold=True, bg_color="#E2E8F0"))
        ws.write_formula(COM["FECHA"], "=TODAY()", F.inp("date"))
        ws.write_blank(COM["PROV"], None, F.inp())
        ws.write(COM["FORMA"], "Transferencia", F.inp())
        ws.write(COM["ESTADOPAGO"], "PAGADA", F.inp())
        ws.write_blank(COM["PAGADO"], None, F.inp("cop"))
        ws.write_blank(COM["TRANSP"], None, F.inp("cop"))
        ws.write_blank(COM["OTROS"], None, F.inp("cop"))
        ws.write_blank(COM["VENCE"], None, F.inp("date"))
        ws.merge_range("C8:I8", "", F.inp())
        self.dv_list(k, COM["PROV"], "=lstProveedores", msg="Seleccione un proveedor registrado en PROVEEDORES.")
        self.dv_list(k, COM["FORMA"], "=lstMetodos")
        self.dv_list(k, COM["ESTADOPAGO"], "=lstEstadosPago")
        ws.write("J6", "← vacío = total si PAGADA, 0 si PENDIENTE", F(italic=True, font_color=GREY, font_size=8))
        heads = ["SKU", "Producto", "Talla", "Color", "Cantidad", "Costo unit.", "Costo total", "Costo real unit. (con fletes)", "Costo actual inv."]
        for i, h in enumerate(heads):
            ws.write(9, 1 + i, h, F.hdr(bg_color=BLUE if h in ("SKU", "Cantidad", "Costo unit.") else TEAL))
        ti = TABLAS["INV"]
        ia = lambda c: f"{q('INV')}!${c}${ti['first']}:${c}${ti['last']}"
        for r in range(COM_L1, COM_LN + 1):
            ws.write_blank(f"B{r}", None, F.inp())
            look = lambda c: f'IF($B{r}="","",IFERROR(INDEX({ia(c)},MATCH($B{r},{ia("A")},0)),"⚠ SKU NO EXISTE"))'
            ws.write_formula(f"C{r}", "=" + look("B"), F.out())
            ws.write_formula(f"D{r}", f'=IF(LEFT(C{r},1)="⚠","",{look("E")})', F.out(align="center"))
            ws.write_formula(f"E{r}", f'=IF(LEFT(C{r},1)="⚠","",{look("F")})', F.out())
            ws.write_blank(f"F{r}", None, F.inp("int"))
            ws.write_blank(f"G{r}", None, F.inp("cop"))
            ws.write_formula(f"H{r}", f'=IF(OR(B{r}="",F{r}="",G{r}=""),"",F{r}*G{r})', F.out("cop"))
            ws.write_formula(f"I{r}", f'=IF(H{r}="","",(H{r}+IF(SUM($H${COM_L1}:$H${COM_LN})=0,0,H{r}/SUM($H${COM_L1}:$H${COM_LN})*(N($C$7)+N($F$7))))/F{r})', F.out("cop"))
            ws.write_formula(f"J{r}", f'=IF($B{r}="","",IFERROR(INDEX({ia("H")},MATCH($B{r},{ia("A")},0)),""))', F.out("cop", font_color=GREY))
        self.dv_list(k, f"B{COM_L1}:B{COM_LN}", "=lstSKU", msg="El SKU debe existir en INVENTARIO. Cree primero el producto.")
        ws.data_validation(f"F{COM_L1}:F{COM_LN}", {"validate": "integer", "criteria": ">", "value": 0, "error_message": "La cantidad debe ser un entero mayor que cero."})
        ws.data_validation(f"G{COM_L1}:G{COM_LN}", {"validate": "decimal", "criteria": ">=", "value": 0})
        ws.write("G19", "Subtotal", lab)
        ws.write_formula(COM["SUBTOTAL"], f"=SUM(H{COM_L1}:H{COM_LN})", F.out("cop", bold=True))
        ws.write("G20", "TOTAL COMPRA", F.lbl(align="right", font_size=11))
        ws.write_formula(COM["TOTAL"], f"={COM['SUBTOTAL']}+N(C7)+N(F7)", F.out("cop", bold=True, font_size=11, bg_color="#DBEAFE"))
        ws.write("B21", "HISTORIAL DE COMPRAS", F(bold=True, font_size=12, font_color=BLUE))
        self.tabla_hdr(k, "COM", freeze=False)
        ws.set_column("A:A", 12)
        self.tabla_valores(k, "COM", self.sim.com if self.sim else [])
        self.proteger(k)

    # ------------------------------------------------------------ COSTOS
    def hoja_cos(self):
        k, ws, F = "COS", self.ws["COS"], self.F
        self.banner(k, "COSTOS Y RENTABILIDAD", "Cuánto gana realmente por producto. Calculadora de precios con margen s/ventas vs. recargo s/costo.")
        ws.set_column("A:A", 17)
        ws.set_column("B:B", 40)
        ws.set_column("C:E", 17)
        ws.set_column("F:F", 3)
        ws.set_column("G:T", 13)
        ws.write("B4", "CALCULADORA DE PRECIOS", F(bold=True, font_size=12, font_color=BLUE))
        ws.write("B5", "Producto (SKU, opcional)", F.lbl())
        ws.write_blank("C5", None, F.inp())
        self.dv_list(k, "C5", "=lstSKU", msg="Seleccione un SKU del inventario o deje vacío para cálculo manual.")
        ws.write_formula("D5", f'=IF(C5="","(cálculo manual)",IFERROR(INDEX({q("INV")}!$B$9:$B$1008,MATCH(C5,{q("INV")}!$A$9:$A$1008,0)),""))', F(italic=True, font_color=GREY))
        for i, h in enumerate(["Concepto", "Desde inventario", "Valor manual (opcional)", "Valor utilizado"]):
            ws.write(5, 1 + i, h, F.hdr(bg_color=BLUE if i == 2 else NAVY))
        ia = lambda c: f"{q('INV')}!${c}$9:${c}$1008"
        lk = lambda c: f'IFERROR(INDEX({ia(c)},MATCH($C$5,{ia("A")},0)),"")'
        items = [(7, "Costo de adquisición (unit.)", f'=IF($C$5="","",{lk("H")})', "cop"),
                 (8, "Transporte / otros de compra (unit.)", f'=IF($C$5="","",{lk("I")})', "cop"),
                 (9, "Empaque (unit.)", f'=IF($C$5="","",IF({lk("V")}="",cfgEmpaque,{lk("V")}))', "cop"),
                 (10, "Comisiones (% del precio de venta)", f'=IF($C$5="","",IF({lk("W")}="",cfgComision,{lk("W")}))', "pct"),
                 (11, "Otros costos variables (unit.)", f'=IF($C$5="","",N({lk("X")}))', "cop"),
                 (12, "Precio de venta actual", f'=IF($C$5="","",{lk("K")})', "cop")]
        for r, lab, fm, nf in items:
            ws.write(f"B{r}", lab, F.lbl(border=1, border_color=LINE))
            ws.write_formula(f"C{r}", fm, F.out(nf, font_color=GREY))
            ws.write_blank(f"D{r}", None, F.inp(nf))
            ws.write_formula(f"E{r}", f'=IF(D{r}<>"",D{r},N(C{r}))', F.out(nf, bold=True))
        ws.write("B13", "Margen de utilidad deseado (s/ventas)", F.lbl(border=1, border_color=LINE))
        ws.write("D13", 0.40, F.inp("pct"))
        ws.write_formula("E13", "=D13", F.out("pct", bold=True))
        ws.write("B14", "Recargo deseado (s/costo, markup)", F.lbl(border=1, border_color=LINE))
        ws.write("D14", 0.60, F.inp("pct"))
        ws.write_formula("E14", "=D14", F.out("pct", bold=True))
        ws.write("B16", "RESULTADOS", F.hdr(align="left"))
        ws.write("E16", "", F.hdr())
        res = [(17, "Costo real unitario (sin comisión)", "=E7+E8+E9+E11", "cop"),
               (18, "Comisión al precio actual ($)", "=E12*E10", "cop"),
               (19, "Costo real total al precio actual", "=E17+E18", "cop"),
               (20, "Ganancia bruta unitaria", "=E12-E19", "cop"),
               (21, "Margen de utilidad (s/ventas)", "=IFERROR(E20/E12,0)", "pct"),
               (22, "Recargo sobre costo (markup)", "=IFERROR(E20/E19,0)", "pct"),
               (23, "Rentabilidad sobre la inversión en mercancía", "=IFERROR(E20/E7,0)", "pct"),
               (24, "Precio mínimo de venta (punto de equilibrio)", "=IFERROR(ROUNDUP(E17/(1-E10)/100,0)*100,0)", "cop"),
               (25, "Precio con el MARGEN deseado", '=IF(1-E10-E13<=0,"N/A",ROUNDUP(E17/(1-E10-E13)/100,0)*100)', "cop"),
               (26, "Precio con el RECARGO deseado", '=IF(1-E10<=0,"N/A",ROUNDUP(E17*(1+E14)/(1-E10)/100,0)*100)', "cop")]
        for r, lab, fm, nf in res:
            ws.merge_range(f"B{r}:D{r}", lab, F.lbl(border=1, border_color=LINE))
            ws.write_formula(f"E{r}", fm, F.out(nf, bold=True, bg_color="#F8FAFC"))
        for i, h in enumerate(["Escenario de precio", "Margen s/ventas", "Precio sugerido", "Ganancia unit."]):
            ws.write(27, 1 + i, h, F.hdr())
        esc = [(29, "Precio competitivo", "cfgMargenComp"), (30, "Precio recomendado (SUGERIDO)", "cfgMargenRec"),
               (31, "Precio premium", "cfgMargenPrem")]
        for r, lab, nm in esc:
            ws.write(f"B{r}", lab, F.lbl(border=1, border_color=LINE, bg_color="#DCFCE7" if r == 30 else None))
            ws.write_formula(f"C{r}", f"={nm}", F.out("pct"))
            ws.write_formula(f"D{r}", f'=IF(1-$E$10-C{r}<=0,"N/A",ROUNDUP($E$17/(1-$E$10-C{r})/100,0)*100)', F.out("cop", bold=True))
            ws.write_formula(f"E{r}", f'=IFERROR(D{r}-$E$17-D{r}*$E$10,"")', F.out("cop"))
        ws.merge_range("B32:E32", "", F(bold=True, font_size=11, align="center", border=1, border_color=LINE))
        ws.write_formula("B32", '=IF(E12=0,"Seleccione un producto o ingrese valores manuales",IF(E12<E19,"⚠ EL PRECIO ACTUAL ESTÁ POR DEBAJO DEL COSTO REAL: SE PIERDE DINERO",IF(E21<cfgMargenMin,"⚠ Margen inferior al mínimo configurado","✔ Precio rentable")))',
                         F(bold=True, font_size=11, align="center", border=1, border_color=LINE))
        ws.conditional_format("B32", {"type": "text", "criteria": "begins with", "value": "⚠", "format": F(bg_color="#FEE2E2", font_color=RED)})
        ws.conditional_format("B32", {"type": "text", "criteria": "begins with", "value": "✔", "format": F(bg_color="#DCFCE7", font_color=GREEN)})
        exp = ["¿MARGEN O RECARGO? NO SON LO MISMO",
               "",
               "Margen s/ventas = (Precio − Costo) ÷ Precio",
               "  → Precio = Costo ÷ (1 − Margen)",
               "Recargo s/costo = (Precio − Costo) ÷ Costo",
               "  → Precio = Costo × (1 + Recargo)",
               "",
               "Ejemplo con costo $50.000 y 40%:",
               "  • Margen 40%  → precio $83.334 (gana $33.334)",
               "  • Recargo 40% → precio $70.000 (gana $20.000,",
               "    que es solo un 28,6% de margen)",
               "",
               "Costo real = adquisición + fletes de compra +",
               "empaque + otros variables + comisión (% del precio).",
               "Precio mínimo = punto donde la ganancia es $0.",
               "",
               "En UTILIDADES el costo de la mercancía vendida usa",
               "adquisición + fletes; empaques y comisiones reales",
               "se registran como EGRESOS (sin contarlos dos veces)."]
        for i, t in enumerate(exp):
            ws.write(4 + i, 6, t, F(bold=(i == 0), font_color=BLUE if i == 0 else "#334155",
                                    bg_color="#EFF6FF", font_size=10 if i else 11))
            for c in range(7, 12):
                ws.write_blank(4 + i, c, None, F(bg_color="#EFF6FF"))
        # tabla por producto
        hdr = 34
        heads = [("SKU", 17), ("Producto", 40), ("Categoría", 17), ("Costo adq.", 17), ("Transp. unit.", 17),
                 ("Empaque", 13), ("Comisión %", 13), ("Otros var.", 13), ("Costo real (sin com.)", 13),
                 ("Precio venta", 13), ("Comisión $", 13), ("Costo total", 13), ("Ganancia unit.", 13),
                 ("Margen s/ventas", 13), ("Recargo s/costo", 13), ("Precio mínimo", 13),
                 ("P. competitivo", 13), ("P. recomendado", 13), ("P. premium", 13), ("Alerta", 18)]
        ws.write(hdr - 2, 6, "", None)
        ws.write("A33", "RENTABILIDAD POR PRODUCTO (todo el inventario)", F(bold=True, font_size=12, font_color=BLUE))
        for i, (h, w) in enumerate(heads):
            ws.write(hdr - 1, i, h, F.hdr(bg_color=TEAL))
        ws.set_column("T:T", 18)
        nfs = ["txt", "txt", "txt", "cop", "cop", "cop", "pct", "cop", "cop", "cop", "cop", "cop", "cop",
               "pct", "pct", "cop", "cop", "cop", "cop", "txt"]
        fms = [F.out(n if n != "txt" else None, border_color="#E2E8F0") for n in nfs]
        I = lambda c, i: f"{q('INV')}!{c}{i}"
        for n in range(1000):
            r = hdr + 1 + n
            i = 9 + n
            A = f"$A{r}"
            def w(ci, fm):
                ws.write_formula(r - 1, ci, fm, fms[ci])
            w(0, f'=IF({I("A",i)}="","",{I("A",i)})')
            w(1, f'=IF({A}="","",{I("B",i)}&" "&{I("E",i)}&" "&{I("F",i)})')
            w(2, f'=IF({A}="","",{I("C",i)})')
            w(3, f'=IF({A}="","",N({I("H",i)}))')
            w(4, f'=IF({A}="","",N({I("I",i)}))')
            w(5, f'=IF({A}="","",IF({I("V",i)}="",cfgEmpaque,{I("V",i)}))')
            w(6, f'=IF({A}="","",IF({I("W",i)}="",cfgComision,{I("W",i)}))')
            w(7, f'=IF({A}="","",N({I("X",i)}))')
            w(8, f'=IF({A}="","",D{r}+E{r}+F{r}+H{r})')
            w(9, f'=IF({A}="","",N({I("K",i)}))')
            w(10, f'=IF({A}="","",J{r}*G{r})')
            w(11, f'=IF({A}="","",I{r}+K{r})')
            w(12, f'=IF({A}="","",J{r}-L{r})')
            w(13, f'=IF({A}="","",IFERROR(M{r}/J{r},0))')
            w(14, f'=IF({A}="","",IFERROR(M{r}/L{r},0))')
            w(15, f'=IF({A}="","",IFERROR(ROUNDUP(I{r}/(1-G{r})/100,0)*100,0))')
            for ci, nm in ((16, "cfgMargenComp"), (17, "cfgMargenRec"), (18, "cfgMargenPrem")):
                w(ci, f'=IF({A}="","",IF(1-G{r}-{nm}<=0,"N/A",ROUNDUP(I{r}/(1-G{r}-{nm})/100,0)*100))')
            w(19, f'=IF({A}="","",IF(J{r}<L{r},"⚠ BAJO COSTO",IF(N{r}<cfgMargenMin,"⚠ MARGEN BAJO","✔ OK")))')
        rr = f"T{hdr+1}:T{hdr+1000}"
        ws.conditional_format(rr, {"type": "text", "criteria": "begins with", "value": "⚠", "format": F(bg_color="#FEE2E2", font_color=RED, bold=True)})
        ws.conditional_format(rr, {"type": "text", "criteria": "begins with", "value": "✔", "format": F(font_color=GREEN, bold=True)})
        ws.autofilter(hdr - 1, 0, hdr + 999, 19)
        self.proteger(k)

    # ------------------------------------------------------------ FACTURACIÓN
    def encabezado_doc(self, k, titulo):
        ws, F = self.ws[k], self.F
        if os.path.exists(self.logo):
            ws.insert_image("A1", self.logo, {"x_scale": 0.62, "y_scale": 0.62, "x_offset": 6, "y_offset": 6,
                                             "object_position": 3, "description": "LogoEmpresa"})
        ws.merge_range("C1:F2", "", F(bold=True, font_size=18, font_color=NAVY))
        ws.write_formula("C1", "=cfgEmpresa", F(bold=True, font_size=18, font_color=NAVY))
        ws.merge_range("C3:F3", "", F(font_size=9, font_color=GREY))
        ws.write_formula("C3", '="NIT: "&cfgNIT&"  ·  "&cfgRegimen', F(font_size=9, font_color=GREY))
        ws.merge_range("C4:F4", "", F(font_size=9, font_color=GREY))
        ws.write_formula("C4", '=cfgDireccion&"  ·  "&cfgCiudad', F(font_size=9, font_color=GREY))
        ws.merge_range("C5:F5", "", F(font_size=9, font_color=GREY))
        ws.write_formula("C5", '="Tel: "&cfgTelefono&"  ·  "&cfgCorreo', F(font_size=9, font_color=GREY))
        ws.merge_range("G1:I1", titulo, F(bold=True, font_size=13, font_color="white", bg_color=NAVY, align="center"))

    def hoja_fac(self):
        k, ws, F = "FAC", self.ws["FAC"], self.F
        ws.hide_gridlines(2)
        ws.set_zoom(100)
        for c, w in zip("ABCDEFGHI", (5, 14, 34, 8, 11, 8, 14, 9, 17)):
            ws.set_column(f"{c}:{c}", w)
        ws.set_column("J:J", 3)
        ws.set_column("K:K", 30)
        ws.set_column("L:M", 20)
        ws.set_column("T:Z", 12, None, {"hidden": True})
        for r in range(0, 5):
            ws.set_row(r, 19)
        self.encabezado_doc(k, "FACTURA DE VENTA")
        lab = F.lbl(font_size=9, align="right", bg_color=LIGHT, border=1, border_color=LINE)
        for r, txt in ((2, "N°"), (3, "Fecha de expedición"), (4, "Estado"), (5, "Vence")):
            ws.merge_range(f"G{r}:H{r}", txt, lab)
        num0 = f"FV-{self.sim.sig['FV']:04d}" if self.sim else "FV-0001"
        ws.write(FAC["NUM"], num0, F.out(bold=True, font_size=13, font_color=RED, align="center"))
        ws.write_formula(FAC["FECHA"], "=TODAY()", F.inp("date", align="center"))
        ws.write(FAC["ESTADO"], "BORRADOR", F.out(bold=True, align="center"))
        ws.write_formula(FAC["VENCE"], f'=IF({FAC["ESTADOPAGO"]}="PAGADA","Contado",IF({FAC["FECHA"]}="","",{FAC["FECHA"]}+cfgDiasCredito))', F.out("date", align="center"))
        ws.write_url("K1", f"internal:'{SH['INI']}'!A1", F(bold=True, font_color="white", bg_color=BLUE, align="center"), string="⌂ INICIO")
        sec = F(bold=True, font_color="white", bg_color=BLUE, font_size=10)
        ws.merge_range("A7:I7", "DATOS DEL CLIENTE", sec)
        L2 = lambda r, t: ws.merge_range(f"A{r}:B{r}", t, F.lbl(font_size=9))
        L2(8, "Cliente")
        ws.merge_range("C8:F8", "", F.inp(bold=True))
        if self.sim:
            ws.write("C8", "CL-0001 | CONSUMIDOR FINAL", F.inp(bold=True))
        self.dv_list(k, "C8", "=lstClientes", msg="Seleccione un cliente registrado (CLIENTES). Para uno nuevo use el botón NUEVO CLIENTE.")
        tc = TABLAS["CLI"]
        ca = lambda c: f"{q('CLI')}!${c}${tc['first']}:${c}${tc['last']}"
        code = '$T$8'
        ws.write_formula("T8", f'=IFERROR(LEFT($C$8,FIND(" |",$C$8)-1),"")')
        lk = lambda c: f'IF({code}="","",IFERROR(INDEX({ca(c)},MATCH({code},{ca("A")},0))&"",""))'
        L2(9, "Nombre")
        ws.merge_range("C9:F9", "", F.out(bold=True))
        ws.write_formula("C9", "=" + lk("B"), F.out(bold=True))
        L2(10, "Dirección")
        ws.merge_range("C10:F10", "", F.out())
        ws.write_formula("C10", "=" + lk("G"), F.out())
        L2(11, "Entregar en")
        ws.merge_range("C11:F11", "", F.inp())
        for r, t, c in ((8, "Documento", "C"), (9, "Teléfono", "D"), (10, "Ciudad", "H"), (11, "WhatsApp", "E")):
            ws.merge_range(f"G{r}:H{r}", t, F.lbl(font_size=9, align="right"))
            ws.write_formula(f"I{r}", "=" + lk(c), F.out())
        ws.write("K11", "← opcional: si es diferente a la dirección del cliente", F(italic=True, font_color=GREY, font_size=8))
        heads = ["#", "Código", "Descripción", "Talla", "Color", "Cant.", "Precio unit.", "Desc. %", "Subtotal"]
        for i, h in enumerate(heads):
            ws.write(12, i, h, F.hdr(bg_color=NAVY))
        ws.write("K13", "Validación (no se imprime)", F.hdr(bg_color=GREY))
        ti = TABLAS["INV"]
        ia = lambda c: f"{q('INV')}!${c}${ti['first']}:${c}${ti['last']}"
        for n, r in enumerate(range(FAC_L1, FAC_LN + 1)):
            zebra = "#F8FAFC" if n % 2 else None
            ws.write(f"A{r}", n + 1, F.out(align="center", font_color=GREY, bg_color=zebra))
            ws.write_blank(f"B{r}", None, F.inp())
            m = f'MATCH($B{r},{ia("A")},0)'
            ws.write_formula(f"C{r}", f'=IF($B{r}="","",IFERROR(INDEX({ia("B")},{m}),"⚠ SKU NO EXISTE"))', F.out(bg_color=zebra))
            ws.write_formula(f"D{r}", f'=IF($B{r}="","",IFERROR(INDEX({ia("E")},{m})&"",""))', F.out(align="center", bg_color=zebra))
            ws.write_formula(f"E{r}", f'=IF($B{r}="","",IFERROR(INDEX({ia("F")},{m})&"",""))', F.out(bg_color=zebra))
            ws.write_blank(f"F{r}", None, F.inp("int", align="center"))
            ws.write_formula(f"G{r}", f'=IF($B{r}="","",IFERROR(INDEX({ia("K")},{m}),0))', F.out("cop", bg_color=zebra))
            ws.write_blank(f"H{r}", None, F.inp("pct", align="center"))
            ws.write_formula(f"I{r}", f'=IF(OR($B{r}="",$F{r}=""),"",ROUND($F{r}*$G{r}*(1-N($H{r})),0))', F.out("cop", bold=True, bg_color=zebra))
            # auxiliares ocultas: T costo real, U categoria, V stock, W cantidad total del SKU en la factura
            ws.write_formula(f"T{r}", f'=IF($B{r}="",0,IFERROR(INDEX({ia("J")},{m}),0))')
            ws.write_formula(f"U{r}", f'=IF($B{r}="","",IFERROR(INDEX({ia("C")},{m}),""))')
            ws.write_formula(f"V{r}", f'=IF($B{r}="",0,IFERROR(INDEX({ia("O")},{m}),0))')
            ws.write_formula(f"W{r}", f'=IF($B{r}="",0,SUMIF($B${FAC_L1}:$B${FAC_LN},$B{r},$F${FAC_L1}:$F${FAC_LN}))')
            ws.write_formula(f"X{r}", f'=IF($B{r}="","",IFERROR(INDEX({ia("T")},{m}),""))')
            ws.write_formula(f"K{r}", f'=IF(OR($B{r}="",{FAC["ESTADO"]}<>"BORRADOR"),"",IF(ISNA({m}),"⛔ SKU no existe",IF($X{r}="INACTIVO","⛔ Producto inactivo",IF(N($F{r})<=0,"Falta cantidad",IF($W{r}>$V{r},"⛔ Stock insuficiente (hay "&$V{r}&")",IF($G{r}*(1-N($H{r}))<$T{r},"⚠ Precio bajo costo real","✔ OK"))))))',
                             F(font_size=9))
        rk = f"K{FAC_L1}:K{FAC_LN}"
        ws.conditional_format(rk, {"type": "text", "criteria": "begins with", "value": "⛔", "format": F(bg_color="#FEE2E2", font_color=RED, bold=True)})
        ws.conditional_format(rk, {"type": "text", "criteria": "begins with", "value": "⚠", "format": F(bg_color="#FEF3C7", font_color=AMBER, bold=True)})
        ws.conditional_format(rk, {"type": "text", "criteria": "begins with", "value": "✔", "format": F(font_color=GREEN)})
        self.dv_list(k, f"B{FAC_L1}:B{FAC_LN}", "=lstSKU", msg="El código debe existir en INVENTARIO.")
        ws.data_validation(f"F{FAC_L1}:F{FAC_LN}", {"validate": "integer", "criteria": ">", "value": 0, "error_message": "Cantidad entera mayor que cero."})
        ws.data_validation(f"H{FAC_L1}:H{FAC_LN}", {"validate": "decimal", "criteria": "between", "minimum": 0, "maximum": 1, "error_message": "Descuento entre 0% y 100%."})
        # pagos y totales
        lab2 = F.lbl(font_size=9, bg_color=LIGHT, border=1, border_color=LINE)
        for r, t in ((30, "Método de pago"), (31, "Estado de pago"), (32, "Valor pagado"), (33, "Saldo pendiente")):
            ws.merge_range(f"A{r}:B{r}", t, lab2)
        ws.merge_range("C30:D30", "Efectivo", F.inp(bold=True))
        ws.merge_range("C31:D31", "PAGADA", F.inp(bold=True))
        ws.merge_range("C32:D32", "", F.inp("cop"))
        ws.merge_range("C33:D33", "", F.out("cop", bold=True))
        ws.write_formula(FAC["PAGADOEF"], f'=IF({FAC["PAGADO"]}<>"",N({FAC["PAGADO"]}),IF({FAC["ESTADOPAGO"]}="PAGADA",{FAC["TOTAL"]},0))')
        ws.write_formula("C33", f'={FAC["TOTAL"]}-{FAC["PAGADOEF"]}', F.out("cop", bold=True))
        ws.write("E32", "← vacío = total si PAGADA", F(italic=True, font_color=GREY, font_size=7))
        self.dv_list(k, "C30", "=lstMetodos")
        self.dv_list(k, "C31", "=lstEstadosPago")
        ws.data_validation("C32", {"validate": "decimal", "criteria": ">=", "value": 0})
        for r, t in ((30, "Subtotal"), (31, "Descuentos"), (32, "TOTAL A PAGAR"), (33, "Unidades")):
            ws.merge_range(f"G{r}:H{r}", t, F.lbl(align="right", font_size=11 if r == 32 else 9))
        ws.write_formula(FAC["BRUTO"], f"=SUMPRODUCT(F{FAC_L1}:F{FAC_LN},G{FAC_L1}:G{FAC_LN})", F.out("cop"))
        ws.write_formula(FAC["DESC"], f"={FAC['BRUTO']}-SUM(I{FAC_L1}:I{FAC_LN})", F.out("cop"))
        ws.write_formula(FAC["TOTAL"], f"=SUM(I{FAC_L1}:I{FAC_LN})", F(bold=True, font_size=13, font_color="white", bg_color=NAVY, nf="cop", border=1))
        ws.write_formula(FAC["UNIDADES"], f"=SUM(F{FAC_L1}:F{FAC_LN})", F.out("int"))
        ws.merge_range("A35:I35", "OBSERVACIONES", sec)
        ws.merge_range("A36:I37", "", F.inp(text_wrap=True, valign="top"))
        ws.merge_range("A39:I39", "", F(font_size=8, font_color=GREY, align="center"))
        ws.write_formula("A39", '="Pagos: "&cfgBanco&" "&cfgTipoCuenta&" N° "&cfgCuenta&" a nombre de "&cfgTitular&"  ·  Nequi "&cfgNequi&"  ·  Daviplata "&cfgDaviplata', F(font_size=8, font_color=GREY, align="center"))
        ws.merge_range("A40:I41", "", F(font_size=7, font_color=GREY, align="center", text_wrap=True))
        ws.write_formula("A40", "=cfgPie", F(font_size=7, font_color=GREY, align="center", text_wrap=True))
        ws.merge_range("B43:D43", "Firma y sello", F(top=1, align="center", font_size=8, font_color=GREY))
        ws.merge_range("F43:I43", "Recibí conforme (cliente)", F(top=1, align="center", font_size=8, font_color=GREY))
        ws.merge_range("A44:I44", "", F(font_size=7, font_color=GREY, align="right"))
        ws.write_formula("A44", f'="Elaborado por: "&{q("INI")}!{INI["USUARIO"]}', F(font_size=7, font_color=GREY, align="right"))
        ws.conditional_format(FAC["ESTADO"], {"type": "cell", "criteria": "==", "value": '"ANULADA"', "format": F(bg_color="#FEE2E2", font_color=RED)})
        ws.conditional_format(FAC["ESTADO"], {"type": "cell", "criteria": "==", "value": '"GUARDADA"', "format": F(bg_color="#DCFCE7", font_color=GREEN)})
        ws.print_area("A1:I44")
        ws.set_paper(1)
        ws.set_portrait()
        ws.fit_to_pages(1, 1)
        ws.center_horizontally()
        ws.set_margins(0.4, 0.4, 0.5, 0.5)
        ws.write("L3", "Atajos: seleccione el cliente y los códigos de los productos; todo lo demás se calcula.", F(italic=True, font_color=GREY, font_size=8))
        self.proteger(k)

    # ------------------------------------------------------------ GUÍAS
    def hoja_gui(self):
        k, ws, F = "GUI", self.ws["GUI"], self.F
        ws.hide_gridlines(2)
        for c, w in zip("ABCDEFGHI", (5, 14, 34, 8, 11, 9, 13, 10, 17)):
            ws.set_column(f"{c}:{c}", w)
        ws.set_column("J:J", 3)
        ws.set_column("K:K", 30)
        ws.set_column("L:M", 20)
        for r in range(0, 5):
            ws.set_row(r, 19)
        self.encabezado_doc(k, "GUÍA DE ENVÍO")
        lab = F.lbl(font_size=9, align="right", bg_color=LIGHT, border=1, border_color=LINE)
        for r, t in ((2, "N° guía"), (3, "Fecha despacho"), (4, "Factura relacionada"), (5, "Estado del envío")):
            ws.merge_range(f"G{r}:H{r}", t, lab)
        num0 = f"GE-{self.sim.sig['GE']:04d}" if self.sim else "GE-0001"
        ws.write(GUI["NUM"], num0, F.out(bold=True, font_size=13, font_color=RED, align="center"))
        ws.write_formula(GUI["FECHA"], "=TODAY()", F.inp("date", align="center"))
        ws.write_blank(GUI["FACTURA"], None, F.inp(align="center"))
        ws.write(GUI["ESTADO"], "PREPARANDO", F.inp(bold=True, align="center"))
        self.dv_list(k, GUI["ESTADO"], "=lstEstadosEnvio")
        ws.write_url("K1", f"internal:'{SH['INI']}'!A1", F(bold=True, font_color="white", bg_color=BLUE, align="center"), string="⌂ INICIO")
        ws.write("K2", "Estado del registro", F.lbl(font_size=9))
        ws.write(GUI["REG"], "BORRADOR", F.out(bold=True, align="center"))
        sec = F(bold=True, font_color="white", bg_color=BLUE)
        ws.merge_range("A7:D7", "REMITENTE", sec)
        ws.merge_range("E7:I7", "DESTINATARIO", sec)
        rem = [(8, "Nombre", "=cfgEmpresa"), (9, "NIT", "=cfgNIT"), (10, "Dirección", '=cfgDireccion&", "&cfgCiudad'), (11, "Teléfono", "=cfgTelefono")]
        for r, t, fm in rem:
            ws.merge_range(f"A{r}:B{r}", t, F.lbl(font_size=9))
            ws.merge_range(f"C{r}:D{r}", "", F.out())
            ws.write_formula(f"C{r}", fm, F.out())
        dest = [(8, "Nombre", "DEST"), (9, "Documento", "DOC"), (10, "Teléfono", "TEL"), (11, "Ciudad", "CIUDAD"), (12, "Departamento", "DEPTO")]
        for r, t, key in dest:
            ws.merge_range(f"E{r}:F{r}", t, F.lbl(font_size=9, align="right"))
            ws.merge_range(f"G{r}:I{r}", "", F.inp(bold=(key == "DEST")))
        ws.merge_range("E13:F14", "Dirección completa", F.lbl(font_size=9, align="right"))
        ws.merge_range("G13:I14", "", F.inp(text_wrap=True, valign="top"))
        self.dv_list(k, GUI["DEPTO"], "=lstDeptos", "warning")
        ws.merge_range("A16:I16", "DATOS DEL TRANSPORTE", sec)
        ws.merge_range("A17:B17", "Transportadora", F.lbl(font_size=9))
        ws.merge_range("C17:D17", "", F.inp(bold=True))
        ws.merge_range("E17:F17", "N° de seguimiento", F.lbl(font_size=9, align="right"))
        ws.merge_range("G17:I17", "", F.inp(bold=True))
        ws.merge_range("A18:B18", "Valor declarado", F.lbl(font_size=9))
        ws.merge_range("C18:D18", "", F.inp("cop"))
        ws.merge_range("E18:F18", "N° de paquetes", F.lbl(font_size=9, align="right"))
        ws.merge_range("G18:I18", "", F.inp("int"))
        self.dv_list(k, GUI["TRANSP"], "=lstTransportadoras", "warning")
        heads = ["#", "Código", "Descripción", "Talla", "Color", "Cantidad"]
        for i, h in enumerate(heads):
            ws.write(19, i, h, F.hdr())
        ws.merge_range("G20:I20", "Verificado / observación", F.hdr())
        ti = TABLAS["INV"]
        ia = lambda c: f"{q('INV')}!${c}${ti['first']}:${c}${ti['last']}"
        for n, r in enumerate(range(GUI_L1, GUI_LN + 1)):
            ws.write(f"A{r}", n + 1, F.out(align="center", font_color=GREY))
            ws.write_blank(f"B{r}", None, F.inp())
            m = f'MATCH($B{r},{ia("A")},0)'
            ws.write_formula(f"C{r}", f'=IF($B{r}="","",IFERROR(INDEX({ia("B")},{m}),"⚠ SKU NO EXISTE"))', F.out())
            ws.write_formula(f"D{r}", f'=IF($B{r}="","",IFERROR(INDEX({ia("E")},{m})&"",""))', F.out(align="center"))
            ws.write_formula(f"E{r}", f'=IF($B{r}="","",IFERROR(INDEX({ia("F")},{m})&"",""))', F.out())
            ws.write_blank(f"F{r}", None, F.inp("int", align="center"))
            ws.merge_range(f"G{r}:I{r}", "", F.out())
        self.dv_list(k, f"B{GUI_L1}:B{GUI_LN}", "=lstSKU")
        ws.data_validation(f"F{GUI_L1}:F{GUI_LN}", {"validate": "integer", "criteria": ">", "value": 0})
        ws.merge_range("D33:E33", "Total unidades", F.lbl(align="right"))
        ws.write_formula(GUI["UNIDADES"], f"=SUM(F{GUI_L1}:F{GUI_LN})", F.out("int", bold=True, align="center"))
        ws.merge_range("A35:I35", "OBSERVACIONES", sec)
        ws.merge_range("A36:I37", "", F.inp(text_wrap=True, valign="top"))
        ws.merge_range("B40:D40", "Despachado por", F(top=1, align="center", font_size=8, font_color=GREY))
        ws.merge_range("F40:I40", "Recibido por (nombre, firma, fecha)", F(top=1, align="center", font_size=8, font_color=GREY))
        ws.merge_range("A42:I42", "", F(font_size=7, font_color=GREY, align="center"))
        ws.write_formula("A42", '="Documento de transporte interno · "&cfgEmpresa&" · "&cfgTelefono', F(font_size=7, font_color=GREY, align="center"))
        ws.print_area("A1:I42")
        ws.set_paper(1)
        ws.fit_to_pages(1, 1)
        ws.center_horizontally()
        ws.set_margins(0.4, 0.4, 0.5, 0.5)
        self.proteger(k)

    # ------------------------------------------------------------ INGRESOS
    def resumen_periodo(self, k, r0, col="H"):
        """Celdas de año/mes para resúmenes; devuelve (inicio, fin) como direcciones absolutas."""
        ws, F = self.ws[k], self.F
        ws.write(f"{col}{r0}", "Año", F.lbl(align="right"))
        c2 = chr(ord(col) + 1)
        c3 = chr(ord(col) + 2)
        c4 = chr(ord(col) + 3)
        ws.write(f"{c2}{r0}", DM.HOY.year, F.inp())
        ws.write(f"{c3}{r0}", "Mes", F.lbl(align="right"))
        ws.write(f"{c4}{r0}", "TODOS", F.inp())
        self.dv_list(k, f"{c4}{r0}", "=lstMeses")
        ini = f'IF(${c4}${r0}="TODOS",DATE(${c2}${r0},1,1),DATE(${c2}${r0},MATCH(${c4}${r0},lstMeses,0),1))'
        fin = f'IF(${c4}${r0}="TODOS",DATE(${c2}${r0}+1,1,1),DATE(${c2}${r0},MATCH(${c4}${r0},lstMeses,0)+1,1))'
        ws.write_formula(f"M{r0}", "=" + ini, F(nf="date", font_color="white"))
        ws.write_formula(f"M{r0+1}", "=" + fin, F(nf="date", font_color="white"))
        return f"$M${r0}", f"$M${r0+1}"

    def form_labels(self, k, items):
        ws, F = self.ws[k], self.F
        for addr, txt in items:
            ws.write(addr, txt, F.lbl(align="right"))

    def hoja_ing(self):
        k, ws, F = "ING", self.ws["ING"], self.F
        self.banner(k, "INGRESOS", "Las ventas y abonos llegan solos desde FACTURACIÓN y CARTERA (no se duplican). Aquí registre solo ingresos adicionales.")
        ws.write("B4", "REGISTRAR INGRESO ADICIONAL", F(bold=True, font_size=12, font_color=BLUE))
        self.form_labels(k, [("B5", "Fecha"), ("B6", "Tipo"), ("B7", "Concepto"), ("B8", "Valor"),
                             ("E5", "Método"), ("E6", "Tercero"), ("E7", "Soporte"), ("E8", "Observación")])
        ws.write_formula(ING["FECHA"], "=TODAY()", F.inp("date"))
        ws.write(ING["TIPO"], "OTRO INGRESO", F.inp())
        ws.write_blank(ING["CONCEPTO"], None, F.inp())
        ws.write_blank(ING["VALOR"], None, F.inp("cop"))
        ws.write(ING["METODO"], "Efectivo", F.inp())
        for a in ("TERCERO", "DOC", "OBS"):
            ws.write_blank(ING[a], None, F.inp())
        self.dv_list(k, ING["TIPO"], "=lstTiposIngreso", msg="Ventas y abonos NO se registran aquí: se generan desde FACTURACIÓN y CARTERA.")
        self.dv_list(k, ING["METODO"], "=lstMetodos")
        ws.data_validation(ING["VALOR"], {"validate": "decimal", "criteria": ">", "value": 0})
        ini, fin = self.resumen_periodo(k, 4)
        V, B, D, Fm, N = (rng("ING", x) for x in ("VALOR", "FECHA", "TIPO", "METODO", "ESTADO"))
        per = f'{B},">="&{ini},{B},"<"&{fin},{N},"ACTIVO"'
        met = lambda i: f"{q('CFG')}!$E${CFG_LIST_FIRST + i}"
        rows = [("Ventas de contado (efectivo)", f'=SUMIFS({V},{per},{D},"VENTA",{Fm},{met(0)})'),
                ("Ventas por transferencia", f'=SUMIFS({V},{per},{D},"VENTA",{Fm},{met(1)})'),
                ("Ventas por Nequi", f'=SUMIFS({V},{per},{D},"VENTA",{Fm},{met(2)})'),
                ("Ventas por Daviplata", f'=SUMIFS({V},{per},{D},"VENTA",{Fm},{met(3)})'),
                ("Ventas por otros medios", f'=SUMIFS({V},{per},{D},"VENTA")-SUM(K6:K9)'),
                ("Abonos de clientes", f'=SUMIFS({V},{per},{D},"ABONO")'),
                ("Devoluciones a clientes", f'=SUMIFS({V},{per},{D},"DEVOLUCIÓN")'),
                ("Otros ingresos", f'=SUMIFS({V},{per})-SUMIFS({V},{per},{D},"VENTA")-SUMIFS({V},{per},{D},"ABONO")-SUMIFS({V},{per},{D},"DEVOLUCIÓN")-SUMIFS({V},{per},{D},"NOTA CRÉDITO")'),
                ("TOTAL DINERO RECIBIDO", "=SUM(K6:K13)")]
        for i, (t, fm) in enumerate(rows):
            r = 6 + i
            tot = i == len(rows) - 1
            ws.merge_range(f"H{r}:J{r}", t, F.lbl(border=1, border_color=LINE, bg_color="#DBEAFE" if tot else LIGHT))
            ws.write_formula(f"K{r}", fm, F.out("cop", bold=tot, bg_color="#DBEAFE" if tot else None))
        ws.set_column("K:K", 16)
        ws.write("B16", "REGISTRO DE INGRESOS", F(bold=True, font_size=12, font_color=BLUE))
        self.tabla_hdr(k, "ING", freeze=False)
        ws.set_column("C:C", 16)
        ws.set_column("F:F", 16)
        self.tabla_valores(k, "ING", self.sim.ing if self.sim else [])
        t = TABLAS["ING"]
        ws.conditional_format(f"A{t['first']}:O{t['first']+20000}", {"type": "formula", "criteria": f'=$N{t["first"]}="ANULADO"', "format": F(font_color="#94A3B8", italic=True)})
        self.proteger(k)

    # ------------------------------------------------------------ EGRESOS
    def hoja_egr(self):
        k, ws, F = "EGR", self.ws["EGR"], self.F
        self.banner(k, "EGRESOS", "Toda salida de dinero. Los pagos de compras y de cartera a proveedores se registran solos.")
        ws.write("B4", "REGISTRAR EGRESO / GASTO", F(bold=True, font_size=12, font_color=BLUE))
        self.form_labels(k, [("B5", "Fecha"), ("B6", "Concepto"), ("B7", "Categoría"), ("B8", "Tipo de gasto"),
                             ("E5", "Valor"), ("E6", "Método"), ("E7", "Responsable"), ("E8", "Comprobante"), ("B9", "Observaciones")])
        ws.write_formula(EGR["FECHA"], "=TODAY()", F.inp("date"))
        ws.write_blank(EGR["CONCEPTO"], None, F.inp())
        ws.write_blank(EGR["CATEGORIA"], None, F.inp())
        ws.write(EGR["TIPO"], "VARIABLE", F.inp())
        ws.write_blank(EGR["VALOR"], None, F.inp("cop"))
        ws.write(EGR["METODO"], "Efectivo", F.inp())
        ws.write_blank(EGR["RESP"], None, F.inp())
        ws.write_blank(EGR["COMPROB"], None, F.inp())
        ws.merge_range("C9:F9", "", F.inp())
        self.dv_list(k, EGR["CATEGORIA"], "=lstCatGasto")
        self.dv_list(k, EGR["TIPO"], "=lstTipoGasto")
        self.dv_list(k, EGR["METODO"], "=lstMetodos")
        self.dv_list(k, EGR["RESP"], "=lstUsuarios", "warning")
        ws.data_validation(EGR["VALOR"], {"validate": "decimal", "criteria": ">", "value": 0})
        ws.set_column("C:C", 30)
        ws.set_column("D:D", 17)
        ws.set_column("F:F", 14)
        ini, fin = self.resumen_periodo(k, 4)
        V, B, Ec, Tp, Ct = (rng("EGR", x) for x in ("VALOR", "FECHA", "ESTADO", "TIPO", "CATEGORIA"))
        per = f'{B},">="&{ini},{B},"<"&{fin},{Ec},"ACTIVO"'
        r = 6
        ws.merge_range(f"H{r-1}:K{r-1}", "POR TIPO", F.hdr())
        for i, tp in enumerate(["FIJO", "VARIABLE", "EXTRAORDINARIO"]):
            ws.merge_range(f"H{r+i}:J{r+i}", f"Gastos {tp.lower()}s", F.lbl(border=1, border_color=LINE, bg_color=LIGHT))
            ws.write_formula(f"K{r+i}", f'=SUMIFS({V},{per},{Tp},"{tp}")', F.out("cop"))
        ws.merge_range("H9:K9", "POR CATEGORÍA", F.hdr())
        for i in range(12):
            rr = 10 + i
            ws.merge_range(f"H{rr}:J{rr}", "", F.lbl(border=1, border_color=LINE, bg_color=LIGHT))
            ws.write_formula(f"H{rr}", f"={q('CFG')}!$K${CFG_LIST_FIRST+i}&\"\"", F.lbl(border=1, border_color=LINE, bg_color=LIGHT))
            ws.write_formula(f"K{rr}", f'=IF(H{rr}="","",SUMIFS({V},{per},{Ct},H{rr}))', F.out("cop"))
        ws.merge_range("H22:J22", "TOTAL EGRESOS", F.lbl(border=1, border_color=LINE, bg_color="#DBEAFE"))
        ws.write_formula("K22", f"=SUMIFS({V},{per})", F.out("cop", bold=True, bg_color="#DBEAFE"))
        ws.write("L22", "(incluye compras de mercancía)", F(italic=True, font_color=GREY, font_size=8))
        ws.set_column("K:K", 16)
        ws.write("B23", "REGISTRO DE EGRESOS", F(bold=True, font_size=12, font_color=BLUE))
        self.tabla_hdr(k, "EGR", freeze=False)
        self.tabla_valores(k, "EGR", self.sim.egr if self.sim else [])
        t = TABLAS["EGR"]
        ws.conditional_format(f"A{t['first']}:P{t['first']+20000}", {"type": "formula", "criteria": f'=$P{t["first"]}="ANULADO"', "format": F(font_color="#94A3B8", italic=True)})
        self.proteger(k)

    # ------------------------------------------------------------ GASTOS FIJOS
    def hoja_gfi(self):
        k, ws, F = "GFI", self.ws["GFI"], self.F
        self.banner(k, "GASTOS FIJOS RECURRENTES", "Defina sus gastos fijos mensuales. Seleccione una fila y pulse PAGAR GASTO SELECCIONADO para registrar el egreso del mes.")
        self.tabla_hdr(k, "GFI")
        C = lambda key, r: cell("GFI", key, r)
        hoy_ini = "DATE(YEAR(TODAY()),MONTH(TODAY()),1)"
        fx = {
            "VENCE": lambda r: f'=IF({C("CONCEPTO",r)}="","",DATE(YEAR(TODAY()),MONTH(TODAY()),MIN(MAX(1,N({C("DIA",r)})),DAY(EOMONTH(TODAY(),0)))))',
            "PAGADO": lambda r: f'=IF({C("CONCEPTO",r)}="","",SUMIFS({rng("EGR","VALOR")},{rng("EGR","DOC")},"GF|"&{C("CONCEPTO",r)},{rng("EGR","ESTADO")},"ACTIVO",{rng("EGR","FECHA")},">="&{hoy_ini},{rng("EGR","FECHA")},"<="&EOMONTH(TODAY(),0)))',
            "ESTADO": lambda r: f'=IF({C("CONCEPTO",r)}="","",IF({C("ACTIVO",r)}="NO","INACTIVO",IF({C("PAGADO",r)}>=N({C("VALOR",r)}),"PAGADO",IF(TODAY()>{C("VENCE",r)},"VENCIDO",IF({C("VENCE",r)}-TODAY()<=5,"PRÓXIMO A VENCER","PENDIENTE")))))',
            "ULTIMO": lambda r: (f'=IF({C("CONCEPTO",r)}="","",IF(COUNTIFS({rng("EGR","DOC")},"GF|"&{C("CONCEPTO",r)},{rng("EGR","ESTADO")},"ACTIVO")=0,"",'
                                 f'SUMPRODUCT(MAX(({q("EGR")}!$M$25:$M$30024="GF|"&{C("CONCEPTO",r)})*({q("EGR")}!$P$25:$P$30024="ACTIVO")*{q("EGR")}!$B$25:$B$30024))))'),
        }
        rows = [list(g) for g in DM.GASTOS_FIJOS] if self.sim else []
        self.tabla_maestra(k, "GFI", rows, fx)
        t = TABLAS["GFI"]
        a, b = t["first"], t["last"]
        ws.data_validation(f"A{a}:A{b}", {"validate": "custom", "value": f"=COUNTIF($A${a}:$A${b},A{a})=1", "error_message": "Concepto repetido."})
        self.dv_list(k, f"B{a}:B{b}", "=lstCatGasto")
        self.dv_list(k, f"F{a}:F{b}", "=lstMetodos")
        self.dv_list(k, f"G{a}:G{b}", "=lstSiNo")
        self.dv_list(k, f"E{a}:E{b}", "=lstUsuarios", "warning")
        ws.data_validation(f"D{a}:D{b}", {"validate": "integer", "criteria": "between", "minimum": 1, "maximum": 31})
        e = f"J{a}:J{b}"
        for v, fm in (("PAGADO", F(font_color=GREEN, bold=True)), ("VENCIDO", F(bg_color="#FEE2E2", font_color=RED, bold=True)),
                      ("PRÓXIMO A VENCER", F(bg_color="#FEF3C7", font_color=AMBER, bold=True))):
            ws.conditional_format(e, {"type": "cell", "criteria": "==", "value": f'"{v}"', "format": fm})
        for addr, lab, fm in (("A4", "Total fijos mensual", f'=SUMIFS($C${a}:$C${b},$G${a}:$G${b},"<>NO")'),
                              ("C4", "Pagado este mes", f"=SUM($I${a}:$I${b})"),
                              ("E4", "Pendiente este mes", f'=SUMIFS($C${a}:$C${b},$J${a}:$J${b},"<>PAGADO",$J${a}:$J${b},"<>INACTIVO")')):
            ws.write(addr, lab, F(font_color=GREY, bg_color="#EEF2FF"))
            ws.write_formula(f"{chr(ord(addr[0])+1)}4", fm, F(bold=True, font_size=12, bg_color="#EEF2FF", nf="cop"))
        self.proteger(k)

    # ------------------------------------------------------------ CARTERA
    def hoja_car(self):
        k, ws, F = "CAR", self.ws["CAR"], self.F
        self.banner(k, "CARTERA", "Cuentas por cobrar (clientes) y por pagar (proveedores). Saldos, mora y estado se calculan solos.")
        ws.write("B4", "REGISTRAR ABONO", F(bold=True, font_size=12, font_color=BLUE))
        self.form_labels(k, [("B5", "Obligación"), ("B6", "Fecha del abono"), ("B7", "Valor del abono"),
                             ("B8", "Método de pago"), ("E6", "Saldo actual"), ("E7", "Observación")])
        ws.merge_range("C5:F5", "", F.inp(bold=True))
        self.dv_list(k, "C5", "=lstPendientes", msg="Seleccione una obligación con saldo pendiente.")
        ws.write_formula(CAR["FECHA"], "=TODAY()", F.inp("date"))
        ws.write_blank(CAR["VALOR"], None, F.inp("cop"))
        ws.write(CAR["METODO"], "Efectivo", F.inp())
        self.dv_list(k, CAR["METODO"], "=lstMetodos")
        t = TABLAS["CAR"]
        a, b = t["first"], t["last"]
        ws.write_formula(CAR["SALDO"], f'=IFERROR(INDEX($J${a}:$J${b},MATCH(LEFT($C$5,FIND(" |",$C$5)-1),$A${a}:$A${b},0)),"")', F.out("cop", bold=True, bg_color="#DBEAFE"))
        ws.write_blank(CAR["OBS"], None, F.inp())
        ws.set_column("F:F", 14)
        # resumen
        J, Bt, M = f"$J${a}:$J${b}", f"$B${a}:$B${b}", f"$M${a}:$M${b}"
        res = [("Por cobrar — saldo total", f'=SUMIFS({J},{Bt},"POR COBRAR")', "cop"),
               ("Por cobrar — vencida", f'=SUMIFS({J},{Bt},"POR COBRAR",{M},"VENCIDA")', "cop"),
               ("Obligaciones vencidas (cobrar)", f'=COUNTIFS({Bt},"POR COBRAR",{M},"VENCIDA")', "int"),
               ("Por pagar — saldo total", f'=SUMIFS({J},{Bt},"POR PAGAR")', "cop"),
               ("Por pagar — vencida", f'=SUMIFS({J},{Bt},"POR PAGAR",{M},"VENCIDA")', "cop"),
               ("Obligaciones vencidas (pagar)", f'=COUNTIFS({Bt},"POR PAGAR",{M},"VENCIDA")', "int")]
        ws.merge_range("H4:K4", "RESUMEN DE CARTERA", F.hdr())
        for i, (lab, fm, nf) in enumerate(res):
            r = 5 + i
            ws.merge_range(f"H{r}:J{r}", lab, F.lbl(border=1, border_color=LINE, bg_color=LIGHT))
            ws.write_formula(f"K{r}", fm, F.out(nf, bold=True, font_color=RED if "vencid" in lab.lower() else None))
        ws.set_column("K:K", 16)
        ws.write("B13", "OBLIGACIONES", F(bold=True, font_size=12, font_color=BLUE))
        self.tabla_hdr(k, "CAR", freeze=False)
        C = lambda key, r: cell("CAR", key, r)
        iv, id_, it, ie = rng("ING", "VALOR"), rng("ING", "DOC"), rng("ING", "TIPO"), rng("ING", "ESTADO")
        ev, ed, eo, ee = rng("EGR", "VALOR"), rng("EGR", "DOC"), rng("EGR", "ORIGEN"), rng("EGR", "ESTADO")
        fx = {
            "ABONOS": lambda r: (f'=IF({C("ID",r)}="","",N({C("PAGOINI",r)})+IF({C("TIPO",r)}="POR COBRAR",'
                                 f'SUMIFS({iv},{id_},{C("DOC",r)},{it},"ABONO",{ie},"ACTIVO")+SUMIFS({iv},{id_},{C("DOC",r)},{it},"NOTA CRÉDITO",{ie},"ACTIVO"),'
                                 f'SUMIFS({ev},{ed},{C("DOC",r)},{eo},"CARTERA",{ee},"ACTIVO")))'),
            "SALDO": lambda r: f'=IF({C("ID",r)}="","",IF({C("REGESTADO",r)}="ANULADA",0,MAX(0,{C("VALORINI",r)}-{C("ABONOS",r)})))',
            "MORA": lambda r: f'=IF({C("ID",r)}="","",IF(AND({C("SALDO",r)}>0,TODAY()>{C("VENCE",r)}),TODAY()-{C("VENCE",r)},0))',
            "ESTADO": lambda r: f'=IF({C("ID",r)}="","",IF({C("REGESTADO",r)}="ANULADA","ANULADA",IF({C("SALDO",r)}<=0,"PAGADA",IF({C("MORA",r)}>0,"VENCIDA",IF({C("ABONOS",r)}>0,"ABONO PARCIAL","PENDIENTE")))))',
            "COMBO": lambda r: f'=IF({C("ID",r)}="","",{C("ID",r)}&" | "&{C("DOC",r)}&" | "&{C("TERCERO",r)}&" | Saldo $"&FIXED(N({C("SALDO",r)}),0))',
            "IDX": lambda r: f'=IF({C("ID",r)}="","",IF(N({C("SALDO",r)})>0,COUNTIF($J${a}:J{r},">0"),""))',
            "LISTA": lambda r: f'=IFERROR(INDEX($P${a}:$P${b},MATCH(ROW()-{a-1},$Q${a}:$Q${b},0)),"")',
        }
        rows = []
        if self.sim:
            for c in self.sim.car:
                # ID,TIPO,CODTER,TERCERO,DOC,FECHA,VALORINI,PAGOINI,(ABONOS),(SALDO),VENCE,(MORA),(ESTADO),REG,OBS
                rows.append([c[0], c[1], c[2], c[3], c[4], c[5], c[6], c[7], None, None, c[8], None, None, c[9], c[10]])
        self.tabla_maestra(k, "CAR", rows, fx)
        e = f"M{a}:M{b}"
        for v, fm in (("VENCIDA", F(bg_color="#FEE2E2", font_color=RED, bold=True)), ("PAGADA", F(font_color=GREEN, bold=True)),
                      ("ABONO PARCIAL", F(font_color=AMBER, bold=True)), ("ANULADA", F(font_color="#94A3B8", italic=True))):
            ws.conditional_format(e, {"type": "cell", "criteria": "==", "value": f'"{v}"', "format": fm})
        ws.conditional_format(f"L{a}:L{b}", {"type": "cell", "criteria": ">", "value": 0, "format": F(font_color=RED, bold=True)})
        self.proteger(k)

    # ------------------------------------------------------------ CAJA
    def hoja_caj(self):
        k, ws, F = "CAJ", self.ws["CAJ"], self.F
        self.banner(k, "CAJA Y BANCOS", "Saldo disponible por cuenta = saldo inicial + entradas − salidas ± traslados. Cierres diarios, semanales y mensuales.")
        ws.set_column("A:A", 3)
        ws.set_column("B:B", 16)
        ws.set_column("C:H", 15)
        ws.set_column("I:N", 13)
        ws.set_column("O:Q", 3)
        ws.set_column("R:R", 11)
        ws.set_column("S:T", 12)
        ws.set_column("U:U", 14)
        ws.set_column("V:V", 26)
        ws.set_column("W:W", 13)
        ws.write("B3", "Saldos iniciales al", F.lbl(align="right"))
        ws.write_datetime(CAJ["FINI"], todt(DM.FECHA_SALDOS), F.inp("date"))
        heads = ["Cuenta", "Saldo inicial", "Entradas", "Salidas", "Traslados (+)", "Traslados (−)", "SALDO DISPONIBLE"]
        for i, h in enumerate(heads):
            ws.write(4, 1 + i, h, F.hdr())
        IV, IG, IN, IB = rng("ING", "VALOR"), rng("ING", "CUENTA"), rng("ING", "ESTADO"), rng("ING", "FECHA")
        EV, EH, EP, EB = rng("EGR", "VALOR"), rng("EGR", "CUENTA"), rng("EGR", "ESTADO"), rng("EGR", "FECHA")
        TR = lambda c: f"${c}$6:${c}$205"
        for i in range(5):
            r = CAJ["CTA1"] + i
            ws.write_formula(f"B{r}", f"={q('CFG')}!$G${CFG_LIST_FIRST+i}&\"\"", F.lbl(border=1, border_color=LINE, bg_color=LIGHT))
            ws.write(f"C{r}", DM.SALDOS_INI[DM.CFG_LISTS["lstCuentas"][2][i]] if self.sim else 0, F.inp("cop"))
            ws.write_formula(f"D{r}", f'=SUMIFS({IV},{IG},$B{r},{IN},"ACTIVO",{IB},">="&$C$3)', F.out("cop"))
            ws.write_formula(f"E{r}", f'=SUMIFS({EV},{EH},$B{r},{EP},"ACTIVO",{EB},">="&$C$3)', F.out("cop"))
            ws.write_formula(f"F{r}", f'=SUMIFS({TR("U")},{TR("T")},$B{r},{TR("R")},">="&$C$3)', F.out("cop"))
            ws.write_formula(f"G{r}", f'=SUMIFS({TR("U")},{TR("S")},$B{r},{TR("R")},">="&$C$3)', F.out("cop"))
            ws.write_formula(f"H{r}", f"=C{r}+D{r}-E{r}+F{r}-G{r}", F.out("cop", bold=True, bg_color="#DCFCE7"))
        ws.write("B11", "TOTAL", F.hdr())
        for c in "CDEFGH":
            ws.write_formula(f"{c}11", f"=SUM({c}6:{c}10)", F.out("cop", bold=True, bg_color="#DBEAFE"))
        ws.conditional_format("H6:H11", {"type": "cell", "criteria": "<", "value": 0, "format": F(bg_color="#FEE2E2", font_color=RED)})
        # cierre
        ws.merge_range("B13:G13", "CIERRE DE CAJA (use los botones DIARIO / SEMANAL / MENSUAL o escriba las fechas)", F.hdr(align="left", bg_color=BLUE))
        ws.write("B14", "Tipo de cierre", F.lbl(align="right"))
        ws.write(CAJ["TIPO"], "DIARIO", F.inp(bold=True))
        self.dv_list(k, CAJ["TIPO"], '"DIARIO,SEMANAL,MENSUAL,PERSONALIZADO"')
        ws.write("B15", "Desde", F.lbl(align="right"))
        ws.write_formula(CAJ["DESDE"], "=TODAY()", F.inp("date"))
        ws.write("B16", "Hasta", F.lbl(align="right"))
        ws.write_formula(CAJ["HASTA"], "=TODAY()", F.inp("date"))
        for i, h in enumerate(["Cuenta", "Saldo inicial período", "Entradas", "Salidas", "Traslados netos", "Saldo final"]):
            ws.write(17, 1 + i, h, F.hdr())
        d0, d1 = "$C$15", "$C$16"
        for i in range(5):
            r = CAJ["PER1"] + i
            rc = CAJ["CTA1"] + i
            ws.write_formula(f"B{r}", f"=B{rc}", F.lbl(border=1, border_color=LINE, bg_color=LIGHT))
            prev = (f'C{rc}+SUMIFS({IV},{IG},$B{r},{IN},"ACTIVO",{IB},">="&$C$3,{IB},"<"&{d0})'
                    f'-SUMIFS({EV},{EH},$B{r},{EP},"ACTIVO",{EB},">="&$C$3,{EB},"<"&{d0})'
                    f'+SUMIFS({TR("U")},{TR("T")},$B{r},{TR("R")},">="&$C$3,{TR("R")},"<"&{d0})'
                    f'-SUMIFS({TR("U")},{TR("S")},$B{r},{TR("R")},">="&$C$3,{TR("R")},"<"&{d0})')
            ws.write_formula(f"C{r}", "=" + prev, F.out("cop"))
            ws.write_formula(f"D{r}", f'=SUMIFS({IV},{IG},$B{r},{IN},"ACTIVO",{IB},">="&{d0},{IB},"<"&{d1}+1)', F.out("cop"))
            ws.write_formula(f"E{r}", f'=SUMIFS({EV},{EH},$B{r},{EP},"ACTIVO",{EB},">="&{d0},{EB},"<"&{d1}+1)', F.out("cop"))
            ws.write_formula(f"F{r}", f'=SUMIFS({TR("U")},{TR("T")},$B{r},{TR("R")},">="&{d0},{TR("R")},"<"&{d1}+1)-SUMIFS({TR("U")},{TR("S")},$B{r},{TR("R")},">="&{d0},{TR("R")},"<"&{d1}+1)', F.out("cop"))
            ws.write_formula(f"G{r}", f"=C{r}+D{r}-E{r}+F{r}", F.out("cop", bold=True, bg_color="#DCFCE7"))
        ws.write(f"B{CAJ['PERTOT']}", "TOTAL", F.hdr())
        for c in "CDEFG":
            ws.write_formula(f"{c}{CAJ['PERTOT']}", f"=SUM({c}{CAJ['PER1']}:{c}{CAJ['PERN']})", F.out("cop", bold=True, bg_color="#DBEAFE"))
        ws.write("B26", "HISTORIAL DE CIERRES", F(bold=True, font_size=12, font_color=BLUE))
        self.tabla_hdr(k, "CIE", freeze=False, filtro=False)
        for i in range(5):
            ws.write_formula(26, 8 + i, f'="Saldo "&B{CAJ["CTA1"]+i}', F.hdr())
        # traslados
        ws.merge_range("R4:W4", "TRASLADOS ENTRE CUENTAS (ej. consignar efectivo al banco)", F.hdr(bg_color=BLUE))
        for i, h in enumerate(["Fecha", "De cuenta", "A cuenta", "Valor", "Observación", "Usuario"]):
            ws.write(4, 17 + i, h, F.hdr())
        trs = self.sim.traslados if self.sim else []
        for n in range(200):
            r = 5 + n
            row = trs[n] if n < len(trs) else [None] * 6
            for i, nf in enumerate(["date", None, None, "cop", None, None]):
                v = row[i]
                fm = F.inp(nf)
                if v is None:
                    ws.write_blank(r, 17 + i, None, fm)
                elif nf == "date":
                    ws.write_datetime(r, 17 + i, todt(v), fm)
                elif isinstance(v, str):
                    ws.write_string(r, 17 + i, v, fm)
                else:
                    ws.write_number(r, 17 + i, v, fm)
        self.dv_list(k, "S6:T205", "=lstCuentas")
        self.dv_list(k, "W6:W205", "=lstUsuarios", "warning")
        ws.data_validation("U6:U205", {"validate": "decimal", "criteria": ">", "value": 0})
        self.proteger(k)

    # ------------------------------------------------------------ UTILIDADES
    def hoja_uti(self):
        k, ws, F = "UTI", self.ws["UTI"], self.F
        self.banner(k, "UTILIDADES", "Ventas facturadas ≠ dinero recibido ≠ ganancia. Aquí se calcula la ganancia REAL: ventas − costo de lo vendido − gastos operativos.")
        ws.set_column("A:A", 3)
        ws.set_column("B:B", 22)
        ws.set_column("C:K", 15)
        ws.write("B3", "Año", F.lbl(align="right"))
        ws.write("C3", DM.HOY.year, F.inp(bold=True))
        heads = ["Mes", "Ventas facturadas", "Costo mercancía vendida", "Utilidad bruta", "Margen bruto",
                 "Gastos operativos", "UTILIDAD NETA", "Margen neto", "Dinero recibido (ventas)", "Utilidad acumulada"]
        for i, h in enumerate(heads):
            ws.write(5, 1 + i, h, F.hdr())
        BS, BC, BF, BE = rng("BD", "SUBTOTAL"), rng("BD", "COSTOTOTAL"), rng("BD", "FECHA"), rng("BD", "ESTADO")
        EV, EB, EP, ED = rng("EGR", "VALOR"), rng("EGR", "FECHA"), rng("EGR", "ESTADO"), rng("EGR", "CATEGORIA")
        IV, IB, IN, IT = rng("ING", "VALOR"), rng("ING", "FECHA"), rng("ING", "ESTADO"), rng("ING", "TIPO")
        meses = DM.CFG_LISTS["lstMeses"][2][:12]
        for m in range(1, 13):
            r = 6 + m
            s, e = f"DATE($C$3,{m},1)", f"DATE($C$3,{m + 1},1)"
            bd = f'{BF},">="&{s},{BF},"<"&{e},{BE},"ACTIVA"'
            eg = f'{EB},">="&{s},{EB},"<"&{e},{EP},"ACTIVO"'
            ig = f'{IB},">="&{s},{IB},"<"&{e},{IN},"ACTIVO"'
            ws.write(f"B{r}", meses[m - 1], F.lbl(border=1, border_color=LINE, bg_color=LIGHT))
            ws.write_formula(f"C{r}", f"=SUMIFS({BS},{bd})", F.out("cop"))
            ws.write_formula(f"D{r}", f"=SUMIFS({BC},{bd})", F.out("cop"))
            ws.write_formula(f"E{r}", f"=C{r}-D{r}", F.out("cop"))
            ws.write_formula(f"F{r}", f"=IFERROR(E{r}/C{r},0)", F.out("pct"))
            ws.write_formula(f"G{r}", f"=SUMIFS({EV},{eg})-SUMIFS({EV},{eg},{ED},cfgCatCompras)", F.out("cop"))
            ws.write_formula(f"H{r}", f"=E{r}-G{r}", F.out("cop", bold=True))
            ws.write_formula(f"I{r}", f"=IFERROR(H{r}/C{r},0)", F.out("pct"))
            ws.write_formula(f"J{r}", f'=SUMIFS({IV},{ig},{IT},"VENTA")+SUMIFS({IV},{ig},{IT},"ABONO")+SUMIFS({IV},{ig},{IT},"DEVOLUCIÓN")', F.out("cop"))
            ws.write_formula(f"K{r}", f"=SUM($H$7:H{r})", F.out("cop"))
        ws.write("B19", "TOTAL AÑO", F.hdr())
        for c in "CDEGHJ":
            ws.write_formula(f"{c}19", f"=SUM({c}7:{c}18)", F.out("cop", bold=True, bg_color="#DBEAFE"))
        ws.write_formula("F19", "=IFERROR(E19/C19,0)", F.out("pct", bold=True, bg_color="#DBEAFE"))
        ws.write_formula("I19", "=IFERROR(H19/C19,0)", F.out("pct", bold=True, bg_color="#DBEAFE"))
        ws.write_formula("K19", "=H19", F.out("cop", bold=True, bg_color="#DBEAFE"))
        ws.conditional_format("H7:H19", {"type": "cell", "criteria": "<", "value": 0, "format": F(font_color=RED)})
        cards = [("M5", "Ventas facturadas (año)", "=C19"), ("M8", "Dinero recibido de ventas (año)", "=J19"),
                 ("M11", "UTILIDAD NETA REAL (año)", "=H19"), ("M14", "Cartera por cobrar (hoy)", f"=SUMIFS({rng('CAR','SALDO')},{rng('CAR','TIPO')},\"POR COBRAR\")")]
        ws.set_column("M:O", 14)
        for addr, lab, fm in cards:
            r = int(addr[1:])
            ws.merge_range(f"M{r}:O{r}", lab, F(font_color=GREY, bg_color="#EEF2FF", font_size=9))
            ws.merge_range(f"M{r+1}:O{r+1}", "", F(bold=True, font_size=15, bg_color="#EEF2FF", nf="cop"))
            ws.write_formula(f"M{r+1}", fm, F(bold=True, font_size=15, bg_color="#EEF2FF", nf="cop"))
        ws.write("M17", "Ventas facturadas: lo vendido (aunque no se haya cobrado).", F(italic=True, font_color=GREY, font_size=8))
        ws.write("M18", "Dinero recibido: lo que efectivamente entró a caja.", F(italic=True, font_color=GREY, font_size=8))
        ws.write("M19", "Utilidad neta: lo que realmente se gana.", F(italic=True, font_color=GREY, font_size=8))
        # por categoría
        ws.write("B21", "RENTABILIDAD POR CATEGORÍA (año seleccionado)", F(bold=True, font_size=12, font_color=BLUE))
        for i, h in enumerate(["Categoría", "Unidades", "Ventas", "Costo", "Utilidad bruta", "Margen"]):
            ws.write(21, 1 + i, h, F.hdr(bg_color=TEAL))
        BK, BH = rng("BD", "CANT"), rng("BD", "CATEGORIA")
        yr = f'{BF},">="&DATE($C$3,1,1),{BF},"<"&DATE($C$3+1,1,1),{BE},"ACTIVA"'
        for i in range(20):
            r = 23 + i
            ws.write_formula(f"B{r}", f"={q('CFG')}!$D${CFG_LIST_FIRST+i}&\"\"", F.out())
            ws.write_formula(f"C{r}", f'=IF(B{r}="","",SUMIFS({BK},{yr},{BH},B{r}))', F.out("int"))
            ws.write_formula(f"D{r}", f'=IF(B{r}="","",SUMIFS({BS},{yr},{BH},B{r}))', F.out("cop"))
            ws.write_formula(f"E{r}", f'=IF(B{r}="","",SUMIFS({BC},{yr},{BH},B{r}))', F.out("cop"))
            ws.write_formula(f"F{r}", f'=IF(B{r}="","",D{r}-E{r})', F.out("cop"))
            ws.write_formula(f"G{r}", f'=IF(B{r}="","",IFERROR(F{r}/D{r},0))', F.out("pct"))
        ws.conditional_format("F23:F42", {"type": "data_bar", "bar_color": "#60A5FA", "bar_solid": True})
        # por producto
        ws.write("B44", "RENTABILIDAD POR PRODUCTO (año seleccionado)", F(bold=True, font_size=12, font_color=BLUE))
        for i, h in enumerate(["SKU", "Producto", "Categoría", "Unidades", "Ventas", "Costo", "Utilidad bruta", "Margen"]):
            ws.write(44, 1 + i, h, F.hdr(bg_color=TEAL))
        BSKU = rng("BD", "SKU")
        I = lambda c, i: f"{q('INV')}!{c}{i}"
        fo = {n: F.out(n if n != "txt" else None, border_color="#E2E8F0") for n in ("txt", "int", "cop", "pct")}
        for n in range(1000):
            r = 46 + n
            i = 9 + n
            ws.write_formula(f"B{r}", f'=IF({I("A",i)}="","",{I("A",i)})', fo["txt"])
            ws.write_formula(f"C{r}", f'=IF(B{r}="","",{I("B",i)}&" "&{I("E",i)}&" "&{I("F",i)})', fo["txt"])
            ws.write_formula(f"D{r}", f'=IF(B{r}="","",{I("C",i)})', fo["txt"])
            ws.write_formula(f"E{r}", f'=IF(B{r}="","",SUMIFS({BK},{yr},{BSKU},B{r}))', fo["int"])
            ws.write_formula(f"F{r}", f'=IF(B{r}="","",SUMIFS({BS},{yr},{BSKU},B{r}))', fo["cop"])
            ws.write_formula(f"G{r}", f'=IF(B{r}="","",SUMIFS({BC},{yr},{BSKU},B{r}))', fo["cop"])
            ws.write_formula(f"H{r}", f'=IF(B{r}="","",F{r}-G{r})', fo["cop"])
            ws.write_formula(f"I{r}", f'=IF(B{r}="","",IFERROR(H{r}/F{r},0))', fo["pct"])
        ws.autofilter(44, 1, 1045, 8)
        ws.conditional_format("H46:H1045", {"type": "data_bar", "bar_color": "#34D399", "bar_solid": True})
        self.proteger(k)

    # ------------------------------------------------------------ REPORTES
    def hoja_rep(self):
        k, ws, F = "REP", self.ws["REP"], self.F
        self.banner(k, "REPORTES", "Elija el reporte y el período, pulse GENERAR. Luego exporte a PDF / Excel o imprima.")
        ws.set_column("A:A", 3)
        ws.set_column("B:P", 15)
        self.form_labels(k, [("B4", "Tipo de reporte"), ("B5", "Fecha desde"), ("B6", "Fecha hasta")])
        ws.merge_range("C4:D4", "Ventas mensual", F.inp(bold=True))
        self.dv_list(k, REP["TIPO"], "=lstReportes")
        ws.write_formula(REP["DESDE"], "=DATE(YEAR(TODAY()),MONTH(TODAY()),1)", F.inp("date"))
        ws.write_formula(REP["HASTA"], "=TODAY()", F.inp("date"))
        ws.write("E5", "Diario: usa 'desde'. Mensual: mes de 'desde'. Anual: año de 'desde'. Otros: el rango.", F(italic=True, font_color=GREY, font_size=8))
        ws.write("B8", "El reporte aparecerá aquí abajo ↓", F(italic=True, font_color=GREY))
        ws.set_paper(1)
        ws.set_landscape()
        ws.fit_to_pages(1, 0)
        ws.repeat_rows(REP["OUT"] - 1, REP["OUT"] + 2)
        self.proteger(k)

    # ------------------------------------------------------------ CALC (oculta)
    def hoja_calc(self):
        k, ws, F = "CALC", self.ws["CALC"], self.F
        ws.write("A1", "HOJA TÉCNICA DE CÁLCULO (no modificar)")
        Ii = lambda a: f"{q('INI')}!{a}"
        ws.write("A2", "Año")
        ws.write_formula("B2", f"=N({Ii(INI['ANIO'])})")
        ws.write("A3", "Mes idx")
        ws.write_formula("B3", f'=IFERROR(MATCH({Ii(INI["MES"])},lstMeses,0),13)')
        ws.write_formula("B3", f'=IF(IFERROR(MATCH({Ii(INI["MES"])},lstMeses,0),13)>12,0,MATCH({Ii(INI["MES"])},lstMeses,0))')
        ws.write("A4", "Categoría crit")
        ws.write_formula("B4", f'=IF(OR({Ii(INI["CAT"])}="",{Ii(INI["CAT"])}="TODAS"),"*",{Ii(INI["CAT"])})')
        ws.write("A5", "Inicio")
        ws.write_formula("B5", "=IF(B3=0,DATE(B2,1,1),DATE(B2,B3,1))", F(nf="date"))
        ws.write("A6", "Fin (excl.)")
        ws.write_formula("B6", "=IF(B3=0,DATE(B2+1,1,1),DATE(B2,B3+1,1))", F(nf="date"))
        ws.write("A7", "Día")
        ws.write_formula("B7", f"=N({Ii(INI['DIA'])})", F(nf="date"))
        ws.write("A8", "Etiqueta")
        ws.write_formula("B8", f'=IF(B3=0,"año "&B2,{Ii(INI["MES"])}&" "&B2)')
        BS, BR, BF, BE, BH, BK, BSKU = (rng("BD", x) for x in ("SUBTOTAL", "UTILIDAD", "FECHA", "ESTADO", "CATEGORIA", "CANT", "SKU"))
        EV, EB, EP, ED = rng("EGR", "VALOR"), rng("EGR", "FECHA"), rng("EGR", "ESTADO"), rng("EGR", "CATEGORIA")
        IV, IB, IN, IT, IM = rng("ING", "VALOR"), rng("ING", "FECHA"), rng("ING", "ESTADO"), rng("ING", "TIPO"), rng("ING", "METODO")
        per = f'{BF},">="&$B$5,{BF},"<"&$B$6,{BE},"ACTIVA"'
        eper = f'{EB},">="&$B$5,{EB},"<"&$B$6,{EP},"ACTIVO"'
        kpis = [("D2", "Gastos op. período", f"=SUMIFS({EV},{eper})-SUMIFS({EV},{eper},{ED},cfgCatCompras)"),
                ("D3", "Ventas período (todas cat.)", f"=SUMIFS({BS},{per})"),
                ("D4", "Utilidad bruta período (todas)", f"=SUMIFS({BR},{per})"),
                ("D5", "Ventas período (cat.)", f"=SUMIFS({BS},{per},{BH},$B$4)"),
                ("D6", "Utilidad bruta período (cat.)", f"=SUMIFS({BR},{per},{BH},$B$4)"),
                ("D7", "Ventas del día (cat.)", f'=SUMIFS({BS},{BF},$B$7,{BE},"ACTIVA",{BH},$B$4)'),
                ("D8", "Ventas año (cat.)", f'=SUMIFS({BS},{BF},">="&DATE($B$2,1,1),{BF},"<"&DATE($B$2+1,1,1),{BE},"ACTIVA",{BH},$B$4)'),
                ("D9", "N° facturas período", f'=COUNTIFS({rng("HFA","FECHA")},">="&$B$5,{rng("HFA","FECHA")},"<"&$B$6,{rng("HFA","ESTADO")},"ACTIVA")'),
                ("D10", "Utilidad neta período", "=M4-M2")]
        for addr, lab, fm in kpis:
            ws.write("L" + addr[1:], lab)
            ws.write_formula("M" + addr[1:], fm)
        # mensual
        ws.write_row("A10", ["Mes", "Etiqueta", "Ventas", "Costo", "Util. bruta", "Gastos op.", "Util. neta", "Ingresos", "Egresos", "Valor inventario"])
        meses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
        MV, MF = rng("MOV", "VALOR"), rng("MOV", "FECHA")
        ti = TABLAS["INV"]
        vini = f"SUMPRODUCT({q('INV')}!$L${ti['first']}:$L${ti['last']},{q('INV')}!$J${ti['first']}:$J${ti['last']})"
        for m in range(1, 13):
            r = 10 + m
            s, e = f"DATE($B$2,{m},1)", f"DATE($B$2,{m+1},1)"
            bd = f'{BF},">="&{s},{BF},"<"&{e},{BE},"ACTIVA"'
            eg = f'{EB},">="&{s},{EB},"<"&{e},{EP},"ACTIVO"'
            ig = f'{IB},">="&{s},{IB},"<"&{e},{IN},"ACTIVO"'
            ws.write(f"A{r}", m)
            ws.write(f"B{r}", meses[m - 1])
            ws.write_formula(f"C{r}", f"=SUMIFS({BS},{bd},{BH},$B$4)")
            ws.write_formula(f"D{r}", f"=SUMIFS({rng('BD','COSTOTOTAL')},{bd},{BH},$B$4)")
            ws.write_formula(f"E{r}", f"=SUMIFS({BR},{bd})")
            ws.write_formula(f"F{r}", f"=SUMIFS({EV},{eg})-SUMIFS({EV},{eg},{ED},cfgCatCompras)")
            ws.write_formula(f"G{r}", f"=E{r}-F{r}")
            ws.write_formula(f"H{r}", f'=SUMIFS({IV},{ig})-SUMIFS({IV},{ig},{IT},"NOTA CRÉDITO")')
            ws.write_formula(f"I{r}", f"=SUMIFS({EV},{eg})")
            ws.write_formula(f"J{r}", f'=IF({s}>TODAY(),NA(),{vini}+SUMIFS({MV},{MF},"<"&{e}))')
        # categorías
        ws.write_row("A30", ["Categoría", "Ventas", "Score"])
        for i in range(20):
            r = 31 + i
            ws.write_formula(f"A{r}", f"={q('CFG')}!$D${CFG_LIST_FIRST+i}&\"\"")
            ws.write_formula(f"B{r}", f'=IF(A{r}="",0,SUMIFS({BS},{per},{BH},A{r}))')
            ws.write_formula(f"C{r}", f'=IF(A{r}="","",B{r}+ROW()/10000000)')
        ws.write_row("E30", ["Top categorías", "Ventas"])
        for kk in range(1, 9):
            r = 30 + kk
            mm = f"MATCH(LARGE($C$31:$C$50,{kk}),$C$31:$C$50,0)"
            ws.write_formula(f"E{r}", f'=IFERROR(INDEX($A$31:$A$50,{mm}),"")')
            ws.write_formula(f"F{r}", f'=IFERROR(INDEX($B$31:$B$50,{mm}),0)')
        # métodos de pago
        ws.write_row("A55", ["Método", "Recaudo período"])
        ig = f'{IB},">="&$B$5,{IB},"<"&$B$6,{IN},"ACTIVO"'
        for i in range(6):
            r = 56 + i
            ws.write_formula(f"A{r}", f"={q('CFG')}!$E${CFG_LIST_FIRST+i}&\"\"")
            ws.write_formula(f"B{r}", f'=IF(A{r}="",0,SUMIFS({IV},{ig},{IM},A{r},{IT},"VENTA")+SUMIFS({IV},{ig},{IM},A{r},{IT},"ABONO"))')
        # productos
        ws.write_row("A100", ["SKU", "Producto", "Activo", "Unid. período", "Score top", "Score bajo", "Idx alerta"])
        for n in range(1000):
            r = 101 + n
            i = 9 + n
            I = lambda c: f"{q('INV')}!{c}{i}"
            ws.write_formula(f"A{r}", f'=IF({I("A")}="","",{I("A")})')
            ws.write_formula(f"B{r}", f'=IF(A{r}="","",{I("B")}&" "&{I("E")}&" "&{I("F")})')
            ws.write_formula(f"C{r}", f'=IF(A{r}="","",{I("U")})')
            ws.write_formula(f"D{r}", f'=IF(A{r}="","",SUMIFS({BK},{BSKU},A{r},{per},{BH},$B$4))')
            ws.write_formula(f"E{r}", f'=IF(A{r}="","",D{r}+ROW()/10000000)')
            ws.write_formula(f"F{r}", f'=IF(OR(A{r}="",C{r}="NO"),"",D{r}+ROW()/10000000)')
            ws.write_formula(f"G{r}", f'=IF(A{r}="","",IF(OR({I("T")}="AGOTADO",{I("T")}="BAJO STOCK"),COUNTIF($G$100:G{r-1},">0")+1,""))')
        ws.write_row("A70", ["Rank", "Más vendidos", "Unidades"])
        ws.write_row("A77", ["Rank", "Menos vendidos", "Unidades"])
        for kk in range(1, 6):
            for base, fn, col in ((70, "LARGE", "E"), (77, "SMALL", "F")):
                r = base + kk
                mm = f"MATCH({fn}(${col}$101:${col}$1100,{kk}),${col}$101:${col}$1100,0)"
                ws.write(f"A{r}", kk)
                ws.write_formula(f"B{r}", f'=IFERROR(INDEX($B$101:$B$1100,{mm}),"-")')
                ws.write_formula(f"C{r}", f"=IFERROR(INDEX($D$101:$D$1100,{mm}),0)")
        ws.write_row("A84", ["#", "Alerta producto", "Stock", "Estado"])
        for kk in range(1, 11):
            r = 84 + kk
            mm = f"MATCH({kk},$G$101:$G$1100,0)"
            ws.write(f"A{r}", kk)
            ws.write_formula(f"B{r}", f'=IFERROR(INDEX($B$101:$B$1100,{mm}),"")')
            ws.write_formula(f"C{r}", f'=IFERROR(INDEX({q("INV")}!$O$9:$O$1008,{mm}),"")')
            ws.write_formula(f"D{r}", f'=IFERROR(INDEX({q("INV")}!$T$9:$T$1008,{mm}),"")')
        # lista de filtro de categorías
        ws.write("Y1", "TODAS")
        for i in range(CFG_LIST_N):
            ws.write_formula(f"Y{2+i}", f"={q('CFG')}!$D${CFG_LIST_FIRST+i}&\"\"")
        ws.hide()

    # ------------------------------------------------------------ INICIO / DASHBOARD
    def hoja_ini(self):
        k, ws, F = "INI", self.ws["INI"], self.F
        ws.hide_gridlines(2)
        ws.set_zoom(90)
        ws.set_column("A:A", 26, F(bg_color="#0F172A"))
        ws.set_column("B:B", 2)
        ws.set_column("C:Q", 10.5)
        ws.set_column("R:R", 2)
        ws.set_column("S:S", 30)
        ws.set_column("T:V", 11)
        ws.set_row(0, 38, F(bg_color=NAVY))
        ws.set_row(1, 18, F(bg_color=NAVY))
        ws.write("A1", "  ◆ SISTEMA EMPRESARIAL", F(bold=True, font_size=13, font_color="white", bg_color="#0F172A"))
        ws.write("C1", "PANEL GERENCIAL", F(bold=True, font_size=20, font_color="white", bg_color=NAVY))
        ws.write_formula("C2", '=cfgEmpresa&"  ·  NIT "&cfgNIT', F(font_color="#CBD5E1", bg_color=NAVY))
        menu = [("MÓDULOS", None), ("Facturación", "FAC"), ("Guías de envío", "GUI"), ("Inventario", "INV"),
                ("Compras", "COM"), ("Proveedores", "PRO"), ("Clientes", "CLI"), ("Costos y rentabilidad", "COS"),
                ("FINANZAS", None), ("Ingresos", "ING"), ("Egresos", "EGR"), ("Gastos fijos", "GFI"),
                ("Cartera", "CAR"), ("Caja y bancos", "CAJ"), ("Utilidades", "UTI"), ("Reportes", "REP"),
                ("SISTEMA", None), ("Configuración", "CFG"), ("Base de datos (ventas)", "BD"),
                ("Historial de facturas", "HFA"), ("Historial de guías", "HGU"), ("Movimientos inventario", "MOV"),
                ("Auditoría", "AUD")]
        for i, (t, sk) in enumerate(menu):
            r = 3 + i
            ws.set_row(r, 22)
            if sk is None:
                ws.write(r, 0, "  " + t, F(bold=True, font_size=8, font_color="#64748B", bg_color="#0F172A"))
            else:
                ws.write_url(r, 0, f"internal:'{SH[sk]}'!A1", F(font_color="white", bg_color="#1E293B", bold=True,
                                                              border=1, border_color="#0F172A", indent=1), string="▸ " + t)
        # filtros
        lab = F.lbl(align="right", font_size=9)
        for addr, t in (("C4", "Año"), ("E4", "Mes"), ("G4", "Categoría"), ("J4", "Día"), ("L4", "Usuario")):
            ws.write(addr, t, lab)
        ws.write(INI["ANIO"], DM.HOY.year, F.inp(bold=True, align="center"))
        ws.write(INI["MES"], DM.CFG_LISTS["lstMeses"][2][DM.HOY.month - 1], F.inp(bold=True))
        ws.merge_range("H4:I4", "TODAS", F.inp(bold=True))
        ws.write_formula(INI["DIA"], "=TODAY()", F.inp("date", bold=True))
        ws.merge_range("M4:N4", "Administrador", F.inp(bold=True))
        self.dv_list(k, INI["MES"], "=lstMeses")
        self.dv_list(k, INI["CAT"], "=lstCatFiltro")
        self.dv_list(k, INI["USUARIO"], "=lstUsuarios", msg="Seleccione el usuario que está operando (CONFIGURACIÓN > Usuarios).")
        ws.data_validation(INI["ANIO"], {"validate": "integer", "criteria": "between", "minimum": 2000, "maximum": 2100})
        Cc = lambda a: f"{q('CALC')}!{a}"
        BS = rng("CAR", "SALDO")
        BT = rng("CAR", "TIPO")
        ti = TABLAS["INV"]
        Ir = lambda c: f"{q('INV')}!${c}${ti['first']}:${c}${ti['last']}"
        cards = [
            ("Ventas del día", f"={Cc('$M$7')}", "cop", f'="Fecha "&DAY({Cc("$B$7")})&"/"&MONTH({Cc("$B$7")})&"/"&YEAR({Cc("$B$7")})', BLUE),
            ("Ventas del período", f"={Cc('$M$5')}", "cop", f'={Cc("$B$8")}', BLUE),
            ("Ventas acumuladas del año", f"={Cc('$M$8')}", "cop", f'="Año "&{Cc("$B$2")}', BLUE),
            ("Utilidad bruta", f"={Cc('$M$6')}", "cop", f'="Margen bruto "&FIXED(IFERROR({Cc("$M$6")}/{Cc("$M$5")},0)*100,1)&"%"', TEAL),
            ("Utilidad neta", f"={Cc('$M$10')}", "cop", '="Todas las categorías"', TEAL),
            ("Gastos operativos", f"={Cc('$M$2')}", "cop", '="Sin compras de mercancía"', RED),
            ("Cartera por cobrar", f'=SUMIFS({BS},{BT},"POR COBRAR")', "cop", f'="Vencida: $"&FIXED(SUMIFS({BS},{BT},"POR COBRAR",{rng("CAR","ESTADO")},"VENCIDA"),0)', AMBER),
            ("Cuentas por pagar", f'=SUMIFS({BS},{BT},"POR PAGAR")', "cop", f'="Vencida: $"&FIXED(SUMIFS({BS},{BT},"POR PAGAR",{rng("CAR","ESTADO")},"VENCIDA"),0)', AMBER),
            ("Valor del inventario", f"=SUMIFS({Ir('Q')},{Ir('C')},{Cc('$B$4')})", "cop", f'=FIXED(SUMIFS({Ir("O")},{Ir("C")},{Cc("$B$4")}),0)&" unidades"', NAVY),
            ("Productos disponibles", f'=COUNTIFS({Ir("T")},"DISPONIBLE",{Ir("C")},{Cc("$B$4")})+COUNTIFS({Ir("T")},"BAJO STOCK",{Ir("C")},{Cc("$B$4")})', "int", f'="de "&COUNTIFS({Ir("C")},{Cc("$B$4")})&" referencias"', NAVY),
            ("Productos agotados", f'=COUNTIFS({Ir("T")},"AGOTADO",{Ir("C")},{Cc("$B$4")})', "int", '="Reponer de inmediato"', RED),
            ("Bajo inventario", f'=COUNTIFS({Ir("T")},"BAJO STOCK",{Ir("C")},{Cc("$B$4")})', "int", '="En o bajo el stock mínimo"', AMBER),
            ("Clientes", f"=COUNTA({q('CLI')}!$A$7:$A$2006)", "int", f'=COUNTIF({q("CLI")}!$O$7:$O$2006,"FRECUENTE")&" frecuentes"', BLUE),
            ("Ticket promedio", f"=IFERROR({Cc('$M$3')}/{Cc('$M$9')},0)", "cop", f'={Cc("$M$9")}&" facturas en el período"', BLUE),
            ("Margen neto", f"=IFERROR({Cc('$M$10')}/{Cc('$M$3')},0)", "pct", '="Utilidad neta / ventas"', TEAL),
        ]
        cols = ["C", "F", "I", "L", "O"]
        for n, (lab_, fm, nf, sub, color) in enumerate(cards):
            r0 = 6 + (n // 5) * 4
            c0 = cols[n % 5]
            c1 = chr(ord(c0) + 2)
            ws.merge_range(f"{c0}{r0}:{c1}{r0}", lab_.upper(), F(font_size=8, bold=True, font_color=GREY, bg_color="white", left=5, left_color=color, top=1, top_color=LINE, right=1, right_color=LINE))
            ws.merge_range(f"{c0}{r0+1}:{c1}{r0+1}", "", F(font_size=17, bold=True, font_color=NAVY, bg_color="white", nf=nf, left=5, left_color=color, right=1, right_color=LINE))
            ws.write_formula(f"{c0}{r0+1}", fm, F(font_size=17, bold=True, font_color=NAVY, bg_color="white", nf=nf, left=5, left_color=color, right=1, right_color=LINE))
            ws.merge_range(f"{c0}{r0+2}:{c1}{r0+2}", "", F(font_size=8, font_color=GREY, bg_color="white", left=5, left_color=color, bottom=1, bottom_color=LINE, right=1, right_color=LINE))
            ws.write_formula(f"{c0}{r0+2}", sub, F(font_size=8, font_color=GREY, bg_color="white", left=5, left_color=color, bottom=1, bottom_color=LINE, right=1, right_color=LINE))
        for r in range(5, 19):
            ws.set_row(r, 20 if (r - 5) % 4 != 1 else 28)
        ws.set_row(9, 8)
        ws.set_row(13, 8)
        # panel de alertas
        ws.merge_range("S6:V6", "⚠ ALERTAS DE INVENTARIO", F.hdr(bg_color=RED))
        ws.write("S7", "Producto", F.hdr(bg_color="#991B1B"))
        ws.write("T7", "Stock", F.hdr(bg_color="#991B1B"))
        ws.merge_range("U7:V7", "Estado", F.hdr(bg_color="#991B1B"))
        for kk in range(10):
            r = 8 + kk
            ws.write_formula(f"S{r}", f"={Cc(f'$B${85+kk}')}", F.out(font_size=9))
            ws.write_formula(f"T{r}", f"={Cc(f'$C${85+kk}')}", F.out("int", font_size=9, align="center"))
            ws.merge_range(f"U{r}:V{r}", "", F.out(font_size=9))
            ws.write_formula(f"U{r}", f"={Cc(f'$D${85+kk}')}", F.out(font_size=9, bold=True))
        ws.conditional_format("U8:U17", {"type": "cell", "criteria": "==", "value": '"AGOTADO"', "format": F(font_color=RED)})
        ws.conditional_format("U8:U17", {"type": "cell", "criteria": "==", "value": '"BAJO STOCK"', "format": F(font_color=AMBER)})
        ws.merge_range("S19:V19", "GASTOS FIJOS DEL MES", F.hdr(bg_color=AMBER))
        ws.write("S20", "Vencidos sin pagar", F.out())
        ws.merge_range("T20:V20", "", F.out("int", bold=True, font_color=RED))
        ws.write_formula("T20", f'=COUNTIF({q("GFI")}!$J$7:$J$106,"VENCIDO")', F.out("int", bold=True, font_color=RED))
        ws.write("S21", "Próximos a vencer (5 días)", F.out())
        ws.merge_range("T21:V21", "", F.out("int", bold=True, font_color=AMBER))
        ws.write_formula("T21", f'=COUNTIF({q("GFI")}!$J$7:$J$106,"PRÓXIMO A VENCER")', F.out("int", bold=True, font_color=AMBER))
        ws.merge_range("S23:V23", "CAJA DISPONIBLE HOY", F.hdr(bg_color=GREEN))
        for i in range(5):
            r = 24 + i
            ws.write_formula(f"S{r}", f"={q('CAJ')}!B{6+i}", F.out())
            ws.merge_range(f"T{r}:V{r}", "", F.out("cop", bold=True))
            ws.write_formula(f"T{r}", f"={q('CAJ')}!H{6+i}", F.out("cop", bold=True))
        ws.write("S29", "TOTAL", F.hdr())
        ws.merge_range("T29:V29", "", F.out("cop", bold=True, bg_color="#DCFCE7"))
        ws.write_formula("T29", f"={q('CAJ')}!H11", F.out("cop", bold=True, bg_color="#DCFCE7"))
        ws.write("S31", "Los indicadores se recalculan solos con cada operación.", F(italic=True, font_color=GREY, font_size=8))
        ws.write("S32", "Filtros: año, mes (o TODOS), categoría y día.", F(italic=True, font_color=GREY, font_size=8))
        # gráficos
        wb = self.wb
        CN = SH["CALC"]

        def chart(kind, title, series, sub=None, legend=True):
            ch = wb.add_chart({"type": kind, **({"subtype": sub} if sub else {})})
            for s in series:
                ch.add_series(s)
            ch.set_title({"name": title, "name_font": {"size": 11, "bold": True, "color": NAVY, "name": FONT}})
            ch.set_legend({"position": "bottom", "font": {"size": 8}} if legend else {"none": True})
            ch.set_chartarea({"border": {"color": LINE}})
            ch.set_size({"width": 560, "height": 290})
            if kind not in ("pie", "doughnut"):
                ch.set_y_axis({"num_format": '"$"#,##0,,"M"' if "Unid" not in title and "unid" not in title else "#,##0",
                               "major_gridlines": {"visible": True, "line": {"color": "#E2E8F0"}},
                               "num_font": {"size": 8}})
                ch.set_x_axis({"num_font": {"size": 8}})
            return ch
        cat = [CN, 10, 1, 21, 1]
        c1 = chart("column", "Ventas mensuales", [{"name": "Ventas", "categories": cat, "values": [CN, 10, 2, 21, 2], "fill": {"color": BLUE}, "gap": 60}], legend=False)
        c2 = chart("column", "Ingresos vs. egresos (caja)", [
            {"name": "Ingresos", "categories": cat, "values": [CN, 10, 7, 21, 7], "fill": {"color": GREEN}},
            {"name": "Egresos", "categories": cat, "values": [CN, 10, 8, 21, 8], "fill": {"color": RED}}])
        c3 = chart("line", "Utilidad neta mensual", [{"name": "Utilidad neta", "categories": cat, "values": [CN, 10, 6, 21, 6],
                                                       "line": {"color": TEAL, "width": 2.5}, "marker": {"type": "circle", "size": 6, "fill": {"color": TEAL}}},
                                                      {"name": "Utilidad bruta", "categories": cat, "values": [CN, 10, 4, 21, 4],
                                                       "line": {"color": "#93C5FD", "width": 1.5, "dash_type": "dash"}}])
        c4 = chart("bar", "Productos más vendidos (unidades)", [{"name": "Unidades", "categories": [CN, 70, 1, 74, 1], "values": [CN, 70, 2, 74, 2],
                                                                "fill": {"color": GREEN}, "data_labels": {"value": True, "font": {"size": 8}}}], legend=False)
        c4.set_y_axis({"reverse": True, "num_font": {"size": 8}})
        c4.set_x_axis({"num_format": "#,##0", "num_font": {"size": 8}})
        c5 = chart("bar", "Productos menos vendidos (unidades)", [{"name": "Unidades", "categories": [CN, 77, 1, 81, 1], "values": [CN, 77, 2, 81, 2],
                                                                  "fill": {"color": AMBER}, "data_labels": {"value": True, "font": {"size": 8}}}], legend=False)
        c5.set_y_axis({"reverse": True, "num_font": {"size": 8}})
        c5.set_x_axis({"num_format": "#,##0", "num_font": {"size": 8}})
        c6 = chart("bar", "Ventas por categoría", [{"name": "Ventas", "categories": [CN, 30, 4, 37, 4], "values": [CN, 30, 5, 37, 5],
                                                  "fill": {"color": "#6366F1"}}], legend=False)
        c6.set_y_axis({"reverse": True, "num_font": {"size": 8}})
        c6.set_x_axis({"num_format": '"$"#,##0,,"M"', "num_font": {"size": 8}})
        c7 = chart("doughnut", "Recaudo por método de pago", [{"name": "Recaudo", "categories": [CN, 55, 0, 60, 0], "values": [CN, 55, 1, 60, 1],
                                                              "data_labels": {"percentage": True, "font": {"size": 8, "color": "white", "bold": True}},
                                                              "points": [{"fill": {"color": c}} for c in (GREEN, BLUE, "#7C3AED", RED, AMBER, GREY)]}])
        c7.set_hole_size(50)
        c8 = chart("line", "Evolución del valor del inventario", [{"name": "Valor inventario", "categories": cat, "values": [CN, 10, 9, 21, 9],
                                                                   "line": {"color": BLUE, "width": 2.5}, "marker": {"type": "circle", "size": 5}}], legend=False)
        c8.show_na_as_empty_cell()
        pos = [("C19", c1), ("J19", c2), ("C35", c3), ("J35", c4), ("C51", c5), ("J51", c6), ("C67", c7), ("J67", c8)]
        for addr, ch in pos:
            ws.insert_chart(addr, ch, {"x_offset": 4, "y_offset": 4, "object_position": 1})
        ws.activate()
        ws.set_first_sheet()
        self.proteger(k)

    # ------------------------------------------------------------
    def logo_demo(self):
        try:
            from PIL import Image, ImageDraw, ImageFont
        except ImportError:
            return
        im = Image.new("RGBA", (170, 140), (255, 255, 255, 0))
        d = ImageDraw.Draw(im)
        d.rounded_rectangle((8, 8, 162, 132), radius=24, fill=(30, 41, 59, 255))
        try:
            fnt = ImageFont.truetype("DejaVuSans-Bold.ttf", 46)
            f2 = ImageFont.truetype("DejaVuSans.ttf", 13)
        except OSError:
            fnt = f2 = ImageFont.load_default()
        d.text((85, 58), "M&E", fill="white", font=fnt, anchor="mm")
        d.text((85, 104), "MODA & ESTILO", fill=(203, 213, 225), font=f2, anchor="mm")
        im.save(self.logo)

    def construir(self):
        self.logo_demo()
        self.nombres()
        self.hoja_cfg()
        self.hoja_inv()
        self.hojas_historial()
        self.hoja_cli()
        self.hoja_pro()
        self.hoja_com()
        self.hoja_cos()
        self.hoja_fac()
        self.hoja_gui()
        self.hoja_ing()
        self.hoja_egr()
        self.hoja_gfi()
        self.hoja_car()
        self.hoja_caj()
        self.hoja_uti()
        self.hoja_rep()
        self.hoja_calc()
        self.hoja_ini()
        for k in ORDEN:
            if k not in ("INI", "CALC", "FAC", "GUI"):
                self.ws[k].set_tab_color({"BD": GREY, "HFA": GREY, "HGU": GREY, "MOV": GREY, "AUD": GREY,
                                          "CFG": "#475569"}.get(k, BLUE))
        self.ws["INI"].set_tab_color(NAVY)
        self.ws["FAC"].set_tab_color(GREEN)
        self.ws["GUI"].set_tab_color(GREEN)
        self.wb.close()
        return {"ok": True, "archivo": self.out, "hojas": len(ORDEN), "demo": self.demo}


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default="../SistemaEmpresarial_base.xlsx")
    ap.add_argument("--sin-demo", action="store_true", help="genera el libro vacío (sin datos de ejemplo)")
    a = ap.parse_args()
    print(json.dumps(Builder(a.out, not a.sin_demo).construir(), ensure_ascii=False))
