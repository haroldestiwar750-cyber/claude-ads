"""Genera ERP_Comercial.xlsx: interfaz Excel del ERP comercial con datos de prueba.

Uso:
    python generar_erp.py [ruta_salida.xlsx]

Salida (stdout, JSON): {"archivo": ..., "hojas": n, "nombres_definidos": n}

Todas las cifras calculadas son fórmulas de Excel; los datos de prueba vienen de
datos_prueba.py, que aplica las mismas reglas que las macros VBA.
"""
import json
import os
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink

import datos_prueba as dp

AQUI = os.path.dirname(os.path.abspath(__file__))
SALIDA = sys.argv[1] if len(sys.argv) > 1 else os.path.join(AQUI, "..", "ERP_Comercial.xlsx")

# ------------------------------------------------------------------ estilo
FONT = "Arial"
NAVY, AZUL, AZUL_CLARO, GRIS, GRIS_CLARO = "1F3864", "2E75B6", "DDEBF7", "595959", "F2F2F2"
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
CALC_FILL = PatternFill("solid", fgColor="F3F6FA")
F_NAVY = PatternFill("solid", fgColor=NAVY)
F_AZUL = PatternFill("solid", fgColor=AZUL)
F_TEAL = PatternFill("solid", fgColor="44546A")
F_CLARO = PatternFill("solid", fgColor=AZUL_CLARO)
THIN = Side(style="thin", color="BFBFBF")
BORDE = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
COP = '"$" #,##0;[Red]-"$" #,##0;"-"'
NUM = '#,##0;[Red]-#,##0;"-"'
PCT = '0.0%;[Red]-0.0%;"-"'
FECHA = "DD/MM/YYYY"
FECHAHORA = "DD/MM/YYYY HH:MM"
UNLOCK = Protection(locked=False)

DATA_MAX = 5000      # filas de BD
DET_MAX = 20000      # filas de detalle / kárdex


def f(size=10, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def q(nombre):
    return "'" + nombre + "'"


wb = Workbook()
wb.remove(wb.active)
HOJAS = {}


def hoja(nombre, code, tab=None):
    ws = wb.create_sheet(nombre)
    ws.sheet_properties.codeName = code
    if tab:
        ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    HOJAS[code] = ws
    return ws


def nombre(nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def rango(nm, ws, col, r1, r2):
    nombre(nm, f"{q(ws.title)}!${col}${r1}:${col}${r2}")


def celda(nm, ws, ref):
    c = ref.replace("$", "")
    col = "".join(ch for ch in c if ch.isalpha())
    row = "".join(ch for ch in c if ch.isdigit())
    nombre(nm, f"{q(ws.title)}!${col}${row}")


def lista_dinamica(nm, ws, col, r1, r2):
    nombre(nm, f'OFFSET({q(ws.title)}!${col}${r1},0,0,MAX(1,COUNTIF({q(ws.title)}!${col}${r1}:${col}${r2},"?*")),1)')


def titulo(ws, texto, ultima_col, subtitulo=""):
    ws.row_dimensions[1].height = 30
    for c in range(1, ultima_col + 1):
        ws.cell(1, c).fill = F_NAVY
        ws.cell(2, c).fill = PatternFill("solid", fgColor="D9E1F2")
    ws.cell(1, 2, texto).font = f(15, True, "FFFFFF")
    ws.cell(1, 2).alignment = Alignment(vertical="center")
    a = ws.cell(2, 2, "◀ MENÚ PRINCIPAL")
    a.hyperlink = Hyperlink(ref="B2", location=f"{q('MENÚ')}!A1")
    a.font = f(9, True, AZUL)
    if subtitulo:
        ws.cell(2, 4, subtitulo).font = f(9, False, GRIS, italic=True)
    ws.column_dimensions["A"].width = 2
    ws.row_dimensions[3].height = 30  # franja de botones (las macros los dibujan aquí)


def encabezados(ws, fila, cols, filtro=True):
    """cols: lista de (titulo, ancho, tipo) con tipo 'in' (editable) o 'calc' (fórmula/sistema)."""
    for i, (t, w, tipo) in enumerate(cols, start=1):
        c = ws.cell(fila, i, t)
        c.font = f(9, True, "FFFFFF")
        c.fill = F_AZUL if tipo == "in" else F_TEAL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDE
        ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[fila].height = 32
    if filtro:
        ws.auto_filter.ref = f"A{fila}:{L(len(cols))}{fila}"
    ws.freeze_panes = ws.cell(fila + 1, 3)


def escribir_filas(ws, fila0, filas, formatos=None):
    formatos = formatos or {}
    for r, fila in enumerate(filas, start=fila0):
        for c, v in enumerate(fila, start=1):
            cell = ws.cell(r, c, v)
            cell.font = f(9)
            cell.border = BORDE
            if c in formatos:
                cell.number_format = formatos[c]


def estilo_columnas(ws, fila0, fila1, cols, formatos=None):
    """Aplica estilo (bloqueo, relleno, formato) a un rango de filas de una tabla."""
    formatos = formatos or {}
    for i, (_, _, tipo) in enumerate(cols, start=1):
        for r in range(fila0, fila1 + 1):
            cell = ws.cell(r, i)
            cell.font = f(9)
            cell.border = BORDE
            if tipo == "in":
                cell.protection = UNLOCK
            else:
                cell.fill = CALC_FILL
            if i in formatos:
                cell.number_format = formatos[i]


def dv_lista(ws, ref, fuente, msg=None):
    dv = DataValidation(type="list", formula1=f"{fuente}", allow_blank=True, showErrorMessage=True,
                        errorTitle="Valor no válido", error="Seleccione un valor de la lista.")
    if msg:
        dv.promptTitle, dv.prompt, dv.showInputMessage = "Ayuda", msg, True
    ws.add_data_validation(dv)
    dv.add(ref)


def dv_unico(ws, ref, primera, rango_abs, campo):
    dv = DataValidation(type="custom", formula1=f"COUNTIF({rango_abs},{primera})=1", allow_blank=True,
                        showErrorMessage=True, errorTitle="Duplicado",
                        error=f"Ya existe un registro con este {campo}. No se permiten duplicados.")
    ws.add_data_validation(dv)
    dv.add(ref)


def nota(ws, ref, texto):
    ws[ref].comment = Comment(texto, "ERP")


def marcador_botones(ws, texto):
    c = ws.cell(3, 2, texto)
    c.font = f(8, False, "808080", italic=True)


# ====================================================== crear todas las hojas (orden de pestañas)
sh = {}
sh["menu"] = hoja("MENÚ", "shMenu", NAVY)
sh["dash"] = hoja("DASHBOARD", "shDash", "C00000")
sh["fact"] = hoja("FACTURACIÓN", "shFact", "00B050")
sh["guia"] = hoja("GUÍAS", "shGuia", "00B050")
sh["inv"] = hoja("INVENTARIO", "shInv", "ED7D31")
sh["compra"] = hoja("COMPRAS", "shCompra", "ED7D31")
sh["cli"] = hoja("CLIENTES", "shCli", "00B050")
sh["prov"] = hoja("PROVEEDORES", "shProv", "ED7D31")
sh["costos"] = hoja("COSTOS Y RENTABILIDAD", "shCostos", "ED7D31")
sh["ing"] = hoja("INGRESOS", "shIng", "7030A0")
sh["egr"] = hoja("EGRESOS", "shEgr", "7030A0")
sh["gf"] = hoja("GASTOS FIJOS", "shGF", "7030A0")
sh["cart"] = hoja("CARTERA", "shCart", "7030A0")
sh["caja"] = hoja("CAJA Y BANCOS", "shCaja", "7030A0")
sh["util"] = hoja("UTILIDADES", "shUtil", "7030A0")
sh["vend"] = hoja("VENDEDORES", "shVend", "00B050")
sh["usr"] = hoja("USUARIOS Y PERMISOS", "shUsr", "808080")
sh["com"] = hoja("COMISIONES", "shCom", "00B050")
sh["canal"] = hoja("CANALES DE VENTA", "shCanal", "00B050")
sh["rep"] = hoja("REPORTES", "shRep", "C00000")
sh["kdx"] = hoja("HISTORIAL MOVIMIENTOS", "shKardex", "ED7D31")
sh["cfg"] = hoja("CONFIGURACIÓN", "shCfg", "808080")
sh["bd"] = hoja("BASE DE DATOS", "shBD", "808080")
sh["dbfac"] = hoja("BD_FACTURAS", "dbFact", "BFBFBF")
sh["dbdet"] = hoja("BD_FACT_DETALLE", "dbFactDet", "BFBFBF")
sh["dbcom"] = hoja("BD_COMPRAS", "dbCompra", "BFBFBF")
sh["dbcomd"] = hoja("BD_COMPRA_DETALLE", "dbCompraDet", "BFBFBF")
sh["dbgui"] = hoja("BD_GUIAS", "dbGuia", "BFBFBF")
sh["dbab"] = hoja("BD_ABONOS", "dbAbono", "BFBFBF")
sh["dbcie"] = hoja("BD_CIERRES", "dbCierre", "BFBFBF")
sh["dblog"] = hoja("BD_BITACORA", "dbLog", "BFBFBF")

S = dp.simular()

# ====================================================== CONFIGURACIÓN
ws = sh["cfg"]
titulo(ws, "CONFIGURACIÓN DEL SISTEMA", 8, "Datos de la empresa, consecutivos y parámetros. Solo Administrador.")
marcador_botones(ws, "Botones: DETECTAR ONEDRIVE · CREAR CARPETAS · RESPALDO AHORA · RESTAURAR COPIA · CAMBIAR MI CLAVE")
ws.column_dimensions["B"].width = 38
ws.column_dimensions["C"].width = 42
ws.column_dimensions["D"].width = 4
ws.column_dimensions["E"].width = 22
ws.column_dimensions["F"].width = 18
ws.column_dimensions["G"].width = 16
ws.column_dimensions["H"].width = 30

ultimo_fac = int(S.facturas[-1][0].split("-")[1])
ultimo_oc = int(S.compras[-1][0].split("-")[1])
ultimo_gu = int(S.guias[-1][0].split("-")[1])

CFG = [
    ("EMPRESA", None, None),
    ("Razón social", "cfg_Empresa", "TIENDA URBANA S.A.S. (ejemplo: cambiar)"),
    ("NIT", "cfg_NIT", "901.000.000-0"),
    ("Dirección", "cfg_Direccion", "Cra 15 # 85-24 Local 3"),
    ("Ciudad", "cfg_Ciudad", "Bogotá D.C."),
    ("Teléfono / WhatsApp", "cfg_Telefono", "300 000 0000"),
    ("Correo", "cfg_Correo", "ventas@tiendaurbana.ejemplo.co"),
    ("Instagram / TikTok", "cfg_Redes", "@tiendaurbana"),
    ("Leyenda del comprobante", "cfg_Leyenda",
     "Documento interno de venta. NO es factura electrónica de venta validada por la DIAN."),
    ("CONSECUTIVOS (siguiente número a usar)", None, None),
    ("Prefijo comprobante de venta", "cfg_PrefFactura", "FV"),
    ("Siguiente número de venta", "cfg_NumFactura", ultimo_fac + 1),
    ("Prefijo compra", "cfg_PrefCompra", "OC"),
    ("Siguiente número de compra", "cfg_NumCompra", ultimo_oc + 1),
    ("Prefijo guía", "cfg_PrefGuia", "GU"),
    ("Siguiente número de guía", "cfg_NumGuia", ultimo_gu + 1),
    ("PARÁMETROS DE COSTOS Y PRECIOS", None, None),
    ("Empaque por unidad (COP)", "cfg_Empaque", dp.EMPAQUE),
    ("Transporte asumido por unidad vendida (COP)", "cfg_TransporteUnit", 0),
    ("% comisión promedio para costeo", "cfg_ComisionProm", dp.COMISION_PROM),
    ("Margen escenario COMPETITIVO", "cfg_MargenComp", 0.25),
    ("Margen escenario RECOMENDADO", "cfg_MargenRec", 0.40),
    ("Margen escenario PREMIUM", "cfg_MargenPrem", 0.55),
    ("PARÁMETROS COMERCIALES", None, None),
    ("Días de crédito a clientes", "cfg_DiasCredito", 15),
    ("Días de crédito de proveedores (por defecto)", "cfg_DiasCreditoProv", 30),
    ("Días sin comprar para cliente INACTIVO", "cfg_DiasInactivo", 90),
    ("Compras mínimas para cliente FRECUENTE", "cfg_ComprasFrec", 3),
    ("Permitir inventario negativo (requiere autorización)", "cfg_PermitirNegativo", "NO"),
    ("SISTEMA Y ONEDRIVE", None, None),
    ("Carpeta raíz del sistema", "cfg_CarpetaSistema", "SISTEMA EMPRESARIAL"),
    ("Ruta OneDrive detectada (la llena la macro)", "cfg_RutaOneDrive", ""),
    ("Copias de seguridad a conservar", "cfg_CopiasConservar", 30),
    ("Respaldo automático cada (horas)", "cfg_HorasRespaldo", 4),
    ("Clave de protección de hojas", "cfg_ClaveHojas", "ERP-Cambiar-2026"),
    ("Versión del sistema", "cfg_Version", "1.0.0"),
    ("SESIÓN ACTUAL (la llena la macro)", None, None),
    ("Usuario", "app_Usuario", "(sin sesión)"),
    ("Rol", "app_Rol", ""),
    ("Vendedor asociado", "app_Vendedor", ""),
]
r = 5
for etiqueta, nm, val in CFG:
    if nm is None:
        r += 1
        c = ws.cell(r, 2, etiqueta)
        c.font = f(10, True, "FFFFFF")
        c.fill = F_TEAL
        ws.cell(r, 3).fill = F_TEAL
    else:
        ws.cell(r, 2, etiqueta).font = f(9)
        c = ws.cell(r, 3, val)
        c.font = f(9, color="0000FF")
        c.border = BORDE
        c.protection = UNLOCK
        if not nm.startswith("app_"):
            c.fill = INPUT_FILL
        else:
            c.fill = CALC_FILL
        if nm in ("cfg_Empaque", "cfg_TransporteUnit"):
            c.number_format = COP
        if nm in ("cfg_ComisionProm", "cfg_MargenComp", "cfg_MargenRec", "cfg_MargenPrem"):
            c.number_format = "0.0%"
        celda(nm, ws, f"C{r}")
    r += 1
nota(ws, "C7", "Reemplace por los datos reales de su empresa.")
nota_clave_row = [i for i, x in enumerate(CFG) if x[1] == "cfg_ClaveHojas"]

# saldos iniciales por cuenta
ws["E6"] = "SALDOS INICIALES POR CUENTA"
ws["E6"].font = f(10, True, "FFFFFF")
for c in "EFG":
    ws[f"{c}6"].fill = F_TEAL
for i, t in enumerate(["Cuenta / método", "Saldo inicial", "Fecha saldo"]):
    c = ws.cell(7, 5 + i, t)
    c.font = f(9, True)
    c.fill = F_CLARO
    c.border = BORDE
for i, m in enumerate(dp.METODOS):
    rr = 8 + i
    ws.cell(rr, 5, m).font = f(9)
    c = ws.cell(rr, 6, dp.SALDOS_INICIALES[m])
    c.number_format, c.fill, c.protection, c.font = COP, INPUT_FILL, UNLOCK, f(9, color="0000FF")
    c = ws.cell(rr, 7, dp.FECHA_SALDO_INICIAL)
    c.number_format, c.fill, c.protection, c.font = FECHA, INPUT_FILL, UNLOCK, f(9, color="0000FF")
    for cc in (5, 6, 7):
        ws.cell(rr, cc).border = BORDE
rango("cfg_Cuentas", ws, "E", 8, 12)
rango("cfg_SaldosIni", ws, "F", 8, 12)
ws["E14"] = ("Nota: la protección de hojas con clave evita errores de digitación, pero NO es seguridad "
             "fuerte: cualquier usuario con conocimientos puede quitarla. La seguridad real llegará con la "
             "base de datos central (Etapa 1).")
ws["E14"].font = f(8, italic=True, color="C00000")
ws["E14"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("E14:H18")

# ====================================================== BASE DE DATOS (catálogos)
ws = sh["bd"]
titulo(ws, "BASE DE DATOS · CATÁLOGOS Y LISTAS", 22,
       "Listas desplegables de todo el sistema. Agregue valores al final de cada columna.")
cat_cols = [
    ("Categorías", dp.CATEGORIAS, "lst_Categorias"),
    ("Tallas", dp.TALLAS, "lst_Tallas"),
    ("Colores", dp.COLORES, "lst_Colores"),
    ("Canales de venta", dp.CANALES, "lst_Canales"),
    ("Métodos de pago / cuentas", dp.METODOS, "lst_Metodos"),
    ("Estados de pago", dp.ESTADOS_PAGO, "lst_EstadosPago"),
    ("Transportadoras", dp.TRANSPORTADORAS, "lst_Transportadoras"),
    ("Tipos de ingreso", dp.CONCEPTOS_ING, "lst_TiposIngreso"),
    ("Conceptos de egreso", dp.CONCEPTOS_EGR, "lst_ConceptosEgr"),
    ("Roles", dp.ROLES, "lst_Roles"),
    ("Estados de guía", dp.ESTADOS_GUIA, "lst_EstadosGuia"),
    ("Tipos movimiento inventario", dp.TIPOS_MOV, "lst_TiposMov"),
    ("Ciudades", dp.CIUDADES, "lst_Ciudades"),
    ("Estados", dp.ESTADOS, "lst_Estados"),
    ("Tipos de documento", dp.TIPOS_DOC, "lst_TiposDoc"),
    ("Sí / No", ["SI", "NO"], "lst_SiNo"),
    ("Tipo cartera", ["CxC", "CxP"], "lst_TipoCartera"),
    ("Tipo de cierre", ["DIARIO", "SEMANAL", "MENSUAL"], "lst_TipoCierre"),
]
for i, (t, vals, nm) in enumerate(cat_cols):
    col = 2 + i
    c = ws.cell(5, col, t)
    c.font, c.fill, c.alignment, c.border = f(9, True, "FFFFFF"), F_AZUL, Alignment(wrap_text=True, horizontal="center"), BORDE
    ws.column_dimensions[L(col)].width = 17
    for j, v in enumerate(vals):
        cc = ws.cell(6 + j, col, v)
        cc.font, cc.border, cc.protection = f(9), BORDE, UNLOCK
    for j in range(len(vals), 40):
        ws.cell(6 + j, col).protection = UNLOCK
    lista_dinamica(nm, ws, L(col), 6, 60)
ws.row_dimensions[5].height = 32
# listas de filtro (Todos + valores)
base = 2 + len(cat_cols) + 1
for k, (t, fuente_col, nm, n) in enumerate([
    ("Filtro canal", "E", "flt_Canales", 30), ("Filtro categoría", "B", "flt_Categorias", 30)]):
    col = base + k
    c = ws.cell(5, col, t)
    c.font, c.fill, c.border = f(9, True, "FFFFFF"), F_TEAL, BORDE
    ws.column_dimensions[L(col)].width = 17
    ws.cell(6, col, "Todos").font = f(9)
    for j in range(n):
        cc = ws.cell(7 + j, col, f'=IF({fuente_col}{6 + j}="","",{fuente_col}{6 + j})')
        cc.font, cc.fill = f(9), CALC_FILL
    lista_dinamica(nm, ws, L(col), 6, 6 + n)
col = base + 2
ws.cell(5, col, "Filtro vendedor").font = f(9, True, "FFFFFF")
ws.cell(5, col).fill = F_TEAL
ws.column_dimensions[L(col)].width = 24
ws.cell(6, col, "Todos").font = f(9)
for j in range(30):
    cc = ws.cell(7 + j, col, f"=IF({q('VENDEDORES')}!O{6 + j}=\"\",\"\",{q('VENDEDORES')}!O{6 + j})")
    cc.font, cc.fill = f(9), CALC_FILL
lista_dinamica("flt_Vendedores", ws, L(col), 6, 36)
FLT_VEND_COL = L(col)

# índice de tablas BD
ws.cell(48, 2, "TABLAS TRANSACCIONALES (las escriben las macros; no editar a mano)").font = f(10, True, NAVY)
for i, (h, desc) in enumerate([
    ("BD_FACTURAS", "Encabezado de cada comprobante de venta"),
    ("BD_FACT_DETALLE", "Líneas vendidas con costo congelado y utilidad"),
    ("BD_COMPRAS", "Encabezado de compras"),
    ("BD_COMPRA_DETALLE", "Líneas compradas con prorrateo de fletes"),
    ("BD_GUIAS", "Guías de envío"),
    ("BD_ABONOS", "Abonos a cuentas por cobrar y por pagar"),
    ("BD_CIERRES", "Cierres de caja diarios / semanales / mensuales"),
    ("BD_BITACORA", "Registro de actividad de todos los usuarios"),
    ("HISTORIAL MOVIMIENTOS", "Kárdex: cada entrada y salida de inventario"),
]):
    c = ws.cell(49 + i, 2, h)
    c.hyperlink = Hyperlink(ref=c.coordinate, location=f"{q(h)}!A1")
    c.font = f(9, True, AZUL)
    ws.cell(49 + i, 4, desc).font = f(9)

# ====================================================== INVENTARIO
ws = sh["inv"]
titulo(ws, "INVENTARIO · SKU POR PRODUCTO, TALLA Y COLOR", 16,
       "Entradas y salidas se calculan desde el HISTORIAL MOVIMIENTOS (kárdex). Costo = promedio ponderado.")
marcador_botones(ws, "Botones: NUEVO PRODUCTO (genera SKU por talla y color) · AJUSTE DE INVENTARIO · ACTUALIZAR")
INV_R1, INV_R2 = 6, 505
inv_cols = [("SKU", 15, "calc"), ("Producto", 26, "in"), ("Categoría", 15, "in"), ("Talla", 7, "in"),
            ("Color", 10, "in"), ("Proveedor", 10, "in"), ("Costo promedio", 12, "calc"),
            ("Precio de venta", 12, "in"), ("Stock inicial", 9, "calc"), ("Entradas", 9, "calc"),
            ("Salidas", 9, "calc"), ("Existencias", 10, "calc"), ("Stock mínimo", 9, "in"),
            ("Valor inventario", 14, "calc"), ("Estado", 11, "calc"), ("Etiqueta (lista)", 34, "calc")]
encabezados(ws, 5, inv_cols)
estilo_columnas(ws, INV_R1, INV_R2, inv_cols, {7: COP, 8: COP, 9: NUM, 10: NUM, 11: NUM, 12: NUM, 13: NUM, 14: COP})
for i, s in enumerate(dp.SKUS):
    r = INV_R1 + i
    for c, v in enumerate([s["sku"], s["prod"], s["cat"], s["talla"], s["color"], s["prov"], S.costo[s["sku"]],
                           s["precio"], s["ini"]], start=1):
        ws.cell(r, c, v)
    ws.cell(r, 13, s["minimo"])
for r in range(INV_R1, INV_R2 + 1):
    ws.cell(r, 10, f'=IF(A{r}="","",SUMIFS(kdx_Ent,kdx_SKU,A{r},kdx_Tipo,"<>INICIAL"))')
    ws.cell(r, 11, f'=IF(A{r}="","",SUMIFS(kdx_Sal,kdx_SKU,A{r}))')
    ws.cell(r, 12, f'=IF(A{r}="","",N(I{r})+J{r}-K{r})')
    ws.cell(r, 14, f'=IF(A{r}="","",L{r}*G{r})')
    ws.cell(r, 15, f'=IF(A{r}="","",IF(L{r}<=0,"AGOTADO",IF(L{r}<=M{r},"BAJO","OK")))')
    ws.cell(r, 16, f'=IF(A{r}="","",A{r}&" - "&B{r}&" "&D{r}&" "&E{r})')
nota(ws, "G5", "Costo promedio ponderado. Lo actualiza la macro al guardar cada compra (incluye fletes prorrateados).")
nota(ws, "I5", "Solo al crear el SKU (queda también en el kárdex como INICIAL, sin sumarse dos veces). Después todo entra por compras, ventas o ajustes autorizados.")
for nm, col in [("inv_SKU", "A"), ("inv_Prod", "B"), ("inv_Cat", "C"), ("inv_Talla", "D"), ("inv_Color", "E"),
                ("inv_Prov", "F"), ("inv_Costo", "G"), ("inv_Precio", "H"), ("inv_Exist", "L"),
                ("inv_Min", "M"), ("inv_Valor", "N"), ("inv_Estado", "O"), ("inv_Label", "P")]:
    rango(nm, ws, col, INV_R1, INV_R2)
lista_dinamica("lst_SKU", ws, "P", INV_R1, INV_R2)
dv_unico(ws, f"A{INV_R1}:A{INV_R2}", f"A{INV_R1}", f"$A${INV_R1}:$A${INV_R2}", "SKU")
dv_lista(ws, f"C{INV_R1}:C{INV_R2}", "lst_Categorias")
dv_lista(ws, f"D{INV_R1}:D{INV_R2}", "lst_Tallas")
dv_lista(ws, f"E{INV_R1}:E{INV_R2}", "lst_Colores")
rng = f"O{INV_R1}:O{INV_R2}"
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"AGOTADO"'],
                              fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True)))
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"BAJO"'],
                              fill=PatternFill("solid", fgColor="FFEB9C"), font=Font(color="9C5700", bold=True)))
ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"OK"'],
                              fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(color="006100")))
# totales
ws["J4"] = "Unidades:"
ws["K4"] = "=SUM(L6:L505)"
ws["M4"] = "Valor total:"
ws["N4"] = "=SUM(N6:N505)"
for c in ("J4", "M4"):
    ws[c].font = f(9, True)
    ws[c].alignment = Alignment(horizontal="right")
ws["K4"].number_format, ws["N4"].number_format = NUM, COP
ws["K4"].font = ws["N4"].font = f(10, True, NAVY)

# ====================================================== HISTORIAL MOVIMIENTOS (kárdex)
ws = sh["kdx"]
titulo(ws, "HISTORIAL DE MOVIMIENTOS DE INVENTARIO (KÁRDEX)", 13,
       "Cada entrada/salida con fecha, usuario, documento y motivo. Solo lo escriben las macros.")
kdx_cols = [("ID", 10, "calc"), ("Fecha", 11, "calc"), ("Hora", 7, "calc"), ("Usuario", 10, "calc"),
            ("Tipo", 16, "calc"), ("SKU", 14, "calc"), ("Producto", 30, "calc"), ("Entrada", 8, "calc"),
            ("Salida", 8, "calc"), ("Costo unit.", 11, "calc"), ("Documento", 12, "calc"),
            ("Motivo", 34, "calc"), ("Autorizado por", 12, "calc")]
encabezados(ws, 5, kdx_cols)
escribir_filas(ws, 6, S.kardex, {2: FECHA, 8: NUM, 9: NUM, 10: COP})
for nm, col in [("kdx_ID", "A"), ("kdx_Fecha", "B"), ("kdx_Tipo", "E"), ("kdx_SKU", "F"), ("kdx_Ent", "H"),
                ("kdx_Sal", "I")]:
    rango(nm, ws, col, 6, DET_MAX)

# ====================================================== CLIENTES
ws = sh["cli"]
titulo(ws, "CLIENTES · CRM", 18, "Ficha individual arriba; tabla maestra abajo. Totales y clasificación automáticos.")
marcador_botones(ws, "Botones: NUEVO CLIENTE · VER FICHA · ACTUALIZAR")
CLI_HR, CLI_R1, CLI_R2 = 16, 17, 516
cli_cols = [("Código", 10, "calc"), ("Nombre / Razón social", 28, "in"), ("Tipo doc.", 7, "in"),
            ("Cédula / NIT", 14, "in"), ("Teléfono", 12, "in"), ("WhatsApp", 12, "in"), ("Correo", 26, "in"),
            ("Dirección", 26, "in"), ("Ciudad", 13, "in"), ("Fecha registro", 11, "in"),
            ("Canal captación", 13, "in"), ("Nº compras", 8, "calc"), ("Total comprado", 13, "calc"),
            ("Saldo pendiente", 13, "calc"), ("Última compra", 11, "calc"), ("Frecuencia (compras/mes)", 11, "calc"),
            ("Clasificación", 14, "calc"), ("Etiqueta (lista)", 34, "calc")]
encabezados(ws, CLI_HR, cli_cols)
estilo_columnas(ws, CLI_R1, CLI_R2, cli_cols, {10: FECHA, 12: NUM, 13: COP, 14: COP, 15: FECHA, 16: "0.0"})
for i, c in enumerate(dp.CLIENTES):
    for j, v in enumerate(c, start=1):
        ws.cell(CLI_R1 + i, j, v)
for r in range(CLI_R1, CLI_R2 + 1):
    ws.cell(r, 12, f'=IF(A{r}="","",COUNTIFS(fac_Cli,A{r},fac_Estado,"VIGENTE"))')
    ws.cell(r, 13, f'=IF(A{r}="","",SUMIFS(fac_Total,fac_Cli,A{r},fac_Estado,"VIGENTE"))')
    ws.cell(r, 14, f'=IF(A{r}="","",SUMIFS(car_Saldo,car_Ter,A{r},car_Tipo,"CxC"))')
    ws.cell(r, 15, f'=IF(OR(A{r}="",L{r}=0),"",SUMPRODUCT(MAX((fac_Cli=A{r})*(fac_Estado="VIGENTE")*fac_Fecha)))')
    ws.cell(r, 16, f'=IF(A{r}="","",L{r}/MAX(1,(TODAY()-J{r})/30))')
    ws.cell(r, 17, f'=IF(A{r}="","",IF(N{r}>0,"CON CARTERA",IF(L{r}=0,IF(TODAY()-J{r}>cfg_DiasInactivo,'
                   f'"INACTIVO","NUEVO"),IF(TODAY()-O{r}>cfg_DiasInactivo,"INACTIVO",IF(L{r}>=cfg_ComprasFrec,'
                   f'"FRECUENTE","NUEVO")))))')
    ws.cell(r, 18, f'=IF(A{r}="","",A{r}&" - "&B{r})')
for nm, col in [("cli_Cod", "A"), ("cli_Nom", "B"), ("cli_TDoc", "C"), ("cli_Doc", "D"), ("cli_Tel", "E"),
                ("cli_WA", "F"), ("cli_Mail", "G"), ("cli_Dir", "H"), ("cli_Ciudad", "I"), ("cli_Fecha", "J"),
                ("cli_Canal", "K"), ("cli_NCompras", "L"), ("cli_Total", "M"), ("cli_Saldo", "N"),
                ("cli_Ultima", "O"), ("cli_Seg", "Q"), ("cli_Label", "R")]:
    rango(nm, ws, col, CLI_R1, CLI_R2)
lista_dinamica("lst_Clientes", ws, "R", CLI_R1, CLI_R2)
dv_unico(ws, f"D{CLI_R1}:D{CLI_R2}", f"D{CLI_R1}", f"$D${CLI_R1}:$D${CLI_R2}", "documento")
dv_lista(ws, f"C{CLI_R1}:C{CLI_R2}", "lst_TiposDoc")
dv_lista(ws, f"I{CLI_R1}:I{CLI_R2}", "lst_Ciudades")
dv_lista(ws, f"K{CLI_R1}:K{CLI_R2}", "lst_Canales")
rng = f"Q{CLI_R1}:Q{CLI_R2}"
for val, bg, fg in [("CON CARTERA", "FFC7CE", "9C0006"), ("INACTIVO", "D9D9D9", "404040"),
                    ("FRECUENTE", "C6EFCE", "006100"), ("NUEVO", "DDEBF7", "1F3864")]:
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{val}"'],
                                  fill=PatternFill("solid", fgColor=bg), font=Font(color=fg, bold=True)))

# ficha individual
ws["B4"] = "FICHA DEL CLIENTE"
ws["B4"].font = f(11, True, NAVY)
ws["D4"] = "Seleccione cliente:"
ws["D4"].font = f(9, True)
ws["D4"].alignment = Alignment(horizontal="right")
ws.merge_cells("E4:H4")
ws["E4"] = dp.CLIENTES[1][0] + " - " + dp.CLIENTES[1][1]
ws["E4"].fill, ws["E4"].protection, ws["E4"].font, ws["E4"].border = INPUT_FILL, UNLOCK, f(10, True, "0000FF"), BORDE
dv_lista(ws, "E4", "lst_Clientes", "Elija un cliente para ver su ficha.")
ws["T4"] = '=IF(E4="","",LEFT(E4,FIND(" - ",E4)-1))'
ws["T4"].font = f(8, color="808080")
ficha = [("Documento", "cli_Doc", None), ("Teléfono", "cli_Tel", None), ("WhatsApp", "cli_WA", None),
         ("Correo", "cli_Mail", None), ("Dirección", "cli_Dir", None), ("Ciudad", "cli_Ciudad", None),
         ("Registro", "cli_Fecha", FECHA), ("Canal captación", "cli_Canal", None),
         ("Nº compras", "cli_NCompras", NUM), ("Total comprado", "cli_Total", COP),
         ("Saldo pendiente", "cli_Saldo", COP), ("Última compra", "cli_Ultima", FECHA),
         ("Clasificación", "cli_Seg", None)]
for i, (lab, nm, fmt) in enumerate(ficha):
    rr = 5 + (i % 7)
    cl = 2 if i < 7 else 6
    ws.cell(rr, cl, lab).font = f(9, True, GRIS)
    c = ws.cell(rr, cl + 1, f'=IFERROR(INDEX({nm},MATCH($T$4,cli_Cod,0)),"")')
    c.font, c.border, c.fill = f(9), BORDE, CALC_FILL
    if fmt:
        c.number_format = fmt
    if cl == 2:
        ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=4)
ws["J4"] = "ÚLTIMAS COMPRAS"
ws["J4"].font = f(10, True, NAVY)
for i, t in enumerate(["Comprobante", "Canal", "Estado", "Total", "Saldo", "Fecha"]):
    c = ws.cell(5, 10 + i, t)
    c.font, c.fill, c.border = f(9, True, "FFFFFF"), F_TEAL, BORDE
ws["T5"] = '=COUNTIF(fac_Cli,$T$4)'
ws["T5"].font = f(8, color="808080")
for k in range(8):
    rr = 6 + k
    ws.cell(rr, 20, f'=IF($T$5-{k}<1,"",MATCH($T$4&"|"&($T$5-{k}),fac_Key,0))').font = f(8, color="808080")
    for i, (nm, fmt) in enumerate([("fac_Num", None), ("fac_Canal", None), ("fac_Estado", None),
                                   ("fac_Total", COP), ("fac_Saldo", COP), ("fac_Fecha", FECHA)]):
        c = ws.cell(rr, 10 + i, f'=IF($T{rr}="","",INDEX({nm},$T{rr}))')
        c.font, c.border = f(9), BORDE
        if fmt:
            c.number_format = fmt
ws["T3"] = "auxiliares ficha"
ws.freeze_panes = None
ws["T3"].font = f(7, color="808080")

# ====================================================== PROVEEDORES
ws = sh["prov"]
titulo(ws, "PROVEEDORES", 20, "Conectado con compras, inventario y cuentas por pagar.")
marcador_botones(ws, "Botones: NUEVO PROVEEDOR · ACTUALIZAR")
PRV_R1, PRV_R2 = 6, 205
prv_cols = [("Código", 9, "calc"), ("Razón social", 30, "in"), ("NIT", 13, "in"), ("Contacto", 16, "in"),
            ("Teléfono", 12, "in"), ("WhatsApp", 12, "in"), ("Correo", 28, "in"), ("Dirección", 20, "in"),
            ("Ciudad", 11, "in"), ("Productos suministrados", 24, "in"), ("Precios / condiciones", 28, "in"),
            ("Banco", 12, "in"), ("Tipo cuenta", 10, "in"), ("Nº cuenta", 15, "in"), ("Nº compras", 8, "calc"),
            ("Total compras", 13, "calc"), ("Abonos registrados", 13, "calc"), ("Saldo pendiente", 13, "calc"),
            ("Última compra", 11, "calc"), ("Etiqueta (lista)", 34, "calc")]
encabezados(ws, 5, prv_cols)
estilo_columnas(ws, PRV_R1, PRV_R2, prv_cols, {15: NUM, 16: COP, 17: COP, 18: COP, 19: FECHA})
for i, p in enumerate(dp.PROVEEDORES):
    for j, v in enumerate(p, start=1):
        ws.cell(PRV_R1 + i, j, v)
for r in range(PRV_R1, PRV_R2 + 1):
    ws.cell(r, 15, f'=IF(A{r}="","",COUNTIFS(com_Prov,A{r},com_Estado,"VIGENTE"))')
    ws.cell(r, 16, f'=IF(A{r}="","",SUMIFS(com_Total,com_Prov,A{r},com_Estado,"VIGENTE"))')
    ws.cell(r, 17, f'=IF(A{r}="","",SUMIFS(ab_Valor,ab_Ter,A{r},ab_Tipo,"CxP"))')
    ws.cell(r, 18, f'=IF(A{r}="","",SUMIFS(car_Saldo,car_Ter,A{r},car_Tipo,"CxP"))')
    ws.cell(r, 19, f'=IF(OR(A{r}="",O{r}=0),"",SUMPRODUCT(MAX((com_Prov=A{r})*(com_Estado="VIGENTE")*com_Fecha)))')
    ws.cell(r, 20, f'=IF(A{r}="","",A{r}&" - "&B{r})')
for nm, col in [("prv_Cod", "A"), ("prv_Nom", "B"), ("prv_NIT", "C"), ("prv_Tel", "E"), ("prv_Ciudad", "I"),
                ("prv_Label", "T")]:
    rango(nm, ws, col, PRV_R1, PRV_R2)
lista_dinamica("lst_Proveedores", ws, "T", PRV_R1, PRV_R2)
dv_unico(ws, f"C{PRV_R1}:C{PRV_R2}", f"C{PRV_R1}", f"$C${PRV_R1}:$C${PRV_R2}", "NIT")
dv_lista(ws, f"I{PRV_R1}:I{PRV_R2}", "lst_Ciudades")

# ====================================================== VENDEDORES
ws = sh["vend"]
titulo(ws, "VENDEDORES · DESEMPEÑO", 15, "Ventas y cumplimiento del mes en curso (según la fecha del equipo).")
marcador_botones(ws, "Botones: NUEVO VENDEDOR · ACTUALIZAR")
VEN_R1, VEN_R2 = 6, 55
ven_cols = [("Código", 9, "calc"), ("Nombre", 22, "in"), ("Identificación", 13, "in"), ("Usuario", 10, "in"),
            ("Rol", 13, "in"), ("Estado", 9, "in"), ("Canal principal", 13, "in"), ("% Comisión", 9, "in"),
            ("Meta mensual", 13, "in"), ("Ventas mes", 13, "calc"), ("Pedidos mes", 8, "calc"),
            ("Cumplimiento", 11, "calc"), ("Ventas históricas", 14, "calc"), ("Utilidad mes", 13, "calc"),
            ("Etiqueta (lista)", 26, "calc")]
encabezados(ws, 5, ven_cols)
estilo_columnas(ws, VEN_R1, VEN_R2, ven_cols, {8: "0.0%", 9: COP, 10: COP, 11: NUM, 12: PCT, 13: COP, 14: COP})
for i, v in enumerate(dp.VENDEDORES):
    for j, x in enumerate(v, start=1):
        ws.cell(VEN_R1 + i, j, x)
MES_INI = "DATE(YEAR(TODAY()),MONTH(TODAY()),1)"
for r in range(VEN_R1, VEN_R2 + 1):
    crit = f'fac_Vend,A{r},fac_Estado,"VIGENTE",fac_Fecha,">="&{MES_INI},fac_Fecha,"<="&TODAY()'
    ws.cell(r, 10, f'=IF(A{r}="","",SUMIFS(fac_Total,{crit}))')
    ws.cell(r, 11, f'=IF(A{r}="","",COUNTIFS({crit}))')
    ws.cell(r, 12, f'=IF(OR(A{r}="",N(I{r})=0),"",J{r}/I{r})')
    ws.cell(r, 13, f'=IF(A{r}="","",SUMIFS(fac_Total,fac_Vend,A{r},fac_Estado,"VIGENTE"))')
    ws.cell(r, 14, f'=IF(A{r}="","",SUMIFS(fac_Util,{crit}))')
    ws.cell(r, 15, f'=IF(A{r}="","",A{r}&" - "&B{r})')
for nm, col in [("ven_Cod", "A"), ("ven_Nom", "B"), ("ven_Usuario", "D"), ("ven_Estado", "F"),
                ("ven_Pct", "H"), ("ven_Meta", "I"), ("ven_Label", "O")]:
    rango(nm, ws, col, VEN_R1, VEN_R2)
lista_dinamica("lst_Vendedores", ws, "O", VEN_R1, VEN_R2)
dv_lista(ws, f"E{VEN_R1}:E{VEN_R2}", "lst_Roles")
dv_lista(ws, f"F{VEN_R1}:F{VEN_R2}", "lst_Estados")
dv_lista(ws, f"G{VEN_R1}:G{VEN_R2}", "lst_Canales")
ws.conditional_formatting.add(f"L{VEN_R1}:L{VEN_R2}", CellIsRule(operator="greaterThanOrEqual", formula=["1"],
                              fill=PatternFill("solid", fgColor="C6EFCE")))
ws.conditional_formatting.add(f"L{VEN_R1}:L{VEN_R2}", CellIsRule(operator="lessThan", formula=["0.5"],
                              fill=PatternFill("solid", fgColor="FFC7CE")))
nota(ws, "N5", "Columna oculta para usuarios sin permiso VER_COSTOS.")

# ====================================================== BD_FACTURAS
ws = sh["dbfac"]
titulo(ws, "BD_FACTURAS · ENCABEZADO DE COMPROBANTES DE VENTA", 24, "Tabla del sistema. No editar a mano.")
fac_cols = [("Nº comprobante", 12, "calc"), ("Fecha", 11, "calc"), ("Cód. cliente", 10, "calc"),
            ("Cliente", 24, "calc"), ("Documento", 12, "calc"), ("Vendedor", 9, "calc"), ("Canal", 12, "calc"),
            ("Forma de pago", 12, "calc"), ("Estado de pago", 13, "calc"), ("Subtotal bruto", 12, "calc"),
            ("Descuento", 11, "calc"), ("Total", 12, "calc"), ("Abono inicial", 12, "calc"),
            ("Saldo inicial", 12, "calc"), ("Costo total", 12, "calc"), ("Utilidad", 12, "calc"),
            ("Estado", 10, "calc"), ("Usuario", 9, "calc"), ("Registrado", 15, "calc"), ("Ruta PDF", 30, "calc"),
            ("Motivo anulación", 26, "calc"), ("Clave única", 14, "calc"), ("Clave cliente|n", 14, "calc"),
            ("Fecha anulación", 11, "calc")]
encabezados(ws, 5, fac_cols)
escribir_filas(ws, 6, S.facturas, {2: FECHA, 10: COP, 11: COP, 12: COP, 13: COP, 14: COP, 15: COP, 16: COP,
                                   19: FECHAHORA, 24: FECHA})
for nm, col in [("fac_Num", "A"), ("fac_Fecha", "B"), ("fac_Cli", "C"), ("fac_CliNom", "D"), ("fac_Vend", "F"),
                ("fac_Canal", "G"), ("fac_Forma", "H"), ("fac_EstPago", "I"), ("fac_Total", "L"),
                ("fac_Saldo", "N"), ("fac_Costo", "O"), ("fac_Util", "P"), ("fac_Estado", "Q"),
                ("fac_Clave", "V"), ("fac_Key", "W")]:
    rango(nm, ws, col, 6, DATA_MAX)
ws.conditional_formatting.add("Q6:Q5000", CellIsRule(operator="notEqual", formula=['"VIGENTE"'],
                              font=Font(color="C00000", bold=True)))

ws = sh["dbdet"]
titulo(ws, "BD_FACT_DETALLE · LÍNEAS VENDIDAS", 19, "Costo unitario congelado al momento de la venta.")
det_cols = [("Nº comprobante", 12, "calc"), ("Fecha", 11, "calc"), ("Canal", 12, "calc"), ("Vendedor", 9, "calc"),
            ("Cliente", 10, "calc"), ("SKU", 14, "calc"), ("Producto", 24, "calc"), ("Categoría", 14, "calc"),
            ("Talla", 6, "calc"), ("Color", 8, "calc"), ("Cantidad", 8, "calc"), ("Precio unit.", 11, "calc"),
            ("Desc. %", 7, "calc"), ("Subtotal", 12, "calc"), ("Costo unit.", 11, "calc"),
            ("Costo total", 12, "calc"), ("Utilidad", 12, "calc"), ("Estado", 10, "calc"), ("Clave línea", 14, "calc")]
encabezados(ws, 5, det_cols)
escribir_filas(ws, 6, S.detalle, {2: FECHA, 11: NUM, 12: COP, 13: "0%", 14: COP, 15: COP, 16: COP, 17: COP})
for nm, col in [("det_Fact", "A"), ("det_Fecha", "B"), ("det_Canal", "C"), ("det_Vend", "D"), ("det_Cli", "E"),
                ("det_SKU", "F"), ("det_Prod", "G"), ("det_Cat", "H"), ("det_Talla", "I"), ("det_Color", "J"),
                ("det_Cant", "K"), ("det_Sub", "N"), ("det_CostoT", "P"), ("det_Util", "Q"),
                ("det_Estado", "R"), ("det_Key", "S")]:
    rango(nm, ws, col, 6, DET_MAX)

ws = sh["dbcom"]
titulo(ws, "BD_COMPRAS · ENCABEZADO DE COMPRAS", 17, "Tabla del sistema. No editar a mano.")
dbc_cols = [("Nº compra", 10, "calc"), ("Fecha", 11, "calc"), ("Proveedor", 9, "calc"), ("Nombre", 28, "calc"),
            ("Factura proveedor", 13, "calc"), ("Subtotal", 12, "calc"), ("Transporte", 11, "calc"),
            ("Costos adicionales", 11, "calc"), ("Total", 12, "calc"), ("Forma de pago", 12, "calc"),
            ("Estado de pago", 13, "calc"), ("Pagado", 12, "calc"), ("Saldo", 12, "calc"), ("Estado", 10, "calc"),
            ("Usuario", 9, "calc"), ("Registrado", 15, "calc"), ("Clave única", 14, "calc")]
encabezados(ws, 5, dbc_cols)
escribir_filas(ws, 6, S.compras, {2: FECHA, 6: COP, 7: COP, 8: COP, 9: COP, 12: COP, 13: COP, 16: FECHAHORA})
for nm, col in [("com_Num", "A"), ("com_Fecha", "B"), ("com_Prov", "C"), ("com_FactProv", "E"),
                ("com_Total", "I"), ("com_Estado", "N")]:
    rango(nm, ws, col, 6, DATA_MAX)

ws = sh["dbcomd"]
titulo(ws, "BD_COMPRA_DETALLE · LÍNEAS COMPRADAS", 11, "Prorrateo de transporte y costos adicionales por valor.")
dbcd_cols = [("Nº compra", 10, "calc"), ("Fecha", 11, "calc"), ("Proveedor", 9, "calc"), ("SKU", 14, "calc"),
             ("Producto", 26, "calc"), ("Cantidad", 8, "calc"), ("Costo unit.", 11, "calc"),
             ("Subtotal", 12, "calc"), ("Prorrateo unit.", 11, "calc"), ("Costo real unit.", 12, "calc"),
             ("Estado", 10, "calc")]
encabezados(ws, 5, dbcd_cols)
escribir_filas(ws, 6, S.compra_det, {2: FECHA, 6: NUM, 7: COP, 8: COP, 9: '"$" #,##0.00', 10: '"$" #,##0.00'})

ws = sh["dbgui"]
titulo(ws, "BD_GUIAS · GUÍAS DE ENVÍO", 17, "Tabla del sistema. El estado se puede actualizar en la columna N.")
dbg_cols = [("Nº guía", 10, "calc"), ("Fecha", 11, "calc"), ("Comprobante", 12, "calc"), ("Cód. cliente", 10, "calc"),
            ("Cliente", 22, "calc"), ("Teléfono", 12, "calc"), ("Dirección", 24, "calc"), ("Ciudad", 12, "calc"),
            ("Transportadora", 14, "calc"), ("Nº seguimiento", 13, "calc"), ("Productos", 34, "calc"),
            ("Unidades", 8, "calc"), ("Valor declarado", 12, "calc"), ("Estado", 12, "in"), ("Usuario", 9, "calc"),
            ("Registrado", 15, "calc"), ("Ruta PDF", 28, "calc")]
encabezados(ws, 5, dbg_cols)
escribir_filas(ws, 6, S.guias, {2: FECHA, 12: NUM, 13: COP, 16: FECHAHORA})
for r in range(6, 6 + len(S.guias)):
    ws.cell(r, 14).protection = UNLOCK
dv_lista(ws, "N6:N5000", "lst_EstadosGuia")
rango("gui_Num", ws, "A", 6, DATA_MAX)
rango("gui_Fact", ws, "C", 6, DATA_MAX)

ws = sh["dbab"]
titulo(ws, "BD_ABONOS · ABONOS A CARTERA", 11, "Abonos a cuentas por cobrar (clientes) y por pagar (proveedores).")
dba_cols = [("ID abono", 10, "calc"), ("Fecha", 11, "calc"), ("ID cartera", 10, "calc"), ("Tipo", 6, "calc"),
            ("Documento", 12, "calc"), ("Tercero", 10, "calc"), ("Valor", 12, "calc"), ("Método", 12, "calc"),
            ("Usuario", 9, "calc"), ("Observación", 24, "calc"), ("Registrado", 15, "calc")]
encabezados(ws, 5, dba_cols)
escribir_filas(ws, 6, S.abonos, {2: FECHA, 7: COP, 11: FECHAHORA})
for nm, col in [("ab_ID", "A"), ("ab_Fecha", "B"), ("ab_Car", "C"), ("ab_Tipo", "D"), ("ab_Ter", "F"),
                ("ab_Valor", "G")]:
    rango(nm, ws, col, 6, DATA_MAX)

ws = sh["dbcie"]
titulo(ws, "BD_CIERRES · CIERRES DE CAJA", 11, "Saldo del sistema vs. saldo contado por cuenta.")
dbci_cols = [("ID", 9, "calc"), ("Tipo", 10, "calc"), ("Fecha cierre", 11, "calc"), ("Desde", 11, "calc"),
             ("Hasta", 11, "calc"), ("Cuenta", 13, "calc"), ("Saldo sistema", 13, "calc"),
             ("Saldo contado", 13, "calc"), ("Diferencia", 12, "calc"), ("Usuario", 9, "calc"),
             ("Observación", 28, "calc")]
encabezados(ws, 5, dbci_cols)
escribir_filas(ws, 6, S.cierres, {3: FECHA, 4: FECHA, 5: FECHA, 7: COP, 8: COP, 9: COP})

ws = sh["dblog"]
titulo(ws, "BD_BITACORA · REGISTRO DE ACTIVIDAD", 8, "Quién hizo qué, cuándo y desde qué equipo.")
dbl_cols = [("ID", 11, "calc"), ("Fecha y hora", 16, "calc"), ("Usuario", 10, "calc"), ("Módulo", 16, "calc"),
            ("Acción", 13, "calc"), ("Documento", 13, "calc"), ("Detalle", 44, "calc"), ("Equipo", 14, "calc")]
encabezados(ws, 5, dbl_cols)
escribir_filas(ws, 6, S.bitacora, {2: FECHAHORA})

# ====================================================== CARTERA
ws = sh["cart"]
titulo(ws, "CARTERA · CUENTAS POR COBRAR Y POR PAGAR", 13,
       "Las crean las macros al guardar ventas/compras a crédito. Abonos parciales con el botón REGISTRAR ABONO.")
marcador_botones(ws, "Botones: REGISTRAR ABONO (fila seleccionada) · ACTUALIZAR")
CAR_R1, CAR_R2 = 6, 2005
car_cols = [("ID", 10, "calc"), ("Tipo", 6, "calc"), ("Documento", 12, "calc"), ("Cód. tercero", 10, "calc"),
            ("Cliente / proveedor", 28, "calc"), ("Fecha", 11, "calc"), ("Valor inicial", 13, "calc"),
            ("Abonos", 13, "calc"), ("Saldo", 13, "calc"), ("Vencimiento", 11, "in"), ("Días de mora", 9, "calc"),
            ("Estado", 11, "calc"), ("Anulado", 8, "calc")]
encabezados(ws, 5, car_cols)
estilo_columnas(ws, CAR_R1, CAR_R2, car_cols, {6: FECHA, 7: COP, 8: COP, 9: COP, 10: FECHA, 11: NUM})
for i, c in enumerate(S.cartera):
    r = CAR_R1 + i
    for j, v in zip([1, 2, 3, 4, 5, 6, 7, 10, 13], c):
        ws.cell(r, j, v)
for r in range(CAR_R1, CAR_R2 + 1):
    ws.cell(r, 8, f'=IF(A{r}="","",SUMIFS(ab_Valor,ab_Car,A{r}))')
    ws.cell(r, 9, f'=IF(A{r}="","",IF(M{r}="SI",0,MAX(0,G{r}-H{r})))')
    ws.cell(r, 11, f'=IF(A{r}="","",IF(I{r}<=0,0,MAX(0,TODAY()-J{r})))')
    ws.cell(r, 12, f'=IF(A{r}="","",IF(M{r}="SI","ANULADO",IF(I{r}<=0,"PAGADO",IF(TODAY()>J{r},"VENCIDO","AL DÍA"))))')
for nm, col in [("car_ID", "A"), ("car_Tipo", "B"), ("car_Doc", "C"), ("car_Ter", "D"), ("car_Saldo", "I"),
                ("car_Mora", "K"), ("car_Estado", "L")]:
    rango(nm, ws, col, CAR_R1, CAR_R2)
rng = f"L{CAR_R1}:L{CAR_R2}"
for val, bg, fg in [("VENCIDO", "FFC7CE", "9C0006"), ("AL DÍA", "FFEB9C", "9C5700"), ("PAGADO", "C6EFCE", "006100"),
                    ("ANULADO", "D9D9D9", "404040")]:
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{val}"'],
                                  fill=PatternFill("solid", fgColor=bg), font=Font(color=fg, bold=True)))
ws["D4"], ws["E4"] = "Por cobrar (CxC):", '=SUMIFS(car_Saldo,car_Tipo,"CxC")'
ws["G4"], ws["H4"] = "Por pagar (CxP):", '=SUMIFS(car_Saldo,car_Tipo,"CxP")'
ws["J4"], ws["K4"] = "Vencido CxC:", '=SUMIFS(car_Saldo,car_Tipo,"CxC",car_Estado,"VENCIDO")'
for c in ("D4", "G4", "J4"):
    ws[c].font, ws[c].alignment = f(9, True), Alignment(horizontal="right")
for c in ("E4", "H4", "K4"):
    ws[c].font, ws[c].number_format = f(10, True, NAVY), COP
ws["K4"].font = f(10, True, "C00000")

# ====================================================== INGRESOS / EGRESOS
ws = sh["ing"]
titulo(ws, "INGRESOS", 10, "Ventas y abonos los registra la macro. Otros ingresos: botón NUEVO INGRESO.")
marcador_botones(ws, "Botones: NUEVO INGRESO · ACTUALIZAR")
ing_cols = [("ID", 10, "calc"), ("Fecha", 11, "calc"), ("Tipo", 13, "calc"), ("Método / cuenta", 13, "calc"),
            ("Valor", 13, "calc"), ("Documento", 12, "calc"), ("Tercero", 12, "calc"), ("Concepto / detalle", 32, "calc"),
            ("Usuario", 9, "calc"), ("Estado", 10, "calc")]
encabezados(ws, 5, ing_cols)
escribir_filas(ws, 6, S.ingresos, {2: FECHA, 5: COP})
for nm, col in [("ing_ID", "A"), ("ing_Fecha", "B"), ("ing_Tipo", "C"), ("ing_Met", "D"), ("ing_Valor", "E"),
                ("ing_Doc", "F"), ("ing_Estado", "J")]:
    rango(nm, ws, col, 6, DATA_MAX)
dv_lista(ws, "C6:C5000", "lst_TiposIngreso")
dv_lista(ws, "D6:D5000", "lst_Metodos")
ws["G4"], ws["H4"] = "Total vigente:", '=SUMIFS(ing_Valor,ing_Estado,"VIGENTE")'
ws["G4"].font, ws["G4"].alignment = f(9, True), Alignment(horizontal="right")
ws["H4"].font, ws["H4"].number_format = f(10, True, NAVY), COP
ws.conditional_formatting.add("J6:J5000", CellIsRule(operator="equal", formula=['"ANULADO"'],
                              font=Font(color="C00000", bold=True)))

ws = sh["egr"]
titulo(ws, "EGRESOS", 10, "Compras y comisiones los registra la macro. Demás gastos: botón NUEVO EGRESO.")
marcador_botones(ws, "Botones: NUEVO EGRESO · ANULAR EGRESO · ACTUALIZAR")
egr_cols = [("ID", 10, "calc"), ("Fecha", 11, "calc"), ("Concepto", 16, "calc"), ("Método / cuenta", 13, "calc"),
            ("Valor", 13, "calc"), ("Documento", 16, "calc"), ("Beneficiario", 20, "calc"), ("Descripción", 30, "calc"),
            ("Usuario", 9, "calc"), ("Estado", 10, "calc")]
encabezados(ws, 5, egr_cols)
escribir_filas(ws, 6, S.egresos, {2: FECHA, 5: COP})
for nm, col in [("egr_ID", "A"), ("egr_Fecha", "B"), ("egr_Con", "C"), ("egr_Met", "D"), ("egr_Valor", "E"),
                ("egr_Doc", "F"), ("egr_Ben", "G"), ("egr_Estado", "J")]:
    rango(nm, ws, col, 6, DATA_MAX)
dv_lista(ws, "C6:C5000", "lst_ConceptosEgr")
dv_lista(ws, "D6:D5000", "lst_Metodos")
ws["G4"], ws["H4"] = "Total vigente:", '=SUMIFS(egr_Valor,egr_Estado,"VIGENTE")'
ws["G4"].font, ws["G4"].alignment = f(9, True), Alignment(horizontal="right")
ws["H4"].font, ws["H4"].number_format = f(10, True, NAVY), COP
ws.conditional_formatting.add("J6:J5000", CellIsRule(operator="equal", formula=['"ANULADO"'],
                              font=Font(color="C00000", bold=True)))

# ====================================================== GASTOS FIJOS
ws = sh["gf"]
titulo(ws, "GASTOS FIJOS MENSUALES", 11, "Plantilla. El botón GENERAR EGRESOS DEL MES crea los egresos sin duplicar.")
marcador_botones(ws, "Botones: NUEVO GASTO FIJO · GENERAR EGRESOS DEL MES")
gf_cols = [("Código", 8, "calc"), ("Concepto", 14, "in"), ("Descripción", 30, "in"), ("Valor mensual", 13, "in"),
           ("Día de pago", 8, "in"), ("Método / cuenta", 13, "in"), ("Beneficiario", 22, "in"), ("Activo", 7, "in"),
           ("Último mes generado", 12, "calc"), ("Vence este mes", 11, "calc"), ("Estado del mes", 12, "calc")]
encabezados(ws, 5, gf_cols)
estilo_columnas(ws, 6, 55, gf_cols, {4: COP, 9: "MM/YYYY", 10: FECHA})
for i, g in enumerate(dp.GASTOS_FIJOS):
    for j, v in enumerate(g, start=1):
        ws.cell(6 + i, j, v)
    ws.cell(6 + i, 9, S.gf_ultimo.get(g[0]))
for r in range(6, 56):
    ws.cell(r, 10, f'=IF(A{r}="","",DATE(YEAR(TODAY()),MONTH(TODAY()),MIN(E{r},DAY(DATE(YEAR(TODAY()),MONTH(TODAY())+1,0)))))')
    ws.cell(r, 11, f'=IF(A{r}="","",IF(H{r}<>"SI","INACTIVO",IF(N(I{r})=DATE(YEAR(TODAY()),MONTH(TODAY()),1),'
                   f'"GENERADO",IF(TODAY()>J{r},"VENCIDO","PENDIENTE"))))')
dv_lista(ws, "B6:B55", "lst_ConceptosEgr")
dv_lista(ws, "F6:F55", "lst_Metodos")
dv_lista(ws, "H6:H55", "lst_SiNo")
for val, bg in [("VENCIDO", "FFC7CE"), ("PENDIENTE", "FFEB9C"), ("GENERADO", "C6EFCE")]:
    ws.conditional_formatting.add("K6:K55", CellIsRule(operator="equal", formula=[f'"{val}"'],
                                  fill=PatternFill("solid", fgColor=bg)))
ws["C4"], ws["D4"] = "Total mensual activo:", '=SUMIFS(D6:D55,H6:H55,"SI")'
ws["C4"].font, ws["C4"].alignment = f(9, True), Alignment(horizontal="right")
ws["D4"].font, ws["D4"].number_format = f(10, True, NAVY), COP
rango("gf_Estado", ws, "K", 6, 55)

# ====================================================== FORMULARIOS: logo
from PIL import Image as PILImage, ImageDraw, ImageFont
logo_path = os.path.join(AQUI, "logo_ejemplo.png")
img = PILImage.new("RGBA", (360, 360), (0, 0, 0, 0))
d = ImageDraw.Draw(img)
d.rounded_rectangle([0, 0, 359, 359], radius=60, fill=(31, 56, 100, 255))
try:
    fnt = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 150)
    fnt2 = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 40)
except OSError:
    fnt = fnt2 = ImageFont.load_default()
d.text((180, 150), "TU", font=fnt, fill=(255, 255, 255, 255), anchor="mm")
d.text((180, 285), "SU LOGO", font=fnt2, fill=(189, 215, 238, 255), anchor="mm")
img.save(logo_path)


def poner_logo(ws, ancla, alto=78):
    im = XLImage(logo_path)
    im.width = im.height = alto
    ws.add_image(im, ancla)


def encabezado_empresa(ws, titulo_doc, num_formula, col_ini="D"):
    ws["D5"] = "=cfg_Empresa"
    ws["D5"].font = f(12, True, NAVY)
    ws["D6"] = '="NIT "&cfg_NIT'
    ws["D7"] = '=cfg_Direccion&" · "&cfg_Ciudad'
    ws["D8"] = '="Tel. "&cfg_Telefono&" · "&cfg_Correo'
    for c in ("D6", "D7", "D8"):
        ws[c].font = f(9, color=GRIS)
    ws.merge_cells("H5:K5")
    ws["H5"] = titulo_doc
    ws["H5"].font = f(12, True, "FFFFFF")
    ws["H5"].fill = F_NAVY
    ws["H5"].alignment = Alignment(horizontal="center", vertical="center")
    for lab, rr in [("No.", 6), ("Fecha", 7), ("Estado", 8)]:
        ws.cell(rr, 8, lab).font = f(9, True)
        ws.merge_cells(start_row=rr, start_column=9, end_row=rr, end_column=11)
        c = ws.cell(rr, 9)
        c.border = BORDE
        c.alignment = Alignment(horizontal="center")
    ws["I6"] = num_formula
    ws["I6"].font = f(12, True, "C00000")
    ws["I7"].number_format = FECHA
    ws["I7"].fill, ws["I7"].protection = INPUT_FILL, UNLOCK
    ws["I7"].font = f(10, color="0000FF")
    ws["I8"] = "NUEVO"
    ws["I8"].font = f(10, True)
    poner_logo(ws, "B5")
    clave_borrador(ws)


def clave_borrador(ws):
    ws["M7"] = "Clave borrador"
    ws["M7"].font = f(7, color="808080")
    ws["M8"] = "BORRADOR-INICIAL"
    ws["M8"].font = f(7, color="808080")


def etiqueta(ws, ref, texto):
    ws[ref] = texto
    ws[ref].font = f(9, True, GRIS)


def entrada(ws, ref, valor=None, merge_to=None, fmt=None, lista=None, ayuda=None):
    if merge_to:
        ws.merge_cells(f"{ref}:{merge_to}")
    c = ws[ref]
    if valor is not None:
        c.value = valor
    c.fill, c.protection, c.border, c.font = INPUT_FILL, UNLOCK, BORDE, f(10, color="0000FF")
    if fmt:
        c.number_format = fmt
    if lista:
        dv_lista(ws, ref if not merge_to else f"{ref}:{merge_to}", lista, ayuda)


def calculada(ws, ref, formula, merge_to=None, fmt=None, bold=False):
    if merge_to:
        ws.merge_cells(f"{ref}:{merge_to}")
    c = ws[ref]
    c.value = formula
    c.fill, c.border, c.font = CALC_FILL, BORDE, f(10, bold)
    if fmt:
        c.number_format = fmt


def anchos(ws, d):
    for k, v in d.items():
        ws.column_dimensions[k].width = v


# ====================================================== FACTURACIÓN
ws = sh["fact"]
titulo(ws, "FACTURACIÓN · COMPROBANTE DE VENTA", 16)
marcador_botones(ws, "Botones: NUEVA FACTURA · GUARDAR · IMPRIMIR · PDF · BUSCAR · ANULAR / DEVOLUCIÓN · LIMPIAR")
anchos(ws, {"B": 13, "C": 30, "D": 22, "E": 7, "F": 10, "G": 7, "H": 14, "I": 12, "J": 8, "K": 14, "L": 3,
            "M": 15, "N": 12, "O": 12, "P": 11})
encabezado_empresa(ws, "COMPROBANTE DE VENTA",
                   '=cfg_PrefFactura&"-"&TEXT(cfg_NumFactura,"000000")&" (pendiente)"')
ws["I7"] = dp.HOY
ws.merge_cells("B9:K9")
ws["B9"] = "=cfg_Leyenda"
ws["B9"].font = f(8, True, "C00000", italic=True)
ws["B9"].alignment = Alignment(horizontal="center")

ejemplo_cli = dp.CLIENTES[7]
etiqueta(ws, "B11", "Cliente*")
entrada(ws, "C11", f"{ejemplo_cli[0]} - {ejemplo_cli[1]}", "F11", lista="lst_Clientes",
        ayuda="Escriba o elija el cliente. Si no existe, créelo en CLIENTES.")
etiqueta(ws, "B12", "Documento")
calculada(ws, "C12", '=IFERROR(INDEX(cli_TDoc,MATCH($M$11,cli_Cod,0))&" "&INDEX(cli_Doc,MATCH($M$11,cli_Cod,0)),"")')
etiqueta(ws, "D12", "Teléfono / WhatsApp")
calculada(ws, "E12", '=IFERROR(INDEX(cli_Tel,MATCH($M$11,cli_Cod,0)),"")', "F12")
etiqueta(ws, "B13", "Dirección")
calculada(ws, "C13", '=IFERROR(INDEX(cli_Dir,MATCH($M$11,cli_Cod,0)),"")', "D13")
etiqueta(ws, "B14", "Ciudad")
calculada(ws, "C14", '=IFERROR(INDEX(cli_Ciudad,MATCH($M$11,cli_Cod,0)),"")')
etiqueta(ws, "D14", "Correo")
calculada(ws, "E14", '=IFERROR(INDEX(cli_Mail,MATCH($M$11,cli_Cod,0)),"")', "F14")
etiqueta(ws, "H11", "Vendedor*")
entrada(ws, "I11", "VEN-01 - Laura Gómez", "K11", lista="lst_Vendedores")
etiqueta(ws, "H12", "Canal*")
entrada(ws, "I12", "WhatsApp", "K12", lista="lst_Canales", ayuda="Obligatorio: canal por el que llegó la venta.")
etiqueta(ws, "H13", "Forma de pago*")
entrada(ws, "I13", "Nequi", "K13", lista="lst_Metodos")
etiqueta(ws, "H14", "Estado de pago*")
entrada(ws, "I14", "ABONO PARCIAL", "K14", lista="lst_EstadosPago")
etiqueta(ws, "H15", "Abono inicial")
entrada(ws, "I15", 100000, "K15", fmt=COP)
nota(ws, "I15", "Solo si el estado de pago es ABONO PARCIAL. Si es PAGADA se toma el total; si es CRÉDITO, cero.")
# auxiliares
ws["M10"] = "AUXILIARES (no se imprimen)"
ws["M10"].font = f(8, True, "808080")
ws["M11"] = '=IF(C11="","",LEFT(C11,FIND(" - ",C11&" - ")-1))'
ws["M12"] = '=IF(I11="","",LEFT(I11,FIND(" - ",I11&" - ")-1))'
for c in ("M11", "M12"):
    ws[c].font = f(8, color="808080")

hdr = ["#", "SKU / producto*", "Producto", "Talla", "Color", "Disp.", "Cantidad*", "Precio unit.", "Desc. %", "Subtotal"]
for i, t in enumerate(hdr):
    c = ws.cell(17, 2 + i, t)
    c.font, c.fill, c.border = f(9, True, "FFFFFF"), F_NAVY, BORDE
    c.alignment = Alignment(horizontal="center", wrap_text=True)
for i, t in enumerate(["SKU", "Costo unit.", "Costo línea", "Validación"]):
    c = ws.cell(17, 13 + i, t)
    c.font, c.fill, c.border = f(8, True, "FFFFFF"), F_TEAL, BORDE
FL1, FL2 = 18, 32
ejemplo_lineas = {18: ("CAMB-M-NEG", 2, 0), 19: ("GORR-U-NEG", 1, 0.1)}
for r in range(FL1, FL2 + 1):
    ws.cell(r, 2, r - FL1 + 1).font = f(9, color=GRIS)
    ws.cell(r, 2).alignment = Alignment(horizontal="center")
    entrada(ws, f"C{r}")
    if r in ejemplo_lineas:
        k, qy, ds = ejemplo_lineas[r]
        sk = dp.SKU[k]
        ws[f"C{r}"] = f'{k} - {sk["prod"]} {sk["talla"]} {sk["color"]}'
        ws[f"H{r}"] = qy
        ws[f"J{r}"] = ds if ds else None
    ws.cell(r, 13, f'=IF(C{r}="","",LEFT(C{r},FIND(" - ",C{r}&" - ")-1))').font = f(8, color="808080")
    calculada(ws, f"D{r}", f'=IF($M{r}="","",IFERROR(INDEX(inv_Prod,MATCH($M{r},inv_SKU,0)),"SKU NO EXISTE"))')
    calculada(ws, f"E{r}", f'=IF($M{r}="","",IFERROR(INDEX(inv_Talla,MATCH($M{r},inv_SKU,0)),""))')
    calculada(ws, f"F{r}", f'=IF($M{r}="","",IFERROR(INDEX(inv_Color,MATCH($M{r},inv_SKU,0)),""))')
    calculada(ws, f"G{r}", f'=IF($M{r}="","",IFERROR(INDEX(inv_Exist,MATCH($M{r},inv_SKU,0)),0))', fmt=NUM)
    entrada(ws, f"H{r}", fmt=NUM)
    calculada(ws, f"I{r}", f'=IF($M{r}="","",IFERROR(INDEX(inv_Precio,MATCH($M{r},inv_SKU,0)),0))', fmt=COP)
    entrada(ws, f"J{r}", fmt="0%")
    calculada(ws, f"K{r}", f'=IF(OR($M{r}="",N(H{r})=0),"",ROUND(H{r}*I{r}*(1-N(J{r})),0))', fmt=COP)
    for cc, frm, fm in [(14, f'=IF($M{r}="","",IFERROR(INDEX(inv_Costo,MATCH($M{r},inv_SKU,0)),0))', COP),
                        (15, f'=IF(K{r}="","",N{r}*H{r})', COP),
                        (16, f'=IF(OR($M{r}="",N(H{r})=0),"",IF(H{r}>G{r},"SIN STOCK","OK"))', None)]:
        c = ws.cell(r, cc, frm)
        c.font = f(8, color="808080")
        if fm:
            c.number_format = fm
    dv_lista(ws, f"C{r}", "lst_SKU")
dvq = DataValidation(type="whole", operator="between", formula1="1", formula2="9999", allow_blank=True,
                     showErrorMessage=True, error="Cantidad entera mayor que cero.")
ws.add_data_validation(dvq)
dvq.add(f"H{FL1}:H{FL2}")
dvd = DataValidation(type="decimal", operator="between", formula1="0", formula2="0.9", allow_blank=True,
                     showErrorMessage=True, error="Descuento entre 0% y 90%.")
ws.add_data_validation(dvd)
dvd.add(f"J{FL1}:J{FL2}")
ws.conditional_formatting.add(f"G{FL1}:G{FL2}", FormulaRule(formula=[f'AND(N(H{FL1})>0,N(H{FL1})>N(G{FL1}))'],
                              fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True)))
ws.conditional_formatting.add(f"P{FL1}:P{FL2}", CellIsRule(operator="equal", formula=['"SIN STOCK"'],
                              font=Font(color="C00000", bold=True)))
etiqueta(ws, "B34", "Observaciones")
entrada(ws, "C34", "Entrega por Servientrega. Cliente paga el resto contra entrega.", "G37")
ws["C34"].alignment = Alignment(wrap_text=True, vertical="top")
for rr, lab, frm, bold in [(34, "Subtotal", f"=SUMPRODUCT(H{FL1}:H{FL2},I{FL1}:I{FL2})", False),
                           (35, "Descuentos", "=K34-K36", False),
                           (36, "TOTAL", f"=SUM(K{FL1}:K{FL2})", True),
                           (37, "Pagado", '=IF(I14="PAGADA",K36,IF(I14="ABONO PARCIAL",MIN(N(I15),K36),0))', False),
                           (38, "Saldo", "=K36-K37", True)]:
    ws.merge_cells(f"I{rr}:J{rr}")
    ws[f"I{rr}"] = lab
    ws[f"I{rr}"].font = f(10 if not bold else 11, True)
    ws[f"I{rr}"].alignment = Alignment(horizontal="right")
    calculada(ws, f"K{rr}", frm, fmt=COP, bold=bold)
ws["K36"].fill = F_CLARO
ws["K36"].font = f(12, True, NAVY)
ws["M34"] = "Utilidad estimada"
ws["M34"].font = f(8, True, "808080")
ws["N34"] = f"=K36-SUM(O{FL1}:O{FL2})"
ws["M35"] = "Margen"
ws["M35"].font = f(8, True, "808080")
ws["N35"] = '=IF(K36=0,"",N34/K36)'
ws["M36"] = "Líneas sin stock"
ws["M36"].font = f(8, True, "808080")
ws["N36"] = f'=COUNTIF(P{FL1}:P{FL2},"SIN STOCK")'
ws["N34"].number_format, ws["N35"].number_format = COP, PCT
for c in ("N34", "N35", "N36"):
    ws[c].font = f(9, True)
ws.merge_cells("B40:F40")
ws["B40"] = "Recibido por: ______________________________"
ws["B40"].font = f(9, color=GRIS)
ws.merge_cells("H40:K40")
ws["H40"] = "=cfg_Redes&\" · ¡Gracias por su compra!\""
ws["H40"].font = f(9, True, AZUL)
ws["H40"].alignment = Alignment(horizontal="right")
ws.print_area = "B4:K41"
ws.page_setup.orientation = "portrait"
ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_options.horizontalCentered = True
ws.freeze_panes = None

# ====================================================== COMPRAS (formulario)
ws = sh["compra"]
titulo(ws, "COMPRAS · REGISTRO DE COMPRA A PROVEEDOR", 14)
marcador_botones(ws, "Botones: NUEVA COMPRA · GUARDAR · BUSCAR · LIMPIAR · IMPRIMIR")
anchos(ws, {"B": 13, "C": 30, "D": 22, "E": 7, "F": 10, "G": 9, "H": 14, "I": 13, "J": 12, "K": 13, "M": 15})
encabezado_empresa(ws, "COMPRA A PROVEEDOR", '=cfg_PrefCompra&"-"&TEXT(cfg_NumCompra,"00000")&" (pendiente)"')
ws["I7"] = dp.HOY
etiqueta(ws, "B11", "Proveedor*")
entrada(ws, "C11", f"{dp.PROVEEDORES[2][0]} - {dp.PROVEEDORES[2][1]}", "F11", lista="lst_Proveedores")
ws["M11"] = '=IF(C11="","",LEFT(C11,FIND(" - ",C11&" - ")-1))'
ws["M11"].font = f(8, color="808080")
etiqueta(ws, "B12", "NIT")
calculada(ws, "C12", '=IFERROR(INDEX(prv_NIT,MATCH($M$11,prv_Cod,0)),"")')
etiqueta(ws, "D12", "Teléfono")
calculada(ws, "E12", '=IFERROR(INDEX(prv_Tel,MATCH($M$11,prv_Cod,0)),"")', "F12")
etiqueta(ws, "B13", "Factura proveedor Nº*")
entrada(ws, "C13", "IAB-401")
nota(ws, "C13", "Número de la factura del proveedor. El sistema no permite registrarla dos veces.")
etiqueta(ws, "H11", "Forma de pago*")
entrada(ws, "I11", "Nequi", "K11", lista="lst_Metodos")
etiqueta(ws, "H12", "Estado de pago*")
entrada(ws, "I12", "PAGADA", "K12", lista="lst_EstadosPago")
etiqueta(ws, "H13", "Valor pagado")
entrada(ws, "I13", None, "K13", fmt=COP)
nota(ws, "I13", "Solo si es ABONO PARCIAL.")
etiqueta(ws, "H14", "Días de crédito")
entrada(ws, "I14", "=cfg_DiasCreditoProv", "K14", fmt="0")
hdr = ["#", "SKU / producto*", "Producto", "Talla", "Color", "Cantidad*", "Costo unit.*", "Subtotal",
       "Flete prorrat. unit.", "Costo real unit."]
for i, t in enumerate(hdr):
    c = ws.cell(17, 2 + i, t)
    c.font, c.fill, c.border = f(9, True, "FFFFFF"), F_NAVY, BORDE
    c.alignment = Alignment(horizontal="center", wrap_text=True)
ws.row_dimensions[17].height = 28
ej_c = {18: ("GORR-U-ROJ", 10, 12500), 19: ("MORR-U-NEG", 6, 40000)}
for r in range(FL1, FL2 + 1):
    ws.cell(r, 2, r - FL1 + 1).font = f(9, color=GRIS)
    entrada(ws, f"C{r}")
    if r in ej_c:
        k, qy, cu = ej_c[r]
        sk = dp.SKU[k]
        ws[f"C{r}"] = f'{k} - {sk["prod"]} {sk["talla"]} {sk["color"]}'
    ws.cell(r, 13, f'=IF(C{r}="","",LEFT(C{r},FIND(" - ",C{r}&" - ")-1))').font = f(8, color="808080")
    calculada(ws, f"D{r}", f'=IF($M{r}="","",IFERROR(INDEX(inv_Prod,MATCH($M{r},inv_SKU,0)),"SKU NO EXISTE"))')
    calculada(ws, f"E{r}", f'=IF($M{r}="","",IFERROR(INDEX(inv_Talla,MATCH($M{r},inv_SKU,0)),""))')
    calculada(ws, f"F{r}", f'=IF($M{r}="","",IFERROR(INDEX(inv_Color,MATCH($M{r},inv_SKU,0)),""))')
    entrada(ws, f"G{r}", ej_c[r][1] if r in ej_c else None, fmt=NUM)
    entrada(ws, f"H{r}", ej_c[r][2] if r in ej_c else None, fmt=COP)
    calculada(ws, f"I{r}", f'=IF(OR($M{r}="",N(G{r})=0),"",G{r}*H{r})', fmt=COP)
    calculada(ws, f"J{r}", f'=IF(I{r}="","",IF($I$34=0,0,($I$35+$I$36)*(I{r}/$I$34)/G{r}))', fmt='"$" #,##0.00')
    calculada(ws, f"K{r}", f'=IF(I{r}="","",H{r}+J{r})', fmt='"$" #,##0.00')
    dv_lista(ws, f"C{r}", "lst_SKU")
dvq2 = DataValidation(type="whole", operator="between", formula1="1", formula2="99999", allow_blank=True,
                      showErrorMessage=True, error="Cantidad entera mayor que cero.")
ws.add_data_validation(dvq2)
dvq2.add(f"G{FL1}:G{FL2}")
etiqueta(ws, "B34", "Observaciones")
entrada(ws, "C34", None, "F37")
for rr, lab, frm, inp in [(34, "Subtotal productos", f"=SUM(I{FL1}:I{FL2})", False),
                          (35, "Transporte", 15000, True),
                          (36, "Costos adicionales", 0, True),
                          (37, "TOTAL", "=I34+I35+I36", False),
                          (38, "Pagado", '=IF(I12="PAGADA",I37,IF(I12="ABONO PARCIAL",MIN(N(I13),I37),0))', False),
                          (39, "Saldo (cuenta por pagar)", "=I37-I38", False)]:
    ws[f"G{rr}"] = lab
    ws[f"G{rr}"].font = f(9, True)
    ws[f"G{rr}"].alignment = Alignment(horizontal="right")
    ws.merge_cells(f"G{rr}:H{rr}")
    if inp:
        entrada(ws, f"I{rr}", frm, fmt=COP)
    else:
        calculada(ws, f"I{rr}", frm, fmt=COP, bold=(rr in (37, 39)))
ws["I37"].fill = F_CLARO
ws.print_area = "B4:K40"
ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.freeze_panes = None

# ====================================================== GUÍAS (formulario)
ws = sh["guia"]
titulo(ws, "GUÍAS DE ENVÍO", 14)
marcador_botones(ws, "Botones: NUEVA GUÍA · GUARDAR · IMPRIMIR · PDF · BUSCAR · LIMPIAR")
anchos(ws, {"B": 13, "C": 22, "D": 30, "E": 8, "F": 10, "G": 9, "H": 15, "I": 12, "J": 10, "K": 14, "M": 14})
encabezado_empresa(ws, "GUÍA DE ENVÍO", '=cfg_PrefGuia&"-"&TEXT(cfg_NumGuia,"00000")&" (pendiente)"')
ws["I7"] = dp.HOY
sin_guia = [x for x in S.facturas if x[16] == "VIGENTE" and x[6] != "Venta directa"
            and x[0] not in {g[2] for g in S.guias}]
fac_ej = sin_guia[-1][0]
etiqueta(ws, "B10", "Comprobante*")
entrada(ws, "C10", fac_ej, ayuda="Escriba el número del comprobante de venta (ej. FV-000031).")
ws["M10"] = '=IFERROR(MATCH(C10,fac_Num,0),"")'
ws["M11"] = '=IF(M10="","",INDEX(fac_Cli,M10))'
ws["M12"] = '=IFERROR(MATCH(M11,cli_Cod,0),"")'
ws["M13"] = ('=IF(C10="","",IF(M10="","NO EXISTE",IF(INDEX(fac_Estado,M10)<>"VIGENTE","COMPROBANTE "&INDEX(fac_Estado,M10),'
             'IF(COUNTIF(gui_Fact,C10)>0,"YA TIENE GUÍA","OK"))))')
for c in ("M10", "M11", "M12", "M13"):
    ws[c].font = f(8, color="808080")
etiqueta(ws, "D10", "Validación")
calculada(ws, "E10", "=M13", "F10", bold=True)
ws.conditional_formatting.add("E10", CellIsRule(operator="notEqual", formula=['"OK"'],
                              font=Font(color="C00000", bold=True)))
for rr, lab, frm, mto in [(11, "Destinatario", '=IF(M10="","",INDEX(fac_CliNom,M10))', "D11"),
                          (12, "Teléfono", '=IF(M12="","",INDEX(cli_Tel,M12))', "D12"),
                          (13, "Dirección", '=IF(M12="","",INDEX(cli_Dir,M12))', "F13"),
                          (14, "Ciudad", '=IF(M12="","",INDEX(cli_Ciudad,M12))', "D14")]:
    etiqueta(ws, f"B{rr}", lab)
    calculada(ws, f"C{rr}", frm, mto)
nota(ws, "C13", "Se toma de la ficha del cliente. Si el envío va a otra dirección, escríbala encima (LIMPIAR restaura la fórmula).")
ws["C13"].protection = UNLOCK
etiqueta(ws, "H10", "Transportadora*")
entrada(ws, "I10", "Interrapidísimo", "K10", lista="lst_Transportadoras")
etiqueta(ws, "H11", "Nº seguimiento")
entrada(ws, "I11", None, "K11")
etiqueta(ws, "H12", "Valor declarado")
entrada(ws, "I12", '=IF(M10="","",INDEX(fac_Total,M10))', "K12", fmt=COP)
etiqueta(ws, "H13", "Estado*")
entrada(ws, "I13", "PENDIENTE", "K13", lista="lst_EstadosGuia")
etiqueta(ws, "H14", "Contra entrega")
calculada(ws, "I14", '=IF(M10="","",INDEX(fac_Saldo,M10))', "K14", fmt=COP)
hdr = ["#", "SKU", "Producto", "Talla", "Color", "Cantidad"]
for i, t in enumerate(hdr):
    c = ws.cell(16, 2 + i, t)
    c.font, c.fill, c.border = f(9, True, "FFFFFF"), F_NAVY, BORDE
for r in range(17, 32):
    n = r - 16
    ws.cell(r, 2, n).font = f(9, color=GRIS)
    key = f'$C$10&"|"&{n}'
    calculada(ws, f"C{r}", f'=IFERROR(INDEX(det_SKU,MATCH({key},det_Key,0)),"")')
    calculada(ws, f"D{r}", f'=IFERROR(INDEX(det_Prod,MATCH({key},det_Key,0)),"")')
    calculada(ws, f"E{r}", f'=IFERROR(INDEX(det_Talla,MATCH({key},det_Key,0)),"")')
    calculada(ws, f"F{r}", f'=IFERROR(INDEX(det_Color,MATCH({key},det_Key,0)),"")')
    calculada(ws, f"G{r}", f'=IFERROR(INDEX(det_Cant,MATCH({key},det_Key,0)),"")', fmt=NUM)
ws["F32"] = "Unidades"
ws["F32"].font = f(9, True)
calculada(ws, "G32", "=SUM(G17:G31)", fmt=NUM, bold=True)
ws.merge_cells("B34:K36")
ws["B34"] = ('="REMITENTE: "&cfg_Empresa&" · "&cfg_Direccion&" · "&cfg_Ciudad&" · Tel. "&cfg_Telefono')
ws["B34"].font = f(9, True)
ws["B34"].alignment = Alignment(wrap_text=True, vertical="top")
ws["B34"].border = BORDE
ws.merge_cells("B38:F38")
ws["B38"] = "Firma transportadora: ______________________"
ws.merge_cells("H38:K38")
ws["H38"] = "Recibido: ______________________"
ws.print_area = "B4:K39"
ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1
ws.sheet_properties.pageSetUpPr.fitToPage = True

# ====================================================== COSTOS Y RENTABILIDAD
ws = sh["costos"]
titulo(ws, "COSTOS Y RENTABILIDAD POR SKU", 17,
       "Margen = utilidad / precio. Recargo = utilidad / costo. Escenarios calculados con los márgenes de CONFIGURACIÓN.")
cst_cols = [("SKU", 15, "calc"), ("Producto", 30, "calc"), ("Costo adquisición", 12, "calc"),
            ("Transporte unit.", 10, "in"), ("Empaque", 9, "in"), ("Otros costos", 9, "in"),
            ("Precio de venta", 12, "calc"), ("Comisión", 10, "calc"), ("Costo unitario real", 12, "calc"),
            ("Utilidad", 11, "calc"), ("Margen s/ventas", 9, "calc"), ("Recargo s/costo", 9, "calc"),
            ("Precio mínimo", 11, "calc"), ("Competitivo", 11, "calc"), ("Recomendado (sugerido)", 12, "calc"),
            ("Premium", 11, "calc"), ("Alerta", 24, "calc")]
encabezados(ws, 5, cst_cols)
estilo_columnas(ws, 6, INV_R2, cst_cols, {3: COP, 4: COP, 5: COP, 6: COP, 7: COP, 8: COP, 9: COP, 10: COP,
                                          11: PCT, 12: PCT, 13: COP, 14: COP, 15: COP, 16: COP})
IV = q("INVENTARIO")
for r in range(6, INV_R2 + 1):
    base_c = f"(C{r}+N(D{r})+N(E{r})+N(F{r}))"
    ws.cell(r, 1, f'=IF({IV}!A{r}="","",{IV}!A{r})')
    ws.cell(r, 2, f'=IF(A{r}="","",{IV}!B{r}&" "&{IV}!D{r}&" "&{IV}!E{r})')
    ws.cell(r, 3, f'=IF(A{r}="","",{IV}!G{r})')
    ws.cell(r, 4, f'=IF(A{r}="","",cfg_TransporteUnit)')
    ws.cell(r, 5, f'=IF(A{r}="","",cfg_Empaque)')
    ws.cell(r, 7, f'=IF(A{r}="","",{IV}!H{r})')
    ws.cell(r, 8, f'=IF(A{r}="","",G{r}*cfg_ComisionProm)')
    ws.cell(r, 9, f'=IF(A{r}="","",{base_c}+H{r})')
    ws.cell(r, 10, f'=IF(A{r}="","",G{r}-I{r})')
    ws.cell(r, 11, f'=IF(OR(A{r}="",N(G{r})=0),"",J{r}/G{r})')
    ws.cell(r, 12, f'=IF(OR(A{r}="",N(I{r})=0),"",J{r}/I{r})')
    for cc, mg in [(13, "0"), (14, "cfg_MargenComp"), (15, "cfg_MargenRec"), (16, "cfg_MargenPrem")]:
        ws.cell(r, cc, f'=IF(A{r}="","",IF(1-{mg}-cfg_ComisionProm<=0,"N/A",'
                       f'ROUNDUP({base_c}/(1-{mg}-cfg_ComisionProm),-2)))')
    ws.cell(r, 17, f'=IF(A{r}="","",IF(G{r}<I{r},"⚠ PRECIO BAJO EL COSTO REAL",IF(K{r}<cfg_MargenComp,'
                   f'"Margen bajo (< competitivo)","OK")))')
ws.conditional_formatting.add(f"Q6:Q{INV_R2}", FormulaRule(formula=['LEFT(Q6,1)="⚠"'],
                              fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True)))
ws.conditional_formatting.add(f"Q6:Q{INV_R2}", CellIsRule(operator="equal", formula=['"Margen bajo (< competitivo)"'],
                              fill=PatternFill("solid", fgColor="FFEB9C")))
nota(ws, "D5", "Por defecto toma CONFIGURACIÓN. Puede escribir un valor propio para un SKU.")
nota(ws, "M5", "Precio en el que la utilidad es cero, considerando la comisión como % del precio.")
nota(ws, "C5", "Costo promedio ponderado (incluye fletes de compra prorrateados).")
ws["B4"] = ("Fórmula de precio por escenario: (costo sin comisión) / (1 − margen objetivo − % comisión), "
            "redondeado a la centena.")
ws["B4"].font = f(8, italic=True, color=GRIS)

# ====================================================== CAJA Y BANCOS
ws = sh["caja"]
titulo(ws, "CAJA Y BANCOS", 10, "Saldos por método de pago. Flujo de caja ≠ utilidad contable (ver abajo).")
marcador_botones(ws, "Botones: CIERRE DIARIO · CIERRE SEMANAL · CIERRE MENSUAL · ACTUALIZAR")
anchos(ws, {"B": 30, "C": 15, "D": 15, "E": 15, "F": 15, "G": 15, "H": 16, "I": 4, "J": 40})
etiqueta(ws, "B4", "Período desde / hasta:")
entrada(ws, "C4", dp.D(2026, 11, 1), fmt=FECHA)
entrada(ws, "D4", dp.D(2026, 11, 30), fmt=FECHA)
for i, t in enumerate(["Cuenta / método", "Saldo inicial", "Ingresos período", "Egresos período",
                       "Ingresos acumulados", "Egresos acumulados", "SALDO AL CORTE"]):
    c = ws.cell(6, 2 + i, t)
    c.font, c.fill, c.border = f(9, True, "FFFFFF"), F_TEAL, BORDE
    c.alignment = Alignment(horizontal="center", wrap_text=True)
CF = q("CONFIGURACIÓN")
for i in range(5):
    r = 7 + i
    calculada(ws, f"B{r}", f"={CF}!E{8 + i}")
    calculada(ws, f"C{r}", f"={CF}!F{8 + i}", fmt=COP)
    per = f',ing_Fecha,">="&$C$4,ing_Fecha,"<="&$D$4'
    calculada(ws, f"D{r}", f'=SUMIFS(ing_Valor,ing_Met,B{r},ing_Estado,"VIGENTE"{per})', fmt=COP)
    calculada(ws, f"E{r}", f'=SUMIFS(egr_Valor,egr_Met,B{r},egr_Estado,"VIGENTE",egr_Fecha,">="&$C$4,'
                           f'egr_Fecha,"<="&$D$4)', fmt=COP)
    calculada(ws, f"F{r}", f'=SUMIFS(ing_Valor,ing_Met,B{r},ing_Estado,"VIGENTE",ing_Fecha,"<="&$D$4)', fmt=COP)
    calculada(ws, f"G{r}", f'=SUMIFS(egr_Valor,egr_Met,B{r},egr_Estado,"VIGENTE",egr_Fecha,"<="&$D$4)', fmt=COP)
    calculada(ws, f"H{r}", f"=C{r}+F{r}-G{r}", fmt=COP, bold=True)
ws["B12"] = "TOTAL"
ws["B12"].font = f(10, True)
for col in "CDEFGH":
    calculada(ws, f"{col}12", f"=SUM({col}7:{col}11)", fmt=COP, bold=True)
    ws[f"{col}12"].fill = F_CLARO
ws.conditional_formatting.add("H7:H11", CellIsRule(operator="lessThan", formula=["0"],
                              fill=PatternFill("solid", fgColor="FFC7CE")))
ws["B15"] = "FLUJO DE CAJA vs. UTILIDAD CONTABLE DEL PERÍODO"
ws["B15"].font = f(11, True, NAVY)
FLUJO = [
    ("Entradas de dinero (ingresos)", "=D12"),
    ("Salidas de dinero (egresos, incluye pago de compras)", "=E12"),
    ("FLUJO DE CAJA NETO", "=C16-C17"),
    ("Ventas netas del período", '=SUMIFS(det_Sub,det_Estado,"VIGENTE",det_Fecha,">="&C4,det_Fecha,"<="&D4)'),
    ("Costo de la mercancía vendida", '=SUMIFS(det_CostoT,det_Estado,"VIGENTE",det_Fecha,">="&C4,det_Fecha,"<="&D4)'),
    ("Gastos operativos (sin compras ni devoluciones)",
     '=SUMIFS(egr_Valor,egr_Estado,"VIGENTE",egr_Fecha,">="&C4,egr_Fecha,"<="&D4,egr_Con,"<>Compras",'
     'egr_Con,"<>Devoluciones a clientes")'),
    ("UTILIDAD NETA CONTABLE", "=C19-C20-C21"),
    ("Diferencia (flujo − utilidad)", "=C18-C22"),
]
for i, (lab, frm) in enumerate(FLUJO):
    r = 16 + i
    ws[f"B{r}"] = lab
    ws[f"B{r}"].font = f(9, lab.isupper())
    calculada(ws, f"C{r}", frm, fmt=COP, bold=lab.isupper())
ws["J16"] = ("¿Por qué son diferentes? El flujo de caja mide el dinero que entró y salió. La utilidad mide lo "
             "que se ganó: las compras de inventario no son gasto hasta que se venden (entran como costo de la "
             "mercancía vendida), las ventas a crédito son utilidad pero aún no son dinero, y los abonos de "
             "ventas anteriores son dinero pero no son utilidad de este período.")
ws["J16"].font = f(9, italic=True, color=GRIS)
ws["J16"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells("J16:J23")
ws["B26"] = "Últimos cierres: ver hoja BD_CIERRES"
ws["B26"].hyperlink = Hyperlink(ref="B26", location=f"{q('BD_CIERRES')}!A1")
ws["B26"].font = f(9, True, AZUL)

# ====================================================== UTILIDADES
ws = sh["util"]
titulo(ws, "UTILIDADES · ESTADO DE RESULTADOS MENSUAL", 14, "Ventas vigentes − costo de lo vendido − gastos operativos.")
anchos(ws, {"B": 12, "C": 11, "D": 11, "E": 14, "F": 14, "G": 14, "H": 9, "I": 14, "J": 14, "K": 9, "L": 14,
            "M": 14, "N": 14})
etiqueta(ws, "B4", "Año:")
entrada(ws, "C4", 2026, fmt="0")
hdr = ["Mes", "Desde", "Hasta", "Ventas netas", "Costo de ventas", "Utilidad bruta", "Margen bruto",
       "Gastos operativos", "Utilidad neta", "Margen neto", "Ingresos caja", "Egresos caja", "Flujo neto"]
for i, t in enumerate(hdr):
    c = ws.cell(6, 2 + i, t)
    c.font, c.fill, c.border = f(9, True, "FFFFFF"), F_TEAL, BORDE
    c.alignment = Alignment(horizontal="center", wrap_text=True)
ws.row_dimensions[6].height = 30
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre",
         "Noviembre", "Diciembre"]
for i, m in enumerate(MESES):
    r = 7 + i
    ws[f"B{r}"] = m
    ws[f"B{r}"].font, ws[f"B{r}"].border = f(9, True), BORDE
    calculada(ws, f"C{r}", f"=DATE($C$4,{i + 1},1)", fmt=FECHA)
    calculada(ws, f"D{r}", f"=DATE($C$4,{i + 2},0)", fmt=FECHA)
    rg = f'det_Fecha,">="&C{r},det_Fecha,"<="&D{r}'
    calculada(ws, f"E{r}", f'=SUMIFS(det_Sub,det_Estado,"VIGENTE",{rg})', fmt=COP)
    calculada(ws, f"F{r}", f'=SUMIFS(det_CostoT,det_Estado,"VIGENTE",{rg})', fmt=COP)
    calculada(ws, f"G{r}", f"=E{r}-F{r}", fmt=COP)
    calculada(ws, f"H{r}", f'=IF(E{r}=0,"",G{r}/E{r})', fmt=PCT)
    calculada(ws, f"I{r}", f'=SUMIFS(egr_Valor,egr_Estado,"VIGENTE",egr_Fecha,">="&C{r},egr_Fecha,"<="&D{r},'
                           f'egr_Con,"<>Compras",egr_Con,"<>Devoluciones a clientes")', fmt=COP)
    calculada(ws, f"J{r}", f"=G{r}-I{r}", fmt=COP, bold=True)
    calculada(ws, f"K{r}", f'=IF(E{r}=0,"",J{r}/E{r})', fmt=PCT)
    calculada(ws, f"L{r}", f'=SUMIFS(ing_Valor,ing_Estado,"VIGENTE",ing_Fecha,">="&C{r},ing_Fecha,"<="&D{r})', fmt=COP)
    calculada(ws, f"M{r}", f'=SUMIFS(egr_Valor,egr_Estado,"VIGENTE",egr_Fecha,">="&C{r},egr_Fecha,"<="&D{r})', fmt=COP)
    calculada(ws, f"N{r}", f"=L{r}-M{r}", fmt=COP)
ws["B19"] = "TOTAL AÑO"
ws["B19"].font = f(10, True)
for col in "EFGIJLMN":
    calculada(ws, f"{col}19", f"=SUM({col}7:{col}18)", fmt=COP, bold=True)
    ws[f"{col}19"].fill = F_CLARO
calculada(ws, "H19", '=IF(E19=0,"",G19/E19)', fmt=PCT, bold=True)
calculada(ws, "K19", '=IF(E19=0,"",J19/E19)', fmt=PCT, bold=True)
ws.conditional_formatting.add("J7:J19", CellIsRule(operator="lessThan", formula=["0"],
                              font=Font(color="C00000", bold=True)))
ch = BarChart()
ch.type = "col"
ch.title = "Utilidad bruta vs. neta por mes"
ch.add_data(Reference(ws, min_col=7, min_row=6, max_row=18), titles_from_data=True)
ch.add_data(Reference(ws, min_col=10, min_row=6, max_row=18), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=2, min_row=7, max_row=18))
ch.height, ch.width = 7.5, 22
ws.add_chart(ch, "B22")
UTIL_CHARTS = [ch]

# ====================================================== COMISIONES
ws = sh["com"]
titulo(ws, "COMISIONES POR VENDEDOR", 12, "Base: ventas vigentes del mes seleccionado × % de comisión del vendedor.")
marcador_botones(ws, "Botones: PAGAR COMISIONES DEL MES (registra egresos sin duplicar)")
anchos(ws, {"B": 10, "C": 22, "D": 14, "E": 14, "F": 10, "G": 14, "H": 14, "I": 11, "J": 14, "K": 14, "L": 12})
etiqueta(ws, "B4", "Mes (día 1):")
entrada(ws, "C4", dp.D(2026, 10, 1), fmt="MMMM YYYY")
ws["E4"] = '="Clave de pago: COM-"&(YEAR(C4)*100+MONTH(C4))&"-<vendedor>"'
ws["E4"].font = f(8, italic=True, color=GRIS)
hdr = ["Código", "Vendedor", "Ventas del mes", "Utilidad del mes", "% comisión", "Comisión", "Meta",
       "Cumplimiento", "Pagado", "Pendiente", "Estado"]
for i, t in enumerate(hdr):
    c = ws.cell(6, 2 + i, t)
    c.font, c.fill, c.border = f(9, True, "FFFFFF"), F_TEAL, BORDE
    c.alignment = Alignment(horizontal="center", wrap_text=True)
VD = q("VENDEDORES")
for i in range(30):
    r = 7 + i
    vr = VEN_R1 + i
    calculada(ws, f"B{r}", f'=IF({VD}!A{vr}="","",{VD}!A{vr})')
    calculada(ws, f"C{r}", f'=IF(B{r}="","",{VD}!B{vr})')
    crit = f'fac_Vend,B{r},fac_Estado,"VIGENTE",fac_Fecha,">="&$C$4,fac_Fecha,"<="&DATE(YEAR($C$4),MONTH($C$4)+1,0)'
    calculada(ws, f"D{r}", f'=IF(B{r}="","",SUMIFS(fac_Total,{crit}))', fmt=COP)
    calculada(ws, f"E{r}", f'=IF(B{r}="","",SUMIFS(fac_Util,{crit}))', fmt=COP)
    calculada(ws, f"F{r}", f'=IF(B{r}="","",{VD}!H{vr})', fmt="0.0%")
    calculada(ws, f"G{r}", f'=IF(B{r}="","",ROUND(D{r}*F{r},0))', fmt=COP, bold=True)
    calculada(ws, f"H{r}", f'=IF(B{r}="","",{VD}!I{vr})', fmt=COP)
    calculada(ws, f"I{r}", f'=IF(OR(B{r}="",N(H{r})=0),"",D{r}/H{r})', fmt=PCT)
    calculada(ws, f"J{r}", f'=IF(B{r}="","",SUMIFS(egr_Valor,egr_Con,"Comisiones",egr_Estado,"VIGENTE",egr_Doc,'
                           f'"COM-"&(YEAR($C$4)*100+MONTH($C$4))&"-"&B{r}))', fmt=COP)
    calculada(ws, f"K{r}", f'=IF(B{r}="","",G{r}-J{r})', fmt=COP)
    calculada(ws, f"L{r}", f'=IF(B{r}="","",IF(G{r}=0,"SIN VENTAS",IF(K{r}<=0,"PAGADA","PENDIENTE")))')
ws.conditional_formatting.add("L7:L36", CellIsRule(operator="equal", formula=['"PENDIENTE"'],
                              fill=PatternFill("solid", fgColor="FFEB9C")))
ws.conditional_formatting.add("L7:L36", CellIsRule(operator="equal", formula=['"PAGADA"'],
                              fill=PatternFill("solid", fgColor="C6EFCE")))
ws["F38"] = "TOTAL"
ws["F38"].font = f(10, True)
for col in "DGJK":
    calculada(ws, f"{col}38", f"=SUM({col}7:{col}36)", fmt=COP, bold=True)

# ====================================================== CANALES DE VENTA
ws = sh["canal"]
titulo(ws, "CANALES DE VENTA · COMPARATIVO", 13, "Todas las ventas exigen canal. Compare ventas y rentabilidad.")
anchos(ws, {"B": 16, "C": 14, "D": 9, "E": 14, "F": 9, "G": 13, "H": 10, "I": 11, "J": 12, "K": 11, "L": 10})
etiqueta(ws, "B4", "Desde / hasta:")
entrada(ws, "C4", dp.D(2026, 10, 1), fmt=FECHA)
entrada(ws, "D4", dp.D(2026, 11, 30), "E4", fmt=FECHA)
hdr = ["Canal", "Ventas", "Pedidos", "Utilidad", "Margen", "Ticket promedio", "Clientes nuevos",
       "Devoluciones", "Cancelaciones", "% de ventas", "Ranking utilidad"]
for i, t in enumerate(hdr):
    c = ws.cell(6, 2 + i, t)
    c.font, c.fill, c.border = f(9, True, "FFFFFF"), F_TEAL, BORDE
    c.alignment = Alignment(horizontal="center", wrap_text=True)
ws.row_dimensions[6].height = 30
BD = q("BASE DE DATOS")
for i in range(10):
    r = 7 + i
    calculada(ws, f"B{r}", f'=IF({BD}!E{6 + i}="","",{BD}!E{6 + i})', bold=True)
    rg = f'fac_Fecha,">="&$C$4,fac_Fecha,"<="&$D$4'
    calculada(ws, f"C{r}", f'=IF(B{r}="","",SUMIFS(fac_Total,fac_Canal,B{r},fac_Estado,"VIGENTE",{rg}))', fmt=COP)
    calculada(ws, f"D{r}", f'=IF(B{r}="","",COUNTIFS(fac_Canal,B{r},fac_Estado,"VIGENTE",{rg}))', fmt=NUM)
    calculada(ws, f"E{r}", f'=IF(B{r}="","",SUMIFS(fac_Util,fac_Canal,B{r},fac_Estado,"VIGENTE",{rg}))', fmt=COP)
    calculada(ws, f"F{r}", f'=IF(OR(B{r}="",N(C{r})=0),"",E{r}/C{r})', fmt=PCT)
    calculada(ws, f"G{r}", f'=IF(OR(B{r}="",N(D{r})=0),"",C{r}/D{r})', fmt=COP)
    calculada(ws, f"H{r}", f'=IF(B{r}="","",COUNTIFS(cli_Canal,B{r},cli_Fecha,">="&$C$4,cli_Fecha,"<="&$D$4))', fmt=NUM)
    calculada(ws, f"I{r}", f'=IF(B{r}="","",COUNTIFS(fac_Canal,B{r},fac_Estado,"DEVUELTA",{rg}))', fmt=NUM)
    calculada(ws, f"J{r}", f'=IF(B{r}="","",COUNTIFS(fac_Canal,B{r},fac_Estado,"ANULADA",{rg}))', fmt=NUM)
    calculada(ws, f"K{r}", f'=IF(OR(B{r}="",$C$17=0),"",C{r}/$C$17)', fmt=PCT)
    calculada(ws, f"L{r}", f'=IF(OR(B{r}="",C{r}=0),"",RANK(E{r},$E$7:$E$16))', fmt="0")
ws["B17"] = "TOTAL"
ws["B17"].font = f(10, True)
for col in "CDEHIJ":
    calculada(ws, f"{col}17", f"=SUM({col}7:{col}16)", fmt=COP if col in "CE" else NUM, bold=True)
calculada(ws, "F17", '=IF(C17=0,"",E17/C17)', fmt=PCT, bold=True)
calculada(ws, "G17", '=IF(D17=0,"",C17/D17)', fmt=COP, bold=True)
ch = BarChart()
ch.type = "col"
ch.title = "Ventas y utilidad por canal"
ch.add_data(Reference(ws, min_col=3, min_row=6, max_row=11), titles_from_data=True)
ch.add_data(Reference(ws, min_col=5, min_row=6, max_row=11), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=2, min_row=7, max_row=11))
ch.height, ch.width = 7.5, 18
ws.add_chart(ch, "B20")
CANAL_CHARTS = [ch]

# ====================================================== USUARIOS Y PERMISOS
ws = sh["usr"]
titulo(ws, "USUARIOS Y PERMISOS", 17, "Solo Administrador. La clave se guarda cifrada (hash); nunca en texto plano.")
marcador_botones(ws, "Botones: NUEVO USUARIO · RESTABLECER CLAVE · CERRAR SESIÓN")
usr_cols = [("Usuario", 11, "in"), ("Nombre", 24, "in"), ("Rol", 20, "in"), ("Vendedor asociado", 10, "in"),
            ("Estado", 9, "in"), ("Clave (hash)", 13, "calc"), ("Último acceso", 15, "calc"),
            ("Cambiar clave al entrar", 9, "calc")]
encabezados(ws, 5, usr_cols, filtro=False)
estilo_columnas(ws, 6, 55, usr_cols, {7: FECHAHORA})
for i, u in enumerate(dp.USUARIOS):
    vals = u[:5] + [dp.hash_clave(u[0], u[5]), None, "SI"]
    for j, v in enumerate(vals, start=1):
        ws.cell(6 + i, j, v)
dv_lista(ws, "C6:C55", "lst_Roles")
dv_lista(ws, "E6:E55", "lst_Estados")
rango("usr_Usuario", ws, "A", 6, 55)
rango("usr_Rol", ws, "C", 6, 55)
rango("usr_Estado", ws, "E", 6, 55)
ws["F6"].font = f(8, color="808080")
ws["J4"] = "MATRIZ DE PERMISOS POR ROL (SI / NO)"
ws["J4"].font = f(10, True, NAVY)
for i, (t, w) in enumerate([("Permiso", 20), ("Descripción", 38)] + [(r_, 12) for r_ in dp.ROLES]):
    c = ws.cell(5, 10 + i, t)
    c.font, c.fill, c.border = f(9, True, "FFFFFF"), F_NAVY, BORDE
    c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.column_dimensions[L(10 + i)].width = w
ws.column_dimensions["I"].width = 3
for i, p in enumerate(dp.PERMISOS):
    r = 6 + i
    for j, v in enumerate(p):
        c = ws.cell(r, 10 + j, v)
        c.font, c.border = f(9, j == 0), BORDE
        if j >= 2:
            c.alignment = Alignment(horizontal="center")
            c.protection = UNLOCK
dv_lista(ws, f"L6:O{5 + len(dp.PERMISOS)}", "lst_SiNo")
for val, bg in [("SI", "C6EFCE"), ("NO", "FFC7CE")]:
    ws.conditional_formatting.add(f"L6:O{5 + len(dp.PERMISOS)}", CellIsRule(operator="equal", formula=[f'"{val}"'],
                                  fill=PatternFill("solid", fgColor=bg)))
nombre("perm_Codigos", f"{q(ws.title)}!$J$6:$J$40")
nombre("perm_Roles", f"{q(ws.title)}!$L$5:$O$5")
nombre("perm_Matriz", f"{q(ws.title)}!$L$6:$O$40")
ws["B58"] = ("Claves iniciales de prueba (cámbielas en el primer ingreso): admin/Admin2026 · aux/Aux2026 · "
             "laura/Laura2026 · andres/Andres2026 · camila/Camila2026")
ws["B58"].font = f(9, True, "C00000")

# ====================================================== REPORTES
ws = sh["rep"]
titulo(ws, "REPORTES", 9, "Elija el período y pulse PDF o EXCEL en la fila del reporte. Se guardan en OneDrive/REPORTES.")
anchos(ws, {"B": 5, "C": 28, "D": 52, "E": 13, "F": 13, "G": 4})
etiqueta(ws, "C4", "Desde / hasta:")
entrada(ws, "D4", dp.D(2026, 11, 1), fmt=FECHA)
entrada(ws, "E4", dp.D(2026, 11, 30), fmt=FECHA)
celda("rep_Desde", ws, "D4")
celda("rep_Hasta", ws, "E4")
REPORTES = [
    ("Ventas diarias", "Comprobantes del día 'hasta'"),
    ("Ventas mensuales", "Comprobantes del mes de la fecha 'hasta'"),
    ("Ventas anuales", "Estado de resultados mensual del año de 'hasta' (hoja UTILIDADES)"),
    ("Ventas por canal", "Comparativo de canales en el período (hoja CANALES DE VENTA)"),
    ("Ventas por vendedor", "Ventas, pedidos y utilidad por vendedor en el período"),
    ("Inventario", "Existencias, valor y alertas por SKU"),
    ("Compras", "Compras registradas en el período"),
    ("Proveedores", "Maestro de proveedores con saldos"),
    ("Clientes", "Maestro de clientes con clasificación y saldos"),
    ("Gastos", "Egresos del período (excluye anulados)"),
    ("Cartera", "Cuentas por cobrar y por pagar con saldo pendiente"),
    ("Flujo de caja", "Saldos por cuenta y flujo del período (hoja CAJA Y BANCOS)"),
    ("Utilidades", "Utilidad bruta y neta por mes (hoja UTILIDADES)"),
]
for i, t in enumerate(["Nº", "Reporte", "Contenido", "PDF", "EXCEL"]):
    c = ws.cell(6, 2 + i, t)
    c.font, c.fill, c.border = f(9, True, "FFFFFF"), F_TEAL, BORDE
for i, (n_, d_) in enumerate(REPORTES):
    r = 7 + i
    ws.row_dimensions[r].height = 24
    ws.cell(r, 2, i + 1).font = f(9, True)
    ws.cell(r, 3, n_).font = f(10, True, NAVY)
    ws.cell(r, 4, d_).font = f(9, color=GRIS)
    for cc in range(2, 7):
        ws.cell(r, cc).border = BORDE
        ws.cell(r, cc).alignment = Alignment(vertical="center")
    ws.cell(r, 5, "(botón PDF)").font = f(8, italic=True, color="808080")
    ws.cell(r, 6, "(botón EXCEL)").font = f(8, italic=True, color="808080")

# ====================================================== DASHBOARD
ws = sh["dash"]
titulo(ws, "DASHBOARD GERENCIAL", 14)
marcador_botones(ws, "Botones: ACTUALIZAR · PDF DEL DASHBOARD")
for c in range(2, 15):
    ws.column_dimensions[L(c)].width = 13
ws.column_dimensions["O"].width = 3
etiqueta(ws, "B4", "Desde")
entrada(ws, "C4", dp.D(2026, 10, 1), fmt=FECHA)
etiqueta(ws, "D4", "Hasta")
entrada(ws, "E4", dp.D(2026, 11, 30), fmt=FECHA)
etiqueta(ws, "F4", "Canal")
entrada(ws, "G4", "Todos", "H4", lista="flt_Canales")
etiqueta(ws, "I4", "Vendedor")
entrada(ws, "J4", "Todos", "K4", lista="flt_Vendedores")
etiqueta(ws, "L4", "Categoría")
entrada(ws, "M4", "Todos", "N4", lista="flt_Categorias")
for c in ("B4", "D4", "F4", "I4", "L4"):
    ws[c].alignment = Alignment(horizontal="right")
# criterios auxiliares (columna R)
ws["Q3"] = "Criterios"
ws["Q3"].font = f(7, color="808080")
ws["R4"] = '=IF(OR(G4="",G4="Todos"),"*",G4)'
ws["R5"] = '=IF(OR(J4="",J4="Todos"),"*",LEFT(J4,FIND(" - ",J4&" - ")-1))'
ws["R6"] = '=IF(OR(M4="",M4="Todos"),"*",M4)'
for c in ("R4", "R5", "R6"):
    ws[c].font = f(7, color="808080")
FD = 'det_Canal,$R$4,det_Vend,$R$5,det_Cat,$R$6,det_Estado,"VIGENTE"'
FH = 'fac_Canal,$R$4,fac_Vend,$R$5,fac_Estado,"VIGENTE"'
PER_D = 'det_Fecha,">="&$C$4,det_Fecha,"<="&$E$4'
PER_H = 'fac_Fecha,">="&$C$4,fac_Fecha,"<="&$E$4'
GASTOS_PER = ('SUMIFS(egr_Valor,egr_Estado,"VIGENTE",egr_Fecha,">="&$C$4,egr_Fecha,"<="&$E$4,'
              'egr_Con,"<>Compras",egr_Con,"<>Devoluciones a clientes")')
KPIS = [
    [("Ventas del día", f'=SUMIFS(det_Sub,{FD},det_Fecha,TODAY())', COP),
     ("Ventas del mes", f'=SUMIFS(det_Sub,{FD},det_Fecha,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),det_Fecha,"<="&TODAY())', COP),
     ("Ventas del período", f'=SUMIFS(det_Sub,{FD},{PER_D})', COP),
     ("Utilidad bruta", f'=SUMIFS(det_Util,{FD},{PER_D})', COP),
     ("Gastos operativos", f'={GASTOS_PER}', COP),
     ("Utilidad neta", "=H8-J8", COP)],
    [("Pedidos", f'=COUNTIFS({FH},{PER_H})', NUM),
     ("Ticket promedio", '=IF(B11=0,0,SUMIFS(fac_Total,' + FH + ',' + PER_H + ')/B11)', COP),
     ("Margen bruto", '=IF(F8=0,0,H8/F8)', PCT),
     ("Clientes nuevos", '=COUNTIFS(cli_Fecha,">="&$C$4,cli_Fecha,"<="&$E$4,cli_Canal,$R$4)', NUM),
     ("Cuentas por cobrar", '=SUMIFS(car_Saldo,car_Tipo,"CxC")', COP),
     ("Cuentas por pagar", '=SUMIFS(car_Saldo,car_Tipo,"CxP")', COP)],
    [("Inventario (valor)", '=SUMIFS(inv_Valor,inv_Cat,$R$6)', COP),
     ("Unidades en stock", '=SUMIFS(inv_Exist,inv_Cat,$R$6)', NUM),
     ("Productos agotados", '=COUNTIFS(inv_Estado,"AGOTADO",inv_Cat,$R$6)', NUM),
     ("Stock bajo", '=COUNTIFS(inv_Estado,"BAJO",inv_Cat,$R$6)', NUM),
     ("Cartera vencida", '=SUMIFS(car_Saldo,car_Tipo,"CxC",car_Estado,"VENCIDO")', COP),
     ("Anulaciones / devol.", f'=COUNTIFS(fac_Canal,$R$4,fac_Vend,$R$5,fac_Estado,"<>VIGENTE",{PER_H})', NUM)],
]
for fila_i, fila in enumerate(KPIS):
    rl = 7 + fila_i * 3
    for k, (lab, frm, fmt) in enumerate(fila):
        c0 = 2 + k * 2
        ws.merge_cells(start_row=rl, start_column=c0, end_row=rl, end_column=c0 + 1)
        ws.merge_cells(start_row=rl + 1, start_column=c0, end_row=rl + 1, end_column=c0 + 1)
        a = ws.cell(rl, c0, lab)
        relleno = F_AZUL if fila_i < 2 else F_TEAL
        a.font, a.fill = f(9, True, "FFFFFF"), relleno
        a.alignment = Alignment(horizontal="center")
        ws.cell(rl, c0 + 1).fill = relleno
        v = ws.cell(rl + 1, c0, frm)
        v.font, v.number_format = f(15, True, NAVY), fmt
        v.alignment = Alignment(horizontal="center", vertical="center")
        v.fill = F_CLARO
        ws.cell(rl + 1, c0 + 1).fill = F_CLARO
        ws.row_dimensions[rl + 1].height = 30
ws.conditional_formatting.add("L8", CellIsRule(operator="lessThan", formula=["0"], font=Font(color="C00000", bold=True, size=15)))
ws.conditional_formatting.add("F14", CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="C00000", bold=True, size=15)))
ws.conditional_formatting.add("J14", CellIsRule(operator="greaterThan", formula=["0"], font=Font(color="C00000", bold=True, size=15)))
ws["B16"] = ("Las ventas del día/mes usan la fecha del equipo. Los gastos y saldos de cartera no dependen de "
             "los filtros de canal/vendedor/categoría.")
ws["B16"].font = f(8, italic=True, color=GRIS)

# tablas de datos para gráficos (fila 52 en adelante)
T0 = 52
ws.cell(T0 - 1, 2, "DATOS DE LOS GRÁFICOS (calculados con los filtros de arriba)").font = f(10, True, NAVY)


def cab(r, c, textos):
    for i, t in enumerate(textos):
        x = ws.cell(r, c + i, t)
        x.font, x.fill, x.border = f(9, True, "FFFFFF"), F_TEAL, BORDE


# T1 canales
cab(T0, 2, ["Canal", "Ventas", "Utilidad"])
for i in range(5):
    r = T0 + 1 + i
    ws.cell(r, 2, f"={BD}!E{6 + i}").font = f(9)
    ws.cell(r, 3, f'=SUMIFS(det_Sub,det_Canal,B{r},det_Vend,$R$5,det_Cat,$R$6,det_Estado,"VIGENTE",{PER_D})').number_format = COP
    ws.cell(r, 4, f'=SUMIFS(det_Util,det_Canal,B{r},det_Vend,$R$5,det_Cat,$R$6,det_Estado,"VIGENTE",{PER_D})').number_format = COP
# T2 últimos 6 meses
cab(T0, 6, ["Mes", "Ventas", "Utilidad bruta", "Gastos", "Utilidad neta"])
for i in range(6):
    r = T0 + 1 + i
    ws.cell(r, 6, f"=DATE(YEAR($E$4),MONTH($E$4)-{5 - i},1)").number_format = "MMM YYYY"
    fin = f"DATE(YEAR(F{r}),MONTH(F{r})+1,0)"
    ws.cell(r, 7, f'=SUMIFS(det_Sub,{FD},det_Fecha,">="&F{r},det_Fecha,"<="&{fin})').number_format = COP
    ws.cell(r, 8, f'=SUMIFS(det_Util,{FD},det_Fecha,">="&F{r},det_Fecha,"<="&{fin})').number_format = COP
    ws.cell(r, 9, f'=SUMIFS(egr_Valor,egr_Estado,"VIGENTE",egr_Fecha,">="&F{r},egr_Fecha,"<="&{fin},'
                  f'egr_Con,"<>Compras",egr_Con,"<>Devoluciones a clientes")').number_format = COP
    ws.cell(r, 10, f"=H{r}-I{r}").number_format = COP
# T3 vendedores
T3 = T0 + 9
cab(T3, 2, ["Vendedor", "Ventas", "Pedidos", "Utilidad", "Meta", "Cumplimiento"])
for i in range(8):
    r = T3 + 1 + i
    vr = VEN_R1 + i
    ws.cell(r, 2, f'=IF({VD}!A{vr}="","",{VD}!B{vr})')
    ws.cell(r, 8, f'=IF({VD}!A{vr}="","",{VD}!A{vr})').font = f(7, color="808080")
    ws.cell(r, 3, f'=IF(H{r}="","",SUMIFS(det_Sub,det_Vend,H{r},det_Canal,$R$4,det_Cat,$R$6,det_Estado,"VIGENTE",{PER_D}))').number_format = COP
    ws.cell(r, 4, f'=IF(H{r}="","",COUNTIFS(fac_Vend,H{r},fac_Canal,$R$4,fac_Estado,"VIGENTE",{PER_H}))')
    ws.cell(r, 5, f'=IF(H{r}="","",SUMIFS(det_Util,det_Vend,H{r},det_Canal,$R$4,det_Cat,$R$6,det_Estado,"VIGENTE",{PER_D}))').number_format = COP
    ws.cell(r, 6, f'=IF(H{r}="","",{VD}!I{vr}*MAX(1,ROUND(($E$4-$C$4+1)/30,0)))').number_format = COP
    ws.cell(r, 7, f'=IF(OR(H{r}="",N(F{r})=0),"",C{r}/F{r})').number_format = PCT
# T4 inventario por categoría
cab(T3, 10, ["Categoría", "Valor inventario", "Unidades"])
for i in range(6):
    r = T3 + 1 + i
    ws.cell(r, 10, f'=IF({BD}!B{6 + i}="","",{BD}!B{6 + i})')
    ws.cell(r, 11, f'=IF(J{r}="","",SUMIFS(inv_Valor,inv_Cat,J{r}))').number_format = COP
    ws.cell(r, 12, f'=IF(J{r}="","",SUMIFS(inv_Exist,inv_Cat,J{r}))').number_format = NUM
# T5 gastos por concepto
T5 = T3 + 11
cab(T5, 2, ["Concepto", "Gasto período"])
conceptos_gasto = [c for c in dp.CONCEPTOS_EGR if c not in ("Compras", "Devoluciones a clientes")]
for i, cg in enumerate(conceptos_gasto):
    r = T5 + 1 + i
    ws.cell(r, 2, cg)
    ws.cell(r, 3, f'=SUMIFS(egr_Valor,egr_Con,B{r},egr_Estado,"VIGENTE",egr_Fecha,">="&$C$4,egr_Fecha,"<="&$E$4)').number_format = COP
for row in ws.iter_rows(min_row=T0, max_row=T5 + 12, min_col=2, max_col=12):
    for c in row:
        if c.value is not None and c.font.color is None or (c.value is not None and not c.font.bold):
            c.font = f(9)
ws["B16"].font = f(8, italic=True, color=GRIS)


def grafico(tipo, titulo_g, datos, cats, ancla, w=15, h=7.2, ejes=True):
    g = tipo()
    if tipo is BarChart:
        g.type = "col"
    g.title = titulo_g
    for d_ in datos:
        g.add_data(d_, titles_from_data=True)
    g.set_categories(cats)
    g.width, g.height = w, h
    ws.add_chart(g, ancla)
    return g


DASH_CHARTS = [
    grafico(BarChart, "Ventas y utilidad por canal",
            [Reference(ws, min_col=3, min_row=T0, max_row=T0 + 5), Reference(ws, min_col=4, min_row=T0, max_row=T0 + 5)],
            Reference(ws, min_col=2, min_row=T0 + 1, max_row=T0 + 5), "B18"),
    grafico(LineChart, "Ventas, gastos y utilidad neta (6 meses)",
            [Reference(ws, min_col=7, min_row=T0, max_row=T0 + 6), Reference(ws, min_col=9, min_row=T0, max_row=T0 + 6),
             Reference(ws, min_col=10, min_row=T0, max_row=T0 + 6)],
            Reference(ws, min_col=6, min_row=T0 + 1, max_row=T0 + 6), "H18"),
    grafico(BarChart, "Rendimiento de vendedores",
            [Reference(ws, min_col=3, min_row=T3, max_row=T3 + 3), Reference(ws, min_col=6, min_row=T3, max_row=T3 + 3)],
            Reference(ws, min_col=2, min_row=T3 + 1, max_row=T3 + 3), "B34", w=11),
    grafico(PieChart, "Inventario por categoría (valor)",
            [Reference(ws, min_col=11, min_row=T3, max_row=T3 + 3)],
            Reference(ws, min_col=10, min_row=T3 + 1, max_row=T3 + 3), "G34", w=8.5, ejes=False),
    grafico(BarChart, "Gastos por concepto",
            [Reference(ws, min_col=3, min_row=T5, max_row=T5 + len(conceptos_gasto))],
            Reference(ws, min_col=2, min_row=T5 + 1, max_row=T5 + len(conceptos_gasto)), "K34", w=10.5),
]
DASH_CHARTS[4].type = "bar"
ws.print_area = "B1:N50"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1
ws.sheet_properties.pageSetUpPr.fitToPage = True

# ====================================================== MENÚ
ws = sh["menu"]
for c in range(1, 15):
    ws.column_dimensions[L(c)].width = 12.5
ws.column_dimensions["A"].width = 2
for r in range(1, 5):
    for c in range(1, 15):
        ws.cell(r, c).fill = F_NAVY
ws.row_dimensions[2].height = 34
ws["B2"] = "=cfg_Empresa"
ws["B2"].font = f(20, True, "FFFFFF")
ws["B3"] = "SISTEMA ERP COMERCIAL Y ADMINISTRATIVO"
ws["B3"].font = f(11, True, "BDD7EE")
ws["B4"] = '="Usuario: "&app_Usuario&"   ·   Rol: "&app_Rol&"   ·   Versión "&cfg_Version'
ws["B4"].font = f(9, False, "FFFFFF")
poner_logo(ws, "M1", 70)
ws.merge_cells("B6:M6")
ws["B6"] = ("Si al abrir solo ve esta hoja: habilite las macros (barra amarilla «Habilitar contenido») e inicie "
            "sesión. Sin macros los botones no funcionan y el sistema queda en modo solo lectura.")
ws["B6"].font = f(9, True, "9C5700")
ws["B6"].fill = PatternFill("solid", fgColor="FFEB9C")
ws["B6"].alignment = Alignment(wrap_text=True, vertical="center")
ws.row_dimensions[6].height = 30
ws.row_dimensions[5].height = 28
SECCIONES = [
    ("VENTAS Y CLIENTES", "00B050", ["FACTURACIÓN", "GUÍAS", "CLIENTES", "CANALES DE VENTA", "VENDEDORES", "COMISIONES"]),
    ("INVENTARIO Y COMPRAS", "ED7D31", ["INVENTARIO", "COMPRAS", "PROVEEDORES", "HISTORIAL MOVIMIENTOS",
                                         "COSTOS Y RENTABILIDAD"]),
    ("FINANZAS", "7030A0", ["INGRESOS", "EGRESOS", "GASTOS FIJOS", "CARTERA", "CAJA Y BANCOS", "UTILIDADES"]),
    ("GERENCIA Y SISTEMA", "C00000", ["DASHBOARD", "REPORTES", "USUARIOS Y PERMISOS", "CONFIGURACIÓN",
                                       "BASE DE DATOS", "BD_BITACORA"]),
]
r = 8
for sec, color, hojas_ in SECCIONES:
    ws.cell(r, 2, sec).font = f(11, True, color)
    r += 1
    for i, h in enumerate(hojas_):
        c0 = 2 + i * 2
        ws.merge_cells(start_row=r, start_column=c0, end_row=r + 1, end_column=c0 + 1)
        c = ws.cell(r, c0, h.replace("BD_BITACORA", "BITÁCORA"))
        c.hyperlink = Hyperlink(ref=c.coordinate, location=f"{q(h)}!A1")
        c.font = f(10, True, "FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for rr in (r, r + 1):
            for cc in (c0, c0 + 1):
                ws.cell(rr, cc).fill = PatternFill("solid", fgColor=color)
                ws.cell(rr, cc).border = Border(left=Side(style="medium", color="FFFFFF"),
                                                right=Side(style="medium", color="FFFFFF"),
                                                top=Side(style="medium", color="FFFFFF"),
                                                bottom=Side(style="medium", color="FFFFFF"))
        ws.row_dimensions[r].height = 22
        ws.row_dimensions[r + 1].height = 22
    r += 3
ALERTAS_R = r + 1
ws.cell(ALERTAS_R, 2, "ALERTAS DEL SISTEMA").font = f(11, True, NAVY)
alertas = [
    ("Productos agotados", '=COUNTIF(inv_Estado,"AGOTADO")', "INVENTARIO"),
    ("Productos con stock bajo", '=COUNTIF(inv_Estado,"BAJO")', "INVENTARIO"),
    ("Cuentas por cobrar vencidas", '=COUNTIFS(car_Tipo,"CxC",car_Estado,"VENCIDO")', "CARTERA"),
    ("Cuentas por pagar vencidas", '=COUNTIFS(car_Tipo,"CxP",car_Estado,"VENCIDO")', "CARTERA"),
    ("Gastos fijos pendientes / vencidos del mes", '=COUNTIF(gf_Estado,"PENDIENTE")+COUNTIF(gf_Estado,"VENCIDO")',
     "GASTOS FIJOS"),
    ("Precios por debajo del costo real", f'=COUNTIF({q("COSTOS Y RENTABILIDAD")}!Q6:Q{INV_R2},"⚠*")',
     "COSTOS Y RENTABILIDAD"),
]
for i, (lab, frm, dest) in enumerate(alertas):
    rr = ALERTAS_R + 1 + i
    ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=5)
    c = ws.cell(rr, 2, lab)
    c.font = f(9, True, AZUL)
    c.hyperlink = Hyperlink(ref=c.coordinate, location=f"{q(dest)}!A1")
    v = ws.cell(rr, 6, frm)
    v.font, v.alignment, v.border = f(11, True), Alignment(horizontal="center"), BORDE
    ws.conditional_formatting.add(v.coordinate, CellIsRule(operator="greaterThan", formula=["0"],
                                  fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True)))
    ws.conditional_formatting.add(v.coordinate, CellIsRule(operator="equal", formula=["0"],
                                  fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(color="006100", bold=True)))
ws.cell(ALERTAS_R, 8, "CONVENCIONES").font = f(11, True, NAVY)
for i, (fill, txt) in enumerate([(INPUT_FILL, "Celda editable (escriba o elija de la lista)"),
                                 (CALC_FILL, "Celda calculada o del sistema (protegida)"),
                                 (F_AZUL, "Encabezado de columna editable"),
                                 (F_TEAL, "Encabezado de columna calculada")]):
    rr = ALERTAS_R + 1 + i
    ws.cell(rr, 8).fill = fill
    ws.cell(rr, 8).border = BORDE
    ws.merge_cells(start_row=rr, start_column=9, end_row=rr, end_column=13)
    ws.cell(rr, 9, txt).font = f(9)
ws.cell(ALERTAS_R + 6, 8, "Moneda: pesos colombianos (COP). Fechas: DD/MM/AAAA.").font = f(8, italic=True, color=GRIS)

# ====================================================== fuente por defecto y verificación
for w in wb.worksheets:
    for row in w.iter_rows():
        for c in row:
            if c.value is not None and c.font is not None and c.font.name != FONT:
                fnt = c.font
                c.font = Font(name=FONT, size=fnt.size or 10, bold=fnt.bold, italic=fnt.italic, color=fnt.color)

for g in DASH_CHARTS + UTIL_CHARTS + CANAL_CHARTS:
    g.roundedCorners = False
    if g.legend is not None:
        g.legend.position = "b"
        g.legend.overlay = False
    if not isinstance(g, PieChart):
        g.x_axis.delete = False
        g.y_axis.delete = False
        g.y_axis.number_format = '"$"#,##0'
    if isinstance(g, LineChart):
        for s_ in g.series:
            s_.smooth = False

wb.active = 0
wb.calculation.fullCalcOnLoad = True
wb.save(SALIDA)
print(json.dumps({"archivo": os.path.abspath(SALIDA), "hojas": len(wb.worksheets),
                  "nombres_definidos": len(wb.defined_names)}, ensure_ascii=False))
