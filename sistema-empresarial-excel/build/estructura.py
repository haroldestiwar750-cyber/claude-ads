"""
Estructura única del Sistema Empresarial (fuente de verdad).

Define nombres de hojas, filas de encabezado, columnas de cada tabla y
direcciones de los formularios. A partir de este archivo se generan:
  * el libro de Excel (construir_libro.py)
  * el módulo VBA modEstructura.bas con las constantes equivalentes
De esta forma las macros y la hoja nunca quedan desalineadas.

Uso CLI:
    python estructura.py --vba ../vba/modEstructura.bas
"""
import argparse
import json

# ---------------------------------------------------------------- hojas
SH = {
    "INI": "INICIO",
    "FAC": "FACTURACIÓN",
    "GUI": "GUÍAS DE ENVÍO",
    "INV": "INVENTARIO",
    "COM": "COMPRAS",
    "PRO": "PROVEEDORES",
    "CLI": "CLIENTES",
    "COS": "COSTOS Y RENTABILIDAD",
    "ING": "INGRESOS",
    "EGR": "EGRESOS",
    "GFI": "GASTOS FIJOS",
    "CAR": "CARTERA",
    "CAJ": "CAJA Y BANCOS",
    "UTI": "UTILIDADES",
    "REP": "REPORTES",
    "CFG": "CONFIGURACIÓN",
    "BD": "BASE DE DATOS",
    "HFA": "HISTORIAL FACTURAS",
    "HGU": "HISTORIAL GUÍAS",
    "MOV": "MOVIMIENTOS INV",
    "AUD": "AUDITORÍA",
    "CALC": "CALC",
}
ORDEN = ["INI", "FAC", "GUI", "INV", "COM", "PRO", "CLI", "COS", "ING", "EGR",
         "GFI", "CAR", "CAJ", "UTI", "REP", "CFG", "BD", "HFA", "HGU", "MOV",
         "AUD", "CALC"]

# formato de columna: txt, cop, int, pct, date, dt, num
# tipo: v = valor escrito por macro, i = captura del usuario, f = fórmula
def T(sheet, hdr, cap, cols):
    return {"sheet": sheet, "hdr": hdr, "first": hdr + 1,
            "last": hdr + cap, "cols": cols}

TABLAS = {
    # ventas detalle (una fila por línea de factura)
    "BD": T("BD", 4, 100000, [
        ("FACTURA", "N° factura", 12, "txt", "v"),
        ("FECHA", "Fecha", 11, "date", "v"),
        ("CODCLI", "Cód. cliente", 11, "txt", "v"),
        ("CLIENTE", "Cliente", 26, "txt", "v"),
        ("IDCLI", "Identificación", 14, "txt", "v"),
        ("SKU", "SKU", 16, "txt", "v"),
        ("PRODUCTO", "Producto", 30, "txt", "v"),
        ("CATEGORIA", "Categoría", 16, "txt", "v"),
        ("TALLA", "Talla", 7, "txt", "v"),
        ("COLOR", "Color", 10, "txt", "v"),
        ("CANT", "Cant.", 7, "int", "v"),
        ("PUNIT", "Precio unit.", 13, "cop", "v"),
        ("DESCPCT", "Desc. %", 8, "pct", "v"),
        ("DESCVAL", "Descuento $", 12, "cop", "v"),
        ("SUBTOTAL", "Subtotal neto", 14, "cop", "v"),
        ("COSTOUNIT", "Costo real unit.", 13, "cop", "v"),
        ("COSTOTOTAL", "Costo total", 13, "cop", "v"),
        ("UTILIDAD", "Utilidad bruta", 14, "cop", "v"),
        ("METODO", "Método pago", 13, "txt", "v"),
        ("ESTADOPAGO", "Estado pago", 13, "txt", "v"),
        ("ESTADO", "Estado", 10, "txt", "v"),
        ("USUARIO", "Usuario", 13, "txt", "v"),
        ("REGISTRO", "Registrado", 16, "dt", "v"),
    ]),
    "HFA": T("HFA", 4, 100000, [
        ("FACTURA", "N° factura", 12, "txt", "v"),
        ("FECHA", "Fecha", 11, "date", "v"),
        ("CODCLI", "Cód. cliente", 11, "txt", "v"),
        ("CLIENTE", "Cliente", 26, "txt", "v"),
        ("IDCLI", "Identificación", 14, "txt", "v"),
        ("TEL", "Teléfono", 13, "txt", "v"),
        ("DIRECCION", "Dirección de entrega", 30, "txt", "v"),
        ("CIUDAD", "Ciudad", 13, "txt", "v"),
        ("DEPTO", "Departamento", 13, "txt", "v"),
        ("BRUTO", "Subtotal bruto", 14, "cop", "v"),
        ("DESCUENTOS", "Descuentos", 12, "cop", "v"),
        ("TOTAL", "Total factura", 14, "cop", "v"),
        ("METODO", "Método pago", 13, "txt", "v"),
        ("ESTADOPAGO", "Estado pago", 13, "txt", "v"),
        ("PAGADO", "Pagado al facturar", 14, "cop", "v"),
        ("VENCE", "Vence", 11, "date", "v"),
        ("ESTADO", "Estado", 10, "txt", "v"),
        ("OBS", "Observaciones", 28, "txt", "v"),
        ("COSTO", "Costo total", 13, "cop", "v"),
        ("UTILIDAD", "Utilidad bruta", 13, "cop", "v"),
        ("USUARIO", "Usuario", 13, "txt", "v"),
        ("REGISTRO", "Registrado", 16, "dt", "v"),
        ("ANULUSR", "Anulada por", 13, "txt", "v"),
        ("ANULFECHA", "Fecha anulación", 16, "dt", "v"),
        ("ANULMOTIVO", "Motivo anulación", 26, "txt", "v"),
        ("PDF", "Archivo PDF", 30, "txt", "v"),
    ]),
    "MOV": T("MOV", 4, 100000, [
        ("ID", "ID movimiento", 13, "txt", "v"),
        ("FECHA", "Fecha y hora", 16, "dt", "v"),
        ("TIPO", "Tipo", 20, "txt", "v"),
        ("DOC", "Documento", 12, "txt", "v"),
        ("SKU", "SKU", 16, "txt", "v"),
        ("PRODUCTO", "Producto", 28, "txt", "v"),
        ("ENTRADA", "Entrada", 9, "int", "v"),
        ("SALIDA", "Salida", 9, "int", "v"),
        ("SALDO", "Stock resultante", 11, "int", "v"),
        ("COSTOUNIT", "Costo unit.", 12, "cop", "v"),
        ("VALOR", "Valor (costo)", 14, "cop", "v"),
        ("USUARIO", "Usuario", 13, "txt", "v"),
        ("MOTIVO", "Motivo / observación", 34, "txt", "v"),
    ]),
    "HGU": T("HGU", 4, 100000, [
        ("GUIA", "N° guía", 12, "txt", "v"),
        ("FECHA", "Fecha despacho", 12, "date", "v"),
        ("FACTURA", "Factura", 12, "txt", "v"),
        ("DEST", "Destinatario", 26, "txt", "v"),
        ("DOC", "Documento", 13, "txt", "v"),
        ("TEL", "Teléfono", 13, "txt", "v"),
        ("CIUDAD", "Ciudad", 13, "txt", "v"),
        ("DEPTO", "Departamento", 13, "txt", "v"),
        ("DIRECCION", "Dirección", 30, "txt", "v"),
        ("TRANSP", "Transportadora", 15, "txt", "v"),
        ("SEGUIMIENTO", "N° seguimiento", 15, "txt", "v"),
        ("PRODUCTOS", "Productos enviados", 40, "txt", "v"),
        ("UNIDADES", "Unidades", 9, "int", "v"),
        ("VALORDECL", "Valor declarado", 14, "cop", "v"),
        ("ESTADO", "Estado envío", 13, "txt", "v"),
        ("OBS", "Observaciones", 26, "txt", "v"),
        ("USUARIO", "Usuario", 13, "txt", "v"),
        ("REGISTRO", "Registrado", 16, "dt", "v"),
        ("PDF", "Archivo PDF", 30, "txt", "v"),
        ("ITEMS", "Detalle (técnico)", 20, "txt", "v"),
        ("PAQUETES", "Paquetes", 9, "int", "v"),
    ]),
    "AUD": T("AUD", 4, 100000, [
        ("FECHA", "Fecha y hora", 17, "dt", "v"),
        ("USUARIO", "Usuario", 14, "txt", "v"),
        ("MODULO", "Módulo", 16, "txt", "v"),
        ("ACCION", "Acción", 22, "txt", "v"),
        ("DOC", "Documento", 14, "txt", "v"),
        ("DETALLE", "Detalle", 60, "txt", "v"),
    ]),
    "INV": T("INV", 8, 1000, [
        ("SKU", "SKU (único)", 17, "txt", "i"),
        ("NOMBRE", "Nombre", 28, "txt", "i"),
        ("CATEGORIA", "Categoría", 16, "txt", "i"),
        ("DESCRIPCION", "Descripción", 26, "txt", "i"),
        ("TALLA", "Talla", 7, "txt", "i"),
        ("COLOR", "Color", 10, "txt", "i"),
        ("PROVEEDOR", "Proveedor", 20, "txt", "i"),
        ("COSTO", "Costo compra unit.", 13, "cop", "i"),
        ("TRANSP", "Transp./otros unit.", 12, "cop", "f"),
        ("COSTOREAL", "Costo real unit.", 13, "cop", "f"),
        ("PRECIO", "Precio venta", 13, "cop", "i"),
        ("STOCKINI", "Stock inicial", 9, "int", "i"),
        ("ENTRADAS", "Entradas", 9, "int", "f"),
        ("SALIDAS", "Salidas", 9, "int", "f"),
        ("STOCK", "Stock actual", 9, "int", "f"),
        ("STOCKMIN", "Stock mínimo", 9, "int", "i"),
        ("VALOR", "Valor inventario", 14, "cop", "f"),
        ("VENDIDAS", "Unid. vendidas", 9, "int", "f"),
        ("ROTACION", "Rotación", 11, "txt", "f"),
        ("ESTADO", "Estado", 12, "txt", "f"),
        ("ACTIVO", "Activo", 7, "txt", "i"),
        ("EMPAQUE", "Empaque unit.", 11, "cop", "i"),
        ("COMISION", "Comisión %", 9, "pct", "i"),
        ("OTROS", "Otros var. unit.", 11, "cop", "i"),
    ]),
    "CLI": T("CLI", 6, 2000, [
        ("CODIGO", "Código", 10, "txt", "i"),
        ("NOMBRE", "Nombre completo", 28, "txt", "i"),
        ("IDENT", "Identificación", 14, "txt", "i"),
        ("TEL", "Teléfono", 13, "txt", "i"),
        ("WHATSAPP", "WhatsApp", 13, "txt", "i"),
        ("CORREO", "Correo", 22, "txt", "i"),
        ("DIRECCION", "Dirección", 28, "txt", "i"),
        ("CIUDAD", "Ciudad", 13, "txt", "i"),
        ("DEPTO", "Departamento", 14, "txt", "i"),
        ("PRIMERA", "Primera compra", 12, "date", "f"),
        ("ULTIMA", "Última compra", 12, "date", "f"),
        ("NFACT", "N° facturas", 9, "int", "f"),
        ("TOTAL", "Total compras", 14, "cop", "f"),
        ("SALDO", "Saldo pendiente", 14, "cop", "f"),
        ("CLASIF", "Clasificación", 13, "txt", "f"),
        ("CARTERA", "Cartera", 12, "txt", "f"),
        ("COMBO", "Selector (técnico)", 30, "txt", "f"),
    ]),
    "PRO": T("PRO", 6, 500, [
        ("CODIGO", "Código", 10, "txt", "i"),
        ("NOMBRE", "Nombre / razón social", 28, "txt", "i"),
        ("NIT", "NIT / identificación", 15, "txt", "i"),
        ("TEL", "Teléfono", 13, "txt", "i"),
        ("CORREO", "Correo", 22, "txt", "i"),
        ("DIRECCION", "Dirección", 26, "txt", "i"),
        ("CIUDAD", "Ciudad", 13, "txt", "i"),
        ("CONTACTO", "Contacto", 18, "txt", "i"),
        ("PRODUCTOS", "Productos suministrados", 30, "txt", "i"),
        ("UNIDADES", "Unid. compradas", 11, "int", "f"),
        ("TOTAL", "Total comprado", 14, "cop", "f"),
        ("SALDO", "Saldo pendiente", 14, "cop", "f"),
        ("ESTADO", "Estado cuenta", 12, "txt", "f"),
    ]),
    "COM": T("COM", 23, 100000, [
        ("NUMERO", "N° compra", 11, "txt", "v"),
        ("FECHA", "Fecha", 11, "date", "v"),
        ("PROVEEDOR", "Proveedor", 22, "txt", "v"),
        ("SKU", "SKU", 16, "txt", "v"),
        ("PRODUCTO", "Producto", 26, "txt", "v"),
        ("CANT", "Cantidad", 9, "int", "v"),
        ("COSTOUNIT", "Costo unit.", 12, "cop", "v"),
        ("COSTOTOTAL", "Costo total", 13, "cop", "v"),
        ("TRANSPORTE", "Transporte", 12, "cop", "v"),
        ("OTROS", "Otros costos", 12, "cop", "v"),
        ("COSTOREAL", "Costo real total", 14, "cop", "v"),
        ("COSTOREALU", "Costo real unit.", 13, "cop", "v"),
        ("FORMAPAGO", "Forma de pago", 13, "txt", "v"),
        ("ESTADOPAGO", "Estado pago", 13, "txt", "v"),
        ("USUARIO", "Usuario", 13, "txt", "v"),
        ("REGISTRO", "Registrado", 16, "dt", "v"),
        ("ESTADO", "Estado", 10, "txt", "v"),
        ("OBS", "Observaciones", 26, "txt", "v"),
    ]),
    "ING": T("ING", 17, 100000, [
        ("ID", "ID", 11, "txt", "v"),
        ("FECHA", "Fecha", 11, "date", "v"),
        ("ORIGEN", "Origen", 13, "txt", "v"),
        ("TIPO", "Tipo", 16, "txt", "v"),
        ("CLASIF", "Clasificación", 24, "txt", "v"),
        ("METODO", "Método pago", 13, "txt", "v"),
        ("CUENTA", "Cuenta caja", 12, "txt", "v"),
        ("DOC", "Documento", 12, "txt", "v"),
        ("TERCERO", "Tercero", 24, "txt", "v"),
        ("CONCEPTO", "Concepto", 30, "txt", "v"),
        ("VALOR", "Valor", 14, "cop", "v"),
        ("USUARIO", "Usuario", 13, "txt", "v"),
        ("REGISTRO", "Registrado", 16, "dt", "v"),
        ("ESTADO", "Estado", 10, "txt", "v"),
        ("OBS", "Observaciones", 26, "txt", "v"),
    ]),
    "EGR": T("EGR", 24, 100000, [
        ("ID", "ID", 11, "txt", "v"),
        ("FECHA", "Fecha", 11, "date", "v"),
        ("CONCEPTO", "Concepto", 30, "txt", "v"),
        ("CATEGORIA", "Categoría", 17, "txt", "v"),
        ("TIPO", "Tipo de gasto", 14, "txt", "v"),
        ("VALOR", "Valor", 14, "cop", "v"),
        ("METODO", "Método pago", 13, "txt", "v"),
        ("CUENTA", "Cuenta caja", 12, "txt", "v"),
        ("RESPONSABLE", "Responsable", 14, "txt", "v"),
        ("COMPROBANTE", "Comprobante", 13, "txt", "v"),
        ("OBS", "Observaciones", 24, "txt", "v"),
        ("ORIGEN", "Origen", 12, "txt", "v"),
        ("DOC", "Documento", 18, "txt", "v"),
        ("USUARIO", "Usuario", 13, "txt", "v"),
        ("REGISTRO", "Registrado", 16, "dt", "v"),
        ("ESTADO", "Estado", 10, "txt", "v"),
    ]),
    "GFI": T("GFI", 6, 100, [
        ("CONCEPTO", "Concepto (único)", 28, "txt", "i"),
        ("CATEGORIA", "Categoría", 17, "txt", "i"),
        ("VALOR", "Valor mensual", 14, "cop", "i"),
        ("DIA", "Día de vencimiento", 11, "int", "i"),
        ("RESPONSABLE", "Responsable", 14, "txt", "i"),
        ("METODO", "Método de pago", 14, "txt", "i"),
        ("ACTIVO", "Activo", 8, "txt", "i"),
        ("VENCE", "Vence este mes", 13, "date", "f"),
        ("PAGADO", "Pagado este mes", 14, "cop", "f"),
        ("ESTADO", "Estado", 17, "txt", "f"),
        ("ULTIMO", "Último pago", 12, "date", "f"),
    ]),
    "CAR": T("CAR", 14, 3000, [
        ("ID", "ID", 10, "txt", "v"),
        ("TIPO", "Tipo", 12, "txt", "v"),
        ("CODTER", "Cód. tercero", 11, "txt", "v"),
        ("TERCERO", "Cliente / proveedor", 26, "txt", "v"),
        ("DOC", "Documento", 12, "txt", "v"),
        ("FECHA", "Fecha deuda", 11, "date", "v"),
        ("VALORINI", "Valor inicial", 14, "cop", "v"),
        ("PAGOINI", "Pago inicial", 13, "cop", "v"),
        ("ABONOS", "Abonos (total)", 14, "cop", "f"),
        ("SALDO", "Saldo pendiente", 14, "cop", "f"),
        ("VENCE", "Vencimiento", 11, "date", "v"),
        ("MORA", "Días mora", 9, "int", "f"),
        ("ESTADO", "Estado", 14, "txt", "f"),
        ("REGESTADO", "Registro", 10, "txt", "v"),
        ("OBS", "Observaciones", 24, "txt", "v"),
        ("COMBO", "Selector (técnico)", 20, "txt", "f"),
        ("IDX", "Idx (técnico)", 6, "int", "f"),
        ("LISTA", "Lista (técnico)", 20, "txt", "f"),
    ]),
    "CIE": T("CAJ", 27, 5000, [
        ("FECHA", "Fecha cierre", 16, "dt", "v"),
        ("TIPO", "Tipo", 13, "txt", "v"),
        ("DESDE", "Desde", 11, "date", "v"),
        ("HASTA", "Hasta", 11, "date", "v"),
        ("ENTRADAS", "Entradas", 14, "cop", "v"),
        ("SALIDAS", "Salidas", 14, "cop", "v"),
        ("SALDO", "Saldo final total", 15, "cop", "v"),
        ("C1", "Cuenta 1", 13, "cop", "v"),
        ("C2", "Cuenta 2", 13, "cop", "v"),
        ("C3", "Cuenta 3", 13, "cop", "v"),
        ("C4", "Cuenta 4", 13, "cop", "v"),
        ("C5", "Cuenta 5", 13, "cop", "v"),
        ("USUARIO", "Usuario", 13, "txt", "v"),
    ]),
}
# la tabla de cierres empieza en la columna B de CAJA Y BANCOS
TABLAS["CIE"]["col0"] = 2
for k, t in TABLAS.items():
    t.setdefault("col0", 1)


def col_index(tabla, key):
    """Índice de columna (1-based) de un campo."""
    t = TABLAS[tabla]
    for i, c in enumerate(t["cols"]):
        if c[0] == key:
            return t["col0"] + i
    raise KeyError(f"{tabla}.{key}")


# ----------------------------------------------- celdas de formularios
FAC = {
    "NUM": "I2", "FECHA": "I3", "ESTADO": "I4", "VENCE": "I5",
    "CLIENTE": "C8", "ID": "I8", "NOMBRE": "C9", "TEL": "I9",
    "DIRCLI": "C10", "CIUDAD": "I10", "ENTREGA": "C11", "WHATSAPP": "I11",
    "METODO": "C30", "ESTADOPAGO": "C31", "PAGADO": "C32", "SALDO": "C33",
    "BRUTO": "I30", "DESC": "I31", "TOTAL": "I32", "UNIDADES": "I33",
    "OBS": "A36", "PAGADOEF": "T30",
}
FAC_L1, FAC_LN = 14, 28  # filas de líneas
GUI = {
    "NUM": "I2", "FECHA": "I3", "FACTURA": "I4", "ESTADO": "I5",
    "DEST": "G8", "DOC": "G9", "TEL": "G10", "CIUDAD": "G11",
    "DEPTO": "G12", "DIR": "G13", "TRANSP": "C17", "SEG": "G17",
    "VALOR": "C18", "PAQ": "G18", "OBS": "A36", "REG": "K3",
    "UNIDADES": "F33",
}
GUI_L1, GUI_LN = 21, 32
COM = {"NUM": "C5", "FECHA": "F5", "PROV": "I5", "FORMA": "C6",
       "ESTADOPAGO": "F6", "PAGADO": "I6", "TRANSP": "C7", "OTROS": "F7",
       "VENCE": "I7", "OBS": "C8", "SUBTOTAL": "H19", "TOTAL": "H20"}
COM_L1, COM_LN = 11, 18
ING = {"FECHA": "C5", "TIPO": "C6", "CONCEPTO": "C7", "VALOR": "C8",
       "METODO": "F5", "TERCERO": "F6", "DOC": "F7", "OBS": "F8"}
EGR = {"FECHA": "C5", "CONCEPTO": "C6", "CATEGORIA": "C7", "TIPO": "C8",
       "VALOR": "F5", "METODO": "F6", "RESP": "F7", "COMPROB": "F8",
       "OBS": "C9"}
CAR = {"OBLIG": "C5", "FECHA": "C6", "VALOR": "C7", "METODO": "C8",
       "SALDO": "F6", "OBS": "F7"}
CAJ = {"FINI": "C3", "TIPO": "C14", "DESDE": "C15", "HASTA": "C16",
       "CTA1": 6, "CTAN": 10, "PER1": 19, "PERN": 23, "PERTOT": 24}
REP = {"TIPO": "C4", "DESDE": "C5", "HASTA": "C6", "OUT": 10}
INI = {"ANIO": "D4", "MES": "F4", "CAT": "H4", "DIA": "K4", "USUARIO": "M4"}

# configuración (columna B)
CFG_ROWS = [
    ("cfgEmpresa", "Nombre de la empresa", "MODA & ESTILO S.A.S.", "txt"),
    ("cfgNIT", "NIT / identificación", "901.234.567-8", "txt"),
    ("cfgRegimen", "Régimen / leyenda tributaria", "No responsable de IVA", "txt"),
    ("cfgDireccion", "Dirección", "Calle 45 # 12-30, Local 102", "txt"),
    ("cfgCiudad", "Ciudad", "Bogotá D.C.", "txt"),
    ("cfgTelefono", "Teléfono", "300 123 4567", "txt"),
    ("cfgCorreo", "Correo", "ventas@modayestilo.co", "txt"),
    ("cfgMoneda", "Moneda", "COP - Peso colombiano", "txt"),
    ("cfgBanco", "Banco", "Bancolombia", "txt"),
    ("cfgTipoCuenta", "Tipo de cuenta", "Ahorros", "txt"),
    ("cfgCuenta", "N° de cuenta", "123-456789-00", "txt"),
    ("cfgTitular", "Titular de la cuenta", "MODA & ESTILO S.A.S.", "txt"),
    ("cfgNequi", "Nequi", "300 123 4567", "txt"),
    ("cfgDaviplata", "Daviplata", "300 123 4567", "txt"),
    ("cfgPrefFactura", "Prefijo de factura", "FV-", "txt"),
    ("cfgSigFactura", "Siguiente N° de factura", 1, "int"),
    ("cfgPrefGuia", "Prefijo de guía", "GE-", "txt"),
    ("cfgSigGuia", "Siguiente N° de guía", 1, "int"),
    ("cfgPrefCompra", "Prefijo de compra", "OC-", "txt"),
    ("cfgSigCompra", "Siguiente N° de compra", 1, "int"),
    ("cfgMargenComp", "Margen precio competitivo (s/ventas)", 0.30, "pct"),
    ("cfgMargenRec", "Margen precio recomendado (s/ventas)", 0.45, "pct"),
    ("cfgMargenPrem", "Margen precio premium (s/ventas)", 0.60, "pct"),
    ("cfgMargenMin", "Margen mínimo aceptable (alerta)", 0.20, "pct"),
    ("cfgDiasCredito", "Días de crédito por defecto", 30, "int"),
    ("cfgFrecuente", "Facturas para cliente frecuente", 3, "int"),
    ("cfgEmpaque", "Costo de empaque por defecto (unid.)", 1500, "cop"),
    ("cfgComision", "Comisión por defecto (% del precio)", 0.0, "pct"),
    ("cfgCatCompras", "Categoría de gasto para compras de mercancía", "Compras", "txt"),
    ("cfgLogo", "Ruta del logo (use el botón CARGAR LOGO)", "", "txt"),
    ("cfgPie", "Pie de factura", "Gracias por su compra. Cambios dentro de los 30 días siguientes con factura y etiquetas. Esta factura de venta se asimila en sus efectos a una letra de cambio (Art. 774 C. Co.).", "txt"),
    ("cfgPermitirNeg", "Permitir inventario negativo (SI/NO)", "NO", "txt"),
    ("cfgClave", "Clave de autorización", "admin123", "secret"),
    ("cfgClaveHojas", "Clave de protección de hojas", "1234", "secret"),
]
CFG_FIRST_ROW = 4  # fila de cfgEmpresa
CFG_LIST_HDR = 4
CFG_LIST_FIRST = 5
CFG_LIST_N = 40
CFG_LISTS = {  # nombre -> (columna, título, valores)
    "lstCategorias": ("D", "Categorías de producto", [
        "Camisetas", "Camisas", "Pantalones", "Jeans", "Vestidos", "Chaquetas",
        "Calzado deportivo", "Calzado formal", "Sandalias", "Bolsos",
        "Cinturones", "Gorras", "Accesorios"]),
    "lstMetodos": ("E", "Métodos de pago", [
        "Efectivo", "Transferencia", "Nequi", "Daviplata", "Tarjeta", "Otro"]),
    "lstMetodoCuenta": ("F", "Cuenta asociada al método", [
        "EFECTIVO", "BANCOS", "NEQUI", "DAVIPLATA", "BANCOS", "OTROS"]),
    "lstCuentas": ("G", "Cuentas de caja (5)", [
        "EFECTIVO", "BANCOS", "NEQUI", "DAVIPLATA", "OTROS"]),
    "lstTransportadoras": ("H", "Transportadoras", [
        "Servientrega", "Interrapidísimo", "Coordinadora", "Envía", "TCC",
        "Deprisa", "4-72", "Mensajero propio"]),
    "lstUsuarios": ("I", "Usuarios / responsables", [
        "Administrador", "Vendedor 1", "Vendedor 2", "Bodega", "Contador"]),
    "lstEstadosPago": ("J", "Estados de pago", [
        "PAGADA", "PENDIENTE", "ABONO PARCIAL"]),
    "lstCatGasto": ("K", "Categorías de gasto", [
        "Arriendo", "Servicios públicos", "Internet", "Publicidad",
        "Transporte", "Empaques", "Nómina", "Comisiones", "Compras",
        "Impuestos", "Mantenimiento", "Otros gastos"]),
    "lstTipoGasto": ("L", "Tipos de gasto", ["FIJO", "VARIABLE", "EXTRAORDINARIO"]),
    "lstEstadosEnvio": ("M", "Estados de envío", [
        "PREPARANDO", "DESPACHADO", "EN TRÁNSITO", "ENTREGADO", "DEVUELTO"]),
    "lstTiposIngreso": ("N", "Tipos de ingreso adicional", [
        "OTRO INGRESO", "APORTE DE SOCIOS", "PRÉSTAMO RECIBIDO",
        "RENDIMIENTOS FINANCIEROS", "VENTA DE ACTIVOS"]),
    "lstTallas": ("O", "Tallas", [
        "XS", "S", "M", "L", "XL", "XXL", "28", "30", "32", "34", "36", "37",
        "38", "39", "40", "41", "42", "43", "ÚNICA"]),
    "lstColores": ("P", "Colores", [
        "Negro", "Blanco", "Azul", "Azul oscuro", "Gris", "Rojo", "Verde",
        "Beige", "Café", "Rosado", "Amarillo", "Multicolor"]),
    "lstDeptos": ("Q", "Departamentos", [
        "Amazonas", "Antioquia", "Arauca", "Atlántico", "Bogotá D.C.", "Bolívar",
        "Boyacá", "Caldas", "Caquetá", "Casanare", "Cauca", "Cesar", "Chocó",
        "Córdoba", "Cundinamarca", "Guainía", "Guaviare", "Huila", "La Guajira",
        "Magdalena", "Meta", "Nariño", "Norte de Santander", "Putumayo",
        "Quindío", "Risaralda", "San Andrés", "Santander", "Sucre", "Tolima",
        "Valle del Cauca", "Vaupés", "Vichada"]),
    "lstReportes": ("R", "Tipos de reporte", [
        "Ventas diario", "Ventas mensual", "Ventas anual", "Inventario",
        "Productos vendidos", "Utilidades", "Gastos", "Cartera", "Compras",
        "Clientes", "Flujo de caja"]),
    "lstMeses": ("S", "Meses", [
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
        "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "TODOS"]),
    "lstSiNo": ("T", "Sí / No", ["SI", "NO"]),
}


def cfg_cell(name):
    for i, r in enumerate(CFG_ROWS):
        if r[0] == name:
            return f"B{CFG_FIRST_ROW + i}"
    raise KeyError(name)


# --------------------------------------------------- generación VBA
def vba_module():
    L = ["Attribute VB_Name = \"modEstructura\"",
         "'==============================================================",
         "' modEstructura - CONSTANTES GENERADAS AUTOMATICAMENTE",
         "' No editar a mano: se generan desde build/estructura.py para que",
         "' las macros y el diseño de las hojas siempre coincidan.",
         "'==============================================================",
         "Option Explicit", ""]
    for k in ORDEN:
        L.append(f'Public Const SH_{k} As String = "{SH[k]}"')
    L.append("")
    for tk, t in TABLAS.items():
        L.append(f"' --- Tabla {tk} ({SH[t['sheet']]})")
        L.append(f"Public Const {tk}_HDR As Long = {t['hdr']}")
        L.append(f"Public Const {tk}_FIRST As Long = {t['first']}")
        L.append(f"Public Const {tk}_LAST As Long = {t['last']}")
        L.append(f"Public Const {tk}_NCOLS As Long = {len(t['cols'])}")
        L.append(f"Public Const {tk}_COL0 As Long = {t['col0']}")
        for c in t["cols"]:
            L.append(f"Public Const {tk}_{c[0]} As Long = {col_index(tk, c[0])}")
        L.append("")
    for name, d in (("FAC", FAC), ("GUI", GUI), ("FCOM", COM), ("FING", ING),
                    ("FEGR", EGR), ("FCAR", CAR), ("FREP", REP), ("FINI", INI)):
        for k, v in d.items():
            if isinstance(v, int):
                L.append(f"Public Const {name}_{k} As Long = {v}")
            else:
                L.append(f'Public Const {name}_{k} As String = "{v}"')
        L.append("")
    for k, v in CAJ.items():
        if isinstance(v, int):
            L.append(f"Public Const FCAJ_{k} As Long = {v}")
        else:
            L.append(f'Public Const FCAJ_{k} As String = "{v}"')
    L += ["", f"Public Const FAC_L1 As Long = {FAC_L1}",
          f"Public Const FAC_LN As Long = {FAC_LN}",
          f"Public Const GUI_L1 As Long = {GUI_L1}",
          f"Public Const GUI_LN As Long = {GUI_LN}",
          f"Public Const FCOM_L1 As Long = {COM_L1}",
          f"Public Const FCOM_LN As Long = {COM_LN}", ""]
    for r in CFG_ROWS:
        L.append(f'Public Const {r[0].upper().replace("CFG", "CFG_")} As String = "{cfg_cell(r[0])}"')
    for name, (colL, _t, _v) in CFG_LISTS.items():
        L.append(f'Public Const LST_{name[3:].upper()} As String = "{colL}"')
    L += [f"Public Const CFG_LIST_FIRST As Long = {CFG_LIST_FIRST}",
          f"Public Const CFG_LIST_N As Long = {CFG_LIST_N}", ""]
    return "\r\n".join(L) + "\r\n"


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--vba", help="ruta de salida de modEstructura.bas")
    a = ap.parse_args()
    if a.vba:
        with open(a.vba, "w", encoding="utf-8", newline="") as f:
            f.write(vba_module())
        print(json.dumps({"ok": True, "archivo": a.vba}))
    else:
        print(json.dumps({"hojas": SH, "tablas": list(TABLAS)}, ensure_ascii=False))
