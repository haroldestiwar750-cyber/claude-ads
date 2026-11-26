"""Datos de prueba del ERP comercial.

Simula, en orden cronológico, las mismas reglas que aplican las macros VBA:
compras (costo promedio ponderado con prorrateo de fletes), ventas (costo
congelado al vender), anulaciones, devoluciones, cartera, abonos, ingresos,
egresos, gastos fijos, comisiones, guías, kárdex y bitácora.

Uso: lo importa generar_erp.py (no se ejecuta solo).
"""
import datetime as dt
import random

D = dt.date
HOY = D(2026, 11, 26)
random.seed(20261126)

# ----------------------------------------------------------------- catálogos
CATEGORIAS = ["Ropa", "Calzado deportivo", "Accesorios"]
TALLAS = ["XS", "S", "M", "L", "XL", "28", "30", "32", "34", "36",
          "35", "36", "37", "38", "39", "40", "41", "42", "43", "Única"]
TALLAS = list(dict.fromkeys(TALLAS))
COLORES = ["Negro", "Blanco", "Gris", "Azul", "Rojo", "Verde", "Beige", "Rosado", "Multicolor"]
CANALES = ["WhatsApp", "Facebook", "Instagram", "TikTok LIVE", "Venta directa"]
METODOS = ["Efectivo", "Transferencia", "Nequi", "Daviplata", "Tarjeta"]
ESTADOS_PAGO = ["PAGADA", "CRÉDITO", "ABONO PARCIAL"]
TRANSPORTADORAS = ["Servientrega", "Interrapidísimo", "Coordinadora", "Envía", "TCC", "Entrega propia"]
CONCEPTOS_ING = ["Ventas", "Abonos", "Otros ingresos"]
CONCEPTOS_EGR = ["Compras", "Arriendo", "Servicios", "Publicidad", "Transporte", "Empaques",
                 "Nómina", "Comisiones", "Impuestos", "Devoluciones a clientes", "Otros gastos"]
ROLES = ["Administrador", "Auxiliar administrativo", "Vendedor", "Consulta"]
ESTADOS_GUIA = ["PENDIENTE", "DESPACHADO", "EN TRÁNSITO", "ENTREGADO", "DEVUELTO"]
TIPOS_MOV = ["INICIAL", "COMPRA", "VENTA", "ANULACIÓN VENTA", "DEVOLUCIÓN", "AJUSTE ENTRADA", "AJUSTE SALIDA"]
CIUDADES = ["Bogotá", "Medellín", "Cali", "Barranquilla", "Cartagena", "Bucaramanga", "Pereira",
            "Manizales", "Cúcuta", "Ibagué", "Santa Marta", "Villavicencio", "Pasto", "Neiva"]
ESTADOS = ["ACTIVO", "INACTIVO"]
TIPOS_DOC = ["CC", "NIT", "CE", "PAS", "TI"]

SALDOS_INICIALES = {"Efectivo": 800000, "Transferencia": 18000000, "Nequi": 1000000,
                    "Daviplata": 300000, "Tarjeta": 0}
FECHA_SALDO_INICIAL = D(2026, 9, 30)

EMPAQUE = 1500
COMISION_PROM = 0.05


def hash_clave(usuario, clave):
    """Mismo algoritmo que HashClave() en modSeguridad.bas."""
    s = usuario.lower() + "|" + clave + "|ERP"
    h = 5381
    for r in range(1, 501):
        for ch in s:
            h = h * 33 + ord(ch) + r
            h = h % 2147483647
    return str(h)


# --------------------------------------------------------------- maestros
PROVEEDORES = [
    ["PRV-001", "Textiles Andinos S.A.S.", "900456789-1", "Carlos Mejía", "6015550101", "3105550101",
     "ventas@textilesandinos.ejemplo.co", "Cra 50 # 12-30", "Bogotá", "Camisetas, jeans, buzos",
     "Pedido mínimo 24 und. Crédito 30 días", "Bancolombia", "Ahorros", "123-456789-01"],
    ["PRV-002", "Calzado Deportivo del Valle S.A.S.", "901234567-2", "Diana Ospina", "6025550202", "3155550202",
     "pedidos@calzadovalle.ejemplo.co", "Cl 15 # 30-45", "Cali", "Tenis running y urbanos",
     "Lista de precios trimestral. Crédito 45 días", "Davivienda", "Corriente", "0550-1234-5678"],
    ["PRV-003", "Importadora Accesorios Bogotá Ltda.", "830987654-3", "Jorge Pardo", "6015550303", "3205550303",
     "info@impaccesorios.ejemplo.co", "Av 68 # 22-10", "Bogotá", "Gorras, morrales, medias",
     "Contado con 3% descuento", "BBVA", "Ahorros", "0013-0099-1122"],
]

PRODUCTOS = [  # base, nombre, categoría, proveedor, tallas, colores, costo, precio, mínimo
    ("CAMB", "Camiseta básica algodón", "Ropa", "PRV-001", ["S", "M", "L"], ["Negro", "Blanco"], 18000, 39900, 4),
    ("JEAN", "Jean slim fit", "Ropa", "PRV-001", ["30", "32", "34"], ["Azul"], 45000, 99900, 3),
    ("BUZO", "Buzo con capota", "Ropa", "PRV-001", ["M", "L"], ["Gris"], 38000, 85000, 2),
    ("TRUN", "Tenis running Pro", "Calzado deportivo", "PRV-002", ["38", "40", "42"], ["Negro"], 95000, 219900, 2),
    ("TURB", "Tenis urbanos clásicos", "Calzado deportivo", "PRV-002", ["37", "39", "41"], ["Blanco"], 80000, 179900, 2),
    ("GORR", "Gorra snapback", "Accesorios", "PRV-003", ["Única"], ["Negro", "Rojo"], 12000, 35000, 5),
    ("MORR", "Morral deportivo", "Accesorios", "PRV-003", ["Única"], ["Negro"], 40000, 89900, 2),
    ("MEDI", "Medias deportivas x3", "Accesorios", "PRV-003", ["Única"], ["Blanco"], 7000, 19900, 8),
]

ABREV_COLOR = {"Negro": "NEG", "Blanco": "BLA", "Gris": "GRI", "Azul": "AZU", "Rojo": "ROJ"}


def construir_skus():
    skus = []
    for base, nom, cat, prv, tallas, colores, costo, precio, minimo in PRODUCTOS:
        for t in tallas:
            for c in colores:
                tcode = "U" if t == "Única" else t
                code = f"{base}-{tcode}-{ABREV_COLOR[c]}"
                skus.append(dict(sku=code, prod=nom, cat=cat, talla=t, color=c, prov=prv,
                                 costo_ref=costo, precio=precio, minimo=minimo, ini=0))
    return skus


SKUS = construir_skus()
SKU = {s["sku"]: s for s in SKUS}

CLIENTES = [
    ["CLI-0001", "María Fernanda López", "CC", "1020304050", "3001112233", "3001112233", "mafe.lopez@ejemplo.com",
     "Cl 85 # 15-20 Apto 301", "Bogotá", D(2026, 6, 15), "Instagram"],
    ["CLI-0002", "Juan Camilo Pérez", "CC", "1036654321", "3012223344", "3012223344", "jcperez@ejemplo.com",
     "Cra 43A # 1-50", "Medellín", D(2026, 7, 20), "WhatsApp"],
    ["CLI-0003", "Valentina Ruiz", "CC", "1144556677", "3023334455", "3023334455", "vale.ruiz@ejemplo.com",
     "Av 6N # 23-45", "Cali", D(2026, 8, 1), "TikTok LIVE"],
    ["CLI-0004", "Distribuciones El Faro S.A.S.", "NIT", "900123456-7", "6053334455", "3044445566",
     "compras@elfaro.ejemplo.co", "Cl 72 # 54-30", "Barranquilla", D(2026, 3, 10), "Venta directa"],
    ["CLI-0005", "Santiago Morales", "CC", "80123456", "3055556677", "3055556677", "smorales@ejemplo.com",
     "Cra 7 # 45-12", "Bogotá", D(2026, 9, 5), "Facebook"],
    ["CLI-0006", "Laura Martínez", "CC", "1098765432", "3066667788", "3066667788", "lauram@ejemplo.com",
     "Cl 36 # 27-15", "Bucaramanga", D(2026, 10, 12), "TikTok LIVE"],
    ["CLI-0007", "Daniel Castro", "CC", "1088223344", "3077778899", "3077778899", "dcastro@ejemplo.com",
     "Cra 8 # 20-33", "Pereira", D(2026, 11, 2), "Instagram"],
    ["CLI-0008", "Carolina Herrera Díaz", "CC", "52987654", "3088889900", "3088889900", "caro.herrera@ejemplo.com",
     "Cl 140 # 9-45", "Bogotá", D(2026, 11, 15), "WhatsApp"],
    ["CLI-0009", "Andrés Felipe Gil", "CC", "1047112233", "3099990011", "3099990011", "afgil@ejemplo.com",
     "Bocagrande Cra 3 # 5-10", "Cartagena", D(2026, 11, 20), "TikTok LIVE"],
    ["CLI-0010", "Paula Andrea Ríos", "CC", "30123987", "3101234567", "3101234567", "paula.rios@ejemplo.com",
     "Cl 65 # 23-40", "Manizales", D(2026, 4, 1), "Facebook"],
]
CLI = {c[0]: c for c in CLIENTES}

VENDEDORES = [
    ["VEN-01", "Laura Gómez", "1032456789", "laura", "Vendedor", "ACTIVO", "WhatsApp", 0.05, 4000000],
    ["VEN-02", "Andrés Rojas", "1019876543", "andres", "Vendedor", "ACTIVO", "TikTok LIVE", 0.04, 5000000],
    ["VEN-03", "Camila Torres", "1001234567", "camila", "Vendedor", "ACTIVO", "Venta directa", 0.05, 3000000],
]
VEN = {v[0]: v for v in VENDEDORES}
VEND_POR_CANAL = {"WhatsApp": "VEN-01", "Facebook": "VEN-01", "Instagram": "VEN-01",
                  "TikTok LIVE": "VEN-02", "Venta directa": "VEN-03"}

USUARIOS = [  # usuario, nombre, rol, vendedor, estado, clave inicial
    ["admin", "Propietario (Administrador)", "Administrador", "", "ACTIVO", "Admin2026"],
    ["aux", "Auxiliar administrativo", "Auxiliar administrativo", "", "ACTIVO", "Aux2026"],
    ["laura", "Laura Gómez", "Vendedor", "VEN-01", "ACTIVO", "Laura2026"],
    ["andres", "Andrés Rojas", "Vendedor", "VEN-02", "ACTIVO", "Andres2026"],
    ["camila", "Camila Torres", "Vendedor", "VEN-03", "ACTIVO", "Camila2026"],
    ["contador", "Contador externo", "Consulta", "", "INACTIVO", "Conta2026"],
]

PERMISOS = [  # código, descripción, admin, aux, vendedor, consulta
    ("VER_DASHBOARD", "Ver dashboard gerencial", "SI", "SI", "NO", "SI"),
    ("FACTURAR", "Crear y guardar facturas", "SI", "SI", "SI", "NO"),
    ("ANULAR_FACTURA", "Anular facturas / registrar devoluciones", "SI", "NO", "NO", "NO"),
    ("VER_COSTOS", "Ver costos, utilidad y márgenes", "SI", "SI", "NO", "SI"),
    ("CLIENTES", "Crear y editar clientes", "SI", "SI", "SI", "NO"),
    ("PROVEEDORES", "Crear y editar proveedores", "SI", "SI", "NO", "NO"),
    ("COMPRAS", "Registrar compras", "SI", "SI", "NO", "NO"),
    ("INVENTARIO", "Crear productos / SKU", "SI", "SI", "NO", "NO"),
    ("AJUSTE_INVENTARIO", "Autorizar ajustes e inventario negativo", "SI", "NO", "NO", "NO"),
    ("GUIAS", "Crear guías de envío", "SI", "SI", "SI", "NO"),
    ("FINANZAS", "Ingresos, egresos, gastos fijos, caja y bancos", "SI", "SI", "NO", "SI"),
    ("CARTERA", "Cartera y registro de abonos", "SI", "SI", "NO", "SI"),
    ("COMISIONES", "Liquidar y pagar comisiones", "SI", "NO", "NO", "NO"),
    ("REPORTES", "Generar reportes PDF / Excel", "SI", "SI", "NO", "SI"),
    ("USUARIOS", "Administrar usuarios y permisos", "SI", "NO", "NO", "NO"),
    ("CONFIGURACION", "Modificar configuración y catálogos", "SI", "NO", "NO", "NO"),
    ("RESPALDOS", "Crear y restaurar copias de seguridad", "SI", "NO", "NO", "NO"),
]

GASTOS_FIJOS = [  # código, concepto, descripción, valor, día, método, beneficiario, activo
    ["GF-01", "Arriendo", "Arriendo local / bodega", 1800000, 5, "Transferencia", "Inmobiliaria Centro", "SI"],
    ["GF-02", "Servicios", "Energía, agua e internet", 350000, 10, "Transferencia", "Empresas de servicios", "SI"],
    ["GF-03", "Otros gastos", "Software y hosting", 120000, 15, "Transferencia", "Proveedores digitales", "SI"],
    ["GF-04", "Nómina", "Salario auxiliar administrativo", 1600000, 30, "Transferencia", "Auxiliar administrativo", "SI"],
]


# --------------------------------------------------------------- simulación
class Sim:
    def __init__(self):
        self.stock = {s: 0 for s in SKU}
        self.costo = {s: float(SKU[s]["costo_ref"]) for s in SKU}
        self.kardex, self.facturas, self.detalle = [], [], []
        self.compras, self.compra_det = [], []
        self.cartera, self.abonos = [], []
        self.ingresos, self.egresos = [], []
        self.guias, self.cierres, self.bitacora = [], [], []
        self.n = dict(fac=0, oc=0, gu=0, k=0, car=0, ab=0, ing=0, egr=0, cie=0, log=0)
        self.seq_cli = {}

    def nid(self, k, pref, width):
        self.n[k] += 1
        return f"{pref}-{self.n[k]:0{width}d}"

    def log(self, fecha, usuario, modulo, accion, doc, detalle):
        self.bitacora.append([self.nid("log", "LOG", 6), dt.datetime.combine(fecha, dt.time(9, 0)) +
                              dt.timedelta(minutes=7 * self.n["log"] % 600), usuario, modulo, accion, doc, detalle,
                              "PC-OFICINA"])

    def kdx(self, fecha, usuario, tipo, sku, ent, sal, costo, doc, motivo, aut=""):
        self.kardex.append([self.nid("k", "K", 6), fecha, "", usuario, tipo, sku,
                            f'{SKU[sku]["prod"]} {SKU[sku]["talla"]} {SKU[sku]["color"]}',
                            ent or None, sal or None, round(costo), doc, motivo, aut])

    def ingreso(self, fecha, tipo, metodo, valor, doc, tercero, concepto, usuario):
        self.ingresos.append([self.nid("ing", "ING", 5), fecha, tipo, metodo, valor, doc, tercero, concepto,
                              usuario, "VIGENTE"])

    def egreso(self, fecha, concepto, metodo, valor, doc, benef, desc, usuario):
        self.egresos.append([self.nid("egr", "EGR", 5), fecha, concepto, metodo, valor, doc, benef, desc,
                             usuario, "VIGENTE"])

    # ---- compras
    def compra(self, fecha, prov, factprov, lineas, transp, adic, forma_estado, metodo, pagado=None):
        num = self.nid("oc", "OC", 5)
        sub = sum(q * c for _, q, c in lineas)
        total = sub + transp + adic
        estado_pago = forma_estado
        if estado_pago == "PAGADA":
            pagado = total
        elif estado_pago == "CRÉDITO":
            pagado = 0
        saldo = total - pagado
        pnom = [p for p in PROVEEDORES if p[0] == prov][0][1]
        self.compras.append([num, fecha, prov, pnom, factprov, sub, transp, adic, total, metodo, estado_pago,
                             pagado, saldo, "VIGENTE", "admin", dt.datetime.combine(fecha, dt.time(10, 0)), ""])
        for sku, q, c in lineas:
            linea_sub = q * c
            prorr = round((transp + adic) * linea_sub / sub / q, 2)
            real = c + prorr
            self.compra_det.append([num, fecha, prov, sku, SKU[sku]["prod"], q, c, linea_sub, prorr, real, "VIGENTE"])
            st = self.stock[sku]
            self.costo[sku] = round((st * self.costo[sku] + q * real) / (st + q), 0) if st > 0 else round(real, 0)
            self.stock[sku] += q
            self.kdx(fecha, "admin", "COMPRA", sku, q, 0, real, num, f"Compra a {prov}")
        if pagado > 0:
            self.egreso(fecha, "Compras", metodo, pagado, num, prov, f"Pago compra {num}", "admin")
        if saldo > 0:
            self.cartera.append([self.nid("car", "CAR", 5), "CxP", num, prov, pnom, fecha, saldo,
                                 fecha + dt.timedelta(days=30), "NO"])
        self.log(fecha, "admin", "COMPRAS", "GUARDAR", num, f"Total {total:,.0f}")
        return num

    # ---- ventas
    def factura(self, fecha, cli, canal, lineas, estado_pago, metodo, abono_pct=0.5):
        num = self.nid("fac", "FV", 6)
        vend = VEND_POR_CANAL[canal]
        usuario = VEN[vend][3]
        bruto = desc_tot = total = costo_t = 0
        for i, (sku, q, d) in enumerate(lineas, 1):
            precio = SKU[sku]["precio"]
            sub = round(q * precio * (1 - d))
            cu = self.costo[sku]
            bruto += q * precio
            total += sub
            costo_t += cu * q
            self.detalle.append([num, fecha, canal, vend, cli, sku, SKU[sku]["prod"], SKU[sku]["cat"],
                                 SKU[sku]["talla"], SKU[sku]["color"], q, precio, d, sub, cu, cu * q,
                                 sub - cu * q, "VIGENTE", f"{num}|{i}"])
            self.stock[sku] -= q
            self.kdx(fecha, usuario, "VENTA", sku, 0, q, cu, num, f"Venta {canal}")
        if estado_pago == "PAGADA":
            abono = total
        elif estado_pago == "CRÉDITO":
            abono = 0
        else:
            abono = round(total * abono_pct, -3)
        saldo = total - abono
        self.seq_cli[cli] = self.seq_cli.get(cli, 0) + 1
        c = CLI[cli]
        self.facturas.append([num, fecha, cli, c[1], c[3], vend, canal, metodo, estado_pago, bruto,
                              bruto - total, total, abono, saldo, costo_t, total - costo_t, "VIGENTE", usuario,
                              dt.datetime.combine(fecha, dt.time(11, 0)), f"FACTURAS/{fecha:%Y}/{fecha:%m}/{num}.pdf",
                              "", f"SIM-{num}", f"{cli}|{self.seq_cli[cli]}", None])
        if abono > 0:
            self.ingreso(fecha, "Ventas", metodo, abono, num, cli, f"Venta {num}", usuario)
        if saldo > 0:
            self.cartera.append([self.nid("car", "CAR", 5), "CxC", num, cli, c[1], fecha, saldo,
                                 fecha + dt.timedelta(days=15), "NO"])
        self.log(fecha, usuario, "FACTURACIÓN", "GUARDAR", num, f"{canal} total {total:,.0f}")
        return num

    def fac(self, num):
        return [f for f in self.facturas if f[0] == num][0]

    def anular(self, num, fecha, motivo, devolucion=False):
        f = self.fac(num)
        f[16] = "DEVUELTA" if devolucion else "ANULADA"
        f[20] = motivo
        f[23] = fecha
        for d in self.detalle:
            if d[0] == num:
                d[17] = f[16]
                self.stock[d[5]] += d[10]
                self.kdx(fecha, "admin", "DEVOLUCIÓN" if devolucion else "ANULACIÓN VENTA", d[5], d[10], 0,
                         d[14], num, motivo, "admin")
        pagado = sum(i[4] for i in self.ingresos if i[5] == num and i[9] == "VIGENTE")
        pagado += sum(a[6] for a in self.abonos if a[4] == num)
        for c in self.cartera:
            if c[2] == num:
                c[8] = "SI"
        if devolucion:
            if pagado > 0:
                self.egreso(fecha, "Devoluciones a clientes", f[7], pagado, num, f[2], f"Devolución {num}", "admin")
        else:
            for i in self.ingresos:
                if i[5] == num:
                    i[9] = "ANULADO"
        self.log(fecha, "admin", "FACTURACIÓN", "DEVOLUCIÓN" if devolucion else "ANULAR", num, motivo)

    def abono(self, fecha, doc, valor, metodo, usuario="aux"):
        car = [c for c in self.cartera if c[2] == doc][0]
        aid = self.nid("ab", "AB", 5)
        self.abonos.append([aid, fecha, car[0], car[1], doc, car[3], valor, metodo, usuario, "Abono",
                            dt.datetime.combine(fecha, dt.time(15, 0))])
        if car[1] == "CxC":
            self.ingreso(fecha, "Abonos", metodo, valor, aid, car[3], f"Abono a {doc}", usuario)
        else:
            self.egreso(fecha, "Compras", metodo, valor, aid, car[3], f"Abono a {doc}", usuario)
        self.log(fecha, usuario, "CARTERA", "ABONO", aid, f"{doc} {valor:,.0f}")

    def saldo_cartera(self, doc):
        car = [c for c in self.cartera if c[2] == doc]
        if not car:
            return 0
        car = car[0]
        return car[6] - sum(a[6] for a in self.abonos if a[2] == car[0])


def simular():
    s = Sim()
    # compras iniciales
    s.compra(D(2026, 10, 1), "PRV-001", "TA-5521",
             [(k, 30, 18000) for k in SKU if k.startswith("CAMB")] +
             [(k, 18, 45000) for k in SKU if k.startswith("JEAN")] +
             [(k, 10, 38000) for k in SKU if k.startswith("BUZO")], 180000, 40000, "PAGADA", "Transferencia")
    s.compra(D(2026, 10, 3), "PRV-002", "CDV-0912",
             [(k, 12, 95000) for k in SKU if k.startswith("TRUN")] +
             [(k, 12, 80000) for k in SKU if k.startswith("TURB")], 240000, 0, "ABONO PARCIAL",
             "Transferencia", pagado=3000000)
    s.compra(D(2026, 10, 4), "PRV-003", "IAB-334",
             [(k, 30, 12000) for k in SKU if k.startswith("GORR")] +
             [("MORR-U-NEG", 12, 40000), ("MEDI-U-BLA", 60, 7000)], 60000, 0, "PAGADA", "Nequi")

    # ventas
    fechas = sorted([D(2026, 10, 2) + dt.timedelta(days=random.randint(0, 54)) for _ in range(82)])
    fechas += [HOY, HOY, HOY, HOY]
    clientes_por_fecha = [c for c in CLIENTES]
    hecho_compra2 = False
    abono_programado = []
    for idx, f in enumerate(fechas):
        if not hecho_compra2 and f >= D(2026, 11, 10):
            s.compra(D(2026, 11, 10), "PRV-001", "TA-5688",
                     [("CAMB-M-NEG", 20, 19000), ("CAMB-L-NEG", 20, 19000), ("JEAN-32-AZU", 10, 46000),
                      ("BUZO-M-GRI", 8, 38000), ("BUZO-L-GRI", 8, 38000)],
                     80000, 0, "CRÉDITO", "Transferencia")
            hecho_compra2 = True
        posibles = [c for c in clientes_por_fecha if c[9] <= f and c[0] != "CLI-0010"]
        cli = random.choice(posibles)
        canal = cli[10] if random.random() < 0.7 else random.choice(CANALES)
        lineas, usados = [], set()
        for _ in range(random.choice([1, 1, 2, 2, 3])):
            disp = [k for k in SKU if s.stock[k] - sum(q for kk, q, _ in lineas if kk == k) > 0 and k not in usados]
            if not disp:
                break
            k = random.choice(disp)
            usados.add(k)
            q = 1 if s.stock[k] < 3 or random.random() < 0.7 else 2
            lineas.append((k, q, random.choice([0, 0, 0, 0.05, 0.1])))
        r = random.random()
        ep = "PAGADA" if r < 0.6 else ("CRÉDITO" if r < 0.82 else "ABONO PARCIAL")
        if cli[0] == "CLI-0004":
            ep = "CRÉDITO"
        metodo = random.choice(["Transferencia", "Nequi", "Daviplata"]) if canal != "Venta directa" else \
            random.choice(["Efectivo", "Efectivo", "Tarjeta"])
        num = s.factura(f, cli[0], canal, lineas, ep, metodo)
        if ep != "PAGADA":
            abono_programado.append((num, f, metodo))

    # anulaciones y devolución
    s.anular("FV-000006", s.fac("FV-000006")[1], "Error en talla al facturar (se re-factura)")
    s.anular("FV-000021", s.fac("FV-000021")[1], "Cliente canceló el pedido en el LIVE")
    f13 = s.fac("FV-000013")
    s.anular("FV-000013", f13[1] + dt.timedelta(days=4), "Devolución: producto no le quedó", devolucion=True)

    # abonos de clientes
    for num, f, metodo in abono_programado:
        fac = s.fac(num)
        if fac[16] != "VIGENTE":
            continue
        fa = f + dt.timedelta(days=random.randint(5, 12))
        if fa >= D(2026, 11, 18):
            continue
        saldo = s.saldo_cartera(num)
        valor = saldo if random.random() < 0.6 else round(saldo / 2, -3)
        s.abono(fa, num, valor, metodo)
    # abono a proveedor
    s.abono(D(2026, 11, 5), "OC-00002", 2000000, "Transferencia", "admin")

    # ajuste autorizado
    s.stock["GORR-U-ROJ"] -= 1
    s.kdx(D(2026, 11, 3), "aux", "AJUSTE SALIDA", "GORR-U-ROJ", 0, 1, s.costo["GORR-U-ROJ"], "AJ-0001",
          "Unidad averiada en bodega", "admin")
    s.log(D(2026, 11, 3), "admin", "INVENTARIO", "AJUSTE", "AJ-0001", "Autoriza salida por avería")

    # gastos fijos generados
    gf_ultimo = {}
    for mes in (10, 11):
        for g in GASTOS_FIJOS:
            fecha = D(2026, mes, min(g[4], 30))
            if fecha > HOY:
                continue
            s.egreso(fecha, g[1], g[5], g[3], f"{g[0]}-2026{mes:02d}", g[6], g[2], "admin")
            gf_ultimo[g[0]] = D(2026, mes, 1)
    # gastos variables
    for fecha, con, met, val, ben, desc in [
        (D(2026, 10, 6), "Empaques", "Efectivo", 150000, "Empaques La 14", "Bolsas y cajas"),
        (D(2026, 10, 15), "Publicidad", "Transferencia", 400000, "Meta Ads", "Campaña Instagram/Facebook"),
        (D(2026, 10, 20), "Publicidad", "Nequi", 300000, "TikTok Ads", "Promoción LIVE"),
        (D(2026, 10, 31), "Transporte", "Efectivo", 90000, "Mensajería local", "Domicilios octubre"),
        (D(2026, 11, 8), "Empaques", "Nequi", 120000, "Empaques La 14", "Bolsas con logo"),
        (D(2026, 11, 12), "Publicidad", "Transferencia", 450000, "Meta Ads", "Campaña Black Friday"),
        (D(2026, 11, 20), "Transporte", "Efectivo", 70000, "Mensajería local", "Domicilios noviembre"),
        (D(2026, 11, 20), "Impuestos", "Transferencia", 85000, "Secretaría de Hacienda", "ICA bimestral"),
    ]:
        s.egreso(fecha, con, met, val, "", ben, desc, "aux")
    s.ingreso(D(2026, 11, 14), "Otros ingresos", "Efectivo", 30000, "", "Servientrega",
              "Reintegro por envío fallido", "aux")

    # comisiones de octubre pagadas
    for v in VENDEDORES:
        ventas = sum(f[11] for f in s.facturas if f[5] == v[0] and f[16] == "VIGENTE" and f[1].month == 10)
        com = round(ventas * v[7])
        if com > 0:
            s.egreso(D(2026, 11, 5), "Comisiones", "Transferencia", com, f"COM-202610-{v[0]}", v[0],
                     f"Comisión octubre {v[1]}", "admin")

    # guías
    transp = ["Servientrega", "Interrapidísimo", "Coordinadora", "Envía"]
    for f in s.facturas:
        if f[6] == "Venta directa" or f[16] != "VIGENTE":
            continue
        if len(s.guias) >= 14:
            break
        gid = s.nid("gu", "GU", 5)
        c = CLI[f[2]]
        prods = [d for d in s.detalle if d[0] == f[0]]
        dias = (HOY - f[1]).days
        estado = "ENTREGADO" if dias > 6 else ("EN TRÁNSITO" if dias > 2 else "DESPACHADO")
        s.guias.append([gid, f[1], f[0], f[2], c[1], c[4], c[7], c[8], random.choice(transp),
                        str(random.randint(700000000, 799999999)),
                        "; ".join(f"{d[5]} x{d[10]}" for d in prods), sum(d[10] for d in prods), f[11],
                        estado, f[17], dt.datetime.combine(f[1], dt.time(16, 0)),
                        f"GUIAS/{f[1]:%Y}/{f[1]:%m}/{gid}.pdf"])
        s.log(f[1], f[17], "GUÍAS", "GUARDAR", gid, f[0])

    # cierre mensual de octubre
    for m in METODOS:
        ing = sum(i[4] for i in s.ingresos if i[3] == m and i[9] == "VIGENTE" and i[1] <= D(2026, 10, 31))
        egr = sum(e[4] for e in s.egresos if e[3] == m and e[9] == "VIGENTE" and e[1] <= D(2026, 10, 31))
        sist = SALDOS_INICIALES[m] + ing - egr
        cont = sist - (5000 if m == "Efectivo" else 0)
        s.cierres.append([s.nid("cie", "CIE", 4), "MENSUAL", D(2026, 10, 31), D(2026, 10, 1), D(2026, 10, 31), m,
                          sist, cont, cont - sist, "admin",
                          "Faltante menor en caja" if cont != sist else ""])
    s.log(D(2026, 10, 31), "admin", "CAJA Y BANCOS", "CIERRE", "CIE-MENSUAL", "Cierre octubre")
    s.gf_ultimo = gf_ultimo
    s.bitacora.sort(key=lambda r: r[1])
    for i, r in enumerate(s.bitacora, 1):
        r[0] = f"LOG-{i:06d}"
    s.kardex.sort(key=lambda r: (r[1], r[0]))
    for i, r in enumerate(s.kardex, 1):
        r[0] = f"K-{i:06d}"
    return s


if __name__ == "__main__":
    s = simular()
    print(len(s.facturas), "facturas", len(s.kardex), "movs", "stock negativo:",
          [k for k, v in s.stock.items() if v < 0])
    print({k: v for k, v in s.stock.items()})
