"""
Generador de datos de demostración coherentes.

Simula exactamente las mismas reglas que aplican las macros VBA
(compras con costo promedio ponderado, ventas con descuento de stock,
cartera, abonos, anulaciones, egresos) para que los datos de ejemplo
sean idénticos a los que produciría el sistema operado por un usuario.

Uso CLI:
    python demo.py            -> imprime un resumen JSON de la simulación
"""
import datetime as dt
import json
import random

from estructura import CFG_LISTS

HOY = dt.date(2026, 11, 26)
METODOS = CFG_LISTS["lstMetodos"][2]
MET_CTA = dict(zip(METODOS, CFG_LISTS["lstMetodoCuenta"][2]))

PROVEEDORES = [
    # codigo, nombre, nit, tel, correo, dir, ciudad, contacto, productos
    ("PR-001", "Textiles del Valle S.A.S.", "900.111.222-3", "602 345 6789", "ventas@textilesvalle.co", "Cra 5 # 20-15", "Cali", "Andrea Gómez", "Camisetas, camisas, vestidos"),
    ("PR-002", "Calzado Andino Ltda.", "900.333.444-5", "604 222 3344", "pedidos@calzadoandino.co", "Cl 30 # 65-20", "Bucaramanga", "Jorge Rueda", "Calzado deportivo y formal, sandalias"),
    ("PR-003", "Confecciones Medellín", "901.555.666-7", "604 444 5566", "info@confmedellin.co", "Cra 52 # 45-10", "Medellín", "Luisa Restrepo", "Jeans, pantalones, chaquetas"),
    ("PR-004", "Accesorios Bogotá", "830.777.888-9", "601 666 7788", "contacto@accbogota.co", "Cl 13 # 8-40", "Bogotá D.C.", "Camilo Pardo", "Bolsos, cinturones, gorras"),
    ("PR-005", "Importadora Moda Global", "900.999.000-1", "605 888 9900", "import@modaglobal.co", "Vía 40 # 70-100", "Barranquilla", "Sofía Char", "Chaquetas y accesorios importados"),
]
P = {p[0]: p[1] for p in PROVEEDORES}

PRODUCTOS = [
    # sku, nombre, cat, desc, talla, color, prov, costo, precio, stockini, min, empaque, comision
    ("CAM-BAS-NEG-M", "Camiseta básica algodón", "Camisetas", "Cuello redondo 100% algodón", "M", "Negro", "PR-001", 18000, 39900, 30, 8),
    ("CAM-BAS-NEG-L", "Camiseta básica algodón", "Camisetas", "Cuello redondo 100% algodón", "L", "Negro", "PR-001", 18000, 39900, 25, 8),
    ("CAM-BAS-BLA-M", "Camiseta básica algodón", "Camisetas", "Cuello redondo 100% algodón", "M", "Blanco", "PR-001", 18000, 39900, 28, 8),
    ("CAM-POL-AZU-L", "Camiseta polo piqué", "Camisetas", "Polo manga corta bordada", "L", "Azul oscuro", "PR-001", 26000, 59900, 18, 5),
    ("CAS-OXF-BLA-M", "Camisa oxford", "Camisas", "Manga larga, slim fit", "M", "Blanco", "PR-001", 38000, 89900, 15, 4),
    ("CAS-LIN-BEI-L", "Camisa de lino", "Camisas", "Manga larga, fresca", "L", "Beige", "PR-001", 45000, 109900, 10, 3),
    ("JEA-SLI-AZU-32", "Jean slim fit", "Jeans", "Denim stretch", "32", "Azul", "PR-003", 52000, 129900, 20, 5),
    ("JEA-SLI-AZU-34", "Jean slim fit", "Jeans", "Denim stretch", "34", "Azul", "PR-003", 52000, 129900, 18, 5),
    ("JEA-MOM-AZU-30", "Jean mom fit dama", "Jeans", "Tiro alto", "30", "Azul", "PR-003", 55000, 139900, 14, 4),
    ("PAN-CHI-BEI-32", "Pantalón chino", "Pantalones", "Gabardina stretch", "32", "Beige", "PR-003", 48000, 115000, 12, 4),
    ("VES-FLO-ROS-S", "Vestido floral", "Vestidos", "Midi en viscosa", "S", "Rosado", "PR-001", 50000, 139900, 10, 3),
    ("VES-NEG-NEG-M", "Vestido negro casual", "Vestidos", "Corte recto", "M", "Negro", "PR-001", 47000, 119900, 8, 3),
    ("CHA-JEA-AZU-L", "Chaqueta de jean", "Chaquetas", "Denim clásico", "L", "Azul", "PR-003", 70000, 179900, 8, 3),
    ("CHA-ROM-NEG-M", "Chaqueta rompevientos", "Chaquetas", "Impermeable liviana", "M", "Negro", "PR-005", 65000, 159900, 6, 2),
    ("TEN-URB-BLA-40", "Tenis urbanos", "Calzado deportivo", "Suela en caucho", "40", "Blanco", "PR-002", 85000, 199900, 12, 4),
    ("TEN-URB-BLA-42", "Tenis urbanos", "Calzado deportivo", "Suela en caucho", "42", "Blanco", "PR-002", 85000, 199900, 10, 4),
    ("TEN-RUN-NEG-41", "Tenis running", "Calzado deportivo", "Amortiguación media", "41", "Negro", "PR-002", 95000, 229900, 8, 3),
    ("ZAP-FOR-CAF-41", "Zapato formal cuero", "Calzado formal", "Cuero genuino", "41", "Café", "PR-002", 110000, 249900, 6, 2),
    ("SAN-PLA-BEI-37", "Sandalia plataforma", "Sandalias", "Dama, correa ajustable", "37", "Beige", "PR-002", 40000, 99900, 10, 3),
    ("BOL-CUE-NEG-U", "Bolso de cuero", "Bolsos", "Bolso de mano mediano", "ÚNICA", "Negro", "PR-004", 75000, 189900, 6, 2),
    ("BOL-MOC-CAF-U", "Morral casual", "Bolsos", "Lona y cuero sintético", "ÚNICA", "Café", "PR-004", 48000, 119900, 7, 2),
    ("CIN-CUE-CAF-U", "Cinturón de cuero", "Cinturones", "Hebilla metálica", "ÚNICA", "Café", "PR-004", 22000, 59900, 15, 4),
    ("GOR-BAS-NEG-U", "Gorra béisbol", "Gorras", "Ajustable, bordada", "ÚNICA", "Negro", "PR-004", 15000, 39900, 20, 5),
    ("ACC-GAF-NEG-U", "Gafas de sol", "Accesorios", "Protección UV400", "ÚNICA", "Negro", "PR-005", 25000, 79900, 3, 3),
]

CLIENTES = [
    ("CL-0001", "CONSUMIDOR FINAL", "222222222222", "", "", "", "Mostrador", "Bogotá D.C.", "Bogotá D.C."),
    ("CL-0002", "María Fernanda López", "1.020.345.678", "310 555 1122", "310 555 1122", "mafe.lopez@gmail.com", "Cra 15 # 85-40 Apto 302", "Bogotá D.C.", "Bogotá D.C."),
    ("CL-0003", "Juan Camilo Restrepo", "1.036.789.012", "311 444 2233", "311 444 2233", "jcrestrepo@hotmail.com", "Cl 10 # 43-25", "Medellín", "Antioquia"),
    ("CL-0004", "Laura Valentina Gómez", "1.144.567.890", "312 333 4455", "312 333 4455", "lauravg@gmail.com", "Av 6N # 23-50", "Cali", "Valle del Cauca"),
    ("CL-0005", "Andrés Felipe Martínez", "80.123.456", "313 222 5566", "313 222 5566", "afmartinez@outlook.com", "Cl 72 # 54-10", "Barranquilla", "Atlántico"),
    ("CL-0006", "Boutique La Elegancia", "900.456.123-4", "601 345 2211", "315 777 8899", "compras@laelegancia.co", "Cl 19 # 7-48 Local 5", "Bogotá D.C.", "Bogotá D.C."),
    ("CL-0007", "Daniela Ruiz Ospina", "1.053.222.111", "314 111 6677", "314 111 6677", "daniruiz@gmail.com", "Cra 23 # 64-12", "Manizales", "Caldas"),
    ("CL-0008", "Carlos Andrés Peña", "1.098.765.432", "316 999 0011", "316 999 0011", "carlospena@gmail.com", "Cl 36 # 27-80", "Bucaramanga", "Santander"),
    ("CL-0009", "Tienda Urbana Express", "901.789.456-1", "604 555 3322", "318 222 4466", "pedidos@urbanaexpress.co", "Cra 70 # 44-21", "Medellín", "Antioquia"),
    ("CL-0010", "Sofía Herrera Castro", "1.010.111.222", "317 888 5544", "317 888 5544", "sofiaherrera@yahoo.com", "Cl 140 # 11-35", "Bogotá D.C.", "Bogotá D.C."),
    ("CL-0011", "Natalia Cárdenas", "1.075.654.321", "319 123 7788", "319 123 7788", "nati.cardenas@gmail.com", "Cra 4 # 12-60", "Neiva", "Huila"),
    ("CL-0012", "Felipe Arango Mejía", "1.017.333.444", "320 456 9900", "320 456 9900", "farango@gmail.com", "Cl 50 # 80-15", "Medellín", "Antioquia"),
]

GASTOS_FIJOS = [
    # concepto, categoria, valor, dia, responsable, metodo, activo
    ("Arriendo local comercial", "Arriendo", 1100000, 5, "Administrador", "Transferencia", "SI"),
    ("Energía y acueducto", "Servicios públicos", 210000, 15, "Administrador", "Transferencia", "SI"),
    ("Internet y telefonía", "Internet", 95000, 10, "Administrador", "Nequi", "SI"),
    ("Auxiliar de ventas (medio tiempo)", "Nómina", 750000, 30, "Administrador", "Transferencia", "SI"),
    ("Software y dominio web", "Otros gastos", 60000, 20, "Administrador", "Tarjeta", "SI"),
]

SALDOS_INI = {"EFECTIVO": 1500000, "BANCOS": 9000000, "NEQUI": 600000,
              "DAVIPLATA": 250000, "OTROS": 0}
FECHA_SALDOS = dt.date(2026, 6, 1)


class Sim:
    """Motor que replica la lógica de negocio de las macros."""

    def __init__(self):
        self.inv = {}
        for p in PRODUCTOS:
            self.inv[p[0]] = {"row": p, "costo": float(p[7]), "precio": p[8],
                              "ini": p[9] * (1 if p[0] == "ACC-GAF-NEG-U" else (4 if p[7] >= 65000 else 4)), "ent": 0, "sal": 0,
                              "transp": 0.0, "qcomp": 0}
        self.cli = {c[0]: c for c in CLIENTES}
        self.bd, self.hfa, self.mov, self.hgu = [], [], [], []
        self.aud, self.com, self.ing, self.egr, self.car = [], [], [], [], []
        self.traslados = []
        self.sig = {"FV": 1, "GE": 1, "OC": 1}
        self.usuario = "Administrador"

    # -------------------------------------------------- utilidades
    def stock(self, sku):
        i = self.inv[sku]
        return i["ini"] + i["ent"] - i["sal"]

    def costo_real(self, sku):
        i = self.inv[sku]
        tu = i["transp"] / i["qcomp"] if i["qcomp"] else 0
        return i["costo"] + tu

    def ts(self, fecha, h=10, m=0):
        return dt.datetime(fecha.year, fecha.month, fecha.day, h, m)

    def mov_add(self, fecha, tipo, doc, sku, ent, sal, motivo, costo=None, h=10):
        i = self.inv[sku]
        i["ent"] += ent
        i["sal"] += sal
        cu = self.costo_real(sku) if costo is None else costo
        n = len(self.mov) + 1
        p = i["row"]
        self.mov.append([f"MOV-{n:06d}", self.ts(fecha, h, n % 60), tipo, doc, sku,
                         f"{p[1]} {p[4]} {p[5]}", ent, sal, self.stock(sku),
                         round(cu, 2), round((ent - sal) * cu, 0), self.usuario, motivo])

    def ing_add(self, fecha, origen, tipo, metodo, doc, tercero, concepto, valor, obs=""):
        n = len(self.ing) + 1
        cta = MET_CTA.get(metodo, "") if tipo != "NOTA CRÉDITO" else ""
        clas = f"{tipo} - {metodo}" if metodo else tipo
        self.ing.append([f"ING-{n:05d}", fecha, origen, tipo, clas, metodo, cta,
                         doc, tercero, concepto, valor, self.usuario,
                         self.ts(fecha, 11), "ACTIVO", obs])

    def egr_add(self, fecha, concepto, cat, tipo, valor, metodo, resp, comp, obs,
                origen, doc):
        n = len(self.egr) + 1
        self.egr.append([f"EGR-{n:05d}", fecha, concepto, cat, tipo, valor, metodo,
                         MET_CTA.get(metodo, "OTROS"), resp, comp, obs, origen, doc,
                         self.usuario, self.ts(fecha, 9), "ACTIVO"])

    def car_add(self, tipo, codter, tercero, doc, fecha, valor, pagoini, vence, obs=""):
        pref = "CXC" if tipo == "POR COBRAR" else "CXP"
        n = sum(1 for c in self.car if c[1] == tipo) + 1
        self.car.append([f"{pref}-{n:04d}", tipo, codter, tercero, doc, fecha,
                         valor, pagoini, vence, "ACTIVA", obs])

    # -------------------------------------------------- operaciones
    def compra(self, fecha, prov, lineas, transp=0, otros=0, forma="Transferencia",
               estado="PAGADA", pagado=None, dias=30):
        num = f"OC-{self.sig['OC']:04d}"
        self.sig["OC"] += 1
        sub = sum(q * c for _, q, c in lineas)
        total = sub + transp + otros
        if estado == "PAGADA":
            pagado = total
        elif estado == "PENDIENTE":
            pagado = 0
        for sku, q, c in lineas:
            i = self.inv[sku]
            share = (q * c / sub) if sub else 1 / len(lineas)
            t_l, o_l = round(transp * share, 0), round(otros * share, 0)
            st = self.stock(sku)
            i["costo"] = round((st * i["costo"] + q * c) / (st + q), 2) if st > 0 else float(c)
            i["transp"] += t_l + o_l
            i["qcomp"] += q
            real = q * c + t_l + o_l
            p = i["row"]
            self.com.append([num, fecha, P[prov], sku, f"{p[1]} {p[4]} {p[5]}", q, c,
                             q * c, t_l, o_l, real, round(real / q, 2), forma, estado,
                             self.usuario, self.ts(fecha, 8), "ACTIVA", ""])
            self.mov_add(fecha, "COMPRA", num, sku, q, 0, f"Compra a {P[prov]}", costo=c, h=8)
        if pagado > 0:
            self.egr_add(fecha, f"Compra {num} - {P[prov]}", "Compras", "VARIABLE",
                         pagado, forma, self.usuario, num, "", "COMPRAS", num)
        if total - pagado > 0:
            self.car_add("POR PAGAR", prov, P[prov], num, fecha, total, pagado,
                         fecha + dt.timedelta(days=dias))
        return num

    def venta(self, fecha, cod, lineas, metodo="Efectivo", estado="PAGADA",
              pagado=None, obs="", entrega=""):
        lineas = [(s, q, d) for s, q, d in lineas if self.stock(s) >= q and q > 0]
        if not lineas:
            return None
        num = f"FV-{self.sig['FV']:04d}"
        self.sig["FV"] += 1
        c = self.cli[cod]
        bruto = desc = total = costo = 0
        for sku, q, d in lineas:
            i = self.inv[sku]
            p = i["row"]
            pu = i["precio"]
            sub = round(q * pu * (1 - d), 0)
            cu = round(self.costo_real(sku), 2)
            bruto += q * pu
            total += sub
            costo += round(q * cu, 0)
            self.bd.append([num, fecha, cod, c[1], c[2], sku, p[1], p[2], p[4], p[5],
                            q, pu, d, q * pu - sub, sub, cu, round(q * cu, 0),
                            sub - round(q * cu, 0), metodo, estado, "ACTIVA",
                            self.usuario, self.ts(fecha, 12)])
            self.mov_add(fecha, "VENTA", num, sku, 0, q, f"Venta a {c[1]}", h=12)
        desc = bruto - total
        if estado == "PAGADA":
            pagado = total
        elif estado == "PENDIENTE":
            pagado = 0
        vence = fecha if estado == "PAGADA" else fecha + dt.timedelta(days=30)
        self.hfa.append([num, fecha, cod, c[1], c[2], c[3], entrega or c[6], c[7], c[8],
                         bruto, desc, total, metodo, estado, pagado, vence, "ACTIVA",
                         obs, costo, total - costo, self.usuario, self.ts(fecha, 12),
                         "", "", "", ""])
        if pagado > 0:
            self.ing_add(fecha, "FACTURACIÓN", "VENTA", metodo, num, c[1],
                         f"Venta factura {num}", pagado)
        if total - pagado > 0:
            self.car_add("POR COBRAR", cod, c[1], num, fecha, total, pagado, vence)
        return num

    def abono_cxc(self, fecha, doc, valor, metodo):
        for c in self.car:
            if c[4] == doc:
                self.ing_add(fecha, "CARTERA", "ABONO", metodo, doc, c[3],
                             f"Abono a {doc}", valor)
                return

    def abono_cxp(self, fecha, doc, valor, metodo):
        for c in self.car:
            if c[4] == doc:
                self.egr_add(fecha, f"Abono a {doc} - {c[3]}", "Compras", "VARIABLE",
                             valor, metodo, self.usuario, "", "", "CARTERA", doc)
                return

    def anular(self, fecha, num, motivo):
        for r in self.bd:
            if r[0] == num and r[20] == "ACTIVA":
                r[20] = "ANULADA"
                self.mov_add(fecha, "ANULACIÓN VENTA", num, r[5], r[10], 0,
                             f"Anulación {num}: {motivo}", costo=r[15], h=16)
        for h in self.hfa:
            if h[0] == num:
                h[16] = "ANULADA"
                h[22], h[23], h[24] = self.usuario, self.ts(fecha, 16), motivo
        for i in self.ing:
            if i[7] == num:
                i[13] = "ANULADO"
        for c in self.car:
            if c[4] == num:
                c[9] = "ANULADA"
        self.aud.append([self.ts(fecha, 16), self.usuario, "FACTURACIÓN",
                         "ANULAR FACTURA", num, motivo])

    def guia(self, fecha, num_fac, transp, estado, seguimiento):
        h = next(x for x in self.hfa if x[0] == num_fac)
        num = f"GE-{self.sig['GE']:04d}"
        self.sig["GE"] += 1
        items = [r for r in self.bd if r[0] == num_fac]
        txt = "; ".join(f"{r[10]} x {r[6]} {r[8]} {r[9]}" for r in items)
        ser = ";".join(f"{r[5]}|{r[10]}" for r in items)
        uni = sum(r[10] for r in items)
        self.hgu.append([num, fecha, num_fac, h[3], h[4], h[5], h[7], h[8], h[6],
                         transp, seguimiento, txt, uni, h[11], estado, "",
                         self.usuario, self.ts(fecha, 15), "", ser, 1])


def generar(seed=7):
    """Ejecuta 6 meses de operación simulada y devuelve el objeto Sim."""
    rnd = random.Random(seed)
    s = Sim()
    skus = [p[0] for p in PRODUCTOS]
    # compras de reposición
    plan_compras = [
        (dt.date(2026, 6, 3), "PR-001", [("CAM-BAS-NEG-M", 20, 17500), ("CAM-BAS-BLA-M", 20, 17500), ("CAM-POL-AZU-L", 10, 26000)], 60000, 0, "Transferencia", "PAGADA"),
        (dt.date(2026, 6, 20), "PR-003", [("JEA-SLI-AZU-32", 12, 51000), ("JEA-SLI-AZU-34", 12, 51000), ("PAN-CHI-BEI-32", 8, 48000)], 85000, 15000, "Transferencia", "PENDIENTE"),
        (dt.date(2026, 7, 15), "PR-002", [("TEN-URB-BLA-40", 8, 84000), ("TEN-URB-BLA-42", 8, 84000), ("SAN-PLA-BEI-37", 6, 40000)], 70000, 0, "Transferencia", "ABONO PARCIAL", 600000),
        (dt.date(2026, 8, 12), "PR-004", [("BOL-CUE-NEG-U", 5, 76000), ("CIN-CUE-CAF-U", 10, 22000), ("GOR-BAS-NEG-U", 15, 14500)], 40000, 0, "Efectivo", "PAGADA"),
        (dt.date(2026, 9, 18), "PR-001", [("CAM-BAS-NEG-L", 20, 18200), ("VES-FLO-ROS-S", 6, 50000), ("CAS-OXF-BLA-M", 8, 38500)], 55000, 0, "Nequi", "PAGADA"),
        (dt.date(2026, 10, 22), "PR-005", [("CHA-ROM-NEG-M", 6, 66000), ("CHA-JEA-AZU-L", 4, 70000)], 45000, 20000, "Transferencia", "PENDIENTE"),
    ]
    egresos_var = [
        ("Pauta Instagram y Facebook", "Publicidad", "VARIABLE", 250000, "Tarjeta"),
        ("Bolsas y cajas de empaque", "Empaques", "VARIABLE", 180000, "Efectivo"),
        ("Domicilios y envíos locales", "Transporte", "VARIABLE", 120000, "Nequi"),
    ]
    d = dt.date(2026, 6, 1)
    compras_idx = 0
    credito = []
    while d <= HOY:
        # compras programadas
        while compras_idx < len(plan_compras) and plan_compras[compras_idx][0] == d:
            pc = plan_compras[compras_idx]
            pag = pc[7] if len(pc) > 7 else None
            num = s.compra(pc[0], pc[1], pc[2], pc[3], pc[4], pc[5], pc[6], pag)
            if pc[6] == "PENDIENTE" and pc[0].month == 6:
                s.abono_cxp(dt.date(2026, 7, 20), num, 1000000, "Transferencia")
                s.abono_cxp(dt.date(2026, 8, 20), num, 1000000, "Transferencia")
            compras_idx += 1
        # gastos fijos mensuales (pagados) salvo el mes actual a partir del día 20
        for gf in GASTOS_FIJOS:
            if d.day == gf[3] or (gf[3] == 30 and d.day == 28):
                if not (d.month == 11 and gf[3] >= 20):
                    s.egr_add(d, f"{gf[0]} {d.month:02d}/{d.year}", gf[1], "FIJO", gf[2],
                              gf[5], gf[4], f"CP-{d.month:02d}{gf[3]:02d}", "",
                              "GASTO FIJO", "GF|" + gf[0])
        if d.day == 8:
            for e in egresos_var:
                s.egr_add(d, f"{e[0]} {d.month:02d}/{d.year}", e[1], e[2],
                          e[3] + rnd.randint(-4, 6) * 10000, e[4], "Administrador",
                          "", "", "MANUAL", "")
        if d == dt.date(2026, 9, 25):
            s.egr_add(d, "Reparación vitrina y luces", "Mantenimiento", "EXTRAORDINARIO",
                      420000, "Efectivo", "Administrador", "FAC-778", "", "MANUAL", "")
        if d == dt.date(2026, 7, 14):
            s.egr_add(d, "Impuesto de industria y comercio", "Impuestos",
                      "EXTRAORDINARIO", 310000, "Transferencia", "Contador", "ICA-2026",
                      "", "MANUAL", "")
        # ventas del día
        nv = rnd.choice([0, 0, 1, 1, 1, 2, 2])
        for _ in range(nv):
            cod = rnd.choice([c[0] for c in CLIENTES])
            lin = []
            for sku in rnd.sample(skus, rnd.choice([1, 1, 2, 2, 3])):
                lin.append((sku, rnd.choice([1, 1, 1, 2, 2, 3]),
                            rnd.choice([0, 0, 0, 0, 0.05, 0.1])))
            met = rnd.choice(["Efectivo", "Efectivo", "Transferencia", "Nequi",
                              "Nequi", "Daviplata", "Tarjeta"])
            est = "PAGADA"
            pag = None
            if cod != "CL-0001" and rnd.random() < 0.22:
                est = rnd.choice(["PENDIENTE", "ABONO PARCIAL"])
            if est == "ABONO PARCIAL":
                tot = sum(q * s.inv[k]["precio"] * (1 - x) for k, q, x in lin)
                pag = round(tot * 0.4 / 1000) * 1000
            num = s.venta(d, cod, lin, met, est, pag)
            if num and est != "PAGADA":
                credito.append((d, num))
        # abonos a créditos con más de 12 días
        for (fd, num) in list(credito):
            if (d - fd).days == 12:
                h = next(x for x in s.hfa if x[0] == num)
                saldo = h[11] - h[14]
                if rnd.random() < 0.7:
                    val = saldo if rnd.random() < 0.6 else round(saldo / 2000) * 1000
                    s.abono_cxc(d, num, val, rnd.choice(["Nequi", "Transferencia", "Efectivo"]))
                credito.remove((fd, num))
        if d == dt.date(2026, 10, 1):
            s.traslados.append([d, "EFECTIVO", "BANCOS", 2000000, "Consignación de efectivo", "Administrador"])
        if d == dt.date(2026, 11, 3):
            s.traslados.append([d, "NEQUI", "BANCOS", 800000, "Traslado Nequi a cuenta", "Administrador"])
        d += dt.timedelta(days=1)
    # anulación de ejemplo y guías
    victima = s.hfa[-6][0]
    s.anular(dt.date(2026, 11, 20), victima, "Error en la talla, cliente desistió")
    activas = [h for h in s.hfa if h[16] == "ACTIVA" and h[2] != "CL-0001"]
    transps = ["Servientrega", "Interrapidísimo", "Coordinadora", "Envía", "TCC"]
    for k, h in enumerate(activas[-6:]):
        est = "ENTREGADO" if k < 3 else ("EN TRÁNSITO" if k < 5 else "PREPARANDO")
        s.guia(h[1] + dt.timedelta(days=1), h[0], transps[k % 5], est,
               f"{7000000000 + k * 13579}")
    return s


if __name__ == "__main__":
    s = generar()
    print(json.dumps({"facturas": len(s.hfa), "lineas": len(s.bd), "compras": len(s.com),
                      "movimientos": len(s.mov), "ingresos": len(s.ing),
                      "egresos": len(s.egr), "cartera": len(s.car), "guias": len(s.hgu),
                      "stock_min": min(s.stock(k) for k in s.inv)}, indent=1))
