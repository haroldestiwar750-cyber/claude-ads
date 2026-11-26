"""
Ejecuta las macros VBA reales del sistema dentro de LibreOffice (modo de
compatibilidad VBA) sobre el libro base, para validar la lógica sin Excel.

Pasos: carga el .xlsx, inyecta todos los módulos .bas, ejecuta
modPruebas.EjecutarPruebas(conservar) y devuelve el informe en JSON.

Uso CLI:
    python ejecutar_vba_lo.py ../SistemaEmpresarial_base.xlsx [--conservar] [--guardar salida.ods]
"""
import argparse
import glob
import json
import os
import re

import uno
from lo_util import arrancar, abrir, prop

VBA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "vba")
ORDEN = ["modEstructura", "modUtil", "modInventario", "modFinanzas", "modFactura", "modGuia",
         "modCompras", "modReportes", "modSistema", "modPruebas"]


def cargar_modulos(doc):
    libs = doc.BasicLibraries
    if not libs.hasByName("Standard"):
        libs.createLibrary("Standard")
    lib = libs.getByName("Standard")
    try:
        libs.VBACompatibilityMode = True
    except Exception:
        pass
    for m in ORDEN:
        src = open(os.path.join(VBA_DIR, m + ".bas"), encoding="utf-8").read()
        src = re.sub(r"^Attribute VB_Name.*\n", "", src, flags=re.M)
        src = "Option VBASupport 1\n" + src.replace("\r\n", "\n")
        if lib.hasByName(m):
            lib.removeByName(m)
        lib.insertByName(m, src)


def correr(doc, modulo, proc, args=()):
    sp = doc.getScriptProvider()
    s = sp.getScript(f"vnd.sun.star.script:Standard.{modulo}.{proc}?language=Basic&location=document")
    return s.invoke(tuple(args), (), ())[0]


def leer_informe(doc):
    sh = doc.Sheets.getByName("REPORTES")
    out = []
    for r in range(13, 120):
        paso = sh.getCellByPosition(1, r).String
        if paso == "" and sh.getCellByPosition(2, r).String == "":
            break
        out.append([sh.getCellByPosition(c, r).String for c in range(1, 6)])
    return sh.getCellByPosition(1, 10).String, out


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("libro")
    ap.add_argument("--conservar", action="store_true")
    ap.add_argument("--guardar")
    a = ap.parse_args()
    proc, ctx, desk = arrancar()
    try:
        doc = abrir(desk, a.libro)
        doc.calculateAll()
        cargar_modulos(doc)
        fallos = correr(doc, "modPruebas", "EjecutarPruebas", (a.conservar,))
        resumen, filas = leer_informe(doc)
        log = correr(doc, "modPruebas", "LeerLog")
        print(json.dumps({"fallos": fallos, "resumen": resumen, "detalle": filas, "log": log.splitlines()[-25:]}, ensure_ascii=False, indent=1))
        if a.guardar:
            doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(a.guardar)), (prop("FilterName", "calc8"),))
        doc.close(True)
    finally:
        try:
            desk.terminate()
        except Exception:
            pass
        proc.wait(timeout=60)
