"""
Empaqueta el sistema para Windows.

Genera dist/SistemaEmpresarial/ con:
  * SistemaEmpresarial_base.xlsx (libro con datos de demostración)
  * vba/*.bas y vba/ThisWorkbook.txt en Windows-1252 + CRLF (formato del editor VBA)
  * INSTALAR.bat, instalar.ps1 (UTF-8 con BOM), MANUAL_DE_USO.pdf / .md
y un ZIP listo para copiar al computador con Windows.

Uso CLI:
    python empaquetar.py            -> JSON con la ruta del ZIP
"""
import glob
import json
import os
import shutil
import subprocess
import sys
import zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
DIST = os.path.join(RAIZ, "dist", "SistemaEmpresarial")


def main():
    subprocess.run([sys.executable, os.path.join(AQUI, "estructura.py"), "--vba",
                    os.path.join(RAIZ, "vba", "modEstructura.bas")], check=True, stdout=subprocess.DEVNULL)
    subprocess.run([sys.executable, os.path.join(AQUI, "construir_libro.py"), "--out",
                    os.path.join(RAIZ, "SistemaEmpresarial_base.xlsx")], check=True, stdout=subprocess.DEVNULL, cwd=AQUI)
    if os.path.exists(DIST):
        shutil.rmtree(DIST)
    os.makedirs(os.path.join(DIST, "vba"))
    shutil.copy(os.path.join(RAIZ, "SistemaEmpresarial_base.xlsx"), DIST)
    for f in glob.glob(os.path.join(RAIZ, "vba", "*.bas")) + [os.path.join(RAIZ, "vba", "ThisWorkbook.txt")]:
        txt = open(f, encoding="utf-8").read().replace("\r\n", "\n").replace("\n", "\r\n")
        with open(os.path.join(DIST, "vba", os.path.basename(f)), "w", encoding="cp1252", newline="") as o:
            o.write(txt)   # falla si hay caracteres no representables: así se detectan antes de entregar
    shutil.copy(os.path.join(RAIZ, "instalador", "INSTALAR.bat"), DIST)
    ps = open(os.path.join(RAIZ, "instalador", "instalar.ps1"), encoding="utf-8").read()
    with open(os.path.join(DIST, "instalar.ps1"), "w", encoding="utf-8-sig", newline="\r\n") as o:
        o.write(ps)
    for m in ("MANUAL_DE_USO.md", "MANUAL_DE_USO.pdf"):
        p = os.path.join(RAIZ, m)
        if os.path.exists(p):
            shutil.copy(p, DIST)
    zpath = os.path.join(RAIZ, "dist", "SistemaEmpresarial.zip")
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _d, files in os.walk(DIST):
            for fn in files:
                full = os.path.join(root, fn)
                z.write(full, os.path.relpath(full, os.path.dirname(DIST)))
    print(json.dumps({"ok": True, "zip": zpath, "archivos": sum(len(f) for _r, _d, f in os.walk(DIST))}))


if __name__ == "__main__":
    main()
