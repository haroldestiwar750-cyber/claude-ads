"""Exporta los módulos VBA de vba_fuente/ (UTF-8) a ../vba/ en ANSI Windows-1252 con CRLF,
que es la codificación que espera el editor de VBA al importar .bas en Excel en español.

Uso: python exportar_vba.py
Salida (JSON): {"archivos": [...], "destino": ...}
"""
import glob, json, os
AQUI = os.path.dirname(os.path.abspath(__file__))
DEST = os.path.join(AQUI, "..", "vba")
os.makedirs(DEST, exist_ok=True)
out = []
for src in sorted(glob.glob(os.path.join(AQUI, "vba_fuente", "*"))):
    txt = open(src, encoding="utf-8").read().replace("\r\n", "\n")
    data = txt.replace("\n", "\r\n").encode("cp1252")   # falla si hay un carácter no representable
    open(os.path.join(DEST, os.path.basename(src)), "wb").write(data)
    out.append(os.path.basename(src))
print(json.dumps({"archivos": out, "destino": os.path.abspath(DEST)}))
