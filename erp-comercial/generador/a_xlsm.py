"""Convierte ERP_Comercial.xlsx (ya recalculado) en ERP_Comercial.xlsm listo para recibir macros.

Solo cambia el tipo de contenido del libro a "habilitado para macros"; el contenido es idéntico.
Uso: python a_xlsm.py   ->  JSON {"xlsm": ruta}
"""
import json, os, zipfile
AQUI = os.path.dirname(os.path.abspath(__file__))
src = os.path.join(AQUI, "..", "ERP_Comercial.xlsx")
dst = os.path.join(AQUI, "..", "ERP_Comercial.xlsm")
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
XLSM = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
    for it in zi.infolist():
        data = zi.read(it.filename)
        if it.filename == "[Content_Types].xml":
            assert XLSX.encode() in data
            data = data.replace(XLSX.encode(), XLSM.encode())
        zo.writestr(it, data)
print(json.dumps({"xlsm": os.path.abspath(dst)}))
