"""Exporta hojas a PDF/PNG para revisión visual. Uso: python render.py libro.xlsx salida_dir"""
import sys, os, subprocess
import uno
from lo_util import arrancar, abrir, prop
from com.sun.star.table import CellRangeAddress
libro, out = sys.argv[1], sys.argv[2]
os.makedirs(out, exist_ok=True)
objetivos = {"FACTURACIÓN": None, "GUÍAS DE ENVÍO": None, "INICIO": "A1:V82", "INVENTARIO": "A1:X40",
             "COSTOS Y RENTABILIDAD": "A1:L45", "CARTERA": "A1:R40", "CAJA Y BANCOS": "A1:W30", "UTILIDADES": "A1:O45"}
proc, ctx, desk = arrancar()
try:
    for sh, area in objetivos.items():
        d = abrir(desk, libro); d.calculateAll()
        sheets = d.Sheets
        for n in sheets.ElementNames:
            if n != sh:
                sheets.getByName(n).IsVisible = False
        s = sheets.getByName(sh)
        if area:
            r = s.getCellRangeByName(area).RangeAddress
            s.setPrintAreas((r,))
            ps = d.StyleFamilies.getByName("PageStyles").getByName(s.PageStyle)
            ps.ScaleToPagesX = 1; ps.ScaleToPagesY = 1; ps.IsLandscape = True
            ps.Width, ps.Height = 27940, 21590
        fn = os.path.join(out, sh.replace(" ", "_") + ".pdf")
        d.storeToURL(uno.systemPathToFileUrl(os.path.abspath(fn)), (prop("FilterName", "calc_pdf_Export"),))
        d.close(True)
        subprocess.run(["pdftoppm", "-png", "-r", "70", "-f", "1", "-l", "1", fn, fn[:-4]])
finally:
    desk.terminate(); proc.wait(timeout=30)
