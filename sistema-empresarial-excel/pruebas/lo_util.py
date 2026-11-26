"""Utilidades para controlar LibreOffice por UNO (recalcular, leer celdas, ejecutar macros).

Uso CLI:
    python lo_util.py errores <archivo.xlsx>   -> JSON con celdas en error por hoja
"""
import json
import os
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue

PORT = 2002


def prop(n, v):
    p = PropertyValue()
    p.Name, p.Value = n, v
    return p


def arrancar():
    """Inicia soffice headless y devuelve el desktop."""
    prof = "file:///tmp/lo_prof_sistema"
    proc = subprocess.Popen(["soffice", "--headless", "--norestore", "--nologo", "--nodefault",
                             f"-env:UserInstallation={prof}",
                             f"--accept=socket,host=localhost,port={PORT};urp;"],
                            stdout=open("/tmp/claude-0/-home-user-claude-ads/56bf4dd0-b601-5add-bb9d-c0f48431541e/scratchpad/lo_out.txt", "w"), stderr=subprocess.STDOUT)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(60):
        try:
            ctx = resolver.resolve(f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext")
            break
        except Exception:
            time.sleep(0.5)
    else:
        raise RuntimeError("No se pudo conectar a LibreOffice")
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return proc, ctx, desktop


def abrir(desktop, path, macros=False):
    url = uno.systemPathToFileUrl(os.path.abspath(path))
    args = [prop("Hidden", True)]
    if macros:
        args.append(prop("MacroExecutionMode", 4))  # ALWAYS_EXECUTE_NO_WARN
    return desktop.loadComponentFromURL(url, "_blank", 0, tuple(args))


def errores(doc, max_por_hoja=15):
    res = {}
    for name in doc.Sheets.ElementNames:
        sh = doc.Sheets.getByName(name)
        cur = sh.createCursor()
        cur.gotoEndOfUsedArea(False)
        lc, lr = cur.RangeAddress.EndColumn, cur.RangeAddress.EndRow
        data = sh.getCellRangeByPosition(0, 0, lc, lr)
        ranges = data.queryFormulaCells(4)  # com.sun.star.sheet.FormulaResult.ERROR = 4
        out = []
        for addr in ranges.RangeAddresses:
            for r in range(addr.StartRow, addr.EndRow + 1):
                for c in range(addr.StartColumn, addr.EndColumn + 1):
                    cl = sh.getCellByPosition(c, r)
                    out.append((cl.AbsoluteName.split(".")[-1], cl.getError(), cl.Formula[:140]))
                    if len(out) >= max_por_hoja:
                        break
                if len(out) >= max_por_hoja:
                    break
        if out:
            res[name] = out
    return res


def val(doc, sheet, addr):
    c = doc.Sheets.getByName(sheet).getCellRangeByName(addr)
    if c.getError():
        return "ERR:" + c.String
    t = c.Type.value
    if t == "TEXT" or (t == "FORMULA" and c.FormulaResultType2 == 2):
        return c.String
    return c.Value


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "errores":
        proc, ctx, desk = arrancar()
        try:
            d = abrir(desk, sys.argv[2])
            d.calculateAll()
            print(json.dumps(errores(d), ensure_ascii=False, indent=1))
            d.close(True)
        finally:
            try:
                desk.terminate()
            except Exception:
                pass
            proc.wait(timeout=30)
