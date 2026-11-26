"""
Pruebas extendidas (complementan las 13 obligatorias) ejecutando las macros
VBA reales en LibreOffice:
  1. EjecutarPruebas(False) debe dejar el libro EXACTAMENTE como estaba
  2. Devolución de cliente, ajuste de inventario, pago de gasto fijo,
     ingreso manual, cierre de caja, búsquedas, 11 reportes, numeración
  3. Borrado de datos de demostración (libro listo para operar)

Uso CLI:
    python pruebas_extendidas.py ../SistemaEmpresarial_base.xlsx
"""
import json
import sys

from lo_util import arrancar, abrir
from ejecutar_vba_lo import cargar_modulos, correr

EXTRA = r'''Option VBASupport 1
Private Function Cnt(ByVal h As String, ByVal k As Long, ByVal f As Long) As Long
    Cnt = UltimaFila(Hoja(h), k, f) - f + 1
End Function

Public Function Conteos() As String
    Conteos = Cnt(SH_BD, 1, BD_FIRST) & "," & Cnt(SH_HFA, 1, HFA_FIRST) & "," & Cnt(SH_MOV, 1, MOV_FIRST) & "," & _
              Cnt(SH_ING, 1, ING_FIRST) & "," & Cnt(SH_EGR, 1, EGR_FIRST) & "," & Cnt(SH_CAR, 1, CAR_FIRST) & "," & _
              Cnt(SH_COM, 1, COM_FIRST) & "," & Cnt(SH_HGU, 1, HGU_FIRST) & "," & Cnt(SH_INV, 1, INV_FIRST) & "," & _
              CfgV(CFG_SIGFACTURA) & "," & CfgV(CFG_SIGGUIA) & "," & CfgV(CFG_SIGCOMPRA)
End Function

Private Sub Q(ParamArray v() As Variant)
    Dim i As Long
    Set gRespuestas = New Collection
    For i = LBound(v) To UBound(v)
        gRespuestas.Add CStr(v(i))
    Next i
End Sub

Public Function Extendidas() As String
    Dim s As String, sku As String, st0 As Double, num As String, ws As Worksheet, f As Long, ult As Long
    Dim e0 As Long, i0 As Long, tipos As Variant, t As Variant, clave As String
    On Error GoTo E
    gTest = True
    clave = Tv(CfgV(CFG_CLAVE))
    ' --- devolución: tomar una factura activa pagada con una línea
    Set ws = Hoja(SH_HFA)
    ult = UltimaFila(ws, 1, HFA_FIRST)
    For f = ult To HFA_FIRST Step -1
        If Tv(ws.Cells(f, HFA_ESTADO).Value) = "ACTIVA" And Tv(ws.Cells(f, HFA_ESTADOPAGO).Value) = "PAGADA" Then Exit For
    Next f
    num = Tv(ws.Cells(f, HFA_FACTURA).Value)
    sku = Tv(Hoja(SH_BD).Cells(BuscarFila(Hoja(SH_BD), BD_FACTURA, BD_FIRST, num), BD_SKU).Value)
    st0 = StockReal(sku)
    i0 = Cnt(SH_ING, 1, ING_FIRST)
    Q num, sku, "1", "Talla equivocada", "Efectivo"
    RegistrarDevolucion
    s = s & "devolucion_stock:" & (StockReal(sku) - st0) & ";"
    s = s & "devolucion_ingreso_neg:" & Nv(Hoja(SH_ING).Cells(i0 + ING_FIRST, ING_VALOR).Value) & ";"
    ' --- ajuste
    st0 = StockReal(sku)
    Q sku, "-2", "Conteo físico", clave
    AjusteInventario
    s = s & "ajuste:" & (StockReal(sku) - st0) & ";"
    ' --- ajuste con clave errada
    Q sku, "5", "x", "mala"
    AjusteInventario
    s = s & "ajuste_rechazado:" & (StockReal(sku) - st0) & ";"
    ' --- gasto fijo
    e0 = Cnt(SH_EGR, 1, EGR_FIRST)
    Q "50000"
    PagarGastoFijo
    s = s & "gastofijo:" & (Cnt(SH_EGR, 1, EGR_FIRST) - e0) & "/" & Tv(Hoja(SH_EGR).Cells(e0 + EGR_FIRST, EGR_DOC).Value) & ";"
    ' --- ingreso manual
    i0 = Cnt(SH_ING, 1, ING_FIRST)
    With Hoja(SH_ING)
        .Range(FING_FECHA).Value = Date
        .Range(FING_TIPO).Value = "APORTE DE SOCIOS"
        .Range(FING_CONCEPTO).Value = "Aporte de prueba"
        .Range(FING_VALOR).Value = 1000000
        .Range(FING_METODO).Value = "Transferencia"
    End With
    RegistrarIngresoManual
    s = s & "ingreso_manual:" & (Cnt(SH_ING, 1, ING_FIRST) - i0) & "/" & Tv(Hoja(SH_ING).Cells(i0 + ING_FIRST, ING_CUENTA).Value) & ";"
    ' --- ingreso manual tipo VENTA debe rechazarse
    i0 = Cnt(SH_ING, 1, ING_FIRST)
    With Hoja(SH_ING)
        .Range(FING_TIPO).Value = "VENTA"
        .Range(FING_CONCEPTO).Value = "x"
        .Range(FING_VALOR).Value = 5
    End With
    RegistrarIngresoManual
    s = s & "ingreso_venta_rechazado:" & (Cnt(SH_ING, 1, ING_FIRST) - i0) & ";"
    ' --- cierre de caja
    CierreMensual
    GuardarCierre
    s = s & "cierre:" & Tv(Hoja(SH_CAJ).Cells(CIE_FIRST, CIE_TIPO).Value) & "/" & Nv(Hoja(SH_CAJ).Cells(CIE_FIRST, CIE_SALDO).Value) & ";"
    ' --- búsquedas
    Q num
    BuscarFactura
    s = s & "buscar:" & Tv(Hoja(SH_FAC).Range(FAC_NUM).Value) & "/" & Tv(Hoja(SH_FAC).Range(FAC_ESTADO).Value) & "/" & Nv(Hoja(SH_FAC).Range(FAC_TOTAL).Value) & "=" & Nv(ws.Cells(f, HFA_TOTAL).Value) & ";"
    Q "GE-0002"
    BuscarGuia
    s = s & "buscar_guia:" & Tv(Hoja(SH_GUI).Range(GUI_NUM).Value) & "/" & Tv(Hoja(SH_GUI).Range(GUI_REG).Value) & ";"
    ' --- nueva factura restaura fórmulas de precio
    NuevaFacturaInterna False
    s = s & "formula_precio:" & (Left(Hoja(SH_FAC).Range("G14").Formula, 1) = "=") & ";"
    ' --- reportes
    tipos = Array("Ventas diario", "Ventas mensual", "Ventas anual", "Inventario", "Productos vendidos", "Utilidades", _
                  "Gastos", "Cartera", "Compras", "Clientes", "Flujo de caja")
    For Each t In tipos
        Hoja(SH_REP).Range(FREP_TIPO).Value = t
        Hoja(SH_REP).Range(FREP_DESDE).Value = DateSerial(2026, 6, 1)
        Hoja(SH_REP).Range(FREP_HASTA).Value = Date
        GenerarReporte
        s = s & "rep[" & t & "]:" & (UltimaFila(Hoja(SH_REP), 2, FREP_OUT) - FREP_OUT) & ";"
    Next t
    ' --- numeración: menor al mínimo rechazado; mayor aceptado
    Q clave, "1", "3"
    CambiarNumeracion
    s = s & "num_rechazada:" & CfgV(CFG_SIGFACTURA) & ";"
    Q clave, "1", "500"
    CambiarNumeracion
    s = s & "num_aceptada:" & CfgV(CFG_SIGFACTURA) & "/" & Tv(Hoja(SH_FAC).Range(FAC_NUM).Value) & ";"
    ' --- borrar datos demo
    Q clave
    LimpiarDatosDemo
    s = s & "demo_borrado:" & Conteos() & "/cli=" & (UltimaFila(Hoja(SH_CLI), 1, CLI_FIRST) - CLI_FIRST + 1) & ";"
    Extendidas = s & "FIN"
    Exit Function
E:
    Extendidas = s & " ERR " & Err.Number & ": " & Err.Description
End Function
'''

if __name__ == "__main__":
    proc, ctx, desk = arrancar()
    out = {}
    try:
        d = abrir(desk, sys.argv[1])
        d.calculateAll()
        cargar_modulos(d)
        d.BasicLibraries.getByName("Standard").insertByName("modExtra", EXTRA)
        antes = correr(d, "modExtra", "Conteos")
        fallos = correr(d, "modPruebas", "EjecutarPruebas", (False,))
        despues = correr(d, "modExtra", "Conteos")
        out["pruebas_obligatorias_fallos"] = fallos
        out["conteos_antes"] = antes
        out["conteos_despues_revertir"] = despues
        out["revertido_exacto"] = antes == despues
        out["extendidas"] = correr(d, "modExtra", "Extendidas").split(";")
        out["log_final"] = correr(d, "modPruebas", "LeerLog").splitlines()[-6:]
        d.close(True)
    finally:
        desk.terminate()
        proc.wait(timeout=60)
    print(json.dumps(out, ensure_ascii=False, indent=1))
