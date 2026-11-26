Attribute VB_Name = "modPruebas"
'==============================================================
' modPruebas - PRUEBAS AUTOMÁTICAS OBLIGATORIAS
'  Ejecuta el proceso completo de punta a punta usando las MISMAS
'  macros que usan los botones (sin cuadros de diálogo), verifica
'  cada resultado y escribe el informe en la hoja REPORTES.
'  Al terminar borra todos los registros de prueba (salvo que se
'  llame EjecutarPruebas True para conservarlos).
'==============================================================
Option Explicit

Private resN As Long, resFallos As Long
Private res() As Variant
Private sH() As String, sK() As Long, sF() As Long, sC0() As Long, sNC() As Long, sN As Long

Public Sub PruebasAutomaticas()
    Dim f As Long
    If Msg("Se ejecutarán las 13 pruebas obligatorias del sistema (producto, compra, precio, factura, inventario," & vbLf & _
           "utilidad, guía, abono, gasto, panel y anulación)." & vbLf & vbLf & _
           "Los registros de prueba se eliminan automáticamente al final. ¿Continuar?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    f = EjecutarPruebas(False)
    Hoja(SH_REP).Activate
    If f = 0 Then
        Msg "TODAS LAS PRUEBAS PASARON (" & resN & " verificaciones)." & vbLf & "El informe quedó en la hoja REPORTES.", vbInformation
    Else
        Msg f & " verificación(es) FALLARON de " & resN & ". Revise el informe en la hoja REPORTES.", vbCritical
    End If
End Sub

Private Sub Chk(ByVal paso As String, ByVal desc As String, ByVal esperado As Variant, ByVal obtenido As Variant)
    Dim ok As Boolean
    resN = resN + 1
    If resN > UBound(res, 1) Then Exit Sub
    If IsNumeric(esperado) And IsNumeric(obtenido) And VarType(esperado) <> vbString Then
        ok = Abs(CDbl(esperado) - CDbl(obtenido)) < 0.51
    Else
        ok = (UCase(CStr(esperado)) = UCase(CStr(obtenido)))
    End If
    res(resN, 1) = paso
    res(resN, 2) = desc
    res(resN, 3) = esperado
    res(resN, 4) = obtenido
    If ok Then
        res(resN, 5) = "OK"
    Else
        res(resN, 5) = "FALLA"
        resFallos = resFallos + 1
    End If
End Sub

Private Sub Snap(ByVal h As String, ByVal k As Long, ByVal first As Long, ByVal c0 As Long, ByVal nc As Long)
    sN = sN + 1
    ReDim Preserve sH(1 To sN): ReDim Preserve sK(1 To sN): ReDim Preserve sF(1 To sN)
    ReDim Preserve sC0(1 To sN): ReDim Preserve sNC(1 To sN)
    sH(sN) = h: sK(sN) = k: sC0(sN) = c0: sNC(sN) = nc
    sF(sN) = UltimaFila(Hoja(h), k, first) + 1
End Sub

Private Function Contar(ByVal h As String, ByVal colBusca As Long, ByVal first As Long, ByVal valor As String, _
                        Optional ByVal colFiltro As Long = 0, Optional ByVal valFiltro As String = "") As Long
    Dim ws As Worksheet, ult As Long, r As Long
    Set ws = Hoja(h)
    ult = UltimaFila(ws, 1, first)
    If ult < first Then ult = UltimaFila(ws, colBusca, first)
    For r = first To ult
        If UCase(Tv(ws.Cells(r, colBusca).Value)) = UCase(valor) Then
            If colFiltro = 0 Then
                Contar = Contar + 1
            ElseIf UCase(Tv(ws.Cells(r, colFiltro).Value)) = UCase(valFiltro) Then
                Contar = Contar + 1
            End If
        End If
    Next r
End Function

Private Function Sumar(ByVal h As String, ByVal colBusca As Long, ByVal first As Long, ByVal valor As String, ByVal colSuma As Long, _
                       Optional ByVal colFiltro As Long = 0, Optional ByVal valFiltro As String = "") As Double
    Dim ws As Worksheet, ult As Long, r As Long
    Set ws = Hoja(h)
    ult = UltimaFila(ws, 1, first)
    For r = first To ult
        If UCase(Tv(ws.Cells(r, colBusca).Value)) = UCase(valor) Then
            If colFiltro = 0 Then
                Sumar = Sumar + Nv(ws.Cells(r, colSuma).Value)
            ElseIf UCase(Tv(ws.Cells(r, colFiltro).Value)) = UCase(valFiltro) Then
                Sumar = Sumar + Nv(ws.Cells(r, colSuma).Value)
            End If
        End If
    Next r
End Function

Private Function StockFormula(ByVal sku As String) As Double
    Dim f As Long
    Recalcular
    f = BuscarFila(Hoja(SH_INV), INV_SKU, INV_FIRST, sku)
    If f > 0 Then StockFormula = Nv(Hoja(SH_INV).Cells(f, INV_STOCK).Value)
End Function

Private Function RedondeoCien(ByVal x As Double) As Double
    RedondeoCien = -Int(-x / 100) * 100
End Function

' Ejecuta todas las pruebas. Devuelve el número de fallas.
Public Function EjecutarPruebas(ByVal conservar As Boolean) As Long
    Dim wi As Worksheet, wf As Worksheet, wc As Worksheet, wg As Worksheet, we As Worksheet, wn As Worksheet, wr As Worksheet
    Dim sku As String, fInv As Long, prov As String, cat As String, sug As Double, esperadoSug As Double
    Dim num As String, num2 As String, total As Double, fCar As Long, idCar As String, cliente As String
    Dim v0Dia As Double, v0Cxc As Double, v0Gas As Double, v1 As Double, empaque As Double, com As Double
    Dim sig(1 To 3) As Variant, ini(1 To 4) As Variant, cos5 As Variant, r As Long, clave As String, filaInvPrueba As Long
    Dim abono As Double

    ReDim res(1 To 80, 1 To 5)
    resN = 0: resFallos = 0: sN = 0
    gTest = True
    gLog = ""
    gNoRespuestas = 0
    Set gRespuestas = New Collection
    On Error GoTo Fallo
    DesprotegerTodo
    Set wi = Hoja(SH_INV): Set wf = Hoja(SH_FAC): Set wc = Hoja(SH_CAR): Set wg = Hoja(SH_GUI)
    Set we = Hoja(SH_EGR): Set wn = Hoja(SH_INI)
    ' ---- instantánea para revertir
    Snap SH_BD, BD_FACTURA, BD_FIRST, 1, BD_NCOLS
    Snap SH_HFA, HFA_FACTURA, HFA_FIRST, 1, HFA_NCOLS
    Snap SH_HGU, HGU_GUIA, HGU_FIRST, 1, HGU_NCOLS
    Snap SH_MOV, MOV_ID, MOV_FIRST, 1, MOV_NCOLS
    Snap SH_AUD, AUD_FECHA, AUD_FIRST, 1, AUD_NCOLS
    Snap SH_COM, COM_NUMERO, COM_FIRST, 1, COM_NCOLS
    Snap SH_ING, ING_ID, ING_FIRST, 1, ING_NCOLS
    Snap SH_EGR, EGR_ID, EGR_FIRST, 1, EGR_NCOLS
    Snap SH_CAR, CAR_ID, CAR_FIRST, 1, CAR_NCOLS
    Snap SH_INV, INV_SKU, INV_FIRST, 1, INV_NCOLS
    sig(1) = CfgV(CFG_SIGFACTURA): sig(2) = CfgV(CFG_SIGGUIA): sig(3) = CfgV(CFG_SIGCOMPRA)
    ini(1) = wn.Range(FINI_ANIO).Value: ini(2) = wn.Range(FINI_MES).Value
    ini(3) = wn.Range(FINI_CAT).Value: ini(4) = wn.Range(FINI_DIA).Formula
    cos5 = Hoja(SH_COS).Range("C5").Value
    clave = Tv(CfgV(CFG_CLAVE))
    empaque = Nv(CfgV(CFG_EMPAQUE))
    com = Nv(CfgV(CFG_COMISION))
    cat = Tv(Hoja(SH_CFG).Range(LST_CATEGORIAS & CFG_LIST_FIRST).Value)
    prov = Tv(Hoja(SH_PRO).Cells(PRO_FIRST, PRO_NOMBRE).Value)
    If prov = "" Then
        Hoja(SH_PRO).Cells(PRO_FIRST, PRO_CODIGO).Value = "PR-TEST"
        Hoja(SH_PRO).Cells(PRO_FIRST, PRO_NOMBRE).Value = "Proveedor de prueba"
        prov = "Proveedor de prueba"
    End If
    cliente = "CL-0002"
    If BuscarFila(Hoja(SH_CLI), CLI_CODIGO, CLI_FIRST, cliente) = 0 Then cliente = "CL-0001"
    wn.Range(FINI_ANIO).Value = Year(Date)
    wn.Range(FINI_MES).Value = Hoja(SH_CFG).Range(LST_MESES & (CFG_LIST_FIRST + Month(Date) - 1)).Value
    wn.Range(FINI_CAT).Value = "TODAS"
    wn.Range(FINI_DIA).Value = Date

    ' ================= 1. Registrar producto =================
    sku = "PRUEBA-" & Format(Now, "hhmmss")
    fInv = UltimaFila(wi, INV_SKU, INV_FIRST) + 1
    filaInvPrueba = fInv
    wi.Cells(fInv, INV_SKU).Value = sku
    wi.Cells(fInv, INV_NOMBRE).Value = "Producto de prueba"
    wi.Cells(fInv, INV_CATEGORIA).Value = cat
    wi.Cells(fInv, INV_TALLA).Value = "M"
    wi.Cells(fInv, INV_COLOR).Value = "Negro"
    wi.Cells(fInv, INV_PROVEEDOR).Value = prov
    wi.Cells(fInv, INV_COSTO).Value = 0
    wi.Cells(fInv, INV_PRECIO).Value = 0
    wi.Cells(fInv, INV_STOCKINI).Value = 0
    wi.Cells(fInv, INV_STOCKMIN).Value = 2
    wi.Cells(fInv, INV_ACTIVO).Value = "SI"
    Chk "1", "Producto creado en INVENTARIO", sku, Tv(wi.Cells(BuscarFila(wi, INV_SKU, INV_FIRST, sku), INV_SKU).Value)

    ' ================= 2. Compra de 20 unidades (costo de compra) =================
    With Hoja(SH_COM)
        .Range(FCOM_FECHA).Value = Date
        .Range(FCOM_PROV).Value = prov
        .Range(FCOM_FORMA).Value = "Efectivo"
        .Range(FCOM_ESTADOPAGO).Value = "PAGADA"
        .Range(FCOM_PAGADO).Value = ""
        .Range(FCOM_TRANSP).Value = 20000
        .Range(FCOM_OTROS).Value = ""
        .Range("B" & FCOM_L1 & ":B" & FCOM_LN).ClearContents
        .Range("F" & FCOM_L1 & ":G" & FCOM_LN).ClearContents
        .Cells(FCOM_L1, 2).Value = sku
        .Cells(FCOM_L1, 6).Value = 20
        .Cells(FCOM_L1, 7).Value = 30000
    End With
    RegistrarCompra
    gTest = True
    Chk "2", "Stock tras la compra (fórmula INVENTARIO)", 20, StockFormula(sku)
    Chk "2", "Stock tras la compra (kardex)", 20, StockReal(sku)
    Chk "2", "Costo de compra registrado", 30000, Nv(wi.Cells(fInv, INV_COSTO).Value)
    Chk "2", "Costo real unitario (incluye flete $1.000/u)", 31000, Nv(wi.Cells(fInv, INV_COSTOREAL).Value)
    Chk "2", "Egreso del pago de la compra", 620000, Sumar(SH_EGR, EGR_ORIGEN, EGR_FIRST, "COMPRAS", EGR_VALOR, EGR_DOC, NumeroCompra(CLng(sig(3))))

    ' ================= 3. Precio sugerido =================
    DesprotegerTodo
    Hoja(SH_COS).Range("C5").Value = sku
    Recalcular
    sug = Nv(Hoja(SH_COS).Range("D30").Value)
    esperadoSug = RedondeoCien((31000 + empaque) / (1 - com - Nv(CfgV(CFG_MARGENREC))))
    Chk "3", "Precio sugerido (margen recomendado)", esperadoSug, sug
    wi.Cells(fInv, INV_PRECIO).Value = sug
    Recalcular
    v0Dia = Nv(wn.Range("C7").Value)
    v0Cxc = Nv(wn.Range("F11").Value)
    v0Gas = Nv(Hoja(SH_CALC).Range("M2").Value)

    ' ================= 4-5. Factura por 3 unidades y guardar =================
    NuevaFacturaInterna False
    DesprotegerTodo
    wf.Range(FAC_CLIENTE).Value = cliente & " | x"
    wf.Cells(FAC_L1, 2).Value = sku
    wf.Cells(FAC_L1, 6).Value = 3
    wf.Range(FAC_METODO).Value = "Efectivo"
    wf.Range(FAC_ESTADOPAGO).Value = "ABONO PARCIAL"
    wf.Range(FAC_PAGADO).Value = 50000
    Recalcular
    num = Tv(wf.Range(FAC_NUM).Value)
    total = 3 * sug
    Chk "4", "Total calculado en el formato de factura", total, Nv(wf.Range(FAC_TOTAL).Value)
    GuardarFactura
    gTest = True
    Chk "5", "Factura guardada (estado del formato)", "GUARDADA", Tv(wf.Range(FAC_ESTADO).Value)
    Chk "5", "Factura en HISTORIAL FACTURAS", 1, Contar(SH_HFA, HFA_FACTURA, HFA_FIRST, num)
    Chk "5", "Consecutivo avanzó en 1", CLng(sig(1)) + 1, CLng(CfgV(CFG_SIGFACTURA))

    ' ================= 6. Inventario 20 -> 17 =================
    Chk "6", "Stock tras la venta (fórmula INVENTARIO)", 17, StockFormula(sku)
    Chk "6", "Stock tras la venta (kardex)", 17, StockReal(sku)
    ' intento de doble guardado (no debe duplicar ni descontar otra vez)
    GuardarFactura
    gTest = True
    Chk "6", "Doble guardado bloqueado: líneas de venta", 1, Contar(SH_BD, BD_FACTURA, BD_FIRST, num)
    Chk "6", "Doble guardado bloqueado: stock sin cambios", 17, StockFormula(sku)

    ' ================= 7. Venta registrada =================
    Chk "7", "Unidades vendidas registradas en BASE DE DATOS", 3, Sumar(SH_BD, BD_FACTURA, BD_FIRST, num, BD_CANT)
    Chk "7", "Valor de la venta", total, Sumar(SH_BD, BD_FACTURA, BD_FIRST, num, BD_SUBTOTAL)
    Chk "7", "Ingreso por el pago inicial (sin duplicar)", 50000, Sumar(SH_ING, ING_DOC, ING_FIRST, num, ING_VALOR, ING_ESTADO, "ACTIVO")
    Chk "7", "Cuenta por cobrar creada", 1, Contar(SH_CAR, CAR_DOC, CAR_FIRST, num)

    ' ================= 8. Utilidad =================
    Chk "8", "Utilidad bruta de la venta (venta - 3 x costo real)", total - 3 * 31000, Sumar(SH_BD, BD_FACTURA, BD_FIRST, num, BD_UTILIDAD)
    Chk "8", "Utilidad en historial de facturas", total - 3 * 31000, Sumar(SH_HFA, HFA_FACTURA, HFA_FIRST, num, HFA_UTILIDAD)

    ' ================= 9. Guía de envío =================
    CargarGuiaDesdeFactura num
    DesprotegerTodo
    Hoja(SH_GUI).Range(GUI_TRANSP).Value = "Servientrega"
    Hoja(SH_GUI).Range(GUI_SEG).Value = "PRUEBA-123"
    If Tv(Hoja(SH_GUI).Range(GUI_TEL).Value) = "" Then Hoja(SH_GUI).Range(GUI_TEL).Value = "3000000000"
    If Tv(Hoja(SH_GUI).Range(GUI_CIUDAD).Value) = "" Then Hoja(SH_GUI).Range(GUI_CIUDAD).Value = "Bogotá D.C."
    If Tv(Hoja(SH_GUI).Range(GUI_DIR).Value) = "" Then Hoja(SH_GUI).Range(GUI_DIR).Value = "Dirección de prueba"
    GuardarGuia
    gTest = True
    Chk "9", "Guía guardada para la factura", 1, Contar(SH_HGU, HGU_FACTURA, HGU_FIRST, num)
    Chk "9", "Unidades en la guía", 3, Sumar(SH_HGU, HGU_FACTURA, HGU_FIRST, num, HGU_UNIDADES)

    ' ================= 10. Abono de cliente =================
    Recalcular
    fCar = BuscarFila(wc, CAR_DOC, CAR_FIRST, num)
    idCar = Tv(wc.Cells(fCar, CAR_ID).Value)
    Chk "10", "Saldo inicial de la cartera", total - 50000, Nv(wc.Cells(fCar, CAR_SALDO).Value)
    abono = 30000
    DesprotegerTodo
    wc.Range(FCAR_OBLIG).Value = idCar & " | " & num
    wc.Range(FCAR_FECHA).Value = Date
    wc.Range(FCAR_VALOR).Value = abono
    wc.Range(FCAR_METODO).Value = "Nequi"
    RegistrarAbono
    gTest = True
    Recalcular
    Chk "10", "Saldo tras el abono", total - 50000 - abono, Nv(wc.Cells(fCar, CAR_SALDO).Value)
    Chk "10", "Abono registrado en INGRESOS", abono, Sumar(SH_ING, ING_DOC, ING_FIRST, num, ING_VALOR, ING_TIPO, "ABONO")

    ' ================= 11. Gasto =================
    DesprotegerTodo
    With we
        .Range(FEGR_FECHA).Value = Date
        .Range(FEGR_CONCEPTO).Value = "Gasto de prueba automatizada"
        .Range(FEGR_CATEGORIA).Value = "Publicidad"
        .Range(FEGR_TIPO).Value = "VARIABLE"
        .Range(FEGR_VALOR).Value = 45000
        .Range(FEGR_METODO).Value = "Efectivo"
    End With
    RegistrarEgresoManual
    gTest = True
    Chk "11", "Gasto registrado en EGRESOS", 45000, Sumar(SH_EGR, EGR_CONCEPTO, EGR_FIRST, "Gasto de prueba automatizada", EGR_VALOR)

    ' ================= 12. Panel gerencial =================
    Recalcular
    Chk "12", "Panel: ventas del día aumentaron", total, Nv(wn.Range("C7").Value) - v0Dia
    Chk "12", "Panel: cartera por cobrar aumentó", total - 50000 - abono, Nv(wn.Range("F11").Value) - v0Cxc
    Chk "12", "Panel: gastos operativos aumentaron", 45000, Nv(Hoja(SH_CALC).Range("M2").Value) - v0Gas

    ' ================= 13. Anular factura de prueba y restaurar inventario =================
    NuevaFacturaInterna False
    DesprotegerTodo
    wf.Range(FAC_CLIENTE).Value = cliente & " | x"
    wf.Cells(FAC_L1, 2).Value = sku
    wf.Cells(FAC_L1, 6).Value = 2
    wf.Range(FAC_METODO).Value = "Efectivo"
    wf.Range(FAC_ESTADOPAGO).Value = "PAGADA"
    num2 = Tv(wf.Range(FAC_NUM).Value)
    GuardarFactura
    gTest = True
    Chk "13", "Segunda factura descuenta 2 unidades", 15, StockFormula(sku)
    Set gRespuestas = New Collection
    gRespuestas.Add num2
    gRespuestas.Add clave
    gRespuestas.Add "Prueba automática de anulación"
    AnularFactura
    gTest = True
    Chk "13", "Factura anulada en el historial", "ANULADA", Tv(Hoja(SH_HFA).Cells(BuscarFila(Hoja(SH_HFA), HFA_FACTURA, HFA_FIRST, num2), HFA_ESTADO).Value)
    Chk "13", "Inventario restaurado tras anular (fórmula)", 17, StockFormula(sku)
    Chk "13", "Inventario restaurado tras anular (kardex)", 17, StockReal(sku)
    Chk "13", "Ingreso de la factura anulada queda ANULADO", 0, Sumar(SH_ING, ING_DOC, ING_FIRST, num2, ING_VALOR, ING_ESTADO, "ACTIVO")
    Set gRespuestas = New Collection
    gRespuestas.Add num2
    gRespuestas.Add clave
    gRespuestas.Add "Segundo intento"
    AnularFactura
    gTest = True
    Chk "13", "Doble anulación bloqueada (stock no se duplica)", 17, StockFormula(sku)
    Set gRespuestas = New Collection
    gRespuestas.Add num
    gRespuestas.Add "clave-incorrecta"
    AnularFactura
    gTest = True
    Chk "13", "Anulación con clave incorrecta rechazada", "ACTIVA", Tv(Hoja(SH_HFA).Cells(BuscarFila(Hoja(SH_HFA), HFA_FACTURA, HFA_FIRST, num), HFA_ESTADO).Value)

Finalizar:
    On Error Resume Next
    gTest = True
    DesprotegerTodo
    If Not conservar Then
        For r = sN To 1 Step -1
            LimpiarFilasDesde Hoja(sH(r)), sK(r), sF(r), sC0(r), sNC(r)
        Next r
        Hoja(SH_CFG).Range(CFG_SIGFACTURA).Value = sig(1)
        Hoja(SH_CFG).Range(CFG_SIGGUIA).Value = sig(2)
        Hoja(SH_CFG).Range(CFG_SIGCOMPRA).Value = sig(3)
        Hoja(SH_COS).Range("C5").Value = cos5
        wn.Range(FINI_ANIO).Value = ini(1): wn.Range(FINI_MES).Value = ini(2): wn.Range(FINI_CAT).Value = ini(3)
        wn.Range(FINI_DIA).Formula = ini(4)
        wc.Range(FCAR_OBLIG).Value = "": wc.Range(FCAR_VALOR).Value = ""
        If prov = "Proveedor de prueba" Then Hoja(SH_PRO).Range(Hoja(SH_PRO).Cells(PRO_FIRST, PRO_CODIGO), Hoja(SH_PRO).Cells(PRO_FIRST, PRO_NOMBRE)).ClearContents
        NuevaFacturaInterna False
        NuevaGuiaInterna False
    End If
    gTest = False
    InformePruebas conservar
    Auditar "SISTEMA", "PRUEBAS AUTOMÁTICAS", "", resN & " verificaciones, " & resFallos & " fallas"
    Terminar
    EjecutarPruebas = resFallos
    Exit Function
Fallo:
    resN = resN + 1
    If resN <= UBound(res, 1) Then
        res(resN, 1) = "ERROR": res(resN, 2) = "Error inesperado: " & Err.Description
        res(resN, 3) = "": res(resN, 4) = "": res(resN, 5) = "FALLA"
    End If
    resFallos = resFallos + 1
    Resume Finalizar
End Function

Private Sub InformePruebas(ByVal conservar As Boolean)
    Dim ws As Worksheet, r0 As Long, i As Long, c As Long, ult As Long
    Set ws = Hoja(SH_REP)
    ult = ws.UsedRange.Row + ws.UsedRange.Rows.Count
    If ult >= FREP_OUT Then ws.Range(ws.Cells(FREP_OUT, 1), ws.Cells(ult + 5, 30)).Clear
    r0 = FREP_OUT
    ws.Cells(r0, 2).Value = "INFORME DE PRUEBAS AUTOMÁTICAS — " & Format(Now, "dd/mm/yyyy hh:mm")
    ws.Cells(r0, 2).Font.Bold = True
    ws.Cells(r0, 2).Font.Size = 14
    ws.Cells(r0 + 1, 2).Value = "Verificaciones: " & resN & "   ·   Fallas: " & resFallos & "   ·   " & _
                                IIf(resFallos = 0, "RESULTADO: TODAS LAS PRUEBAS PASARON", "RESULTADO: HAY FALLAS") & _
                                IIf(conservar, "   (registros de prueba conservados)", "   (registros de prueba eliminados)")
    ws.Cells(r0 + 1, 2).Font.Bold = True
    If resFallos = 0 Then
        ws.Cells(r0 + 1, 2).Font.Color = RGB(21, 128, 61)
    Else
        ws.Cells(r0 + 1, 2).Font.Color = RGB(185, 28, 28)
    End If
    For c = 1 To 5
        ws.Cells(r0 + 3, 1 + c).Value = Array("Paso", "Verificación", "Esperado", "Obtenido", "Resultado")(c - 1)
        ws.Cells(r0 + 3, 1 + c).Font.Bold = True
        ws.Cells(r0 + 3, 1 + c).Font.Color = RGB(255, 255, 255)
        ws.Cells(r0 + 3, 1 + c).Interior.Color = RGB(30, 41, 59)
    Next c
    For i = 1 To resN
        If i > UBound(res, 1) Then Exit For
        For c = 1 To 5
            ws.Cells(r0 + 3 + i, 1 + c).Value = res(i, c)
        Next c
        If res(i, 5) = "OK" Then
            ws.Cells(r0 + 3 + i, 6).Font.Color = RGB(21, 128, 61)
        Else
            ws.Cells(r0 + 3 + i, 6).Font.Color = RGB(185, 28, 28)
        End If
        ws.Cells(r0 + 3 + i, 6).Font.Bold = True
    Next i
    ws.Columns("C").ColumnWidth = 52
End Sub

' Devuelve la bitácora de mensajes del modo prueba (diagnóstico)
Public Function LeerLog() As String
    LeerLog = gLog
End Function
