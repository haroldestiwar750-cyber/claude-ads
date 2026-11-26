Attribute VB_Name = "modReportes"
'==============================================================
' modReportes - REPORTES IMPRIMIBLES Y EXPORTABLES
'  11 reportes generados desde los datos reales. Se pueden
'  exportar a PDF y a Excel (.xlsx) en ESCRITORIO\SISTEMA
'  EMPRESARIAL\REPORTES, o imprimir.
'==============================================================
Option Explicit

Private wsR As Worksheet
Private filaR As Long

Private Sub LimpiarSalida()
    Dim ult As Long
    Set wsR = Hoja(SH_REP)
    ult = wsR.UsedRange.Row + wsR.UsedRange.Rows.Count
    If ult >= FREP_OUT Then
        wsR.Range(wsR.Cells(FREP_OUT, 1), wsR.Cells(ult + 5, 30)).Clear
    End If
    filaR = FREP_OUT
End Sub

Private Sub Titulo(ByVal txt As String, ByVal subTxt As String)
    With wsR.Cells(filaR, 2)
        .Value = Tv(CfgV(CFG_EMPRESA)) & " — " & UCase(txt)
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(30, 41, 59)
    End With
    wsR.Cells(filaR + 1, 2).Value = subTxt & "   ·   Generado: " & Format(Now, "dd/mm/yyyy hh:mm") & " por " & UsuarioActual()
    wsR.Cells(filaR + 1, 2).Font.Color = RGB(100, 116, 139)
    filaR = filaR + 3
End Sub

' Escribe encabezados y datos. fmts: "txt|date|cop|int|pct" separados por coma. totCols: columnas (1-based) a totalizar.
Private Sub Tabla(ByVal encabezados As String, ByVal fmts As String, datos As Variant, ByVal n As Long, Optional ByVal totCols As String = "")
    Dim h() As String, f() As String, i As Long, c As Long, nc As Long, r0 As Long, txt As Variant
    h = Split(encabezados, "|")
    f = Split(fmts, ",")
    nc = UBound(h) + 1
    r0 = filaR
    For c = 1 To nc
        With wsR.Cells(r0, 1 + c)
            .Value = h(c - 1)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(30, 41, 59)
            .WrapText = True
            .HorizontalAlignment = xlCenter
        End With
    Next c
    If n > 0 Then
        wsR.Range(wsR.Cells(r0 + 1, 2), wsR.Cells(r0 + n, 1 + nc)).Value = datos
        For c = 1 To nc
            Select Case Trim(f(c - 1))
                Case "cop": wsR.Range(wsR.Cells(r0 + 1, 1 + c), wsR.Cells(r0 + n + 1, 1 + c)).NumberFormat = "$ #,##0;[Red]-$ #,##0"
                Case "int": wsR.Range(wsR.Cells(r0 + 1, 1 + c), wsR.Cells(r0 + n + 1, 1 + c)).NumberFormat = "#,##0"
                Case "pct": wsR.Range(wsR.Cells(r0 + 1, 1 + c), wsR.Cells(r0 + n + 1, 1 + c)).NumberFormat = "0.0%"
                Case "date": wsR.Range(wsR.Cells(r0 + 1, 1 + c), wsR.Cells(r0 + n + 1, 1 + c)).NumberFormat = "dd/mm/yyyy"
            End Select
        Next c
        wsR.Range(wsR.Cells(r0, 2), wsR.Cells(r0 + n, 1 + nc)).Borders.LineStyle = xlContinuous
        wsR.Range(wsR.Cells(r0, 2), wsR.Cells(r0 + n, 1 + nc)).Borders.Color = RGB(203, 213, 225)
    End If
    filaR = r0 + n + 1
    If totCols <> "" And n > 0 Then
        wsR.Cells(filaR, 2).Value = "TOTAL"
        For Each txt In Split(totCols, ",")
            c = CLng(txt)
            wsR.Cells(filaR, 1 + c).Formula = "=SUM(" & wsR.Cells(r0 + 1, 1 + c).Address(False, False) & ":" & wsR.Cells(r0 + n, 1 + c).Address(False, False) & ")"
        Next txt
        With wsR.Range(wsR.Cells(filaR, 2), wsR.Cells(filaR, 1 + nc))
            .Font.Bold = True
            .Interior.Color = RGB(219, 234, 254)
        End With
        filaR = filaR + 1
    End If
    If n = 0 Then
        wsR.Cells(filaR, 2).Value = "(Sin registros en el período)"
        wsR.Cells(filaR, 2).Font.Italic = True
        filaR = filaR + 1
    End If
    filaR = filaR + 1
End Sub

Private Function EnRango(ByVal v As Variant, ByVal d0 As Date, ByVal d1 As Date) As Boolean
    If IsDate(v) Or IsNumeric(v) Then
        If Nv(v) > 0 Then EnRango = (Int(Nv(v)) >= CDbl(d0) And Int(Nv(v)) <= CDbl(d1))
    End If
End Function

Public Sub GenerarReporte()
    Dim tipo As String, d0 As Date, d1 As Date, subTxt As String, ws As Worksheet
    Set ws = Hoja(SH_REP)
    Recalcular
    tipo = Tv(ws.Range(FREP_TIPO).Value)
    If Not IsDate(ws.Range(FREP_DESDE).Value) Then Msg "Escriba la fecha 'desde'.", vbExclamation: Exit Sub
    d0 = CDate(ws.Range(FREP_DESDE).Value)
    If IsDate(ws.Range(FREP_HASTA).Value) Then
        d1 = CDate(ws.Range(FREP_HASTA).Value)
    Else
        d1 = d0
    End If
    Select Case tipo
        Case "Ventas diario": d1 = d0: subTxt = "Día " & FechaCorta(d0)
        Case "Ventas mensual": d0 = DateSerial(Year(d0), Month(d0), 1): d1 = DateSerial(Year(d0), Month(d0) + 1, 0): subTxt = "Mes " & Format(d0, "mm/yyyy")
        Case "Ventas anual", "Utilidades": d0 = DateSerial(Year(d0), 1, 1): d1 = DateSerial(Year(d0), 12, 31): subTxt = "Año " & Year(d0)
        Case Else: subTxt = "Del " & FechaCorta(d0) & " al " & FechaCorta(d1)
    End Select
    If d1 < d0 Then Msg "La fecha 'hasta' es anterior a 'desde'.", vbExclamation: Exit Sub
    Iniciar
    On Error GoTo Fallo
    LimpiarSalida
    Select Case tipo
        Case "Ventas diario", "Ventas mensual", "Ventas anual": RepVentas tipo, d0, d1, subTxt
        Case "Inventario": RepInventario
        Case "Productos vendidos": RepProductos d0, d1, subTxt
        Case "Utilidades": RepUtilidades Year(d0)
        Case "Gastos": RepGastos d0, d1, subTxt
        Case "Cartera": RepCartera
        Case "Compras": RepCompras d0, d1, subTxt
        Case "Clientes": RepClientes
        Case "Flujo de caja": RepFlujo d0, d1, subTxt
        Case Else
            Terminar
            Msg "Seleccione un tipo de reporte de la lista.", vbExclamation
            Exit Sub
    End Select
    wsR.Columns("B:P").ColumnWidth = 15
    wsR.Columns("C").ColumnWidth = 24
    wsR.PageSetup.PrintArea = wsR.Range(wsR.Cells(FREP_OUT, 2), wsR.Cells(filaR, 14)).Address
    Auditar "REPORTES", "GENERAR", tipo, subTxt
    Terminar
    If Not gTest Then wsR.Activate: wsR.Cells(FREP_OUT, 2).Select
    Exit Sub
Fallo:
    Terminar
    Msg "No se pudo generar el reporte: " & Err.Description, vbCritical
End Sub

Private Sub RepVentas(ByVal tipo As String, ByVal d0 As Date, ByVal d1 As Date, ByVal subTxt As String)
    Dim wh As Worksheet, ult As Long, r As Long, n As Long, d() As Variant, act As Long, anul As Long, tot As Double
    Set wh = Hoja(SH_HFA)
    ult = UltimaFila(wh, HFA_FACTURA, HFA_FIRST)
    ReDim d(1 To Application.Max(1, ult - HFA_FIRST + 1), 1 To 10)
    For r = HFA_FIRST To ult
        If EnRango(wh.Cells(r, HFA_FECHA).Value, d0, d1) Then
            n = n + 1
            d(n, 1) = wh.Cells(r, HFA_FACTURA).Value
            d(n, 2) = wh.Cells(r, HFA_FECHA).Value
            d(n, 3) = wh.Cells(r, HFA_CLIENTE).Value
            d(n, 4) = wh.Cells(r, HFA_METODO).Value
            d(n, 5) = wh.Cells(r, HFA_ESTADOPAGO).Value
            d(n, 6) = wh.Cells(r, HFA_ESTADO).Value
            If Tv(wh.Cells(r, HFA_ESTADO).Value) = "ACTIVA" Then
                d(n, 7) = Nv(wh.Cells(r, HFA_TOTAL).Value)
                d(n, 8) = Nv(wh.Cells(r, HFA_PAGADO).Value)
                d(n, 9) = Nv(wh.Cells(r, HFA_COSTO).Value)
                d(n, 10) = Nv(wh.Cells(r, HFA_UTILIDAD).Value)
                act = act + 1: tot = tot + Nv(wh.Cells(r, HFA_TOTAL).Value)
            Else
                d(n, 7) = 0: d(n, 8) = 0: d(n, 9) = 0: d(n, 10) = 0
                anul = anul + 1
            End If
        End If
    Next r
    Titulo "Reporte de ventas (" & Mid(tipo, 8) & ")", subTxt
    wsR.Cells(filaR, 2).Value = "Facturas activas: " & act & "   ·   Anuladas: " & anul & "   ·   Ventas: " & Moneda(tot) & _
                                "   ·   Ticket promedio: " & Moneda(IIf(act > 0, tot / IIf(act = 0, 1, act), 0))
    filaR = filaR + 2
    Tabla "Factura|Fecha|Cliente|Método|Estado pago|Estado|Total|Pagado al facturar|Costo|Utilidad bruta", _
          "txt,date,txt,txt,txt,txt,cop,cop,cop,cop", d, n, "7,8,9,10"
End Sub

Private Sub RepInventario()
    Dim wi As Worksheet, ult As Long, r As Long, n As Long, d() As Variant
    Set wi = Hoja(SH_INV)
    ult = UltimaFila(wi, INV_SKU, INV_FIRST)
    ReDim d(1 To Application.Max(1, ult - INV_FIRST + 1), 1 To 11)
    For r = INV_FIRST To ult
        n = n + 1
        d(n, 1) = wi.Cells(r, INV_SKU).Value
        d(n, 2) = wi.Cells(r, INV_NOMBRE).Value
        d(n, 3) = wi.Cells(r, INV_CATEGORIA).Value
        d(n, 4) = wi.Cells(r, INV_TALLA).Value
        d(n, 5) = wi.Cells(r, INV_COLOR).Value
        d(n, 6) = Nv(wi.Cells(r, INV_STOCK).Value)
        d(n, 7) = Nv(wi.Cells(r, INV_STOCKMIN).Value)
        d(n, 8) = Nv(wi.Cells(r, INV_COSTOREAL).Value)
        d(n, 9) = Nv(wi.Cells(r, INV_PRECIO).Value)
        d(n, 10) = Nv(wi.Cells(r, INV_VALOR).Value)
        d(n, 11) = wi.Cells(r, INV_ESTADO).Value
    Next r
    Titulo "Reporte de inventario", "Existencias al " & FechaCorta(Date)
    Tabla "SKU|Producto|Categoría|Talla|Color|Stock|Mínimo|Costo real|Precio|Valor inventario|Estado", _
          "txt,txt,txt,txt,txt,int,int,cop,cop,cop,txt", d, n, "6,10"
End Sub

Private Sub RepProductos(ByVal d0 As Date, ByVal d1 As Date, ByVal subTxt As String)
    Dim wb As Worksheet, ult As Long, r As Long, n As Long, i As Long, j As Long, sku As String
    Dim k() As String, d() As Variant, tmp As Variant, c As Long
    Set wb = Hoja(SH_BD)
    ult = UltimaFila(wb, BD_FACTURA, BD_FIRST)
    ReDim k(1 To 2000): ReDim d(1 To 2000, 1 To 8)
    For r = BD_FIRST To ult
        If Tv(wb.Cells(r, BD_ESTADO).Value) = "ACTIVA" And EnRango(wb.Cells(r, BD_FECHA).Value, d0, d1) Then
            sku = Tv(wb.Cells(r, BD_SKU).Value)
            j = 0
            For i = 1 To n
                If k(i) = sku Then j = i: Exit For
            Next i
            If j = 0 And n < 2000 Then
                n = n + 1: j = n: k(n) = sku
                d(j, 1) = sku
                d(j, 2) = Tv(wb.Cells(r, BD_PRODUCTO).Value) & " " & Tv(wb.Cells(r, BD_TALLA).Value) & " " & Tv(wb.Cells(r, BD_COLOR).Value)
                d(j, 3) = wb.Cells(r, BD_CATEGORIA).Value
            End If
            If j > 0 Then
                d(j, 4) = Nv(d(j, 4)) + Nv(wb.Cells(r, BD_CANT).Value)
                d(j, 5) = Nv(d(j, 5)) + Nv(wb.Cells(r, BD_SUBTOTAL).Value)
                d(j, 6) = Nv(d(j, 6)) + Nv(wb.Cells(r, BD_COSTOTOTAL).Value)
                d(j, 7) = Nv(d(j, 7)) + Nv(wb.Cells(r, BD_UTILIDAD).Value)
            End If
        End If
    Next r
    For i = 1 To n
        If Nv(d(i, 5)) <> 0 Then
            d(i, 8) = Nv(d(i, 7)) / Nv(d(i, 5))
        Else
            d(i, 8) = 0
        End If
    Next i
    ' ordenar por unidades (descendente)
    For i = 1 To n - 1
        For j = i + 1 To n
            If Nv(d(j, 4)) > Nv(d(i, 4)) Then
                For c = 1 To 8
                    tmp = d(i, c): d(i, c) = d(j, c): d(j, c) = tmp
                Next c
            End If
        Next j
    Next i
    Titulo "Reporte de productos vendidos", subTxt & "  (ordenado de mayor a menor rotación)"
    Tabla "SKU|Producto|Categoría|Unidades|Ventas|Costo|Utilidad bruta|Margen", "txt,txt,txt,int,cop,cop,cop,pct", d, n, "4,5,6,7"
End Sub

Private Sub RepUtilidades(ByVal anio As Long)
    Dim wb As Worksheet, we As Worksheet, ult As Long, r As Long, m As Long, d(1 To 12, 1 To 8) As Variant
    Dim f As Variant, cat As String
    Set wb = Hoja(SH_BD)
    Set we = Hoja(SH_EGR)
    cat = UCase(Tv(CfgV(CFG_CATCOMPRAS)))
    For m = 1 To 12
        d(m, 1) = Format(DateSerial(anio, m, 1), "mmmm")
        d(m, 2) = 0: d(m, 3) = 0: d(m, 5) = 0
    Next m
    ult = UltimaFila(wb, BD_FACTURA, BD_FIRST)
    For r = BD_FIRST To ult
        f = wb.Cells(r, BD_FECHA).Value
        If Tv(wb.Cells(r, BD_ESTADO).Value) = "ACTIVA" And IsDate(f) Then
            If Year(f) = anio Then
                m = Month(f)
                d(m, 2) = d(m, 2) + Nv(wb.Cells(r, BD_SUBTOTAL).Value)
                d(m, 3) = d(m, 3) + Nv(wb.Cells(r, BD_COSTOTOTAL).Value)
            End If
        End If
    Next r
    ult = UltimaFila(we, EGR_ID, EGR_FIRST)
    For r = EGR_FIRST To ult
        f = we.Cells(r, EGR_FECHA).Value
        If Tv(we.Cells(r, EGR_ESTADO).Value) = "ACTIVO" And IsDate(f) And UCase(Tv(we.Cells(r, EGR_CATEGORIA).Value)) <> cat Then
            If Year(f) = anio Then d(Month(f), 5) = d(Month(f), 5) + Nv(we.Cells(r, EGR_VALOR).Value)
        End If
    Next r
    For m = 1 To 12
        d(m, 4) = d(m, 2) - d(m, 3)
        d(m, 6) = d(m, 4) - d(m, 5)
        If d(m, 2) <> 0 Then
            d(m, 7) = d(m, 6) / d(m, 2)
        Else
            d(m, 7) = 0
        End If
        If m = 1 Then
            d(m, 8) = d(m, 6)
        Else
            d(m, 8) = d(m - 1, 8) + d(m, 6)
        End If
    Next m
    Titulo "Estado de resultados (utilidades)", "Año " & anio & "  ·  Gastos operativos sin compras de mercancía"
    Tabla "Mes|Ventas|Costo mercancía vendida|Utilidad bruta|Gastos operativos|Utilidad neta|Margen neto|Utilidad acumulada", _
          "txt,cop,cop,cop,cop,cop,pct,cop", d, 12, "2,3,4,5,6"
End Sub

Private Sub RepGastos(ByVal d0 As Date, ByVal d1 As Date, ByVal subTxt As String)
    Dim we As Worksheet, ult As Long, r As Long, n As Long, d() As Variant
    Set we = Hoja(SH_EGR)
    ult = UltimaFila(we, EGR_ID, EGR_FIRST)
    ReDim d(1 To Application.Max(1, ult - EGR_FIRST + 1), 1 To 8)
    For r = EGR_FIRST To ult
        If Tv(we.Cells(r, EGR_ESTADO).Value) = "ACTIVO" And EnRango(we.Cells(r, EGR_FECHA).Value, d0, d1) Then
            n = n + 1
            d(n, 1) = we.Cells(r, EGR_ID).Value
            d(n, 2) = we.Cells(r, EGR_FECHA).Value
            d(n, 3) = we.Cells(r, EGR_CONCEPTO).Value
            d(n, 4) = we.Cells(r, EGR_CATEGORIA).Value
            d(n, 5) = we.Cells(r, EGR_TIPO).Value
            d(n, 6) = we.Cells(r, EGR_METODO).Value
            d(n, 7) = we.Cells(r, EGR_RESPONSABLE).Value
            d(n, 8) = Nv(we.Cells(r, EGR_VALOR).Value)
        End If
    Next r
    Titulo "Reporte de gastos y egresos", subTxt
    Tabla "ID|Fecha|Concepto|Categoría|Tipo|Método|Responsable|Valor", "txt,date,txt,txt,txt,txt,txt,cop", d, n, "8"
End Sub

Private Sub RepCartera()
    Dim wc As Worksheet, ult As Long, r As Long, n As Long, d() As Variant
    Set wc = Hoja(SH_CAR)
    ult = UltimaFila(wc, CAR_ID, CAR_FIRST)
    ReDim d(1 To Application.Max(1, ult - CAR_FIRST + 1), 1 To 10)
    For r = CAR_FIRST To ult
        If Tv(wc.Cells(r, CAR_REGESTADO).Value) = "ACTIVA" And Nv(wc.Cells(r, CAR_SALDO).Value) > 0 Then
            n = n + 1
            d(n, 1) = wc.Cells(r, CAR_TIPO).Value
            d(n, 2) = wc.Cells(r, CAR_TERCERO).Value
            d(n, 3) = wc.Cells(r, CAR_DOC).Value
            d(n, 4) = wc.Cells(r, CAR_FECHA).Value
            d(n, 5) = Nv(wc.Cells(r, CAR_VALORINI).Value)
            d(n, 6) = Nv(wc.Cells(r, CAR_ABONOS).Value)
            d(n, 7) = Nv(wc.Cells(r, CAR_SALDO).Value)
            d(n, 8) = wc.Cells(r, CAR_VENCE).Value
            d(n, 9) = Nv(wc.Cells(r, CAR_MORA).Value)
            d(n, 10) = wc.Cells(r, CAR_ESTADO).Value
        End If
    Next r
    Titulo "Reporte de cartera", "Obligaciones con saldo al " & FechaCorta(Date) & "  ·  Por cobrar: " & _
          Moneda(Nv(Hoja(SH_CAR).Range("K5").Value)) & "  ·  Por pagar: " & Moneda(Nv(Hoja(SH_CAR).Range("K8").Value))
    Tabla "Tipo|Cliente / proveedor|Documento|Fecha|Valor inicial|Abonos|Saldo|Vence|Días mora|Estado", _
          "txt,txt,txt,date,cop,cop,cop,date,int,txt", d, n, "5,6,7"
End Sub

Private Sub RepCompras(ByVal d0 As Date, ByVal d1 As Date, ByVal subTxt As String)
    Dim wc As Worksheet, ult As Long, r As Long, n As Long, d() As Variant
    Set wc = Hoja(SH_COM)
    ult = UltimaFila(wc, COM_NUMERO, COM_FIRST)
    ReDim d(1 To Application.Max(1, ult - COM_FIRST + 1), 1 To 9)
    For r = COM_FIRST To ult
        If EnRango(wc.Cells(r, COM_FECHA).Value, d0, d1) Then
            n = n + 1
            d(n, 1) = wc.Cells(r, COM_NUMERO).Value
            d(n, 2) = wc.Cells(r, COM_FECHA).Value
            d(n, 3) = wc.Cells(r, COM_PROVEEDOR).Value
            d(n, 4) = wc.Cells(r, COM_SKU).Value
            d(n, 5) = wc.Cells(r, COM_PRODUCTO).Value
            d(n, 6) = Nv(wc.Cells(r, COM_CANT).Value)
            d(n, 7) = Nv(wc.Cells(r, COM_COSTOREAL).Value)
            d(n, 8) = wc.Cells(r, COM_FORMAPAGO).Value
            d(n, 9) = wc.Cells(r, COM_ESTADOPAGO).Value
        End If
    Next r
    Titulo "Reporte de compras", subTxt
    Tabla "N° compra|Fecha|Proveedor|SKU|Producto|Cantidad|Costo real total|Forma de pago|Estado pago", _
          "txt,date,txt,txt,txt,int,cop,txt,txt", d, n, "6,7"
End Sub

Private Sub RepClientes()
    Dim wc As Worksheet, ult As Long, r As Long, n As Long, d() As Variant
    Set wc = Hoja(SH_CLI)
    ult = UltimaFila(wc, CLI_CODIGO, CLI_FIRST)
    ReDim d(1 To Application.Max(1, ult - CLI_FIRST + 1), 1 To 10)
    For r = CLI_FIRST To ult
        n = n + 1
        d(n, 1) = wc.Cells(r, CLI_CODIGO).Value
        d(n, 2) = wc.Cells(r, CLI_NOMBRE).Value
        d(n, 3) = wc.Cells(r, CLI_IDENT).Value
        d(n, 4) = wc.Cells(r, CLI_TEL).Value
        d(n, 5) = wc.Cells(r, CLI_CIUDAD).Value
        d(n, 6) = Nv(wc.Cells(r, CLI_NFACT).Value)
        d(n, 7) = Nv(wc.Cells(r, CLI_TOTAL).Value)
        d(n, 8) = Nv(wc.Cells(r, CLI_SALDO).Value)
        d(n, 9) = wc.Cells(r, CLI_ULTIMA).Value
        d(n, 10) = wc.Cells(r, CLI_CLASIF).Value
    Next r
    Titulo "Reporte de clientes", "Al " & FechaCorta(Date)
    Tabla "Código|Nombre|Identificación|Teléfono|Ciudad|Facturas|Total compras|Saldo pendiente|Última compra|Clasificación", _
          "txt,txt,txt,txt,txt,int,cop,cop,date,txt", d, n, "6,7,8"
End Sub

Private Sub RepFlujo(ByVal d0 As Date, ByVal d1 As Date, ByVal subTxt As String)
    Dim wi As Worksheet
    Dim we As Worksheet
    Dim wj As Worksheet
    Dim ult As Long
    Dim r As Long
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim d() As Variant
    Dim saldo As Double
    Dim fIni As Date
    Dim tmp As Variant
    Dim v As Double
    Dim f As Variant
    Set wi = Hoja(SH_ING): Set we = Hoja(SH_EGR): Set wj = Hoja(SH_CAJ)
    fIni = CDate(wj.Range(FCAJ_FINI).Value)
    saldo = Nv(wj.Range("C11").Value)
    ReDim d(1 To 20000, 1 To 7)
    ult = UltimaFila(wi, ING_ID, ING_FIRST)
    For r = ING_FIRST To ult
        f = wi.Cells(r, ING_FECHA).Value
        If Tv(wi.Cells(r, ING_ESTADO).Value) = "ACTIVO" And Tv(wi.Cells(r, ING_CUENTA).Value) <> "" And IsDate(f) Then
            v = Nv(wi.Cells(r, ING_VALOR).Value)
            If CDate(f) < d0 And CDate(f) >= fIni Then
                saldo = saldo + v
            ElseIf EnRango(f, d0, d1) And n < 20000 Then
                n = n + 1
                d(n, 1) = f: d(n, 2) = "ENTRADA": d(n, 3) = wi.Cells(r, ING_CONCEPTO).Value
                d(n, 4) = wi.Cells(r, ING_CUENTA).Value: d(n, 5) = v: d(n, 6) = 0
            End If
        End If
    Next r
    ult = UltimaFila(we, EGR_ID, EGR_FIRST)
    For r = EGR_FIRST To ult
        f = we.Cells(r, EGR_FECHA).Value
        If Tv(we.Cells(r, EGR_ESTADO).Value) = "ACTIVO" And IsDate(f) Then
            v = Nv(we.Cells(r, EGR_VALOR).Value)
            If CDate(f) < d0 And CDate(f) >= fIni Then
                saldo = saldo - v
            ElseIf EnRango(f, d0, d1) And n < 20000 Then
                n = n + 1
                d(n, 1) = f: d(n, 2) = "SALIDA": d(n, 3) = we.Cells(r, EGR_CONCEPTO).Value
                d(n, 4) = we.Cells(r, EGR_CUENTA).Value: d(n, 5) = 0: d(n, 6) = v
            End If
        End If
    Next r
    ' ordenar por fecha (inserción)
    For i = 2 To n
        For j = i To 2 Step -1
            If Nv(d(j - 1, 1)) > Nv(d(j, 1)) Then
                For c = 1 To 6
                    tmp = d(j, c)
                    d(j, c) = d(j - 1, c)
                    d(j - 1, c) = tmp
                Next c
            Else
                j = 2
            End If
        Next j
    Next i
    Titulo "Flujo de caja", subTxt & "  ·  Saldo inicial del período: " & Moneda(saldo) & "  (traslados entre cuentas no afectan el total)"
    For i = 1 To n
        saldo = saldo + Nv(d(i, 5)) - Nv(d(i, 6))
        d(i, 7) = saldo
    Next i
    Tabla "Fecha|Tipo|Concepto|Cuenta|Entrada|Salida|Saldo acumulado", "date,txt,txt,txt,cop,cop,cop", d, n, "5,6"
End Sub

'-------------------------------------------------------------- exportar / imprimir
Public Sub ReportePDF()
    Dim ws As Worksheet, ruta As String
    Set ws = Hoja(SH_REP)
    If Tv(ws.Cells(FREP_OUT, 2).Value) = "" Then Msg "Primero genere un reporte.", vbExclamation: Exit Sub
    On Error GoTo Fallo
    ruta = Carpeta("REPORTES") & Application.PathSeparator & NombreArchivo(Tv(ws.Range(FREP_TIPO).Value) & " " & Format(Now, "yyyy-mm-dd hhmm")) & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ruta, Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=True
    Msg "Reporte exportado:" & vbLf & ruta, vbInformation
    Exit Sub
Fallo:
    Msg "No se pudo exportar a PDF: " & Err.Description, vbExclamation
End Sub

Public Sub ReporteExcel()
    Dim ws As Worksheet, wb As Workbook, ruta As String, ult As Long
    Set ws = Hoja(SH_REP)
    If Tv(ws.Cells(FREP_OUT, 2).Value) = "" Then Msg "Primero genere un reporte.", vbExclamation: Exit Sub
    On Error GoTo Fallo
    ult = ws.UsedRange.Row + ws.UsedRange.Rows.Count
    ruta = Carpeta("REPORTES") & Application.PathSeparator & NombreArchivo(Tv(ws.Range(FREP_TIPO).Value) & " " & Format(Now, "yyyy-mm-dd hhmm")) & ".xlsx"
    Application.ScreenUpdating = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    ws.Range(ws.Cells(FREP_OUT, 2), ws.Cells(ult, 16)).Copy
    wb.Worksheets(1).Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    wb.Worksheets(1).Range("A1").PasteSpecial xlPasteFormats
    wb.Worksheets(1).Columns("A:O").AutoFit
    Application.CutCopyMode = False
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=ruta, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close False
    Application.ScreenUpdating = True
    Msg "Reporte exportado a Excel:" & vbLf & ruta, vbInformation
    Exit Sub
Fallo:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Msg "No se pudo exportar a Excel: " & Err.Description, vbExclamation
End Sub

Public Sub ImprimirReporte()
    Dim ws As Worksheet
    Set ws = Hoja(SH_REP)
    If Tv(ws.Cells(FREP_OUT, 2).Value) = "" Then Msg "Primero genere un reporte.", vbExclamation: Exit Sub
    On Error Resume Next
    ws.PrintOut Copies:=1
End Sub
