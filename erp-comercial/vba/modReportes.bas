Attribute VB_Name = "modReportes"
'==============================================================================
' modReportes - 13 reportes imprimibles exportables a PDF y a Excel (.xlsx).
' Se guardan en <OneDrive>\SISTEMA EMPRESARIAL\REPORTES.
' Los botones de la hoja REPORTES se llaman btnRep_NN_PDF / btnRep_NN_XLS.
'==============================================================================
Option Explicit

Public Sub ReporteClick()
    Dim partes As Variant
    On Error GoTo FALLO
    partes = Split(Application.Caller, "_")      ' btnRep_05_PDF
    GenerarReporte CLng(partes(1)), CStr(partes(2))
    Exit Sub
FALLO:
    MsgBox "Use los botones de la hoja REPORTES.", vbExclamation, APP_NOMBRE
End Sub

Public Sub PDFDashboard()
    Dim ruta As String
    If Not RequierePermiso("VER_DASHBOARD") Then Exit Sub
    Application.Calculate
    ruta = modRespaldo.ExportarPDF(WS("DASHBOARD"), "REPORTES", "DASHBOARD_" & Format$(Now, "yyyymmdd_hhnn"), Date)
    If ruta <> "" Then MsgBox "PDF generado:" & vbCrLf & ruta, vbInformation, APP_NOMBRE
End Sub

Public Sub ActualizarTodo()
    Application.CalculateFull
End Sub

Public Sub GenerarReporte(ByVal n As Long, ByVal formato As String)
    Dim desde As Date, hasta As Date, titulo As String, wb As Workbook, dst As Worksheet, fila As Long
    Dim raiz As String, ruta As String, nombreArchivo As String, verCostos As Boolean

    If Not RequierePermiso("REPORTES") Then Exit Sub
    desde = CDate(Cfg("rep_Desde"))
    hasta = CDate(Cfg("rep_Hasta"))
    If hasta < desde Then MsgBox "La fecha 'hasta' es menor que 'desde'.", vbExclamation, APP_NOMBRE: Exit Sub
    verCostos = TienePermiso("VER_COSTOS")
    raiz = modRespaldo.CarpetaSistema(True)
    If raiz = "" Then Exit Sub

    On Error GoTo FALLO
    InicioMacro
    Application.Calculation = xlCalculationAutomatic
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set dst = wb.Worksheets(1)
    dst.Cells.Font.Name = "Arial"
    dst.Cells.Font.Size = 9
    Select Case n
        Case 1
            titulo = "Ventas diarias " & Format$(hasta, "dd/mm/yyyy")
            fila = Encabezado(dst, titulo, hasta, hasta)
            Listar WS("BD_FACTURAS"), dst, fila, ColsFacturas(verCostos), 2, hasta, hasta, 0, "", True
        Case 2
            titulo = "Ventas mensuales " & Format$(hasta, "mm/yyyy")
            fila = Encabezado(dst, titulo, DateSerial(Year(hasta), Month(hasta), 1), DateSerial(Year(hasta), Month(hasta) + 1, 0))
            Listar WS("BD_FACTURAS"), dst, fila, ColsFacturas(verCostos), 2, DateSerial(Year(hasta), Month(hasta), 1), _
                   DateSerial(Year(hasta), Month(hasta) + 1, 0), 0, "", True
        Case 3, 13
            titulo = IIf(n = 3, "Ventas anuales ", "Utilidades ") & Year(hasta)
            If Not verCostos And n = 13 Then Err.Raise vbObjectError + 1, , "Requiere permiso VER_COSTOS"
            WS("UTILIDADES").Range("C4").Value = Year(hasta)
            Application.Calculate
            fila = Encabezado(dst, titulo, DateSerial(Year(hasta), 1, 1), DateSerial(Year(hasta), 12, 31))
            CopiarBloque WS("UTILIDADES").Range("B6:N19"), dst, fila
        Case 4
            titulo = "Ventas por canal"
            WS("CANALES DE VENTA").Range("C4").Value = desde
            WS("CANALES DE VENTA").Range("D4").Value = hasta
            Application.Calculate
            fila = Encabezado(dst, titulo, desde, hasta)
            CopiarBloque WS("CANALES DE VENTA").Range("B6:L17"), dst, fila
        Case 5
            titulo = "Ventas por vendedor"
            fila = Encabezado(dst, titulo, desde, hasta)
            PorVendedor dst, fila, desde, hasta, verCostos
        Case 6
            titulo = "Inventario al " & Format$(Date, "dd/mm/yyyy")
            fila = Encabezado(dst, titulo, Date, Date)
            If verCostos Then
                Listar WS("INVENTARIO"), dst, fila, Array(1, 2, 3, 4, 5, 6, 7, 8, 12, 13, 14, 15), 0, 0, 0, 0, "", False
            Else
                Listar WS("INVENTARIO"), dst, fila, Array(1, 2, 3, 4, 5, 8, 12, 13, 15), 0, 0, 0, 0, "", False
            End If
        Case 7
            titulo = "Compras"
            fila = Encabezado(dst, titulo, desde, hasta)
            Listar WS("BD_COMPRAS"), dst, fila, Array(1, 2, 4, 5, 9, 10, 11, 12, 13, 14), 2, desde, hasta, 0, "", False
        Case 8
            titulo = "Proveedores"
            fila = Encabezado(dst, titulo, Date, Date)
            Listar WS("PROVEEDORES"), dst, fila, Array(1, 2, 3, 4, 5, 9, 10, 15, 16, 17, 18, 19), 0, 0, 0, 0, "", False
        Case 9
            titulo = "Clientes"
            fila = Encabezado(dst, titulo, Date, Date)
            Listar WS("CLIENTES"), dst, fila, Array(1, 2, 4, 5, 9, 10, 11, 12, 13, 14, 15, 17), 0, 0, 0, 0, "", False, _
                   FILA_DATOS_CLI
        Case 10
            titulo = "Gastos y egresos"
            fila = Encabezado(dst, titulo, desde, hasta)
            Listar WS("EGRESOS"), dst, fila, Array(1, 2, 3, 4, 5, 6, 7, 8, 10), 2, desde, hasta, 10, "VIGENTE", False
        Case 11
            titulo = "Cartera con saldo pendiente"
            fila = Encabezado(dst, titulo, Date, Date)
            Listar WS("CARTERA"), dst, fila, Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), 0, 0, 0, -9, "", False
        Case 12
            titulo = "Flujo de caja"
            WS("CAJA Y BANCOS").Range("C4").Value = desde
            WS("CAJA Y BANCOS").Range("D4").Value = hasta
            Application.Calculate
            fila = Encabezado(dst, titulo, desde, hasta)
            CopiarBloque WS("CAJA Y BANCOS").Range("B6:H12"), dst, fila
            CopiarBloque WS("CAJA Y BANCOS").Range("B15:C23"), dst, dst.Cells(dst.Rows.Count, 2).End(xlUp).Row + 2
        Case Else
            Err.Raise vbObjectError + 2, , "Reporte no definido"
    End Select
    dst.Columns.AutoFit
    With dst.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterFooter = "Página &P de &N"
        .LeftFooter = APP_NOMBRE & " · " & UsuarioActual()
    End With
    nombreArchivo = "REP_" & Format$(n, "00") & "_" & Replace(SinTildes(Split(titulo, " ")(0) & "_" & _
                    Split(titulo & " x", " ")(1)), "/", "-") & "_" & Format$(Now, "yyyymmdd_hhnnss")
    ruta = raiz & "\REPORTES\" & nombreArchivo
    Application.DisplayAlerts = False
    If UCase$(formato) = "PDF" Then
        ruta = ruta & ".pdf"
        dst.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ruta, OpenAfterPublish:=False
        wb.Close SaveChanges:=False
    Else
        ruta = ruta & ".xlsx"
        wb.SaveAs Filename:=ruta, FileFormat:=51
        wb.Close SaveChanges:=False
    End If
    Application.DisplayAlerts = True
    FinMacro
    ThisWorkbook.Activate
    Bitacora "REPORTES", "GENERAR " & UCase$(formato), CStr(n), titulo
    If MsgBox("Reporte generado:" & vbCrLf & ruta & vbCrLf & vbCrLf & "¿Abrirlo?", vbQuestion + vbYesNo, APP_NOMBRE) = vbYes Then
        ThisWorkbook.FollowHyperlink ruta
    End If
    Exit Sub
FALLO:
    Application.DisplayAlerts = True
    FinMacro
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    MsgBox "No se pudo generar el reporte: " & Err.Description, vbExclamation, APP_NOMBRE
End Sub

Private Function ColsFacturas(ByVal verCostos As Boolean) As Variant
    If verCostos Then
        ColsFacturas = Array(1, 2, 4, 6, 7, 8, 9, 12, 13, 14, 16, 17)
    Else
        ColsFacturas = Array(1, 2, 4, 6, 7, 8, 9, 12, 13, 14, 17)
    End If
End Function

Private Function Encabezado(ByVal dst As Worksheet, ByVal titulo As String, ByVal desde As Date, ByVal hasta As Date) As Long
    dst.Range("A1").Value = Cfg("cfg_Empresa")
    dst.Range("A1").Font.Size = 14
    dst.Range("A1").Font.Bold = True
    dst.Range("A2").Value = UCase$(titulo)
    dst.Range("A2").Font.Bold = True
    dst.Range("A3").Value = "Período: " & Format$(desde, "dd/mm/yyyy") & " a " & Format$(hasta, "dd/mm/yyyy") & _
                            "   ·   Generado: " & Format$(Now, "dd/mm/yyyy hh:mm") & " por " & UsuarioActual()
    dst.Range("A3").Font.Color = RGB(89, 89, 89)
    Encabezado = 5
End Function

' Copia valores y formatos de un bloque calculado
Private Sub CopiarBloque(ByVal src As Range, ByVal dst As Worksheet, ByVal fila As Long)
    src.Copy
    dst.Cells(fila, 1).PasteSpecial xlPasteValuesAndNumberFormats
    dst.Cells(fila, 1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub

' Lista filas de una tabla con columnas elegidas, filtro de fecha y de estado.
' colEstado < 0 => filtrar saldo > 0 en la columna abs(colEstado). Suma columnas con formato de moneda.
Private Sub Listar(ByVal src As Worksheet, ByVal dst As Worksheet, ByVal fila As Long, ByVal cols As Variant, _
                   ByVal colFecha As Long, ByVal desde As Date, ByVal hasta As Date, ByVal colEstado As Long, _
                   ByVal estado As String, ByVal soloNoAnulado As Boolean, Optional ByVal fila0 As Long = FILA_DATOS)
    Dim hdr As Long, ult As Long, r As Long, j As Long, out As Long, incluir As Boolean, datos As Variant
    Dim totales() As Double, esMoneda() As Boolean
    hdr = fila0 - 1
    ult = UltimaFila(src, 1)
    ReDim totales(LBound(cols) To UBound(cols))
    ReDim esMoneda(LBound(cols) To UBound(cols))
    For j = LBound(cols) To UBound(cols)
        dst.Cells(fila, j - LBound(cols) + 1).Value = src.Cells(hdr, cols(j)).Value
        esMoneda(j) = (InStr(src.Cells(fila0, cols(j)).NumberFormat, "$") > 0)
    Next j
    With dst.Range(dst.Cells(fila, 1), dst.Cells(fila, UBound(cols) - LBound(cols) + 1))
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(68, 84, 106)
        .WrapText = True
    End With
    out = fila
    If ult < fila0 Then Exit Sub
    datos = src.Range(src.Cells(fila0, 1), src.Cells(ult, src.Cells(hdr, src.Columns.Count).End(xlToLeft).Column)).Value
    For r = 1 To UBound(datos, 1)
        incluir = Not Vacio(datos(r, 1))
        If incluir And colFecha > 0 Then
            If Not IsDate(datos(r, colFecha)) Then
                incluir = False
            ElseIf CDate(datos(r, colFecha)) < desde Or CDate(datos(r, colFecha)) > hasta Then
                incluir = False
            End If
        End If
        If incluir And colEstado > 0 Then incluir = (CStr(datos(r, colEstado)) = estado)
        If incluir And colEstado < 0 Then incluir = (NumV(datos(r, -colEstado)) > 0)
        If incluir Then
            out = out + 1
            For j = LBound(cols) To UBound(cols)
                dst.Cells(out, j - LBound(cols) + 1).Value = datos(r, cols(j))
                dst.Cells(out, j - LBound(cols) + 1).NumberFormat = src.Cells(fila0, cols(j)).NumberFormat
                If esMoneda(j) And IsNumeric(datos(r, cols(j))) Then
                    If soloNoAnulado Then
                        If CStr(datos(r, 17)) = "VIGENTE" Then totales(j) = totales(j) + NumV(datos(r, cols(j)))
                    Else
                        totales(j) = totales(j) + NumV(datos(r, cols(j)))
                    End If
                End If
            Next j
        End If
    Next r
    out = out + 1
    dst.Cells(out, 1).Value = "TOTAL" & IIf(soloNoAnulado, " (solo vigentes)", "") & " · " & (out - fila - 1) & " registros"
    dst.Rows(out).Font.Bold = True
    For j = LBound(cols) To UBound(cols)
        If esMoneda(j) Then
            dst.Cells(out, j - LBound(cols) + 1).Value = totales(j)
            dst.Cells(out, j - LBound(cols) + 1).NumberFormat = """$"" #,##0"
        End If
    Next j
End Sub

Private Sub PorVendedor(ByVal dst As Worksheet, ByVal fila As Long, ByVal desde As Date, ByVal hasta As Date, _
                        ByVal verCostos As Boolean)
    Dim hV As Worksheet, r As Long, out As Long, wf As WorksheetFunction, cod As String, ventas As Double, ped As Double
    Set hV = WS("VENDEDORES")
    Set wf = Application.WorksheetFunction
    dst.Cells(fila, 1).Resize(1, 7).Value = Array("Código", "Vendedor", "Pedidos", "Ventas", "Ticket promedio", _
                                                  IIf(verCostos, "Utilidad", ""), "Comisión estimada")
    dst.Rows(fila).Font.Bold = True
    out = fila
    For r = FILA_DATOS To UltimaFila(hV, 1)
        cod = CStr(hV.Cells(r, 1).Value)
        If cod <> "" Then
            ventas = wf.SumIfs(NR("fac_Total"), NR("fac_Vend"), cod, NR("fac_Estado"), "VIGENTE", _
                               NR("fac_Fecha"), ">=" & CLng(desde), NR("fac_Fecha"), "<=" & CLng(hasta))
            ped = wf.CountIfs(NR("fac_Vend"), cod, NR("fac_Estado"), "VIGENTE", _
                              NR("fac_Fecha"), ">=" & CLng(desde), NR("fac_Fecha"), "<=" & CLng(hasta))
            out = out + 1
            dst.Cells(out, 1).Value = cod
            dst.Cells(out, 2).Value = hV.Cells(r, 2).Value
            dst.Cells(out, 3).Value = ped
            dst.Cells(out, 4).Value = ventas
            dst.Cells(out, 5).Value = IIf(ped > 0, ventas / IIf(ped > 0, ped, 1), 0)
            If verCostos Then dst.Cells(out, 6).Value = wf.SumIfs(NR("fac_Util"), NR("fac_Vend"), cod, _
                NR("fac_Estado"), "VIGENTE", NR("fac_Fecha"), ">=" & CLng(desde), NR("fac_Fecha"), "<=" & CLng(hasta))
            dst.Cells(out, 7).Value = Round(ventas * NumV(hV.Cells(r, 8).Value), 0)
            dst.Range(dst.Cells(out, 4), dst.Cells(out, 7)).NumberFormat = """$"" #,##0"
        End If
    Next r
End Sub
