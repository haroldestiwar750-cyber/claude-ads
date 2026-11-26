Attribute VB_Name = "modSistema"
'==============================================================
' modSistema - INTERFAZ, SEGURIDAD Y MANTENIMIENTO
'  * CrearInterfaz: dibuja todos los botones (la ejecuta el instalador)
'  * Copias de seguridad y recuperación de datos
'  * Claves, numeración, logo, limpieza de datos de demostración
'==============================================================
Option Explicit

Private Function ColorRGB(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
    ColorRGB = RGB(r, g, b)
End Function

Private Sub Boton(ByVal ws As Worksheet, ByVal texto As String, ByVal macro As String, ByVal celda As String, _
                  ByVal color As Long, Optional ByVal ancho As Single = 150, Optional ByVal alto As Single = 30, _
                  Optional ByVal dx As Single = 0, Optional ByVal dy As Single = 0)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range(celda).Left + dx, ws.Range(celda).Top + dy, ancho, alto)
    s.Name = "btn_" & Replace(macro, " ", "")
    s.Fill.ForeColor.RGB = color
    s.Line.Visible = msoFalse
    On Error Resume Next
    s.Adjustments.Item(1) = 0.25
    s.Shadow.Visible = msoFalse
    On Error GoTo 0
    With s.TextFrame2
        .TextRange.Text = texto
        .TextRange.Font.Size = 10
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 2: .MarginRight = 2
    End With
    s.Placement = xlFreeFloating
    s.OnAction = macro
End Sub

Private Sub BorrarBotones(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left(ws.Shapes(i).Name, 4) = "btn_" Then ws.Shapes(i).Delete
    Next i
End Sub

' Crea (o recrea) todos los botones del sistema. Puede ejecutarse cuantas veces se quiera.
Public Sub CrearInterfaz()
    Dim ws As Worksheet, azul As Long, verde As Long, rojo As Long, gris As Long, naranja As Long, oscuro As Long, morado As Long
    azul = ColorRGB(29, 78, 216): verde = ColorRGB(21, 128, 61): rojo = ColorRGB(185, 28, 28): gris = ColorRGB(100, 116, 139)
    naranja = ColorRGB(217, 119, 6): oscuro = ColorRGB(30, 41, 59): morado = ColorRGB(109, 40, 217)
    Application.ScreenUpdating = False
    DesprotegerTodo
    For Each ws In ThisWorkbook.Worksheets
        BorrarBotones ws
    Next ws
    ' FACTURACIÓN (a la derecha del formato imprimible)
    Set ws = Hoja(SH_FAC)
    Boton ws, "NUEVA FACTURA", "NuevaFactura", "K2", azul, 150, 30
    Boton ws, "GUARDAR FACTURA", "GuardarFactura", "K2", verde, 150, 30, 158
    Boton ws, "IMPRIMIR", "ImprimirFactura", "K4", oscuro, 150, 30
    Boton ws, "LIMPIAR", "LimpiarFactura", "K4", gris, 150, 30, 158
    Boton ws, "BUSCAR FACTURA", "BuscarFactura", "K6", naranja, 150, 30
    Boton ws, "ANULAR FACTURA", "AnularFactura", "K6", rojo, 150, 30, 158
    Boton ws, "GENERAR GUÍA", "GenerarGuiaDeFacturaActual", "K8", morado, 150, 30
    Boton ws, "NUEVO CLIENTE", "NuevoCliente", "K8", gris, 150, 30, 158
    ' GUÍAS
    Set ws = Hoja(SH_GUI)
    Boton ws, "NUEVA GUÍA", "NuevaGuia", "K5", azul
    Boton ws, "GUARDAR GUÍA", "GuardarGuia", "K5", verde, 150, 30, 158
    Boton ws, "IMPRIMIR GUÍA", "ImprimirGuia", "K8", oscuro
    Boton ws, "LIMPIAR", "LimpiarGuia", "K8", gris, 150, 30, 158
    Boton ws, "BUSCAR GUÍA", "BuscarGuia", "K11", naranja
    Boton ws, "DESDE FACTURA", "GuiaDesdeFactura", "K11", morado, 150, 30, 158
    ' INVENTARIO
    Set ws = Hoja(SH_INV)
    Boton ws, "NUEVO PRODUCTO", "NuevoProducto", "Z9", azul, 170
    Boton ws, "DEVOLUCIÓN CLIENTE", "RegistrarDevolucion", "Z9", naranja, 170, 30, 0, 38
    Boton ws, "AJUSTE (CLAVE)", "AjusteInventario", "Z9", rojo, 170, 30, 0, 76
    Boton ws, "VER MOVIMIENTOS", "VerMovimientosSKU", "Z9", oscuro, 170, 30, 0, 114
    Boton ws, "QUITAR FILTROS", "QuitarFiltrosHoja", "Z9", gris, 170, 30, 0, 152
    ' COMPRAS
    Set ws = Hoja(SH_COM)
    Boton ws, "REGISTRAR COMPRA", "RegistrarCompra", "L5", verde, 170, 34
    Boton ws, "LIMPIAR", "LimpiarCompra", "L5", gris, 170, 30, 0, 42
    Boton ws, "NUEVO PROVEEDOR", "NuevoProveedor", "L5", azul, 170, 30, 0, 80
    ' PROVEEDORES / CLIENTES
    Boton Hoja(SH_PRO), "NUEVO PROVEEDOR", "NuevoProveedor", "D3", azul, 170, 26
    Boton Hoja(SH_CLI), "NUEVO CLIENTE", "NuevoCliente", "J3", azul, 160, 26
    Boton Hoja(SH_CLI), "HISTORIAL CLIENTE", "VerHistorialCliente", "J3", oscuro, 160, 26, 168
    Boton Hoja(SH_CLI), "QUITAR FILTROS", "QuitarFiltrosHoja", "J3", gris, 140, 26, 336
    Boton Hoja(SH_HFA), "QUITAR FILTROS", "QuitarFiltrosHoja", "E2", gris, 140, 22
    Boton Hoja(SH_MOV), "QUITAR FILTROS", "QuitarFiltrosHoja", "E2", gris, 140, 22
    ' INGRESOS / EGRESOS / GASTOS FIJOS / CARTERA
    Boton Hoja(SH_ING), "REGISTRAR INGRESO", "RegistrarIngresoManual", "B10", verde, 170, 30
    Boton Hoja(SH_ING), "ANULAR SELECCIONADO", "AnularIngreso", "B10", rojo, 170, 30, 178
    Boton Hoja(SH_EGR), "REGISTRAR EGRESO", "RegistrarEgresoManual", "B11", verde, 170, 30
    Boton Hoja(SH_EGR), "ANULAR SELECCIONADO", "AnularEgreso", "B11", rojo, 170, 30, 178
    Boton Hoja(SH_GFI), "PAGAR GASTO SELECCIONADO", "PagarGastoFijo", "H3", verde, 220, 28
    Boton Hoja(SH_CAR), "REGISTRAR ABONO", "RegistrarAbono", "B10", verde, 170, 30
    Boton Hoja(SH_CAR), "VER SOLO VENCIDA", "FiltrarCarteraVencida", "B10", rojo, 150, 30, 178
    Boton Hoja(SH_CAR), "QUITAR FILTROS", "QuitarFiltrosHoja", "B10", gris, 130, 30, 336
    ' CAJA
    Set ws = Hoja(SH_CAJ)
    Boton ws, "DIARIO", "CierreDiario", "D14", azul, 95, 26
    Boton ws, "SEMANAL", "CierreSemanal", "D14", azul, 95, 26, 100
    Boton ws, "MENSUAL", "CierreMensual", "D14", azul, 95, 26, 200
    Boton ws, "GUARDAR CIERRE", "GuardarCierre", "D16", verde, 140, 26
    Boton ws, "PDF CIERRE", "PDFCierre", "D16", oscuro, 120, 26, 150
    ' REPORTES
    Set ws = Hoja(SH_REP)
    Boton ws, "GENERAR REPORTE", "GenerarReporte", "F4", verde, 160, 30
    Boton ws, "EXPORTAR PDF", "ReportePDF", "F4", oscuro, 140, 30, 168
    Boton ws, "EXPORTAR EXCEL", "ReporteExcel", "F4", verde, 140, 30, 316
    Boton ws, "IMPRIMIR", "ImprimirReporte", "F4", gris, 110, 30, 464
    ' CONFIGURACIÓN
    Set ws = Hoja(SH_CFG)
    Boton ws, "CARGAR LOGO", "CargarLogo", "A44", azul, 200, 28
    Boton ws, "CAMBIAR NUMERACIÓN", "CambiarNumeracion", "A44", naranja, 200, 28, 0, 34
    Boton ws, "CAMBIAR CLAVES", "CambiarClaves", "A44", rojo, 200, 28, 0, 68
    Boton ws, "COPIA DE SEGURIDAD", "CopiaSeguridad", "B44", verde, 200, 28
    Boton ws, "RESTAURAR COPIA", "RestaurarCopia", "B44", oscuro, 200, 28, 0, 34
    Boton ws, "EJECUTAR PRUEBAS", "PruebasAutomaticas", "B44", morado, 200, 28, 0, 68
    Boton ws, "BORRAR DATOS DEMO", "LimpiarDatosDemo", "B44", rojo, 200, 28, 0, 102
    ' INICIO
    Set ws = Hoja(SH_INI)
    Boton ws, "ACTUALIZAR", "ActualizarPanel", "O3", azul, 105, 24, 0, 4
    Boton ws, "HOY", "IrHoy", "O3", gris, 70, 24, 110, 4
    Boton ws, "RESPALDO", "CopiaSeguridad", "O3", verde, 100, 24, 185, 4
    ProtegerTodo
    Application.ScreenUpdating = True
    Hoja(SH_INI).Activate
End Sub

'-------------------------------------------------------------- panel
Public Sub ActualizarPanel()
    On Error Resume Next
    Application.CalculateFull
    Msg "Indicadores actualizados.", vbInformation
End Sub

Public Sub IrHoy()
    Dim ws As Worksheet
    Set ws = Hoja(SH_INI)
    Iniciar
    ws.Range(FINI_ANIO).Value = Year(Date)
    ws.Range(FINI_MES).Value = Hoja(SH_CFG).Range(LST_MESES & (CFG_LIST_FIRST + Month(Date) - 1)).Value
    ws.Range(FINI_DIA).Formula = "=TODAY()"
    Terminar
End Sub

'-------------------------------------------------------------- copias de seguridad
Public Sub CopiaSeguridad()
    CopiaSeguridadInterna False, "manual"
End Sub

Public Function CopiaSeguridadInterna(ByVal silencioso As Boolean, ByVal motivo As String) As String
    Dim ruta As String, ext As String
    If gTest Then Exit Function
    On Error GoTo Fallo
    ext = Mid(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "."))
    ruta = Carpeta("COPIAS DE SEGURIDAD") & Application.PathSeparator & "Respaldo " & Format(Now, "yyyy-mm-dd hhmmss") & ext
    ThisWorkbook.SaveCopyAs ruta
    Auditar "SISTEMA", "COPIA DE SEGURIDAD", "", motivo & " -> " & ruta
    CopiaSeguridadInterna = ruta
    If Not silencioso Then Msg "Copia de seguridad creada:" & vbLf & ruta, vbInformation
    Exit Function
Fallo:
    If Not silencioso Then Msg "No se pudo crear la copia: " & Err.Description, vbExclamation
End Function

' Copia de seguridad automática diaria (al cerrar)
Public Sub CopiaAutomaticaDiaria()
    Dim f As String
    On Error Resume Next
    f = Dir(Carpeta("COPIAS DE SEGURIDAD") & Application.PathSeparator & "Respaldo " & Format(Date, "yyyy-mm-dd") & "*")
    If f = "" Then CopiaSeguridadInterna True, "automática diaria"
End Sub

' Recupera los DATOS de una copia de seguridad dentro de este mismo archivo.
Public Sub RestaurarCopia()
    Dim archivo As Variant, src As Workbook, n As Long
    On Error Resume Next
    ChDir Carpeta("COPIAS DE SEGURIDAD")
    On Error GoTo 0
    archivo = Application.GetOpenFilename("Copias del sistema (*.xlsm;*.xlsx),*.xlsm;*.xlsx", , "Seleccione la copia de seguridad a restaurar")
    If VarType(archivo) = vbBoolean Then Exit Sub
    If Msg("Se REEMPLAZARÁN todos los datos actuales (ventas, inventario, clientes, finanzas, configuración)" & vbLf & _
           "por los de la copia:" & vbLf & archivo & vbLf & vbLf & _
           "Antes se creará automáticamente una copia del estado actual. ¿Continuar?", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    If Not Autorizar("Restaurar una copia de seguridad") Then Exit Sub
    CopiaSeguridadInterna True, "antes de restaurar"
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set src = Workbooks.Open(Filename:=archivo, ReadOnly:=True, UpdateLinks:=0)
    Iniciar
    On Error GoTo Fallo
    n = CopiarDatosDesde(src)
    src.Close False
    Auditar "SISTEMA", "RESTAURAR COPIA", "", CStr(archivo)
    Terminar
    Msg "Restauración terminada (" & n & " bloques de datos). Guarde el archivo para conservar los cambios.", vbInformation
    Exit Sub
Fallo:
    Dim e As String
    e = Err.Description
    On Error Resume Next
    src.Close False
    Terminar
    Msg "Error al restaurar: " & e & vbLf & "Puede recuperar el estado anterior desde la copia creada justo antes.", vbCritical
End Sub

Private Sub CopiarBloque(ByVal src As Workbook, ByVal hojaN As String, ByVal r0 As Long, ByVal r1 As Long, _
                         ByVal c0 As Long, ByVal c1 As Long, ByVal soloValores As Boolean)
    Dim s As Worksheet, d As Worksheet, r As Long, c As Long, v As Variant
    Set s = src.Worksheets(hojaN)
    Set d = ThisWorkbook.Worksheets(hojaN)
    For c = c0 To c1
        If Left(CStr(d.Cells(r0, c).Formula), 1) <> "=" Or Not soloValores Then
            v = s.Range(s.Cells(r0, c), s.Cells(r1, c)).Value
            If Left(CStr(d.Cells(r0, c).Formula), 1) <> "=" Then d.Range(d.Cells(r0, c), d.Cells(r1, c)).Value = v
        End If
    Next c
End Sub

Private Function CopiarDatosDesde(ByVal src As Workbook) As Long
    Dim ult As Long, n As Long
    ' tablas transaccionales (solo valores)
    Dim spec As Variant, it As Variant
    spec = Array(Array(SH_BD, BD_FIRST, BD_NCOLS, 1), Array(SH_HFA, HFA_FIRST, HFA_NCOLS, 1), _
                 Array(SH_HGU, HGU_FIRST, HGU_NCOLS, 1), Array(SH_MOV, MOV_FIRST, MOV_NCOLS, 1), _
                 Array(SH_AUD, AUD_FIRST, AUD_NCOLS, 1), Array(SH_COM, COM_FIRST, COM_NCOLS, 1), _
                 Array(SH_ING, ING_FIRST, ING_NCOLS, 1), Array(SH_EGR, EGR_FIRST, EGR_NCOLS, 1), _
                 Array(SH_CAR, CAR_FIRST, CAR_NCOLS, 1), Array(SH_CAJ, CIE_FIRST, CIE_NCOLS, CIE_COL0))
    For Each it In spec
        ult = Application.Max(UltimaFila(src.Worksheets(it(0)), it(3), it(1)), UltimaFila(ThisWorkbook.Worksheets(it(0)), it(3), it(1)))
        If ult >= it(1) Then
            LimpiarFilasDesde ThisWorkbook.Worksheets(it(0)), it(3), it(1), it(3), it(2)
            CopiarBloque src, it(0), it(1), ult, it(3), it(3) + it(2) - 1, True
            n = n + 1
        End If
    Next it
    ' maestros (capturas, sin tocar fórmulas)
    CopiarBloque src, SH_INV, INV_FIRST, INV_LAST, 1, INV_NCOLS, True
    CopiarBloque src, SH_CLI, CLI_FIRST, CLI_LAST, 1, CLI_NCOLS, True
    CopiarBloque src, SH_PRO, PRO_FIRST, PRO_LAST, 1, PRO_NCOLS, True
    CopiarBloque src, SH_GFI, GFI_FIRST, GFI_LAST, 1, GFI_NCOLS, True
    CopiarBloque src, SH_CFG, 4, 44, 2, 2, True
    CopiarBloque src, SH_CFG, CFG_LIST_FIRST, CFG_LIST_FIRST + CFG_LIST_N - 1, 4, 20, True
    CopiarBloque src, SH_CAJ, 3, 3, 3, 3, True
    CopiarBloque src, SH_CAJ, 6, 10, 3, 3, True
    CopiarBloque src, SH_CAJ, 6, 205, 18, 23, True
    CopiarDatosDesde = n + 9
End Function

'-------------------------------------------------------------- claves y numeración
Public Sub CambiarClaves()
    Dim op As String, c1 As String, c2 As String
    If Not Autorizar("Cambiar claves del sistema") Then Exit Sub
    op = Pedir("¿Qué clave desea cambiar?" & vbLf & "1 = Clave de autorización (anulaciones, ajustes)" & vbLf & _
               "2 = Clave de protección de hojas", "Cambiar claves", "1")
    If op <> "1" And op <> "2" Then Exit Sub
    c1 = Pedir("Escriba la NUEVA clave (mínimo 4 caracteres):", "Cambiar claves")
    If Len(c1) < 4 Then Msg "Clave demasiado corta. No se realizó el cambio.", vbExclamation: Exit Sub
    c2 = Pedir("Confirme la nueva clave:", "Cambiar claves")
    If c1 <> c2 Then Msg "Las claves no coinciden. No se realizó el cambio.", vbExclamation: Exit Sub
    Iniciar     ' desprotege con la clave actual
    If op = "1" Then
        Hoja(SH_CFG).Range(CFG_CLAVE).Value = c1
    Else
        Hoja(SH_CFG).Range(CFG_CLAVEHOJAS).Value = c1
    End If
    Auditar "SEGURIDAD", "CAMBIO DE CLAVE", IIf(op = "1", "AUTORIZACIÓN", "HOJAS"), ""
    Terminar    ' protege con la clave nueva
    Msg "Clave actualizada. Guárdela en un lugar seguro.", vbInformation
End Sub

Private Function MaxConsecutivo(ByVal ws As Worksheet, ByVal col As Long, ByVal primera As Long, ByVal prefijo As String) As Long
    Dim s As String
    s = SiguienteID(ws, col, primera, prefijo, 1)
    MaxConsecutivo = CLng(Val(Mid(s, Len(prefijo) + 1))) - 1
End Function

Public Sub CambiarNumeracion()
    Dim op As String, s As String, nuevo As Long, minimo As Long, celda As String, nombre As String
    If Not Autorizar("Cambiar la numeración de documentos") Then Exit Sub
    op = Pedir("Numeración a cambiar:" & vbLf & "1 = Facturas" & vbLf & "2 = Guías" & vbLf & "3 = Compras", "Numeración", "1")
    Select Case op
        Case "1": celda = CFG_SIGFACTURA: nombre = "facturas": minimo = MaxConsecutivo(Hoja(SH_HFA), HFA_FACTURA, HFA_FIRST, Tv(CfgV(CFG_PREFFACTURA))) + 1
        Case "2": celda = CFG_SIGGUIA: nombre = "guías": minimo = MaxConsecutivo(Hoja(SH_HGU), HGU_GUIA, HGU_FIRST, Tv(CfgV(CFG_PREFGUIA))) + 1
        Case "3": celda = CFG_SIGCOMPRA: nombre = "compras": minimo = MaxConsecutivo(Hoja(SH_COM), COM_NUMERO, COM_FIRST, Tv(CfgV(CFG_PREFCOMPRA))) + 1
        Case Else: Exit Sub
    End Select
    s = Pedir("Siguiente número actual de " & nombre & ": " & CfgV(celda) & vbLf & _
              "Mínimo permitido (para no repetir números ya usados): " & minimo & vbLf & vbLf & "Nuevo siguiente número:", "Numeración", CStr(CfgV(celda)))
    If s = "" Then Exit Sub
    nuevo = CLng(NumTexto(s))
    If nuevo < minimo Or nuevo < 1 Then Msg "El número debe ser mayor o igual a " & minimo & ". No se realizó el cambio.", vbExclamation: Exit Sub
    Iniciar
    Hoja(SH_CFG).Range(celda).Value = nuevo
    Auditar "CONFIGURACIÓN", "CAMBIAR NUMERACIÓN", nombre, CfgV(celda) & " -> " & nuevo
    If op = "1" Then NuevaFacturaInterna False
    If op = "2" Then NuevaGuiaInterna False
    Terminar
    Msg "Numeración de " & nombre & " actualizada. Siguiente: " & nuevo, vbInformation
End Sub

'-------------------------------------------------------------- logo
Public Sub CargarLogo()
    Dim archivo As Variant, ws As Worksheet, i As Long, s As Shape, h As Variant
    archivo = Application.GetOpenFilename("Imágenes (*.png;*.jpg;*.jpeg;*.gif;*.bmp),*.png;*.jpg;*.jpeg;*.gif;*.bmp", , "Seleccione el logo de la empresa")
    If VarType(archivo) = vbBoolean Then Exit Sub
    Iniciar
    On Error GoTo Fallo
    For Each h In Array(SH_FAC, SH_GUI)
        Set ws = Hoja(CStr(h))
        For i = ws.Shapes.Count To 1 Step -1
            If ws.Shapes(i).Type = msoPicture Then
                If ws.Shapes(i).TopLeftCell.Row <= 2 And ws.Shapes(i).TopLeftCell.Column <= 2 Then ws.Shapes(i).Delete
            End If
        Next i
        Set s = ws.Shapes.AddPicture(CStr(archivo), msoFalse, msoTrue, ws.Range("A1").Left + 4, ws.Range("A1").Top + 4, -1, -1)
        s.LockAspectRatio = msoTrue
        s.Height = ws.Range("A1:A5").Height - 8
        If s.Width > ws.Range("A1:B1").Width - 6 Then s.Width = ws.Range("A1:B1").Width - 6
        s.Name = "LogoEmpresa"
        s.Placement = xlMove
    Next h
    Hoja(SH_CFG).Range(CFG_LOGO).Value = CStr(archivo)
    Auditar "CONFIGURACIÓN", "CARGAR LOGO", "", CStr(archivo)
    Terminar
    Msg "Logo actualizado en la factura y en la guía.", vbInformation
    Exit Sub
Fallo:
    Terminar
    Msg "No se pudo cargar el logo: " & Err.Description, vbExclamation
End Sub

'-------------------------------------------------------------- datos de demostración
Public Sub LimpiarDatosDemo()
    Dim it As Variant, ws As Worksheet, r As Long, c As Long
    If Msg("Esto BORRARÁ todos los datos de demostración: ventas, compras, inventario, clientes, proveedores," & vbLf & _
           "ingresos, egresos, cartera, guías y movimientos. La configuración y las listas se conservan." & vbLf & vbLf & _
           "Úselo una sola vez, antes de empezar a operar con datos reales. ¿Continuar?", vbYesNo + vbCritical) <> vbYes Then Exit Sub
    If Msg("¿Está COMPLETAMENTE seguro? Se creará una copia de seguridad antes de borrar.", vbYesNo + vbCritical) <> vbYes Then Exit Sub
    If Not Autorizar("Borrar los datos de demostración") Then Exit Sub
    CopiaSeguridadInterna True, "antes de borrar datos demo"
    Iniciar
    For Each it In Array(Array(SH_BD, BD_FIRST, BD_NCOLS, 1), Array(SH_HFA, HFA_FIRST, HFA_NCOLS, 1), _
                         Array(SH_HGU, HGU_FIRST, HGU_NCOLS, 1), Array(SH_MOV, MOV_FIRST, MOV_NCOLS, 1), _
                         Array(SH_COM, COM_FIRST, COM_NCOLS, 1), Array(SH_ING, ING_FIRST, ING_NCOLS, 1), _
                         Array(SH_EGR, EGR_FIRST, EGR_NCOLS, 1), Array(SH_CAR, CAR_FIRST, CAR_NCOLS, 1), _
                         Array(SH_CAJ, CIE_FIRST, CIE_NCOLS, CIE_COL0), Array(SH_AUD, AUD_FIRST, AUD_NCOLS, 1))
        LimpiarFilasDesde Hoja(it(0)), it(3), it(1), it(3), it(2)
    Next it
    For Each it In Array(Array(SH_INV, INV_FIRST, INV_LAST, INV_NCOLS), Array(SH_PRO, PRO_FIRST, PRO_LAST, PRO_NCOLS), _
                         Array(SH_GFI, GFI_FIRST, GFI_LAST, GFI_NCOLS), Array(SH_CLI, CLI_FIRST + 1, CLI_LAST, CLI_NCOLS))
        Set ws = Hoja(it(0))
        QuitarFiltros ws
        For c = 1 To it(3)
            If Left(CStr(ws.Cells(it(1), c).Formula), 1) <> "=" Then ws.Range(ws.Cells(it(1), c), ws.Cells(it(2), c)).ClearContents
        Next c
    Next it
    Set ws = Hoja(SH_CAJ)
    ws.Range("C6:C10").Value = 0
    ws.Range(FCAJ_FINI).Value = Date
    ws.Range("R6:W205").ClearContents
    Hoja(SH_CFG).Range(CFG_SIGFACTURA).Value = 1
    Hoja(SH_CFG).Range(CFG_SIGGUIA).Value = 1
    Hoja(SH_CFG).Range(CFG_SIGCOMPRA).Value = 1
    Auditar "SISTEMA", "BORRAR DATOS DEMO", "", "Inicio de operación real"
    NuevaFacturaInterna False
    NuevaGuiaInterna False
    Terminar
    Msg "Datos de demostración eliminados." & vbLf & vbLf & "Siguientes pasos:" & vbLf & _
        "1) CONFIGURACIÓN: datos de su empresa, logo y claves" & vbLf & "2) CAJA Y BANCOS: saldos iniciales" & vbLf & _
        "3) PROVEEDORES y CLIENTES" & vbLf & "4) INVENTARIO: productos (o regístrelos con COMPRAS)", vbInformation
End Sub
