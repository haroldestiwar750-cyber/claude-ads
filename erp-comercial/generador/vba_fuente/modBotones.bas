Attribute VB_Name = "modBotones"
'==============================================================================
' modBotones - Dibuja los botones de cada modulo (franja de la fila 3) y los
' de la hoja REPORTES. Se puede ejecutar de nuevo en cualquier momento.
'==============================================================================
Option Explicit

Private Const AZUL As Long = 11957550      ' RGB(46,117,182)
Private Const VERDE As Long = 5287936      ' RGB(0,176,80)
Private Const ROJO As Long = 192           ' RGB(192,0,0)
Private Const GRIS As Long = 6908265       ' RGB(105,105,105)

Private mX As Double

Public Sub CrearBotones()
    Dim sh As Worksheet
    Application.ScreenUpdating = False
    For Each sh In ThisWorkbook.Worksheets
        BorrarBotones sh
    Next sh
    Barra "MENU", Array("INICIAR SESIÓN|IniciarSesion|V", "CERRAR SESIÓN|CerrarSesion|G", _
        "CAMBIAR MI CLAVE|CambiarMiClave|A", "RESPALDO AHORA|RespaldoAhora|A")
    Barra "DASHBOARD", Array("ACTUALIZAR|ActualizarTodo|A", "PDF DEL DASHBOARD|PDFDashboard|A")
    Barra "FACTURACION", Array("NUEVA FACTURA|NuevaFactura|A", "GUARDAR|GuardarFactura|V", "IMPRIMIR|ImprimirFactura|A", _
        "PDF|PDFFactura|A", "BUSCAR|BuscarFactura|A", "ANULAR / DEVOLUCIÓN|AnularFactura|R", "LIMPIAR|LimpiarFactura|G", _
        "+ CLIENTE|ClienteRapido|A")
    Barra "GUIAS", Array("NUEVA GUÍA|NuevaGuia|A", "GUARDAR|GuardarGuia|V", "IMPRIMIR|ImprimirGuia|A", "PDF|PDFGuia|A", _
        "BUSCAR|BuscarGuia|A", "LIMPIAR|LimpiarGuia|G", "CAMBIAR ESTADO|ActualizarEstadoGuia|A")
    Barra "INVENTARIO", Array("NUEVO PRODUCTO (SKU)|NuevoProducto|V", "AJUSTE DE INVENTARIO|AjusteInventario|R", _
        "ACTUALIZAR|ActualizarTodo|A")
    Barra "COMPRAS", Array("NUEVA COMPRA|NuevaCompra|A", "GUARDAR|GuardarCompra|V", "BUSCAR|BuscarCompra|A", _
        "IMPRIMIR|ImprimirCompra|A", "PDF|PDFCompra|A", "LIMPIAR|LimpiarCompra|G")
    Barra "CLIENTES", Array("NUEVO CLIENTE|BotonNuevoCliente|V", "ACTUALIZAR|ActualizarTodo|A")
    Barra "PROVEEDORES", Array("NUEVO PROVEEDOR|NuevoProveedor|V", "ACTUALIZAR|ActualizarTodo|A")
    Barra "INGRESOS", Array("NUEVO INGRESO (otros)|NuevoIngreso|V")
    Barra "EGRESOS", Array("NUEVO EGRESO|NuevoEgreso|V", "ANULAR EGRESO (fila seleccionada)|AnularEgreso|R")
    Barra "GASTOS FIJOS", Array("GENERAR EGRESOS DEL MES|GenerarGastosFijos|V")
    Barra "CARTERA", Array("REGISTRAR ABONO|RegistrarAbono|V", "ACTUALIZAR|ActualizarTodo|A")
    Barra "CAJA Y BANCOS", Array("CIERRE DIARIO|CierreDiario|V", "CIERRE SEMANAL|CierreSemanal|A", _
        "CIERRE MENSUAL|CierreMensual|A")
    Barra "VENDEDORES", Array("NUEVO VENDEDOR|NuevoVendedor|V")
    Barra "USUARIOS Y PERMISOS", Array("NUEVO USUARIO|NuevoUsuario|V", "RESTABLECER CLAVE (fila)|RestablecerClave|R", _
        "CAMBIAR MI CLAVE|CambiarMiClave|A")
    Barra "COMISIONES", Array("PAGAR COMISIONES DEL MES|PagarComisiones|V")
    Barra "CONFIGURACION", Array("DETECTAR ONEDRIVE|DetectarOneDrive|A", "CREAR CARPETAS|CrearCarpetas|A", _
        "RESPALDO AHORA|RespaldoAhora|V", "RESTAURAR COPIA|RestaurarCopia|R", "DESPROTEGER HOJAS|DesprotegerTodo|G")
    Barra "BD_GUIAS", Array("CAMBIAR ESTADO DE GUÍA|ActualizarEstadoGuia|A")
    BotonesReportes
    Application.ScreenUpdating = True
End Sub

Private Sub BorrarBotones(ByVal sh As Worksheet)
    Dim i As Long
    DesprotegerHoja sh
    For i = sh.Shapes.Count To 1 Step -1
        If Left$(sh.Shapes(i).Name, 3) = "btn" Then sh.Shapes(i).Delete
    Next i
End Sub

Private Function ColorDe(ByVal c As String) As Long
    Select Case c
        Case "V": ColorDe = VERDE
        Case "R": ColorDe = ROJO
        Case "G": ColorDe = GRIS
        Case Else: ColorDe = AZUL
    End Select
End Function

Private Sub Boton(ByVal sh As Worksheet, ByVal nombre As String, ByVal texto As String, ByVal macro As String, _
                  ByVal color As Long, ByVal x As Double, ByVal y As Double, ByVal ancho As Double, ByVal alto As Double)
    Dim s As Shape
    Set s = sh.Shapes.AddShape(5, x, y, ancho, alto)    ' 5 = rectangulo redondeado
    s.Name = nombre
    s.OnAction = macro
    s.Placement = xlFreeFloating
    s.Fill.ForeColor.RGB = color
    s.Line.Visible = msoFalse
    With s.TextFrame2
        .TextRange.Text = texto
        .TextRange.Font.Size = 8.5
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Name = "Arial"
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 2
        .MarginRight = 2
        .WordWrap = msoFalse
    End With
End Sub

Private Sub Barra(ByVal nombreHoja As String, ByVal defs As Variant)
    Dim sh As Worksheet, d As Variant, p As Variant, ancho As Double, i As Long
    On Error GoTo Fin
    Set sh = WS(nombreHoja)
    DesprotegerHoja sh
    If Clave(nombreHoja) <> "MENU" Then sh.Range("B3").ClearContents
    If Clave(nombreHoja) = "MENU" Then
        mX = sh.Range("B5").Left
    Else
        mX = sh.Range("B3").Left
    End If
    For Each d In defs
        i = i + 1
        p = Split(d, "|")
        ancho = 14 + Len(p(0)) * 5.6
        If ancho < 70 Then ancho = 70
        If Clave(nombreHoja) = "MENU" Then
            Boton sh, "btn" & i, CStr(p(0)), CStr(p(1)), ColorDe(CStr(p(2))), mX, sh.Range("B5").Top + 4, ancho, 20
        Else
            Boton sh, "btn" & i, CStr(p(0)), CStr(p(1)), ColorDe(CStr(p(2))), mX, sh.Range("B3").Top + 3, ancho, 23
        End If
        mX = mX + ancho + 5
    Next d
Fin:
    ProtegerHoja sh
End Sub

Private Sub BotonesReportes()
    Dim sh As Worksheet, r As Long, n As Long, c As Range
    Set sh = WS("REPORTES")
    DesprotegerHoja sh
    For r = 7 To 19
        n = r - 6
        Set c = sh.Range("E" & r)
        c.ClearContents
        Boton sh, "btnRep_" & Format$(n, "00") & "_PDF", "PDF", "ReporteClick", ROJO, c.Left + 6, c.Top + 3, c.Width - 12, c.Height - 6
        Set c = sh.Range("F" & r)
        c.ClearContents
        Boton sh, "btnRep_" & Format$(n, "00") & "_XLS", "EXCEL", "ReporteClick", VERDE, c.Left + 6, c.Top + 3, c.Width - 12, c.Height - 6
    Next r
    ProtegerHoja sh
End Sub
