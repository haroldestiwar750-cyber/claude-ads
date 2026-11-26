Attribute VB_Name = "modGuia"
'==============================================================================
' modGuia - Guias de envio ligadas a comprobantes de venta.
' Botones: NUEVA GUIA, GUARDAR, IMPRIMIR, PDF, BUSCAR, LIMPIAR.
'==============================================================================
Option Explicit

Private Const C_CLAVE As String = "M8"

Private Function HG() As Worksheet
    Set HG = WS("GUIAS")
End Function

Public Sub NuevaGuia()
    If Not RequierePermiso("GUIAS") Then Exit Sub
    LimpiarFormulario
    HG().Activate
    HG().Range("C10").Select
End Sub

Public Sub LimpiarGuia()
    LimpiarFormulario
End Sub

Private Sub LimpiarFormulario()
    Dim sh As Worksheet
    Set sh = HG()
    Application.EnableEvents = False
    sh.Range("C10,I11").ClearContents
    sh.Range("C13").Formula = "=IF(M12="""","""",INDEX(cli_Dir,M12))"
    sh.Range("I12").Formula = "=IF(M10="""","""",INDEX(fac_Total,M10))"
    sh.Range("I13").Value = "PENDIENTE"
    sh.Range("I6").Formula = "=cfg_PrefGuia&""-""&TEXT(cfg_NumGuia,""00000"")&"" (pendiente)"""
    sh.Range("I7").Value = Date
    sh.Range("I8").Value = "NUEVO"
    sh.Range(C_CLAVE).Value = ClaveUnica()
    Application.EnableEvents = True
End Sub

Public Sub GuardarGuia()
    Dim sh As Worksheet, hG As Worksheet, num As String, n As Long, fac As String, r As Long, prods As String
    Dim unidades As Double, fila As Long, ruta As String, rf As Long, hFac As Worksheet

    If Not RequierePermiso("GUIAS") Then Exit Sub
    Set sh = HG()
    Application.Calculate
    If sh.Range("I8").Value <> "NUEVO" Then
        MsgBox "Esta guía ya está registrada. Use NUEVA GUÍA.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    fac = UCase$(Trim$(CStr(sh.Range("C10").Value)))
    If CStr(sh.Range("M13").Value) <> "OK" Then
        If CStr(sh.Range("M13").Value) = "YA TIENE GUÍA" Then
            If MsgBox("El comprobante " & fac & " ya tiene una guía. ¿Crear otra (envío adicional)?", _
                      vbQuestion + vbYesNo, APP_NOMBRE) = vbNo Then Exit Sub
        Else
            MsgBox "Comprobante no válido: " & sh.Range("M13").Value, vbExclamation, APP_NOMBRE
            Exit Sub
        End If
    End If
    If Vacio(sh.Range("I10").Value) Then MsgBox "Elija la transportadora.", vbExclamation, APP_NOMBRE: Exit Sub
    If Vacio(sh.Range("C13").Value) Or Vacio(sh.Range("C14").Value) Then
        MsgBox "Falta dirección o ciudad de entrega.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    Set hG = WS("BD_GUIAS")
    For r = 17 To 31
        If Not Vacio(sh.Range("C" & r).Value) Then
            prods = prods & IIf(prods = "", "", "; ") & sh.Range("C" & r).Value & " x" & sh.Range("G" & r).Value
            unidades = unidades + NumV(sh.Range("G" & r).Value)
        End If
    Next r

    On Error GoTo FALLO
    InicioMacro
    IniciarTransaccion
    n = CLng(Cfg("cfg_NumGuia"))
    Do
        num = CStr(Cfg("cfg_PrefGuia")) & "-" & Format$(n, "00000")
        If BuscarFila(hG, 1, num) = 0 Then Exit Do
        n = n + 1
    Loop
    EscribirCelda NR("cfg_NumGuia").Worksheet, NR("cfg_NumGuia").Row, NR("cfg_NumGuia").Column, n + 1
    fila = SiguienteFila(hG)
    EscribirFila hG, fila, Array(num, CDate(sh.Range("I7").Value), fac, sh.Range("M11").Value, sh.Range("C11").Value, _
        sh.Range("C12").Value, sh.Range("C13").Value, sh.Range("C14").Value, sh.Range("I10").Value, _
        CStr(sh.Range("I11").Value), prods, unidades, NumV(sh.Range("I12").Value), sh.Range("I13").Value, _
        UsuarioActual(), Now, "")
    hG.Cells(fila, 14).Locked = False
    Bitacora "GUÍAS", "GUARDAR", num, fac & " · " & sh.Range("I10").Value
    CerrarTransaccion
    Application.EnableEvents = False
    sh.Range("I6").Value = num
    sh.Range("I8").Value = "VIGENTE"
    FinMacro
    ruta = modRespaldo.ExportarPDF(sh, "GUIAS", num, CDate(sh.Range("I7").Value))
    If ruta <> "" Then
        Application.EnableEvents = False
        hG.Cells(fila, 17).Value = modRespaldo.RutaRelativa(ruta)
        Application.EnableEvents = True
    End If
    GuardarLibro
    MsgBox "Guía " & num & " guardada." & IIf(ruta <> "", vbCrLf & "PDF: " & ruta, ""), vbInformation, APP_NOMBRE
    Exit Sub
FALLO:
    DeshacerTransaccion
    FinMacro
    ErrorGeneral "GuardarGuia"
End Sub

Public Sub BuscarGuia()
    Dim num As String, hG As Worksheet, rf As Long, sh As Worksheet
    Set sh = HG()
    Set hG = WS("BD_GUIAS")
    num = UCase$(PedirTexto("Número de guía o de comprobante (ej. GU-00003 o FV-000012):", _
                            CStr(hG.Cells(UltimaFila(hG, 1), 1).Value)))
    If num = "" Then Exit Sub
    rf = BuscarFila(hG, 1, num)
    If rf = 0 Then rf = BuscarFila(hG, 3, num)
    If rf = 0 Then MsgBox "No se encontró la guía " & num & ".", vbExclamation, APP_NOMBRE: Exit Sub
    LimpiarFormulario
    Application.EnableEvents = False
    sh.Range("I6").Value = hG.Cells(rf, 1).Value
    sh.Range("I7").Value = hG.Cells(rf, 2).Value
    sh.Range("I8").Value = "VIGENTE · " & hG.Cells(rf, 14).Value
    sh.Range("C10").Value = hG.Cells(rf, 3).Value
    sh.Range("C13").Value = hG.Cells(rf, 7).Value
    sh.Range("I10").Value = hG.Cells(rf, 9).Value
    sh.Range("I11").Value = hG.Cells(rf, 10).Value
    sh.Range("I12").Value = hG.Cells(rf, 13).Value
    sh.Range("I13").Value = hG.Cells(rf, 14).Value
    Application.EnableEvents = True
    sh.Activate
End Sub

' Cambia el estado de la guia seleccionada en BD_GUIAS (o pedida por numero)
Public Sub ActualizarEstadoGuia()
    Dim hG As Worksheet, num As String, rf As Long, est As String
    If Not RequierePermiso("GUIAS") Then Exit Sub
    Set hG = WS("BD_GUIAS")
    num = UCase$(PedirTexto("Número de guía:"))
    If num = "" Then Exit Sub
    rf = BuscarFila(hG, 1, num)
    If rf = 0 Then MsgBox "No existe " & num & ".", vbExclamation, APP_NOMBRE: Exit Sub
    est = ElegirDeLista("Nuevo estado para " & num & ":", ListaCatalogo("Estados de guía"), CStr(hG.Cells(rf, 14).Value))
    If est = "" Then Exit Sub
    Application.EnableEvents = False
    hG.Cells(rf, 14).Value = est
    Application.EnableEvents = True
    Bitacora "GUÍAS", "ESTADO", num, est
    GuardarLibro
End Sub

Public Sub ImprimirGuia()
    HG().PrintOut Copies:=1, Preview:=True
End Sub

Public Sub PDFGuia()
    Dim ruta As String
    If HG().Range("I8").Value = "NUEVO" Then MsgBox "Guarde la guía primero.", vbExclamation, APP_NOMBRE: Exit Sub
    ruta = modRespaldo.ExportarPDF(HG(), "GUIAS", CStr(HG().Range("I6").Value), CDate(HG().Range("I7").Value))
    If ruta <> "" Then MsgBox "PDF generado:" & vbCrLf & ruta, vbInformation, APP_NOMBRE
End Sub
