Attribute VB_Name = "modGuia"
'==============================================================
' modGuia - GUÍAS DE ENVÍO
'  NUEVA · GUARDAR (o actualizar estado) · IMPRIMIR · LIMPIAR ·
'  BUSCAR · GENERAR DESDE FACTURA
'==============================================================
Option Explicit

Public Function NumeroGuia(ByVal n As Long) As String
    NumeroGuia = Tv(CfgV(CFG_PREFGUIA)) & Format(n, "0000")
End Function

Private Sub LimpiarCamposGuia(ByVal ws As Worksheet)
    Dim a As Variant
    For Each a In Array(GUI_FACTURA, GUI_DEST, GUI_DOC, GUI_TEL, GUI_CIUDAD, GUI_DEPTO, GUI_DIR, _
                        GUI_TRANSP, GUI_SEG, GUI_VALOR, GUI_PAQ, GUI_OBS)
        ws.Range(a).Value = ""
    Next a
    ws.Range("B" & GUI_L1 & ":B" & GUI_LN).ClearContents
    ws.Range("F" & GUI_L1 & ":F" & GUI_LN).ClearContents
End Sub

Public Sub NuevaGuia()
    NuevaGuiaInterna True
End Sub

Public Sub NuevaGuiaInterna(ByVal confirmar As Boolean)
    Dim ws As Worksheet
    Set ws = Hoja(SH_GUI)
    If confirmar And Tv(ws.Range(GUI_REG).Value) = "BORRADOR" And Tv(ws.Range(GUI_DEST).Value) <> "" Then
        If Msg("La guía actual no se ha guardado. ¿Descartarla?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    Iniciar
    LimpiarCamposGuia ws
    ws.Range(GUI_NUM).Value = NumeroGuia(CLng(Nv(CfgV(CFG_SIGGUIA))))
    ws.Range(GUI_FECHA).Value = Date
    ws.Range(GUI_ESTADO).Value = "PREPARANDO"
    ws.Range(GUI_PAQ).Value = 1
    ws.Range(GUI_REG).Value = "BORRADOR"
    Terminar
End Sub

Public Sub LimpiarGuia()
    Dim ws As Worksheet
    Set ws = Hoja(SH_GUI)
    If Tv(ws.Range(GUI_REG).Value) <> "BORRADOR" Then
        NuevaGuiaInterna False
        Exit Sub
    End If
    If Msg("¿Borrar los datos capturados de esta guía?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Iniciar
    LimpiarCamposGuia ws
    Terminar
End Sub

Public Sub GuardarGuia()
    Dim ws As Worksheet, wh As Worksheet, wg As Worksheet, num As String, fac As String, r As Long, f As Long
    Dim errores As String, nL As Long, txt As String, ser As String, uni As Double, a As Variant, rr As Long
    Dim ruta As String, sku As String, q As Double, wi As Worksheet, fi As Long
    Set ws = Hoja(SH_GUI)
    Set wg = Hoja(SH_HGU)
    Recalcular
    num = Tv(ws.Range(GUI_NUM).Value)
    ' --- guía ya guardada: solo actualiza estado / seguimiento
    If Tv(ws.Range(GUI_REG).Value) = "GUARDADA" Then
        f = BuscarFila(wg, HGU_GUIA, HGU_FIRST, num)
        If f = 0 Then Msg "No se encontró la guía " & num & " en el historial.", vbExclamation: Exit Sub
        If Msg("La guía " & num & " ya está guardada. ¿Actualizar estado del envío, transportadora, seguimiento y observaciones?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
        Iniciar
        wg.Cells(f, HGU_ESTADO).Value = Tv(ws.Range(GUI_ESTADO).Value)
        wg.Cells(f, HGU_TRANSP).Value = Tv(ws.Range(GUI_TRANSP).Value)
        wg.Cells(f, HGU_SEGUIMIENTO).Value = Tv(ws.Range(GUI_SEG).Value)
        wg.Cells(f, HGU_OBS).Value = Tv(ws.Range(GUI_OBS).Value)
        Auditar "GUÍAS", "ACTUALIZAR GUÍA", num, Tv(ws.Range(GUI_ESTADO).Value)
        Terminar
        Msg "Guía " & num & " actualizada (" & Tv(ws.Range(GUI_ESTADO).Value) & ").", vbInformation
        Exit Sub
    End If
    If num <> NumeroGuia(CLng(Nv(CfgV(CFG_SIGGUIA)))) Then Msg "El número no corresponde al consecutivo. Pulse NUEVA GUÍA.", vbExclamation: Exit Sub
    If BuscarFila(wg, HGU_GUIA, HGU_FIRST, num) > 0 Then Msg "La guía " & num & " ya existe.", vbCritical: Exit Sub
    If Tv(ws.Range(GUI_DEST).Value) = "" Then errores = errores & vbLf & "• Nombre del destinatario."
    If Tv(ws.Range(GUI_TEL).Value) = "" Then errores = errores & vbLf & "• Teléfono."
    If Tv(ws.Range(GUI_CIUDAD).Value) = "" Then errores = errores & vbLf & "• Ciudad."
    If Tv(ws.Range(GUI_DIR).Value) = "" Then errores = errores & vbLf & "• Dirección completa."
    If Tv(ws.Range(GUI_TRANSP).Value) = "" Then errores = errores & vbLf & "• Empresa transportadora."
    If Not IsDate(ws.Range(GUI_FECHA).Value) Then errores = errores & vbLf & "• Fecha de despacho."
    Set wi = Hoja(SH_INV)
    For r = GUI_L1 To GUI_LN
        sku = Tv(ws.Cells(r, 2).Value)
        If sku <> "" Then
            q = Nv(ws.Cells(r, 6).Value)
            fi = BuscarFila(wi, INV_SKU, INV_FIRST, sku)
            If fi = 0 Then errores = errores & vbLf & "• Línea " & (r - GUI_L1 + 1) & ": SKU inexistente."
            If q <= 0 Then errores = errores & vbLf & "• Línea " & (r - GUI_L1 + 1) & ": cantidad."
            nL = nL + 1
            uni = uni + q
            If fi > 0 Then txt = txt & IIf(txt = "", "", "; ") & q & " x " & Tv(wi.Cells(fi, INV_NOMBRE).Value) & " " & _
                               Tv(wi.Cells(fi, INV_TALLA).Value) & " " & Tv(wi.Cells(fi, INV_COLOR).Value)
            ser = ser & IIf(ser = "", "", ";") & sku & "|" & q
        End If
    Next r
    If nL = 0 Then errores = errores & vbLf & "• Agregue al menos un producto."
    If errores <> "" Then Msg "Faltan datos obligatorios:" & errores, vbExclamation: Exit Sub
    fac = UCase(Tv(ws.Range(GUI_FACTURA).Value))
    If fac <> "" Then
        Set wh = Hoja(SH_HFA)
        f = BuscarFila(wh, HFA_FACTURA, HFA_FIRST, fac)
        If f = 0 Then Msg "La factura relacionada " & fac & " no existe.", vbExclamation: Exit Sub
        If Tv(wh.Cells(f, HFA_ESTADO).Value) = "ANULADA" Then Msg "La factura " & fac & " está ANULADA.", vbExclamation: Exit Sub
        If BuscarFila(wg, HGU_FACTURA, HGU_FIRST, fac) > 0 Then
            If Msg("Ya existe una guía para la factura " & fac & ". ¿Crear otra (envío adicional)?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
        End If
    End If
    Iniciar
    On Error GoTo Fallo
    TxIniciar
    TxTabla SH_HGU, HGU_GUIA, HGU_FIRST, 1, HGU_NCOLS
    rr = NuevaFila(wg, HGU_GUIA, HGU_FIRST, HGU_LAST)
    a = NuevaFilaArr(HGU_NCOLS)
    a(1, HGU_GUIA) = num
    a(1, HGU_FECHA) = CDate(ws.Range(GUI_FECHA).Value)
    a(1, HGU_FACTURA) = fac
    a(1, HGU_DEST) = Tv(ws.Range(GUI_DEST).Value)
    a(1, HGU_DOC) = Tv(ws.Range(GUI_DOC).Value)
    a(1, HGU_TEL) = Tv(ws.Range(GUI_TEL).Value)
    a(1, HGU_CIUDAD) = Tv(ws.Range(GUI_CIUDAD).Value)
    a(1, HGU_DEPTO) = Tv(ws.Range(GUI_DEPTO).Value)
    a(1, HGU_DIRECCION) = Tv(ws.Range(GUI_DIR).Value)
    a(1, HGU_TRANSP) = Tv(ws.Range(GUI_TRANSP).Value)
    a(1, HGU_SEGUIMIENTO) = Tv(ws.Range(GUI_SEG).Value)
    a(1, HGU_PRODUCTOS) = txt
    a(1, HGU_UNIDADES) = uni
    a(1, HGU_VALORDECL) = Nv(ws.Range(GUI_VALOR).Value)
    a(1, HGU_ESTADO) = Tv(ws.Range(GUI_ESTADO).Value)
    a(1, HGU_OBS) = Tv(ws.Range(GUI_OBS).Value)
    a(1, HGU_USUARIO) = UsuarioActual()
    a(1, HGU_REGISTRO) = Now
    a(1, HGU_ITEMS) = ser
    a(1, HGU_PAQUETES) = Nv(ws.Range(GUI_PAQ).Value)
    EscribirFila wg, rr, 1, a
    CfgSet CFG_SIGGUIA, CLng(Nv(CfgV(CFG_SIGGUIA))) + 1
    ws.Range(GUI_REG).Value = "GUARDADA"
    Auditar "GUÍAS", "GUARDAR GUÍA", num, "Factura " & fac & " · " & Tv(ws.Range(GUI_DEST).Value)
    TxConfirmar
    Terminar
    ruta = ExportarPDF(ws, "GUIAS", num & " - " & Tv(ws.Range(GUI_DEST).Value))
    If ruta <> "" Then
        DesprotegerTodo
        wg.Cells(rr, HGU_PDF).Value = ruta
        ProtegerTodo
    End If
    Msg "Guía " & num & " guardada." & IIf(ruta <> "", vbLf & "PDF: " & ruta, ""), vbInformation
    Exit Sub
Fallo:
    Dim e As String
    e = Err.Description
    TxRevertir
    Terminar
    Msg "La guía NO se guardó: " & e, vbCritical
End Sub

Public Sub ImprimirGuia()
    Dim ws As Worksheet
    Set ws = Hoja(SH_GUI)
    If Tv(ws.Range(GUI_DEST).Value) = "" Then Msg "La guía está vacía.", vbExclamation: Exit Sub
    If gTest Then Exit Sub
    On Error GoTo Fallo
    DesprotegerTodo
    PrepararImpresion ws, True
    ws.PrintOut Copies:=1
    PrepararImpresion ws, False
    ProtegerTodo
    Exit Sub
Fallo:
    PrepararImpresion ws, False
    ProtegerTodo
    Msg "No se pudo imprimir: " & Err.Description, vbExclamation
End Sub

' Llena una nueva guía con los datos de una factura guardada
Public Sub GuiaDesdeFactura()
    Dim num As String, wf As Worksheet
    Set wf = Hoja(SH_FAC)
    If Tv(wf.Range(FAC_ESTADO).Value) = "GUARDADA" Then num = Tv(wf.Range(FAC_NUM).Value)
    num = UCase(Pedir("Número de la factura para generar la guía:", "Guía desde factura", num))
    If num = "" Then Exit Sub
    CargarGuiaDesdeFactura num
End Sub

Public Sub CargarGuiaDesdeFactura(ByVal num As String)
    Dim ws As Worksheet, wh As Worksheet, wb As Worksheet, f As Long, ult As Long, r As Long, k As Long
    Dim skus() As String, cants() As Double, n As Long, i As Long, s As String, found As Boolean
    Set wh = Hoja(SH_HFA)
    f = BuscarFila(wh, HFA_FACTURA, HFA_FIRST, num)
    If f = 0 Then Msg "La factura " & num & " no existe.", vbExclamation: Exit Sub
    If Tv(wh.Cells(f, HFA_ESTADO).Value) <> "ACTIVA" Then Msg "La factura " & num & " está ANULADA.", vbExclamation: Exit Sub
    NuevaGuiaInterna False
    Set ws = Hoja(SH_GUI)
    Set wb = Hoja(SH_BD)
    ReDim skus(1 To 50): ReDim cants(1 To 50)
    ult = UltimaFila(wb, BD_FACTURA, BD_FIRST)
    For r = BD_FIRST To ult
        If UCase(Tv(wb.Cells(r, BD_FACTURA).Value)) = num And Tv(wb.Cells(r, BD_ESTADO).Value) = "ACTIVA" Then
            s = Tv(wb.Cells(r, BD_SKU).Value)
            found = False
            For i = 1 To n
                If skus(i) = s Then cants(i) = cants(i) + Nv(wb.Cells(r, BD_CANT).Value): found = True
            Next i
            If Not found And n < 50 Then n = n + 1: skus(n) = s: cants(n) = Nv(wb.Cells(r, BD_CANT).Value)
        End If
    Next r
    Iniciar
    ws.Range(GUI_FACTURA).Value = num
    ws.Range(GUI_DEST).Value = Tv(wh.Cells(f, HFA_CLIENTE).Value)
    ws.Range(GUI_DOC).Value = Tv(wh.Cells(f, HFA_IDCLI).Value)
    ws.Range(GUI_TEL).Value = Tv(wh.Cells(f, HFA_TEL).Value)
    ws.Range(GUI_CIUDAD).Value = Tv(wh.Cells(f, HFA_CIUDAD).Value)
    ws.Range(GUI_DEPTO).Value = Tv(wh.Cells(f, HFA_DEPTO).Value)
    ws.Range(GUI_DIR).Value = Tv(wh.Cells(f, HFA_DIRECCION).Value)
    ws.Range(GUI_VALOR).Value = Nv(wh.Cells(f, HFA_TOTAL).Value)
    k = GUI_L1
    For i = 1 To n
        If cants(i) > 0 And k <= GUI_LN Then
            ws.Cells(k, 2).Value = skus(i)
            ws.Cells(k, 6).Value = cants(i)
            k = k + 1
        End If
    Next i
    Terminar
    On Error Resume Next
    If Not gTest Then
        ws.Activate
        ws.Range(GUI_TRANSP).Select
        Msg "Guía preparada desde la factura " & num & "." & vbLf & "Complete la transportadora y el número de seguimiento y pulse GUARDAR GUÍA.", vbInformation
    End If
End Sub

Public Sub BuscarGuia()
    Dim txt As String, wg As Worksheet, f As Long, ult As Long, r As Long, lista As String, n As Long, ultNum As String
    txt = Pedir("Buscar guía: escriba el N° de guía, el N° de factura o parte del nombre del destinatario.", "Buscar guía")
    If txt = "" Then Exit Sub
    Set wg = Hoja(SH_HGU)
    f = BuscarFila(wg, HGU_GUIA, HGU_FIRST, txt)
    If f > 0 Then CargarGuia Tv(wg.Cells(f, HGU_GUIA).Value): Exit Sub
    ult = UltimaFila(wg, HGU_GUIA, HGU_FIRST)
    For r = HGU_FIRST To ult
        If UCase(Tv(wg.Cells(r, HGU_FACTURA).Value)) = UCase(txt) Or InStr(UCase(Tv(wg.Cells(r, HGU_DEST).Value)), UCase(txt)) > 0 Then
            n = n + 1
            ultNum = Tv(wg.Cells(r, HGU_GUIA).Value)
            If n <= 25 Then lista = lista & vbLf & ultNum & "  " & FechaCorta(wg.Cells(r, HGU_FECHA).Value) & "  " & _
                                    Tv(wg.Cells(r, HGU_FACTURA).Value) & "  " & Left(Tv(wg.Cells(r, HGU_DEST).Value), 22) & "  " & Tv(wg.Cells(r, HGU_ESTADO).Value)
        End If
    Next r
    If n = 0 Then
        Msg "No se encontraron guías para: " & txt, vbInformation
    ElseIf n = 1 Then
        CargarGuia ultNum
    Else
        txt = Pedir(n & " guías encontradas:" & lista & vbLf & vbLf & "Escriba el número de la guía:", "Resultados", ultNum)
        If txt <> "" Then CargarGuia UCase(txt)
    End If
End Sub

Public Sub CargarGuia(ByVal num As String)
    Dim ws As Worksheet, wg As Worksheet, f As Long, items() As String, p() As String, i As Long, k As Long
    Set ws = Hoja(SH_GUI)
    Set wg = Hoja(SH_HGU)
    f = BuscarFila(wg, HGU_GUIA, HGU_FIRST, num)
    If f = 0 Then Msg "No existe la guía " & num, vbExclamation: Exit Sub
    Iniciar
    LimpiarCamposGuia ws
    ws.Range(GUI_NUM).Value = Tv(wg.Cells(f, HGU_GUIA).Value)
    ws.Range(GUI_FECHA).Value = wg.Cells(f, HGU_FECHA).Value
    ws.Range(GUI_FACTURA).Value = Tv(wg.Cells(f, HGU_FACTURA).Value)
    ws.Range(GUI_ESTADO).Value = Tv(wg.Cells(f, HGU_ESTADO).Value)
    ws.Range(GUI_DEST).Value = Tv(wg.Cells(f, HGU_DEST).Value)
    ws.Range(GUI_DOC).Value = Tv(wg.Cells(f, HGU_DOC).Value)
    ws.Range(GUI_TEL).Value = Tv(wg.Cells(f, HGU_TEL).Value)
    ws.Range(GUI_CIUDAD).Value = Tv(wg.Cells(f, HGU_CIUDAD).Value)
    ws.Range(GUI_DEPTO).Value = Tv(wg.Cells(f, HGU_DEPTO).Value)
    ws.Range(GUI_DIR).Value = Tv(wg.Cells(f, HGU_DIRECCION).Value)
    ws.Range(GUI_TRANSP).Value = Tv(wg.Cells(f, HGU_TRANSP).Value)
    ws.Range(GUI_SEG).Value = Tv(wg.Cells(f, HGU_SEGUIMIENTO).Value)
    ws.Range(GUI_VALOR).Value = Nv(wg.Cells(f, HGU_VALORDECL).Value)
    ws.Range(GUI_PAQ).Value = Nv(wg.Cells(f, HGU_PAQUETES).Value)
    ws.Range(GUI_OBS).Value = Tv(wg.Cells(f, HGU_OBS).Value)
    If Tv(wg.Cells(f, HGU_ITEMS).Value) <> "" Then
        items = Split(Tv(wg.Cells(f, HGU_ITEMS).Value), ";")
        k = GUI_L1
        For i = LBound(items) To UBound(items)
            p = Split(items(i), "|")
            If UBound(p) >= 1 And k <= GUI_LN Then
                ws.Cells(k, 2).Value = p(0)
                ws.Cells(k, 6).Value = Val(p(1))
                k = k + 1
            End If
        Next i
    End If
    ws.Range(GUI_REG).Value = "GUARDADA"
    Terminar
    On Error Resume Next
    If Not gTest Then ws.Activate
End Sub

' Botón en FACTURACIÓN: genera la guía de la factura en pantalla
Public Sub GenerarGuiaDeFacturaActual()
    Dim wf As Worksheet
    Set wf = Hoja(SH_FAC)
    If Tv(wf.Range(FAC_ESTADO).Value) <> "GUARDADA" Then
        Msg "Primero guarde la factura (o búsquela). Solo se generan guías de facturas guardadas y activas.", vbExclamation
        Exit Sub
    End If
    CargarGuiaDesdeFactura Tv(wf.Range(FAC_NUM).Value)
End Sub
