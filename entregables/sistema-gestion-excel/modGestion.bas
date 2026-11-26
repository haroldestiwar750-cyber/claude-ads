Attribute VB_Name = "modGestion"
'==============================================================================
'  SISTEMA INTEGRAL DE GESTIÓN EMPRESARIAL - MÓDULO DE MACROS
'------------------------------------------------------------------------------
'  Botones de FACTURACION : ImprimirFactura, LimpiarFactura, GuardarFactura,
'                           NuevaFactura
'  Botones de GUIAS       : CargarFacturaEnGuia, ImprimirGuia, LimpiarGuia,
'                           GuardarGuia, NuevaGuia
'  Utilidades             : CrearBotones (crea y asigna todos los botones),
'                           ProtegerHojas / DesprotegerHojas
'
'  IMPORTANTE: las direcciones de celda de este módulo coinciden EXACTAMENTE
'  con la plantilla Sistema_Gestion_Empresarial.xlsx. Si mueve filas o
'  columnas en FACTURACION o GUIAS, actualice las constantes de abajo.
'  Las fórmulas se escriben con .Formula (en inglés y con coma), que Excel
'  traduce solo al idioma de su instalación.
'==============================================================================
Option Explicit

'----------------------------- NOMBRES DE HOJAS -------------------------------
Private Const H_FAC As String = "FACTURACION"
Private Const H_GUIA As String = "GUIAS"
Private Const H_CFG As String = "CONFIG"
Private Const H_HFAC As String = "HIST_FACTURAS"
Private Const H_HDET As String = "HIST_DETALLE"
Private Const H_HGUIA As String = "HIST_GUIAS"
Private Const H_CART As String = "CARTERA"
Private Const H_CAJA As String = "INGRESOS_EGRESOS"
Private Const H_INV As String = "INVENTARIO"
Private Const H_CLI As String = "CLIENTES"

Private Const FILA_DATOS As Long = 5          'primera fila de datos en tablas

'----------------------------- FACTURA ----------------------------------------
Private Const FAC_AREA As String = "$A$1:$H$44"   'área de impresión
Private Const FAC_NUM As String = "H3"            'consecutivo numérico
Private Const FAC_NUMTXT As String = "H2"         'prefijo + consecutivo
Private Const FAC_FECHA As String = "B7"
Private Const FAC_PAGO As String = "D7"
Private Const FAC_VENCE As String = "G7"
Private Const FAC_CLI As String = "B10"
Private Const FAC_CLINOM As String = "C10"
Private Const FAC_CIUDAD As String = "E10"
Private Const FAC_OBS As String = "A36"
Private Const FAC_SUB As String = "H35"
Private Const FAC_IVA As String = "H36"
Private Const FAC_TOT As String = "H37"
Private Const FAC_ESTADO As String = "K3"
Private Const FAC_F1 As Long = 14                 'primera fila de ítems
Private Const FAC_F2 As Long = 33                 'última fila de ítems

'----------------------------- GUÍA -------------------------------------------
Private Const GUI_AREA As String = "$A$1:$H$44"
Private Const GUI_NUM As String = "H3"
Private Const GUI_NUMTXT As String = "H2"
Private Const GUI_FECHA As String = "B7"
Private Const GUI_FACT As String = "D7"
Private Const GUI_TRANSP As String = "G7"
Private Const GUI_VALOR As String = "B8"
Private Const GUI_RASTREO As String = "G8"
Private Const GUI_NOMBRE As String = "B11"
Private Const GUI_ID As String = "E11"
Private Const GUI_TEL As String = "H11"
Private Const GUI_DIR As String = "B12"
Private Const GUI_CIUDAD As String = "E12"
Private Const GUI_DEPTO As String = "H12"
Private Const GUI_OBS As String = "A38"
Private Const GUI_BULTOS As String = "F35"
Private Const GUI_ESTADO As String = "K3"
Private Const GUI_F1 As Long = 15
Private Const GUI_F2 As Long = 34

'----------------------------- PROTECCIÓN -------------------------------------
Private Const CLAVE As String = ""        'escriba una clave si desea, ej. "1234"

'==============================================================================
'                               FACTURACIÓN
'==============================================================================

'--- BOTÓN: IMPRIMIR SOLO FACTURA ---------------------------------------------
Public Sub ImprimirFactura()
    Dim ws As Worksheet, resp As VbMsgBoxResult
    Set ws = ThisWorkbook.Worksheets(H_FAC)

    If Not TieneItems(ws, FAC_F1, FAC_F2) Then
        MsgBox "La factura no tiene productos.", vbExclamation, "Imprimir factura"
        Exit Sub
    End If

    ConfigurarPagina ws, FAC_AREA

    resp = MsgBox("¿Cómo desea continuar?" & vbCrLf & vbCrLf & _
                  "SÍ  = Imprimir directamente" & vbCrLf & _
                  "NO  = Ver vista previa" & vbCrLf & _
                  "CANCELAR = Salir", vbYesNoCancel + vbQuestion, "Imprimir factura " & ws.Range(FAC_NUMTXT).Value)
    Select Case resp
        Case vbYes: ws.PrintOut Copies:=1, Collate:=True
        Case vbNo:  ws.PrintPreview
    End Select
End Sub

'--- BOTÓN: LIMPIAR FACTURA ---------------------------------------------------
Public Sub LimpiarFactura()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(H_FAC)
    If MsgBox("¿Borrar los datos de la factura actual?" & vbCrLf & _
              "(No se borran fórmulas, encabezado ni el número de factura)", _
              vbYesNo + vbQuestion, "Limpiar factura") = vbNo Then Exit Sub
    LimpiarFacturaInterno ws
End Sub

'--- BOTÓN: NUEVA FACTURA -----------------------------------------------------
Public Sub NuevaFactura()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(H_FAC)

    If ws.Range(FAC_ESTADO).Value <> "GUARDADA" And TieneItems(ws, FAC_F1, FAC_F2) Then
        If MsgBox("La factura actual NO ha sido guardada." & vbCrLf & _
                  "¿Desea descartarla y crear una nueva?", vbYesNo + vbExclamation, _
                  "Nueva factura") = vbNo Then Exit Sub
    End If

    LimpiarFacturaInterno ws
    ws.Range(FAC_NUM).Value = SiguienteConsecutivo(H_HFAC, ThisWorkbook.Worksheets(H_CFG).Range("B14").Value)
    ws.Activate
    ws.Range(FAC_CLI).Select
End Sub

'--- BOTÓN: GUARDAR FACTURA (historial + cartera/caja + PDF/Excel) ------------
Public Sub GuardarFactura()
    Dim ws As Worksheet, wsH As Worksheet, wsD As Worksheet, wsCfg As Worksheet
    Dim wsInv As Worksheet, wsCart As Worksheet, wsCaja As Worksheet
    Dim r As Long, fh As Long, fd As Long, n As Long
    Dim numero As Long, numTxt As String, fecha As Date, cliente As String
    Dim sku As String, cant As Double, costoU As Double, costoTot As Double
    Dim m As Variant, avisos As String, carpeta As String, archivo As String
    Dim formato As String, rutas As String

    On Error GoTo Fallo
    Set ws = ThisWorkbook.Worksheets(H_FAC)
    Set wsH = ThisWorkbook.Worksheets(H_HFAC)
    Set wsD = ThisWorkbook.Worksheets(H_HDET)
    Set wsCfg = ThisWorkbook.Worksheets(H_CFG)
    Set wsInv = ThisWorkbook.Worksheets(H_INV)
    Set wsCart = ThisWorkbook.Worksheets(H_CART)
    Set wsCaja = ThisWorkbook.Worksheets(H_CAJA)
    Application.Calculate

    '---------- 1. VALIDACIONES ----------
    numero = Val(ws.Range(FAC_NUM).Value)
    numTxt = CStr(ws.Range(FAC_NUMTXT).Value)
    If numero <= 0 Then Err.Raise vbObjectError + 1, , "No hay número de factura. Use NUEVA FACTURA."
    If Not IsError(Application.Match(numero, wsH.Columns(1), 0)) Then
        Err.Raise vbObjectError + 2, , "La factura " & numTxt & " YA fue guardada." & vbCrLf & _
                  "Use NUEVA FACTURA para generar el siguiente consecutivo."
    End If
    If Not IsDate(ws.Range(FAC_FECHA).Value) Then Err.Raise vbObjectError + 3, , "Digite una fecha válida en " & FAC_FECHA & "."
    fecha = CDate(ws.Range(FAC_FECHA).Value)
    cliente = CStr(ws.Range(FAC_CLINOM).Value)
    If Trim$(CStr(ws.Range(FAC_CLI).Value)) = "" Then Err.Raise vbObjectError + 4, , "Seleccione el NIT/CC del cliente."
    If cliente = "" Or cliente = "CLIENTE NO REGISTRADO" Then _
        Err.Raise vbObjectError + 5, , "El cliente no existe en la hoja CLIENTES. Regístrelo primero."
    If Trim$(CStr(ws.Range(FAC_PAGO).Value)) = "" Then Err.Raise vbObjectError + 6, , "Seleccione la forma de pago (Contado / Crédito)."
    If Not TieneItems(ws, FAC_F1, FAC_F2) Then Err.Raise vbObjectError + 7, , "La factura no tiene productos."

    For r = FAC_F1 To FAC_F2
        sku = Trim$(CStr(ws.Cells(r, "B").Value))
        If sku = "" And Not IsEmpty(ws.Cells(r, "D").Value) Then _
            Err.Raise vbObjectError + 8, , "Fila " & r & ": hay cantidad pero falta el SKU."
        If sku <> "" Then
            If ws.Cells(r, "C").Value = "SKU NO EXISTE" Then _
                Err.Raise vbObjectError + 9, , "Fila " & r & ": el SKU " & sku & " no existe en PRODUCTOS."
            If Not IsNumeric(ws.Cells(r, "D").Value) Or Val(ws.Cells(r, "D").Value) <= 0 Then _
                Err.Raise vbObjectError + 10, , "Fila " & r & ": la cantidad debe ser mayor que cero."
            If IsNumeric(ws.Cells(r, "I").Value) Then
                If CDbl(ws.Cells(r, "D").Value) > CDbl(ws.Cells(r, "I").Value) Then _
                    avisos = avisos & "  - " & sku & ": pide " & ws.Cells(r, "D").Value & _
                             ", disponible " & ws.Cells(r, "I").Value & vbCrLf
            End If
        End If
    Next r
    If avisos <> "" Then
        If MsgBox("Stock insuficiente:" & vbCrLf & avisos & vbCrLf & _
                  "¿Desea guardar de todas formas? (el inventario quedará negativo)", _
                  vbYesNo + vbExclamation, "Inventario") = vbNo Then Exit Sub
    End If

    Application.ScreenUpdating = False

    '---------- 2. COPIA EN EL ESCRITORIO (primero: si falla, no se registra nada) ----------
    carpeta = CarpetaDestino("Facturas", fecha)
    archivo = carpeta & "\" & LimpiarNombre(numTxt & "_" & cliente)
    formato = UCase$(Trim$(CStr(wsCfg.Range("B18").Value)))
    If formato = "" Then formato = "PDF"

    ConfigurarPagina ws, FAC_AREA
    If formato = "PDF" Or formato = "AMBOS" Then
        ExportarPDF ws, archivo & ".pdf"
        rutas = archivo & ".pdf"
    End If
    If formato = "EXCEL" Or formato = "AMBOS" Then
        ExportarExcel ws, archivo & ".xlsx", "I:K"
        rutas = rutas & IIf(rutas = "", "", " | ") & archivo & ".xlsx"
    End If

    '---------- 3. DETALLE (HIST_DETALLE) ----------
    fd = SiguienteFila(wsD, 1)
    For r = FAC_F1 To FAC_F2
        sku = Trim$(CStr(ws.Cells(r, "B").Value))
        If sku <> "" Then
            cant = CDbl(ws.Cells(r, "D").Value)
            m = Application.Match(sku, wsInv.Columns(1), 0)
            If IsError(m) Then costoU = 0 Else costoU = Val(wsInv.Cells(m, "J").Value)
            wsD.Cells(fd, 1).Value = numTxt
            wsD.Cells(fd, 2).Value = fecha
            wsD.Cells(fd, 3).NumberFormat = "@"
            wsD.Cells(fd, 3).Value = sku
            wsD.Cells(fd, 4).Value = ws.Cells(r, "C").Value
            wsD.Cells(fd, 5).Value = cant
            wsD.Cells(fd, 6).Value = ws.Cells(r, "E").Value
            wsD.Cells(fd, 7).Value = ws.Cells(r, "F").Value
            wsD.Cells(fd, 8).Value = ws.Cells(r, "G").Value
            wsD.Cells(fd, 9).Value = ws.Cells(r, "H").Value
            wsD.Cells(fd, 10).Value = costoU
            wsD.Cells(fd, 11).Value = costoU * cant
            wsD.Cells(fd, 12).Value = ws.Cells(r, "F").Value - costoU * cant
            wsD.Cells(fd, 13).NumberFormat = "@"
            wsD.Cells(fd, 13).Value = CStr(ws.Range(FAC_CLI).Value)
            wsD.Cells(fd, 14).Value = numero
            wsD.Range(wsD.Cells(fd, 1), wsD.Cells(fd, 14)).Font.Name = "Arial"
            costoTot = costoTot + costoU * cant
            fd = fd + 1
        End If
    Next r
    FormatearFilasHist wsD

    '---------- 4. ENCABEZADO (HIST_FACTURAS) ----------
    fh = SiguienteFila(wsH, 1)
    wsH.Cells(fh, 1).Value = numero
    wsH.Cells(fh, 2).Value = numTxt
    wsH.Cells(fh, 3).Value = fecha
    wsH.Cells(fh, 4).NumberFormat = "@"
    wsH.Cells(fh, 4).Value = CStr(ws.Range(FAC_CLI).Value)
    wsH.Cells(fh, 5).Value = cliente
    wsH.Cells(fh, 6).Value = ws.Range(FAC_CIUDAD).Value
    wsH.Cells(fh, 7).Value = ws.Range(FAC_SUB).Value
    wsH.Cells(fh, 8).Value = ws.Range(FAC_IVA).Value
    wsH.Cells(fh, 9).Value = ws.Range(FAC_TOT).Value
    wsH.Cells(fh, 10).Value = ws.Range(FAC_PAGO).Value
    wsH.Cells(fh, 11).Value = ws.Range(FAC_VENCE).Value
    wsH.Cells(fh, 12).Value = costoTot
    wsH.Cells(fh, 13).Value = ws.Range(FAC_SUB).Value - costoTot
    wsH.Cells(fh, 14).Value = rutas
    wsH.Cells(fh, 15).Value = Now
    FormatearFilasHist wsH

    '---------- 5. CARTERA (crédito) o CAJA (contado) ----------
    If EsCredito(ws.Range(FAC_PAGO).Value) Then
        n = SiguienteFila(wsCart, 1)
        wsCart.Cells(n, 1).Value = numTxt
        wsCart.Cells(n, 2).NumberFormat = "@"
        wsCart.Cells(n, 2).Value = CStr(ws.Range(FAC_CLI).Value)
        wsCart.Cells(n, 3).Value = cliente
        wsCart.Cells(n, 4).Value = fecha
        wsCart.Cells(n, 5).Value = ws.Range(FAC_VENCE).Value
        wsCart.Cells(n, 6).Value = ws.Range(FAC_TOT).Value
        AsegurarFormulasCartera wsCart, n
    Else
        n = SiguienteFila(wsCaja, 1)
        wsCaja.Cells(n, 1).Value = fecha
        wsCaja.Cells(n, 2).Value = "Venta contado"
        wsCaja.Cells(n, 4).Value = "Venta factura " & numTxt
        wsCaja.Cells(n, 5).Value = numTxt
        wsCaja.Cells(n, 6).Value = cliente
        wsCaja.Cells(n, 7).Value = ws.Range(FAC_TOT).Value
        AsegurarFormulasCaja wsCaja, n
    End If

    '---------- 6. CIERRE ----------
    ws.Range(FAC_ESTADO).Value = "GUARDADA"
    Application.ScreenUpdating = True
    ThisWorkbook.Save

    If MsgBox("Factura " & numTxt & " guardada correctamente." & vbCrLf & vbCrLf & _
              "Copia en: " & rutas & vbCrLf & vbCrLf & _
              "¿Desea crear una NUEVA factura?", vbYesNo + vbInformation, "Guardar factura") = vbYes Then
        NuevaFactura
    End If
    Exit Sub

Fallo:
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbCritical, "No se pudo guardar la factura"
End Sub

'--- Limpieza interna (sin preguntar) -----------------------------------------
Private Sub LimpiarFacturaInterno(ws As Worksheet)
    Application.ScreenUpdating = False
    BorrarSoloValores ws.Range(FAC_CLI)
    BorrarSoloValores ws.Range("B" & FAC_F1 & ":B" & FAC_F2)
    BorrarSoloValores ws.Range("D" & FAC_F1 & ":D" & FAC_F2)
    ws.Range(FAC_OBS).MergeArea.ClearContents
    ws.Range(FAC_FECHA).Value = Date
    ws.Range(FAC_PAGO).Value = "Contado"
    RestaurarFormulasFactura ws
    ws.Range(FAC_ESTADO).Value = "BORRADOR"
    Application.ScreenUpdating = True
End Sub

'--- Reescribe las fórmulas de la factura por si alguien las sobrescribió ------
Private Sub RestaurarFormulasFactura(ws As Worksheet)
    Dim a As String, b As String
    a = FAC_F1: b = FAC_F2
    ws.Range("C" & a & ":C" & b).Formula = "=IF(B" & a & "="""","""",IFERROR(INDEX(PRODUCTOS!$B:$B,MATCH(B" & a & ",PRODUCTOS!$A:$A,0)),""SKU NO EXISTE""))"
    ws.Range("E" & a & ":E" & b).Formula = "=IF(B" & a & "="""","""",IFERROR(INDEX(PRODUCTOS!$F:$F,MATCH(B" & a & ",PRODUCTOS!$A:$A,0)),0))"
    ws.Range("F" & a & ":F" & b).Formula = "=IF(B" & a & "="""","""",D" & a & "*E" & a & ")"
    ws.Range("G" & a & ":G" & b).Formula = "=IF(B" & a & "="""","""",IF(IFERROR(INDEX(PRODUCTOS!$I:$I,MATCH(B" & a & ",PRODUCTOS!$A:$A,0)),""NO"")=""SI"",ROUND(F" & a & "*CONFIG!$B$12,0),0))"
    ws.Range("H" & a & ":H" & b).Formula = "=IF(B" & a & "="""","""",F" & a & "+G" & a & ")"
    ws.Range("I" & a & ":I" & b).Formula = "=IF(B" & a & "="""","""",IFERROR(INDEX(INVENTARIO!$G:$G,MATCH(B" & a & ",INVENTARIO!$A:$A,0)),""""))"
    ws.Range(FAC_SUB).Formula = "=SUM(F" & a & ":F" & b & ")"
    ws.Range(FAC_IVA).Formula = "=SUM(G" & a & ":G" & b & ")"
    ws.Range(FAC_TOT).Formula = "=SUM(H" & a & ":H" & b & ")"
    ws.Range(FAC_NUMTXT).Formula = "=CONFIG!$B$13&H3"
End Sub

'==============================================================================
'                               GUÍAS DE ENVÍO
'==============================================================================

'--- BOTÓN: CARGAR DATOS DESDE LA FACTURA ASOCIADA ----------------------------
Public Sub CargarFacturaEnGuia()
    Dim ws As Worksheet, wsH As Worksheet, wsD As Worksheet, wsC As Worksheet
    Dim fac As String, m As Variant, mc As Variant, r As Long, fila As Long, idCli As String

    Set ws = ThisWorkbook.Worksheets(H_GUIA)
    Set wsH = ThisWorkbook.Worksheets(H_HFAC)
    Set wsD = ThisWorkbook.Worksheets(H_HDET)
    Set wsC = ThisWorkbook.Worksheets(H_CLI)

    fac = Trim$(CStr(ws.Range(GUI_FACT).Value))
    If fac = "" Then MsgBox "Escriba o seleccione la factura asociada en " & GUI_FACT & ".", vbExclamation: Exit Sub
    m = Application.Match(fac, wsH.Columns(2), 0)
    If IsError(m) Then MsgBox "La factura " & fac & " no está en HIST_FACTURAS (¿la guardó?).", vbExclamation: Exit Sub

    Application.ScreenUpdating = False
    idCli = CStr(wsH.Cells(m, 4).Value)
    ws.Range(GUI_NOMBRE).Value = wsH.Cells(m, 5).Value
    ws.Range(GUI_ID).NumberFormat = "@"
    ws.Range(GUI_ID).Value = idCli
    ws.Range(GUI_CIUDAD).Value = wsH.Cells(m, 6).Value
    ws.Range(GUI_VALOR).Value = wsH.Cells(m, 7).Value
    mc = Application.Match(idCli, wsC.Columns(1), 0)
    If Not IsError(mc) Then
        ws.Range(GUI_TEL).NumberFormat = "@"
        ws.Range(GUI_TEL).Value = CStr(wsC.Cells(mc, 3).Value)
        ws.Range(GUI_DIR).Value = wsC.Cells(mc, 5).Value
        ws.Range(GUI_DEPTO).Value = wsC.Cells(mc, 7).Value
    End If

    BorrarSoloValores ws.Range("B" & GUI_F1 & ":B" & GUI_F2)
    BorrarSoloValores ws.Range("D" & GUI_F1 & ":D" & GUI_F2)
    BorrarSoloValores ws.Range("F" & GUI_F1 & ":H" & GUI_F2)
    fila = GUI_F1
    For r = FILA_DATOS To SiguienteFila(wsD, 1) - 1
        If CStr(wsD.Cells(r, 1).Value) = fac Then
            If fila > GUI_F2 Then
                MsgBox "La factura tiene más de " & (GUI_F2 - GUI_F1 + 1) & " ítems; solo se cargaron los primeros.", vbInformation
                Exit For
            End If
            ws.Cells(fila, "B").NumberFormat = "@"
            ws.Cells(fila, "B").Value = CStr(wsD.Cells(r, 3).Value)
            ws.Cells(fila, "D").Value = wsD.Cells(r, 5).Value
            fila = fila + 1
        End If
    Next r
    Application.ScreenUpdating = True
    MsgBox "Datos de la factura " & fac & " cargados. Complete transportadora, bultos y rastreo.", vbInformation
End Sub

'--- BOTÓN: IMPRIMIR GUÍA -----------------------------------------------------
Public Sub ImprimirGuia()
    Dim ws As Worksheet, resp As VbMsgBoxResult
    Set ws = ThisWorkbook.Worksheets(H_GUIA)
    ConfigurarPagina ws, GUI_AREA
    resp = MsgBox("SÍ = Imprimir   |   NO = Vista previa   |   CANCELAR = Salir", _
                  vbYesNoCancel + vbQuestion, "Imprimir guía " & ws.Range(GUI_NUMTXT).Value)
    Select Case resp
        Case vbYes: ws.PrintOut Copies:=1
        Case vbNo:  ws.PrintPreview
    End Select
End Sub

'--- BOTÓN: LIMPIAR GUÍA ------------------------------------------------------
Public Sub LimpiarGuia()
    If MsgBox("¿Borrar los datos de la guía actual?", vbYesNo + vbQuestion, "Limpiar guía") = vbNo Then Exit Sub
    LimpiarGuiaInterno ThisWorkbook.Worksheets(H_GUIA)
End Sub

'--- BOTÓN: NUEVA GUÍA --------------------------------------------------------
Public Sub NuevaGuia()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(H_GUIA)
    If ws.Range(GUI_ESTADO).Value <> "GUARDADA" And TieneItems(ws, GUI_F1, GUI_F2) Then
        If MsgBox("La guía actual no se ha guardado. ¿Descartarla?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If
    LimpiarGuiaInterno ws
    ws.Range(GUI_NUM).Value = SiguienteConsecutivo(H_HGUIA, ThisWorkbook.Worksheets(H_CFG).Range("B16").Value)
    ws.Activate
    ws.Range(GUI_FACT).Select
End Sub

'--- BOTÓN: GUARDAR GUÍA (historial + PDF en el Escritorio) -------------------
Public Sub GuardarGuia()
    Dim ws As Worksheet, wsH As Worksheet, n As Long, numero As Long
    Dim numTxt As String, archivo As String, fecha As Date

    On Error GoTo Fallo
    Set ws = ThisWorkbook.Worksheets(H_GUIA)
    Set wsH = ThisWorkbook.Worksheets(H_HGUIA)
    Application.Calculate

    numero = Val(ws.Range(GUI_NUM).Value)
    numTxt = CStr(ws.Range(GUI_NUMTXT).Value)
    If numero <= 0 Then Err.Raise vbObjectError + 20, , "No hay número de guía. Use NUEVA GUÍA."
    If Not IsError(Application.Match(numero, wsH.Columns(1), 0)) Then _
        Err.Raise vbObjectError + 21, , "La guía " & numTxt & " ya fue guardada. Use NUEVA GUÍA."
    If Not IsDate(ws.Range(GUI_FECHA).Value) Then Err.Raise vbObjectError + 22, , "Digite la fecha de despacho."
    If Trim$(CStr(ws.Range(GUI_NOMBRE).Value)) = "" Or Trim$(CStr(ws.Range(GUI_DIR).Value)) = "" _
       Or Trim$(CStr(ws.Range(GUI_CIUDAD).Value)) = "" Then _
        Err.Raise vbObjectError + 23, , "Faltan datos del destinatario (nombre, dirección o ciudad)."
    If Trim$(CStr(ws.Range(GUI_TRANSP).Value)) = "" Then Err.Raise vbObjectError + 24, , "Seleccione la transportadora."
    If Not TieneItems(ws, GUI_F1, GUI_F2) Then Err.Raise vbObjectError + 25, , "La guía no tiene productos."
    fecha = CDate(ws.Range(GUI_FECHA).Value)

    Application.ScreenUpdating = False
    archivo = CarpetaDestino("Guias", fecha) & "\" & LimpiarNombre(numTxt & "_" & ws.Range(GUI_NOMBRE).Value) & ".pdf"
    ConfigurarPagina ws, GUI_AREA
    ExportarPDF ws, archivo

    n = SiguienteFila(wsH, 1)
    wsH.Cells(n, 1).Value = numero
    wsH.Cells(n, 2).Value = numTxt
    wsH.Cells(n, 3).Value = fecha
    wsH.Cells(n, 4).Value = ws.Range(GUI_FACT).Value
    wsH.Cells(n, 5).Value = ws.Range(GUI_NOMBRE).Value
    wsH.Cells(n, 6).Value = ws.Range(GUI_CIUDAD).Value
    wsH.Cells(n, 7).Value = ws.Range(GUI_TRANSP).Value
    wsH.Cells(n, 8).Value = ws.Range(GUI_RASTREO).Value
    wsH.Cells(n, 9).Value = ws.Range(GUI_BULTOS).Value
    wsH.Cells(n, 10).Value = archivo
    wsH.Cells(n, 11).Value = Now
    FormatearFilasHist wsH

    ws.Range(GUI_ESTADO).Value = "GUARDADA"
    Application.ScreenUpdating = True
    ThisWorkbook.Save
    MsgBox "Guía " & numTxt & " guardada." & vbCrLf & "PDF: " & archivo, vbInformation, "Guardar guía"
    Exit Sub
Fallo:
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbCritical, "No se pudo guardar la guía"
End Sub

Private Sub LimpiarGuiaInterno(ws As Worksheet)
    Application.ScreenUpdating = False
    BorrarSoloValores ws.Range(GUI_FACT & "," & GUI_VALOR & "," & GUI_TEL & "," & GUI_DEPTO)
    ws.Range(GUI_TRANSP).MergeArea.ClearContents
    ws.Range(GUI_RASTREO).MergeArea.ClearContents
    ws.Range(GUI_NOMBRE).MergeArea.ClearContents
    ws.Range(GUI_ID).MergeArea.ClearContents
    ws.Range(GUI_DIR).MergeArea.ClearContents
    ws.Range(GUI_CIUDAD).MergeArea.ClearContents
    ws.Range(GUI_OBS).MergeArea.ClearContents
    BorrarSoloValores ws.Range("B" & GUI_F1 & ":B" & GUI_F2)
    BorrarSoloValores ws.Range("D" & GUI_F1 & ":D" & GUI_F2)
    BorrarSoloValores ws.Range("F" & GUI_F1 & ":H" & GUI_F2)
    ws.Range(GUI_FECHA).Value = Date
    ws.Range(GUI_ESTADO).Value = "BORRADOR"
    Application.ScreenUpdating = True
End Sub

'==============================================================================
'                     FUNCIONES DE APOYO (no se asignan a botones)
'==============================================================================

'Ruta: <Escritorio del usuario>\<carpeta CONFIG B17>\<subcarpeta>\<AAAA-MM>
'Usa la carpeta especial "Desktop" de Windows, que funciona aunque el Escritorio
'esté redirigido a OneDrive; si falla, usa %USERPROFILE%\Desktop.
Private Function CarpetaDestino(ByVal subcarpeta As String, ByVal fecha As Date) As String
    Dim escritorio As String, base As String
    On Error Resume Next
    escritorio = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    On Error GoTo 0
    If escritorio = "" Then escritorio = Environ$("USERPROFILE") & "\Desktop"
    base = Trim$(CStr(ThisWorkbook.Worksheets(H_CFG).Range("B17").Value))
    If base = "" Then base = "Gestion_Empresa"
    CarpetaDestino = escritorio & "\" & LimpiarNombre(base) & "\" & subcarpeta & "\" & Format$(fecha, "yyyy-mm")
    CrearCarpeta CarpetaDestino
End Function

'Crea la carpeta y todas las intermedias que no existan
Private Sub CrearCarpeta(ByVal ruta As String)
    Dim fso As Object, padre As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(ruta) Then Exit Sub
    padre = fso.GetParentFolderName(ruta)
    If padre <> "" And Not fso.FolderExists(padre) Then CrearCarpeta padre
    fso.CreateFolder ruta
End Sub

Private Sub ExportarPDF(ws As Worksheet, ByVal archivo As String)
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=archivo, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End Sub

'Copia la hoja a un libro nuevo, la convierte en valores, quita botones,
'validaciones, formatos condicionales y vínculos, y la guarda como .xlsx
Private Sub ExportarExcel(ws As Worksheet, ByVal archivo As String, ByVal colsAux As String)
    Dim wb As Workbook, sh As Worksheet, l As Variant
    ws.Copy
    Set wb = ActiveWorkbook
    Set sh = wb.Worksheets(1)
    On Error Resume Next
    sh.Unprotect Password:=CLAVE
    On Error GoTo 0
    sh.UsedRange.Copy
    sh.UsedRange.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    On Error Resume Next
    Dim k As Long
    For k = sh.Shapes.Count To 1 Step -1: sh.Shapes(k).Delete: Next k
    sh.Cells.Validation.Delete
    sh.Cells.FormatConditions.Delete
    sh.Columns(colsAux).Delete
    If Not IsEmpty(wb.LinkSources(xlExcelLinks)) Then
        For Each l In wb.LinkSources(xlExcelLinks)
            wb.BreakLink Name:=l, Type:=xlLinkTypeExcelLinks
        Next l
    End If
    On Error GoTo 0
    sh.Range("A1").Select
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=archivo, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ThisWorkbook.Activate
End Sub

Private Sub ConfigurarPagina(ws As Worksheet, ByVal area As String)
    On Error Resume Next             'por si no hay impresora instalada
    With ws.PageSetup
        .PrintArea = area
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .PrintGridlines = False
    End With
    On Error GoTo 0
End Sub

'Siguiente consecutivo = mayor número guardado + 1 (o el inicial de CONFIG)
Private Function SiguienteConsecutivo(ByVal hojaHist As String, ByVal inicial As Variant) As Long
    Dim mx As Double
    mx = Application.WorksheetFunction.Max(ThisWorkbook.Worksheets(hojaHist).Columns(1))
    If mx < Val(inicial) - 1 Then mx = Val(inicial) - 1
    SiguienteConsecutivo = CLng(mx) + 1
End Function

'Primera fila vacía de una tabla según la columna indicada (mínimo FILA_DATOS)
Private Function SiguienteFila(ws As Worksheet, ByVal col As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, col).End(xlUp).Row + 1
    If r < FILA_DATOS Then r = FILA_DATOS
    SiguienteFila = r
End Function

Private Function TieneItems(ws As Worksheet, ByVal f1 As Long, ByVal f2 As Long) As Boolean
    TieneItems = Application.WorksheetFunction.CountA(ws.Range("B" & f1 & ":B" & f2)) > 0
End Function

Private Function EsCredito(ByVal v As Variant) As Boolean
    EsCredito = (LCase$(Left$(Trim$(CStr(v)), 2)) = "cr")
End Function

'Borra solo constantes: nunca toca una celda que tenga fórmula
Private Sub BorrarSoloValores(rng As Range)
    Dim c As Range
    For Each c In rng.Cells
        If Not c.HasFormula Then
            If c.MergeCells Then c.MergeArea.ClearContents Else c.ClearContents
        End If
    Next c
End Sub

Private Function LimpiarNombre(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, ch, "_")
    Next ch
    LimpiarNombre = Left$(Trim$(s), 80)
End Function

'Copia el formato de la fila 5 a las filas nuevas del historial
Private Sub FormatearFilasHist(ws As Worksheet)
    Dim ultima As Long, ncol As Long
    ultima = SiguienteFila(ws, 1) - 1
    If ultima <= FILA_DATOS Then Exit Sub
    ncol = ws.Cells(FILA_DATOS - 1, ws.Columns.Count).End(xlToLeft).Column
    ws.Range(ws.Cells(FILA_DATOS, 1), ws.Cells(FILA_DATOS, ncol)).Copy
    ws.Range(ws.Cells(FILA_DATOS + 1, 1), ws.Cells(ultima, ncol)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub

'Si la fila supera las fórmulas precargadas, las copia de la fila de arriba
Private Sub AsegurarFormulasCartera(ws As Worksheet, ByVal fila As Long)
    If Not ws.Cells(fila, "G").HasFormula Then ws.Range("G" & fila - 1 & ":K" & fila - 1).Copy ws.Range("G" & fila)
End Sub

Private Sub AsegurarFormulasCaja(ws As Worksheet, ByVal fila As Long)
    If Not ws.Cells(fila, "C").HasFormula Then
        ws.Range("C" & fila - 1).Copy ws.Range("C" & fila)
        ws.Range("H" & fila - 1 & ":J" & fila - 1).Copy ws.Range("H" & fila)
    End If
End Sub

'==============================================================================
'           CREAR BOTONES (ejecutar UNA vez con Alt+F8 > CrearBotones)
'==============================================================================
Public Sub CrearBotones()
    CrearBotonesHoja ThisWorkbook.Worksheets(H_FAC), _
        Array("IMPRIMIR FACTURA", "LIMPIAR FACTURA", "GUARDAR (Historial + PDF)", "NUEVA FACTURA"), _
        Array("ImprimirFactura", "LimpiarFactura", "GuardarFactura", "NuevaFactura")
    CrearBotonesHoja ThisWorkbook.Worksheets(H_GUIA), _
        Array("CARGAR DESDE FACTURA", "IMPRIMIR GUÍA", "LIMPIAR GUÍA", "GUARDAR GUÍA (PDF)", "NUEVA GUÍA"), _
        Array("CargarFacturaEnGuia", "ImprimirGuia", "LimpiarGuia", "GuardarGuia", "NuevaGuia")
    MsgBox "Botones creados en FACTURACION y GUIAS.", vbInformation
End Sub

Private Sub CrearBotonesHoja(ws As Worksheet, textos As Variant, macros As Variant)
    Dim i As Long, b As Object, celda As Range
    For i = ws.Shapes.Count To 1 Step -1            'elimina botones anteriores
        If Left$(ws.Shapes(i).Name, 4) = "btn_" Then ws.Shapes(i).Delete
    Next i
    For i = LBound(textos) To UBound(textos)
        Set celda = ws.Range("K" & (5 + i * 3))
        Set b = ws.Buttons.Add(celda.Left + 2, celda.Top, celda.Width - 4, 34)
        b.Name = "btn_" & macros(i)
        b.Caption = textos(i)
        b.OnAction = macros(i)
        b.Font.Bold = True
        b.Font.Size = 10
        b.PrintObject = False                       'el botón nunca se imprime
        b.Placement = xlFreeFloating
    Next i
End Sub

'==============================================================================
'      PROTECCIÓN OPCIONAL: bloquea fórmulas; solo se editan celdas amarillas
'==============================================================================
Public Sub ProtegerHojas()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect Password:=CLAVE, UserInterfaceOnly:=True, AllowFiltering:=True, _
                   AllowFormattingColumns:=True
    Next ws
End Sub

Public Sub DesprotegerHojas()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:=CLAVE
    Next ws
End Sub
