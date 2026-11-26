Attribute VB_Name = "modFactura"
'==============================================================
' modFactura - FACTURACIÓN
'  NUEVA · GUARDAR · IMPRIMIR · LIMPIAR · BUSCAR · ANULAR
'  Reglas de integridad:
'   - El número solo avanza al GUARDAR (sin saltos por borradores)
'   - Una factura guardada no puede volver a guardarse ni descontar
'     inventario dos veces (se verifica historial y kardex)
'   - Guardar es una transacción: si algo falla se revierte todo
'==============================================================
Option Explicit

Public Function NumeroFactura(ByVal n As Long) As String
    NumeroFactura = Tv(CfgV(CFG_PREFFACTURA)) & Format(n, "0000")
End Function

Private Function CodigoDeCombo(ByVal combo As String) As String
    Dim p As Long
    p = InStr(combo, " |")
    If p > 0 Then
        CodigoDeCombo = Trim(Left(combo, p - 1))
    Else
        CodigoDeCombo = Trim(combo)
    End If
End Function

Private Function ComboCliente(ByVal cod As String) As String
    Dim ws As Worksheet, f As Long
    Set ws = Hoja(SH_CLI)
    f = BuscarFila(ws, CLI_CODIGO, CLI_FIRST, cod)
    If f > 0 Then ComboCliente = Tv(ws.Cells(f, CLI_CODIGO).Value) & " | " & Tv(ws.Cells(f, CLI_NOMBRE).Value)
End Function

Private Function TieneDatos(ByVal ws As Worksheet) As Boolean
    Dim r As Long
    For r = FAC_L1 To FAC_LN
        If Tv(ws.Cells(r, 2).Value) <> "" Then TieneDatos = True: Exit Function
    Next r
End Function

' Restaura las fórmulas de precio unitario (columna G) desde la plantilla oculta (columna Y)
Private Sub RestaurarFormulasPrecio(ByVal ws As Worksheet)
    Dim r As Long, f As String
    For r = FAC_L1 To FAC_LN
        f = "=IF($B" & r & "="""","""",IFERROR(INDEX('" & SH_INV & "'!$K$" & INV_FIRST & ":$K$" & INV_LAST & _
            ",MATCH($B" & r & ",'" & SH_INV & "'!$A$" & INV_FIRST & ":$A$" & INV_LAST & ",0)),0))"
        If Left(CStr(ws.Cells(r, 7).Formula), 1) <> "=" Then ws.Cells(r, 7).Formula = f
    Next r
End Sub

Public Sub LimpiarCamposFactura(ByVal ws As Worksheet)
    Dim a As Variant
    ws.Range("B" & FAC_L1 & ":B" & FAC_LN).ClearContents
    ws.Range("F" & FAC_L1 & ":F" & FAC_LN).ClearContents
    ws.Range("H" & FAC_L1 & ":H" & FAC_LN).ClearContents
    For Each a In Array(FAC_ENTREGA, FAC_PAGADO, FAC_OBS)
        ws.Range(a).Value = ""
    Next a
    RestaurarFormulasPrecio ws
End Sub

'-------------------------------------------------------------- NUEVA
Public Sub NuevaFactura()
    NuevaFacturaInterna True
End Sub

Public Sub NuevaFacturaInterna(ByVal confirmar As Boolean)
    Dim ws As Worksheet
    Set ws = Hoja(SH_FAC)
    If confirmar And Tv(ws.Range(FAC_ESTADO).Value) = "BORRADOR" And TieneDatos(ws) Then
        If Msg("La factura actual no se ha guardado. ¿Descartarla y empezar una nueva?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    Iniciar
    LimpiarCamposFactura ws
    ws.Range(FAC_NUM).Value = NumeroFactura(CLng(Nv(CfgV(CFG_SIGFACTURA))))
    ws.Range(FAC_FECHA).Value = Date
    ws.Range(FAC_ESTADO).Value = "BORRADOR"
    ws.Range(FAC_METODO).Value = "Efectivo"
    ws.Range(FAC_ESTADOPAGO).Value = "PAGADA"
    ws.Range(FAC_CLIENTE).Value = ComboCliente("CL-0001")
    Terminar
    On Error Resume Next
    If Not gTest Then ws.Activate: ws.Range(FAC_CLIENTE).Select
End Sub

'-------------------------------------------------------------- LIMPIAR
Public Sub LimpiarFactura()
    Dim ws As Worksheet
    Set ws = Hoja(SH_FAC)
    If Tv(ws.Range(FAC_ESTADO).Value) <> "BORRADOR" Then
        NuevaFacturaInterna False
        Exit Sub
    End If
    If TieneDatos(ws) Then
        If Msg("¿Borrar los datos capturados de esta factura?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    Iniciar
    LimpiarCamposFactura ws
    Terminar
End Sub

'-------------------------------------------------------------- GUARDAR
Public Sub GuardarFactura()
    Dim ws As Worksheet, wi As Worksheet, wc As Worksheet, wb As Worksheet, wh As Worksheet
    Dim num As String, esperado As String, codcli As String, fCli As Long, r As Long, i As Long, j As Long
    Dim nL As Long, sku() As String, cant() As Double, desc() As Double, fInv() As Long, fila() As Long
    Dim errores As String, bajoCosto As String, faltante As String, q As Double, totSku As Double, st As Double
    Dim metodo As String, estPago As String, total As Double, pagado As Double, txtPag As String
    Dim fecha As Date, nombre As String, entrega As String, bruto As Double, costoT As Double
    Dim pu As Double, subt As Double, cu As Double, ct As Double, a As Variant, rr As Long, ruta As String
    Dim vence As Date

    Set ws = Hoja(SH_FAC)
    Recalcular
    If Tv(ws.Range(FAC_ESTADO).Value) <> "BORRADOR" Then
        Msg "Esta factura ya está " & Tv(ws.Range(FAC_ESTADO).Value) & ". No se puede guardar otra vez." & vbLf & _
            "Pulse NUEVA FACTURA para registrar otra venta.", vbExclamation
        Exit Sub
    End If
    num = Tv(ws.Range(FAC_NUM).Value)
    esperado = NumeroFactura(CLng(Nv(CfgV(CFG_SIGFACTURA))))
    If num <> esperado Then
        Msg "El número de la factura (" & num & ") no corresponde al consecutivo vigente (" & esperado & ")." & vbLf & _
            "Pulse NUEVA FACTURA.", vbExclamation
        Exit Sub
    End If
    Set wh = Hoja(SH_HFA)
    If BuscarFila(wh, HFA_FACTURA, HFA_FIRST, num) > 0 Or ExisteMovimiento(num, "VENTA") Then
        Msg "La factura " & num & " YA EXISTE en el historial. No se guardará dos veces.", vbCritical
        Exit Sub
    End If
    ' ---- cliente
    codcli = CodigoDeCombo(Tv(ws.Range(FAC_CLIENTE).Value))
    Set wc = Hoja(SH_CLI)
    fCli = BuscarFila(wc, CLI_CODIGO, CLI_FIRST, codcli)
    If codcli = "" Or fCli = 0 Then
        Msg "Seleccione un cliente válido de la lista (o créelo con NUEVO CLIENTE).", vbExclamation
        Exit Sub
    End If
    ' ---- líneas
    Set wi = Hoja(SH_INV)
    ReDim sku(1 To 15): ReDim cant(1 To 15): ReDim desc(1 To 15): ReDim fInv(1 To 15): ReDim fila(1 To 15)
    For r = FAC_L1 To FAC_LN
        If Tv(ws.Cells(r, 2).Value) <> "" Then
            nL = nL + 1
            sku(nL) = UCase(Tv(ws.Cells(r, 2).Value))
            fila(nL) = r
            fInv(nL) = BuscarFila(wi, INV_SKU, INV_FIRST, sku(nL))
            q = Nv(ws.Cells(r, 6).Value)
            cant(nL) = q
            desc(nL) = Nv(ws.Cells(r, 8).Value)
            If fInv(nL) = 0 Then
                errores = errores & vbLf & "• Línea " & (r - FAC_L1 + 1) & ": el código " & sku(nL) & " no existe."
            ElseIf Tv(wi.Cells(fInv(nL), INV_ACTIVO).Value) = "NO" Then
                errores = errores & vbLf & "• Línea " & (r - FAC_L1 + 1) & ": el producto " & sku(nL) & " está INACTIVO."
            End If
            If q <= 0 Or q <> Int(q) Then errores = errores & vbLf & "• Línea " & (r - FAC_L1 + 1) & ": cantidad no válida."
            If desc(nL) < 0 Or desc(nL) > 1 Then errores = errores & vbLf & "• Línea " & (r - FAC_L1 + 1) & ": descuento fuera de 0% - 100%."
        ElseIf Tv(ws.Cells(r, 6).Value) <> "" Or Tv(ws.Cells(r, 8).Value) <> "" Then
            errores = errores & vbLf & "• Línea " & (r - FAC_L1 + 1) & ": tiene cantidad/descuento pero no tiene código."
        End If
    Next r
    If nL = 0 Then errores = errores & vbLf & "• La factura no tiene productos."
    If errores <> "" Then
        Msg "No se puede guardar la factura. Corrija lo siguiente:" & errores, vbExclamation
        Exit Sub
    End If
    ' ---- existencias (sumando el mismo SKU en varias líneas)
    For i = 1 To nL
        totSku = 0
        For j = 1 To nL
            If sku(j) = sku(i) Then totSku = totSku + cant(j)
        Next j
        st = StockReal(sku(i))
        If totSku > st And InStr(faltante, sku(i) & ":") = 0 Then
            faltante = faltante & vbLf & "• " & sku(i) & ": pide " & totSku & ", hay " & st
        End If
        pu = Nv(ws.Cells(fila(i), 7).Value)
        If pu * (1 - desc(i)) < Nv(wi.Cells(fInv(i), INV_COSTOREAL).Value) Then
            bajoCosto = bajoCosto & vbLf & "• " & sku(i) & ": precio neto " & Moneda(pu * (1 - desc(i))) & _
                        " < costo real " & Moneda(Nv(wi.Cells(fInv(i), INV_COSTOREAL).Value))
        End If
    Next i
    If faltante <> "" Then
        If UCase(Tv(CfgV(CFG_PERMITIRNEG))) <> "SI" Then
            If Msg("STOCK INSUFICIENTE:" & faltante & vbLf & vbLf & _
                   "El inventario quedaría NEGATIVO. ¿Desea autorizar la venta con clave?", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
            If Not Autorizar("Vender con inventario negativo en la factura " & num) Then Exit Sub
        End If
    End If
    If bajoCosto <> "" Then
        If Msg("ALERTA: productos vendidos por DEBAJO DEL COSTO REAL:" & bajoCosto & vbLf & vbLf & _
               "Esta venta genera pérdida. ¿Desea continuar?", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    End If
    ' ---- pago
    metodo = Tv(ws.Range(FAC_METODO).Value)
    estPago = UCase(Tv(ws.Range(FAC_ESTADOPAGO).Value))
    total = Nv(ws.Range(FAC_TOTAL).Value)
    txtPag = Tv(ws.Range(FAC_PAGADO).Value)
    pagado = Nv(ws.Range(FAC_PAGADO).Value)
    If Not EnLista(LST_METODOS, metodo) Then Msg "Seleccione un método de pago válido.", vbExclamation: Exit Sub
    If total <= 0 Then Msg "El total de la factura debe ser mayor que cero.", vbExclamation: Exit Sub
    Select Case estPago
        Case "PAGADA"
            If txtPag <> "" And Abs(pagado - total) > 0.5 Then
                Msg "Si la factura está PAGADA, el valor pagado debe ser igual al total (" & Moneda(total) & ") o quedar vacío." & vbLf & _
                    "Si el cliente pagó una parte, elija ABONO PARCIAL.", vbExclamation
                Exit Sub
            End If
            pagado = total
        Case "PENDIENTE"
            If pagado <> 0 Then Msg "Una factura PENDIENTE no debe tener valor pagado. Use ABONO PARCIAL.", vbExclamation: Exit Sub
            pagado = 0
        Case "ABONO PARCIAL"
            If pagado <= 0 Or pagado >= total Then
                Msg "Para ABONO PARCIAL escriba un valor pagado mayor que 0 y menor que el total (" & Moneda(total) & ").", vbExclamation
                Exit Sub
            End If
        Case Else
            Msg "Seleccione el estado de pago (PAGADA, PENDIENTE o ABONO PARCIAL).", vbExclamation
            Exit Sub
    End Select
    If Not IsDate(ws.Range(FAC_FECHA).Value) Then Msg "Escriba una fecha de expedición válida.", vbExclamation: Exit Sub
    fecha = CDate(ws.Range(FAC_FECHA).Value)
    If fecha > Date Then
        If Msg("La fecha de la factura es futura (" & FechaCorta(fecha) & "). ¿Continuar?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    If Msg("¿Guardar la factura " & num & " por " & Moneda(total) & "?" & vbLf & _
           "Se descontará el inventario y se registrará el pago (" & estPago & ").", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    ' ================= escritura (transacción) =================
    Iniciar
    On Error GoTo Fallo
    TxIniciar
    TxTabla SH_BD, BD_FACTURA, BD_FIRST, 1, BD_NCOLS
    TxTabla SH_HFA, HFA_FACTURA, HFA_FIRST, 1, HFA_NCOLS
    TxTabla SH_MOV, MOV_ID, MOV_FIRST, 1, MOV_NCOLS
    TxTabla SH_ING, ING_ID, ING_FIRST, 1, ING_NCOLS
    TxTabla SH_CAR, CAR_ID, CAR_FIRST, 1, CAR_NCOLS
    TxValor ws, FAC_ESTADO
    nombre = Tv(wc.Cells(fCli, CLI_NOMBRE).Value)
    entrega = Tv(ws.Range(FAC_ENTREGA).Value)
    If entrega = "" Then entrega = Tv(wc.Cells(fCli, CLI_DIRECCION).Value)
    Set wb = Hoja(SH_BD)
    For i = 1 To nL
        pu = Nv(ws.Cells(fila(i), 7).Value)
        subt = Round(cant(i) * pu * (1 - desc(i)), 0)
        cu = Round(Nv(wi.Cells(fInv(i), INV_COSTOREAL).Value), 2)
        ct = Round(cant(i) * cu, 0)
        rr = NuevaFila(wb, BD_FACTURA, BD_FIRST, BD_LAST)
        a = NuevaFilaArr(BD_NCOLS)
        a(1, BD_FACTURA) = num
        a(1, BD_FECHA) = fecha
        a(1, BD_CODCLI) = codcli
        a(1, BD_CLIENTE) = nombre
        a(1, BD_IDCLI) = Tv(wc.Cells(fCli, CLI_IDENT).Value)
        a(1, BD_SKU) = Tv(wi.Cells(fInv(i), INV_SKU).Value)
        a(1, BD_PRODUCTO) = Tv(wi.Cells(fInv(i), INV_NOMBRE).Value)
        a(1, BD_CATEGORIA) = Tv(wi.Cells(fInv(i), INV_CATEGORIA).Value)
        a(1, BD_TALLA) = Tv(wi.Cells(fInv(i), INV_TALLA).Value)
        a(1, BD_COLOR) = Tv(wi.Cells(fInv(i), INV_COLOR).Value)
        a(1, BD_CANT) = cant(i)
        a(1, BD_PUNIT) = pu
        a(1, BD_DESCPCT) = desc(i)
        a(1, BD_DESCVAL) = cant(i) * pu - subt
        a(1, BD_SUBTOTAL) = subt
        a(1, BD_COSTOUNIT) = cu
        a(1, BD_COSTOTOTAL) = ct
        a(1, BD_UTILIDAD) = subt - ct
        a(1, BD_METODO) = metodo
        a(1, BD_ESTADOPAGO) = estPago
        a(1, BD_ESTADO) = "ACTIVA"
        a(1, BD_USUARIO) = UsuarioActual()
        a(1, BD_REGISTRO) = Now
        EscribirFila wb, rr, 1, a
        RegistrarMovimiento "VENTA", num, sku(i), 0, cant(i), "Venta a " & nombre, cu
        bruto = bruto + cant(i) * pu
        costoT = costoT + ct
    Next i
    If estPago = "PAGADA" Then
        vence = fecha
    Else
        vence = fecha + Nv(CfgV(CFG_DIASCREDITO))
    End If
    rr = NuevaFila(wh, HFA_FACTURA, HFA_FIRST, HFA_LAST)
    a = NuevaFilaArr(HFA_NCOLS)
    a(1, HFA_FACTURA) = num
    a(1, HFA_FECHA) = fecha
    a(1, HFA_CODCLI) = codcli
    a(1, HFA_CLIENTE) = nombre
    a(1, HFA_IDCLI) = Tv(wc.Cells(fCli, CLI_IDENT).Value)
    a(1, HFA_TEL) = Tv(wc.Cells(fCli, CLI_TEL).Value)
    a(1, HFA_DIRECCION) = entrega
    a(1, HFA_CIUDAD) = Tv(wc.Cells(fCli, CLI_CIUDAD).Value)
    a(1, HFA_DEPTO) = Tv(wc.Cells(fCli, CLI_DEPTO).Value)
    a(1, HFA_BRUTO) = bruto
    a(1, HFA_DESCUENTOS) = bruto - total
    a(1, HFA_TOTAL) = total
    a(1, HFA_METODO) = metodo
    a(1, HFA_ESTADOPAGO) = estPago
    a(1, HFA_PAGADO) = pagado
    a(1, HFA_VENCE) = vence
    a(1, HFA_ESTADO) = "ACTIVA"
    a(1, HFA_OBS) = Tv(ws.Range(FAC_OBS).Value)
    a(1, HFA_COSTO) = costoT
    a(1, HFA_UTILIDAD) = total - costoT
    a(1, HFA_USUARIO) = UsuarioActual()
    a(1, HFA_REGISTRO) = Now
    EscribirFila wh, rr, 1, a
    If pagado > 0 Then RegistrarIngreso fecha, "FACTURACIÓN", "VENTA", metodo, num, nombre, "Venta factura " & num, pagado, ""
    If total - pagado > 0 Then CrearObligacion "POR COBRAR", codcli, nombre, num, fecha, total, pagado, vence, ""
    CfgSet CFG_SIGFACTURA, CLng(Nv(CfgV(CFG_SIGFACTURA))) + 1
    ws.Range(FAC_ESTADO).Value = "GUARDADA"
    Auditar "FACTURACIÓN", "GUARDAR FACTURA", num, nombre & " · " & Moneda(total) & " · " & estPago
    TxConfirmar
    Terminar
    ' ---- copia PDF (fuera de la transacción: si falla, la venta igual queda registrada)
    ruta = ExportarPDF(ws, "FACTURAS", num & " - " & nombre)
    If ruta <> "" Then
        DesprotegerTodo
        wh.Cells(rr, HFA_PDF).Value = ruta
        ProtegerTodo
    End If
    Msg "Factura " & num & " guardada." & vbLf & _
        "Total: " & Moneda(total) & "   Pagado: " & Moneda(pagado) & "   Saldo: " & Moneda(total - pagado) & vbLf & _
        IIf(ruta <> "", "PDF: " & ruta, IIf(gTest, "", "(No se pudo generar el PDF; puede imprimirla con IMPRIMIR FACTURA)")), vbInformation
    Exit Sub
Fallo:
    Dim e As String
    e = Err.Description
    TxRevertir
    Terminar
    Msg "Ocurrió un error y la factura NO se guardó (se revirtieron todos los cambios)." & vbLf & e, vbCritical
End Sub

'-------------------------------------------------------------- IMPRIMIR
Public Sub ImprimirFactura()
    Dim ws As Worksheet
    Set ws = Hoja(SH_FAC)
    If Not TieneDatos(ws) Then Msg "La factura no tiene productos.", vbExclamation: Exit Sub
    If Tv(ws.Range(FAC_ESTADO).Value) = "BORRADOR" Then
        If Msg("La factura aún NO está guardada (borrador). ¿Imprimir de todos modos?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
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

'-------------------------------------------------------------- BUSCAR
Public Sub BuscarFactura()
    Dim txt As String, wh As Worksheet, f As Long, ult As Long, r As Long, d As Date, lista As String
    Dim n As Long, ultNum As String, crit As String
    txt = Pedir("Buscar factura. Escriba:" & vbLf & "• El número (ej. FV-0012)" & vbLf & _
               "• Una fecha (dd/mm/aaaa)" & vbLf & "• Parte del nombre o código del cliente", "Buscar factura")
    If txt = "" Then Exit Sub
    Set wh = Hoja(SH_HFA)
    f = BuscarFila(wh, HFA_FACTURA, HFA_FIRST, txt)
    If f = 0 And IsNumeric(txt) Then f = BuscarFila(wh, HFA_FACTURA, HFA_FIRST, NumeroFactura(CLng(txt)))
    If f > 0 Then
        CargarFactura Tv(wh.Cells(f, HFA_FACTURA).Value)
        Exit Sub
    End If
    d = FechaTexto(txt)
    crit = UCase(txt)
    ult = UltimaFila(wh, HFA_FACTURA, HFA_FIRST)
    For r = HFA_FIRST To ult
        If (CDbl(d) > 0 And Int(Nv(wh.Cells(r, HFA_FECHA).Value)) = CDbl(d)) Or _
           (CDbl(d) = 0 And (InStr(UCase(Tv(wh.Cells(r, HFA_CLIENTE).Value)), crit) > 0 Or _
                              UCase(Tv(wh.Cells(r, HFA_CODCLI).Value)) = crit)) Then
            n = n + 1
            ultNum = Tv(wh.Cells(r, HFA_FACTURA).Value)
            If n <= 25 Then lista = lista & vbLf & ultNum & "  " & FechaCorta(wh.Cells(r, HFA_FECHA).Value) & "  " & _
                                    Left(Tv(wh.Cells(r, HFA_CLIENTE).Value), 22) & "  " & Moneda(Nv(wh.Cells(r, HFA_TOTAL).Value)) & _
                                    IIf(Tv(wh.Cells(r, HFA_ESTADO).Value) = "ANULADA", "  [ANULADA]", "")
        End If
    Next r
    If n = 0 Then
        Msg "No se encontraron facturas para: " & txt, vbInformation
    ElseIf n = 1 Then
        CargarFactura ultNum
    Else
        txt = Pedir(n & " facturas encontradas" & IIf(n > 25, " (se muestran 25)", "") & ":" & lista & vbLf & vbLf & _
                   "Escriba el número de la factura que desea abrir:", "Resultados", ultNum)
        If txt <> "" Then
            If BuscarFila(wh, HFA_FACTURA, HFA_FIRST, txt) > 0 Then
                CargarFactura UCase(txt)
            Else
                Msg "Número no encontrado.", vbExclamation
            End If
        End If
    End If
End Sub

' Muestra en el formato una factura guardada (con los precios históricos)
Public Sub CargarFactura(ByVal num As String)
    Dim ws As Worksheet, wh As Worksheet, wb As Worksheet, f As Long, ult As Long, r As Long, k As Long
    Set ws = Hoja(SH_FAC)
    Set wh = Hoja(SH_HFA)
    Set wb = Hoja(SH_BD)
    f = BuscarFila(wh, HFA_FACTURA, HFA_FIRST, num)
    If f = 0 Then Msg "No existe la factura " & num, vbExclamation: Exit Sub
    Iniciar
    LimpiarCamposFactura ws
    ws.Range(FAC_NUM).Value = Tv(wh.Cells(f, HFA_FACTURA).Value)
    ws.Range(FAC_FECHA).Value = wh.Cells(f, HFA_FECHA).Value
    ws.Range(FAC_CLIENTE).Value = Tv(wh.Cells(f, HFA_CODCLI).Value) & " | " & Tv(wh.Cells(f, HFA_CLIENTE).Value)
    ws.Range(FAC_ENTREGA).Value = Tv(wh.Cells(f, HFA_DIRECCION).Value)
    ws.Range(FAC_METODO).Value = Tv(wh.Cells(f, HFA_METODO).Value)
    ws.Range(FAC_ESTADOPAGO).Value = Tv(wh.Cells(f, HFA_ESTADOPAGO).Value)
    If Tv(wh.Cells(f, HFA_ESTADOPAGO).Value) = "ABONO PARCIAL" Then ws.Range(FAC_PAGADO).Value = Nv(wh.Cells(f, HFA_PAGADO).Value)
    ws.Range(FAC_OBS).Value = Tv(wh.Cells(f, HFA_OBS).Value)
    ult = UltimaFila(wb, BD_FACTURA, BD_FIRST)
    k = FAC_L1
    For r = BD_FIRST To ult
        If UCase(Tv(wb.Cells(r, BD_FACTURA).Value)) = UCase(num) And Nv(wb.Cells(r, BD_CANT).Value) > 0 And k <= FAC_LN Then
            ws.Cells(k, 2).Value = Tv(wb.Cells(r, BD_SKU).Value)
            ws.Cells(k, 6).Value = Nv(wb.Cells(r, BD_CANT).Value)
            ws.Cells(k, 7).Value = Nv(wb.Cells(r, BD_PUNIT).Value)   ' precio histórico
            If Nv(wb.Cells(r, BD_DESCPCT).Value) > 0 Then ws.Cells(k, 8).Value = Nv(wb.Cells(r, BD_DESCPCT).Value)
            k = k + 1
        End If
    Next r
    If Tv(wh.Cells(f, HFA_ESTADO).Value) = "ANULADA" Then
        ws.Range(FAC_ESTADO).Value = "ANULADA"
    Else
        ws.Range(FAC_ESTADO).Value = "GUARDADA"
    End If
    Terminar
    On Error Resume Next
    If Not gTest Then ws.Activate
End Sub

'-------------------------------------------------------------- ANULAR
Public Sub AnularFactura()
    Dim ws As Worksheet, wh As Worksheet, wb As Worksheet, wi As Worksheet, wc As Worksheet
    Dim num As String, f As Long, ult As Long, r As Long, motivo As String, q As Double, n As Long
    Set ws = Hoja(SH_FAC)
    If Tv(ws.Range(FAC_ESTADO).Value) = "GUARDADA" Then num = Tv(ws.Range(FAC_NUM).Value)
    num = UCase(Pedir("Número de la factura a ANULAR:", "Anular factura", num))
    If num = "" Then Exit Sub
    Set wh = Hoja(SH_HFA)
    f = BuscarFila(wh, HFA_FACTURA, HFA_FIRST, num)
    If f = 0 Then Msg "La factura " & num & " no existe.", vbExclamation: Exit Sub
    If Tv(wh.Cells(f, HFA_ESTADO).Value) = "ANULADA" Then Msg "La factura " & num & " ya estaba ANULADA.", vbExclamation: Exit Sub
    If Msg("¿ANULAR la factura " & num & "?" & vbLf & _
           "Cliente: " & Tv(wh.Cells(f, HFA_CLIENTE).Value) & vbLf & "Total: " & Moneda(Nv(wh.Cells(f, HFA_TOTAL).Value)) & vbLf & vbLf & _
           "• Se devolverán las unidades al inventario" & vbLf & "• Se anularán los ingresos y la cartera asociados" & vbLf & _
           "• El registro se conserva en el historial marcado como ANULADA", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    If Not Autorizar("Anular la factura " & num) Then Exit Sub
    motivo = Pedir("Motivo de la anulación (obligatorio):", "Anular factura")
    If motivo = "" Then Msg "Debe indicar el motivo. La anulación fue cancelada.", vbExclamation: Exit Sub
    CopiaSeguridadInterna True, "antes de anular " & num
    Iniciar
    On Error GoTo Fallo
    TxIniciar
    TxTabla SH_MOV, MOV_ID, MOV_FIRST, 1, MOV_NCOLS
    Set wb = Hoja(SH_BD)
    ult = UltimaFila(wb, BD_FACTURA, BD_FIRST)
    For r = BD_FIRST To ult
        If UCase(Tv(wb.Cells(r, BD_FACTURA).Value)) = num And Tv(wb.Cells(r, BD_ESTADO).Value) = "ACTIVA" Then
            q = Nv(wb.Cells(r, BD_CANT).Value)
            Modificar wb, r, BD_ESTADO, "ANULADA"
            If q > 0 Then
                RegistrarMovimiento "ANULACIÓN VENTA", num, Tv(wb.Cells(r, BD_SKU).Value), q, 0, "Anulación " & num & ": " & motivo, Nv(wb.Cells(r, BD_COSTOUNIT).Value)
            ElseIf q < 0 Then
                RegistrarMovimiento "ANULACIÓN DEVOLUCIÓN", num, Tv(wb.Cells(r, BD_SKU).Value), 0, -q, "Anulación " & num & ": " & motivo, Nv(wb.Cells(r, BD_COSTOUNIT).Value)
            End If
            n = n + 1
        End If
    Next r
    Modificar wh, f, HFA_ESTADO, "ANULADA"
    Modificar wh, f, HFA_ANULUSR, UsuarioActual()
    Modificar wh, f, HFA_ANULFECHA, Now
    Modificar wh, f, HFA_ANULMOTIVO, motivo
    Set wi = Hoja(SH_ING)
    ult = UltimaFila(wi, ING_ID, ING_FIRST)
    For r = ING_FIRST To ult
        If UCase(Tv(wi.Cells(r, ING_DOC).Value)) = num And Tv(wi.Cells(r, ING_ESTADO).Value) = "ACTIVO" Then
            Modificar wi, r, ING_ESTADO, "ANULADO"
        End If
    Next r
    Set wc = Hoja(SH_CAR)
    ult = UltimaFila(wc, CAR_ID, CAR_FIRST)
    For r = CAR_FIRST To ult
        If UCase(Tv(wc.Cells(r, CAR_DOC).Value)) = num Then Modificar wc, r, CAR_REGESTADO, "ANULADA"
    Next r
    If UCase(Tv(ws.Range(FAC_NUM).Value)) = num Then Modificar ws, ws.Range(FAC_ESTADO).Row, ws.Range(FAC_ESTADO).Column, "ANULADA"
    Auditar "FACTURACIÓN", "ANULAR FACTURA", num, motivo
    TxConfirmar
    Terminar
    Msg "Factura " & num & " ANULADA." & vbLf & n & " línea(s) revertidas; las unidades regresaron al inventario.", vbInformation
    Exit Sub
Fallo:
    Dim e As String
    e = Err.Description
    TxRevertir
    Terminar
    Msg "No se pudo anular (no se cambió nada): " & e, vbCritical
End Sub

'-------------------------------------------------------------- cliente rápido
Public Sub NuevoCliente()
    Dim ws As Worksheet, r As Long, cod As String
    Set ws = Hoja(SH_CLI)
    QuitarFiltros ws
    r = UltimaFila(ws, CLI_CODIGO, CLI_FIRST) + 1
    If r > CLI_LAST Then Msg "La tabla de clientes está llena.", vbExclamation: Exit Sub
    cod = SiguienteID(ws, CLI_CODIGO, CLI_FIRST, "CL-", 4)
    Iniciar
    ws.Cells(r, CLI_CODIGO).Value = cod
    Terminar
    If gTest Then Exit Sub
    ws.Activate
    ws.Cells(r, CLI_NOMBRE).Select
    Msg "Se creó el código " & cod & " en la fila " & r & "." & vbLf & _
        "Complete nombre, identificación, teléfono, WhatsApp, dirección, ciudad y departamento." & vbLf & _
        "Luego vuelva a FACTURACIÓN y selecciónelo en la lista de clientes.", vbInformation
End Sub

Public Sub VerHistorialCliente()
    Dim cod As String, ws As Worksheet
    On Error Resume Next
    If ActiveSheet.Name = SH_CLI And ActiveCell.Row >= CLI_FIRST Then cod = Tv(Hoja(SH_CLI).Cells(ActiveCell.Row, CLI_CODIGO).Value)
    On Error GoTo 0
    cod = Pedir("Código del cliente:", "Historial de facturas", cod)
    If cod = "" Then Exit Sub
    Set ws = Hoja(SH_HFA)
    ws.Activate
    On Error Resume Next
    ws.Unprotect Tv(CfgV(CFG_CLAVEHOJAS))
    QuitarFiltros ws
    ws.Range(ws.Cells(HFA_HDR, 1), ws.Cells(HFA_HDR, HFA_NCOLS)).AutoFilter Field:=HFA_CODCLI, Criteria1:=cod
    ProtegerTodo
End Sub
