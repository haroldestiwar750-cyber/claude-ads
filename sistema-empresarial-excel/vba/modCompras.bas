Attribute VB_Name = "modCompras"
'==============================================================
' modCompras - REGISTRO DE COMPRAS DE MERCANCÍA
'  Una compra: aumenta el inventario (kardex), recalcula el costo
'  promedio ponderado, reparte fletes/otros costos por valor,
'  registra el pago en EGRESOS y el saldo en CARTERA (por pagar).
'==============================================================
Option Explicit

Public Function NumeroCompra(ByVal n As Long) As String
    NumeroCompra = Tv(CfgV(CFG_PREFCOMPRA)) & Format(n, "0000")
End Function

Public Sub RegistrarCompra()
    Dim ws As Worksheet
    Dim wi As Worksheet
    Dim wp As Worksheet
    Dim wc As Worksheet
    Dim num As String
    Dim fecha As Variant
    Dim prov As String
    Dim forma As String
    Dim estPago As String
    Dim transp As Double
    Dim otros As Double
    Dim pagado As Double
    Dim txtPag As String
    Dim vence As Variant
    Dim r As Long
    Dim nL As Long
    Dim sku() As String
    Dim cant() As Double
    Dim costo() As Double
    Dim fInv() As Long
    Dim errores As String
    Dim subtotal As Double
    Dim total As Double
    Dim i As Long
    Dim share As Double
    Dim tL As Double
    Dim oL As Double
    Dim costoRealL As Double
    Dim st As Double
    Dim costoPrevio As Double
    Dim cNuevo As Double
    Dim rr As Long
    Dim a As Variant
    Dim fProv As Long
    Dim codProv As String
    Dim saldo As Double

    Set ws = Hoja(SH_COM)
    Recalcular
    num = NumeroCompra(CLng(Nv(CfgV(CFG_SIGCOMPRA))))
    fecha = ws.Range(FCOM_FECHA).Value
    prov = Tv(ws.Range(FCOM_PROV).Value)
    forma = Tv(ws.Range(FCOM_FORMA).Value)
    estPago = UCase(Tv(ws.Range(FCOM_ESTADOPAGO).Value))
    transp = Nv(ws.Range(FCOM_TRANSP).Value)
    otros = Nv(ws.Range(FCOM_OTROS).Value)
    txtPag = Tv(ws.Range(FCOM_PAGADO).Value)
    pagado = Nv(ws.Range(FCOM_PAGADO).Value)
    vence = ws.Range(FCOM_VENCE).Value
    If Not IsDate(fecha) Then errores = errores & vbLf & "• Fecha no válida."
    Set wp = Hoja(SH_PRO)
    fProv = BuscarFila(wp, PRO_NOMBRE, PRO_FIRST, prov)
    If prov = "" Or fProv = 0 Then
        errores = errores & vbLf & "• Seleccione un proveedor registrado."
    Else
        codProv = Tv(wp.Cells(fProv, PRO_CODIGO).Value)
    End If
    If Not EnLista(LST_METODOS, forma) Then errores = errores & vbLf & "• Forma de pago no válida."
    If transp < 0 Or otros < 0 Then errores = errores & vbLf & "• Transporte y otros costos no pueden ser negativos."
    Set wi = Hoja(SH_INV)
    ReDim sku(1 To 20): ReDim cant(1 To 20): ReDim costo(1 To 20): ReDim fInv(1 To 20)
    For r = FCOM_L1 To FCOM_LN
        If Tv(ws.Cells(r, 2).Value) <> "" Then
            nL = nL + 1
            sku(nL) = UCase(Tv(ws.Cells(r, 2).Value))
            fInv(nL) = BuscarFila(wi, INV_SKU, INV_FIRST, sku(nL))
            cant(nL) = Nv(ws.Cells(r, 6).Value)
            costo(nL) = Nv(ws.Cells(r, 7).Value)
            If fInv(nL) = 0 Then errores = errores & vbLf & "• Línea " & (r - FCOM_L1 + 1) & ": el SKU " & sku(nL) & " no existe (créelo en INVENTARIO)."
            If cant(nL) <= 0 Or cant(nL) <> Int(cant(nL)) Then errores = errores & vbLf & "• Línea " & (r - FCOM_L1 + 1) & ": cantidad no válida."
            If Tv(ws.Cells(r, 7).Value) = "" Or costo(nL) < 0 Then errores = errores & vbLf & "• Línea " & (r - FCOM_L1 + 1) & ": falta el costo unitario."
            subtotal = subtotal + cant(nL) * costo(nL)
        ElseIf Tv(ws.Cells(r, 6).Value) <> "" Or Tv(ws.Cells(r, 7).Value) <> "" Then
            errores = errores & vbLf & "• Línea " & (r - FCOM_L1 + 1) & ": tiene cantidad/costo pero no tiene SKU."
        End If
    Next r
    If nL = 0 Then errores = errores & vbLf & "• La compra no tiene productos."
    total = subtotal + transp + otros
    Select Case estPago
        Case "PAGADA"
            If txtPag <> "" And Abs(pagado - total) > 0.5 Then errores = errores & vbLf & "• Si está PAGADA, el valor pagado debe ser el total (" & Moneda(total) & ") o quedar vacío."
            pagado = total
        Case "PENDIENTE"
            If pagado <> 0 Then errores = errores & vbLf & "• Una compra PENDIENTE no lleva valor pagado (use ABONO PARCIAL)."
            pagado = 0
        Case "ABONO PARCIAL"
            If pagado <= 0 Or pagado >= total Then errores = errores & vbLf & "• En ABONO PARCIAL el valor pagado debe estar entre 0 y el total."
        Case Else
            errores = errores & vbLf & "• Seleccione el estado de pago."
    End Select
    If errores <> "" Then
        Msg "No se puede registrar la compra:" & errores, vbExclamation
        Exit Sub
    End If
    If Not IsDate(vence) Then vence = CDate(fecha) + Nv(CfgV(CFG_DIASCREDITO))
    If Msg("¿Registrar la compra " & num & " a " & prov & "?" & vbLf & _
           nL & " producto(s) · Total " & Moneda(total) & " · " & estPago, vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    Iniciar
    On Error GoTo Fallo
    TxIniciar
    TxTabla SH_COM, COM_NUMERO, COM_FIRST, 1, COM_NCOLS
    TxTabla SH_MOV, MOV_ID, MOV_FIRST, 1, MOV_NCOLS
    TxTabla SH_EGR, EGR_ID, EGR_FIRST, 1, EGR_NCOLS
    TxTabla SH_CAR, CAR_ID, CAR_FIRST, 1, CAR_NCOLS
    If BuscarFila(ws, COM_NUMERO, COM_FIRST, num) > 0 Then Err.Raise vbObjectError + 700, , "La compra " & num & " ya existe."
    For i = 1 To nL
        If subtotal > 0 Then
            share = cant(i) * costo(i) / subtotal
        Else
            share = 1 / nL
        End If
        tL = Round(transp * share, 0)
        oL = Round(otros * share, 0)
        costoRealL = cant(i) * costo(i) + tL + oL
        ' costo promedio ponderado
        st = StockReal(sku(i))
        costoPrevio = Nv(wi.Cells(fInv(i), INV_COSTO).Value)
        If st > 0 Then
            cNuevo = Round((st * costoPrevio + cant(i) * costo(i)) / (st + cant(i)), 2)
        Else
            cNuevo = costo(i)
        End If
        TxValor wi, wi.Cells(fInv(i), INV_COSTO).Address(False, False)
        wi.Cells(fInv(i), INV_COSTO).Value = cNuevo
        rr = NuevaFila(ws, COM_NUMERO, COM_FIRST, COM_LAST)
        a = NuevaFilaArr(COM_NCOLS)
        a(1, COM_NUMERO) = num
        a(1, COM_FECHA) = CDate(fecha)
        a(1, COM_PROVEEDOR) = prov
        a(1, COM_SKU) = Tv(wi.Cells(fInv(i), INV_SKU).Value)
        a(1, COM_PRODUCTO) = Tv(wi.Cells(fInv(i), INV_NOMBRE).Value) & " " & Tv(wi.Cells(fInv(i), INV_TALLA).Value) & " " & Tv(wi.Cells(fInv(i), INV_COLOR).Value)
        a(1, COM_CANT) = cant(i)
        a(1, COM_COSTOUNIT) = costo(i)
        a(1, COM_COSTOTOTAL) = cant(i) * costo(i)
        a(1, COM_TRANSPORTE) = tL
        a(1, COM_OTROS) = oL
        a(1, COM_COSTOREAL) = costoRealL
        a(1, COM_COSTOREALU) = Round(costoRealL / cant(i), 2)
        a(1, COM_FORMAPAGO) = forma
        a(1, COM_ESTADOPAGO) = estPago
        a(1, COM_USUARIO) = UsuarioActual()
        a(1, COM_REGISTRO) = Now
        a(1, COM_ESTADO) = "ACTIVA"
        a(1, COM_OBS) = Tv(ws.Range(FCOM_OBS).Value)
        EscribirFila ws, rr, 1, a
        RegistrarMovimiento "COMPRA", num, sku(i), cant(i), 0, "Compra a " & prov, Round(costoRealL / cant(i), 2)
    Next i
    If pagado > 0 Then RegistrarEgreso CDate(fecha), "Compra " & num & " - " & prov, Tv(CfgV(CFG_CATCOMPRAS)), "VARIABLE", _
                                       pagado, forma, UsuarioActual(), num, "", "COMPRAS", num
    saldo = total - pagado
    If saldo > 0 Then CrearObligacion "POR PAGAR", codProv, prov, num, CDate(fecha), total, pagado, CDate(vence), ""
    CfgSet CFG_SIGCOMPRA, CLng(Nv(CfgV(CFG_SIGCOMPRA))) + 1
    Auditar "COMPRAS", "REGISTRAR COMPRA", num, prov & " · " & Moneda(total) & " · " & estPago
    TxConfirmar
    LimpiarFormularioCompra ws
    Terminar
    Msg "Compra " & num & " registrada." & vbLf & "El inventario y los costos fueron actualizados." & _
        IIf(saldo > 0, vbLf & "Saldo por pagar: " & Moneda(saldo) & " (ver CARTERA).", ""), vbInformation
    Exit Sub
Fallo:
    Dim e As String
    e = Err.Description
    TxRevertir
    Terminar
    Msg "La compra NO se registró (se revirtieron los cambios): " & e, vbCritical
End Sub

Private Sub LimpiarFormularioCompra(ByVal ws As Worksheet)
    ws.Range("B" & FCOM_L1 & ":B" & FCOM_LN).ClearContents
    ws.Range("F" & FCOM_L1 & ":G" & FCOM_LN).ClearContents
    ws.Range(FCOM_PAGADO).Value = ""
    ws.Range(FCOM_TRANSP).Value = ""
    ws.Range(FCOM_OTROS).Value = ""
    ws.Range(FCOM_VENCE).Value = ""
    ws.Range(FCOM_OBS).Value = ""
End Sub

Public Sub LimpiarCompra()
    If Msg("¿Limpiar el formulario de compra?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Iniciar
    LimpiarFormularioCompra Hoja(SH_COM)
    Terminar
End Sub

Public Sub NuevoProveedor()
    Dim ws As Worksheet, r As Long, cod As String
    Set ws = Hoja(SH_PRO)
    QuitarFiltros ws
    r = UltimaFila(ws, PRO_NOMBRE, PRO_FIRST) + 1
    If r > PRO_LAST Then Msg "La tabla de proveedores está llena.", vbExclamation: Exit Sub
    cod = SiguienteID(ws, PRO_CODIGO, PRO_FIRST, "PR-", 3)
    Iniciar
    ws.Cells(r, PRO_CODIGO).Value = cod
    Terminar
    If gTest Then Exit Sub
    ws.Activate
    ws.Cells(r, PRO_NOMBRE).Select
    Msg "Código " & cod & " asignado en la fila " & r & ". Complete los datos del proveedor.", vbInformation
End Sub
