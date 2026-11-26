Attribute VB_Name = "modInventario"
'==============================================================
' modInventario - KARDEX Y OPERACIONES DE INVENTARIO
'  Todo cambio de existencias pasa por RegistrarMovimiento, que
'  deja huella en MOVIMIENTOS INV (quién, cuándo, por qué).
'  El stock de INVENTARIO es una fórmula sobre esos movimientos,
'  por eso nunca se descuenta dos veces ni se "pierde" una unidad.
'==============================================================
Option Explicit

' Registra un movimiento. Devuelve el stock resultante.
Public Function RegistrarMovimiento(ByVal tipo As String, ByVal doc As String, ByVal sku As String, _
                                    ByVal entrada As Double, ByVal salida As Double, ByVal motivo As String, _
                                    Optional ByVal costoUnit As Double = -1) As Double
    Dim ws As Worksheet, wi As Worksheet, fi As Long, r As Long, a As Variant, saldo As Double
    Set wi = Hoja(SH_INV)
    fi = BuscarFila(wi, INV_SKU, INV_FIRST, sku)
    If fi = 0 Then Err.Raise vbObjectError + 610, "RegistrarMovimiento", "El SKU " & sku & " no existe en INVENTARIO."
    If entrada < 0 Or salida < 0 Then Err.Raise vbObjectError + 611, "RegistrarMovimiento", "Cantidades negativas no permitidas."
    If costoUnit < 0 Then costoUnit = Nv(wi.Cells(fi, INV_COSTO).Value) + Nv(wi.Cells(fi, INV_TRANSP).Value)
    saldo = StockReal(sku) + entrada - salida
    Set ws = Hoja(SH_MOV)
    r = NuevaFila(ws, MOV_ID, MOV_FIRST, MOV_LAST)
    a = NuevaFilaArr(MOV_NCOLS)
    a(1, MOV_ID) = SiguienteID(ws, MOV_ID, MOV_FIRST, "MOV-", 6)
    a(1, MOV_FECHA) = Now
    a(1, MOV_TIPO) = tipo
    a(1, MOV_DOC) = doc
    a(1, MOV_SKU) = Tv(wi.Cells(fi, INV_SKU).Value)
    a(1, MOV_PRODUCTO) = Tv(wi.Cells(fi, INV_NOMBRE).Value) & " " & Tv(wi.Cells(fi, INV_TALLA).Value) & " " & Tv(wi.Cells(fi, INV_COLOR).Value)
    a(1, MOV_ENTRADA) = entrada
    a(1, MOV_SALIDA) = salida
    a(1, MOV_SALDO) = saldo
    a(1, MOV_COSTOUNIT) = Round(costoUnit, 2)
    a(1, MOV_VALOR) = Round((entrada - salida) * costoUnit, 0)
    a(1, MOV_USUARIO) = UsuarioActual()
    a(1, MOV_MOTIVO) = motivo
    EscribirFila ws, r, 1, a
    RegistrarMovimiento = saldo
End Function

Public Function ExisteMovimiento(ByVal doc As String, ByVal tipo As String) As Boolean
    Dim ws As Worksheet, ult As Long, aD As Variant, aT As Variant, i As Long
    Set ws = Hoja(SH_MOV)
    ult = UltimaFila(ws, MOV_ID, MOV_FIRST)
    If ult < MOV_FIRST Then Exit Function
    aD = LeerColumna(ws, MOV_DOC, MOV_FIRST, ult)
    aT = LeerColumna(ws, MOV_TIPO, MOV_FIRST, ult)
    For i = 1 To UBound(aD, 1)
        If UCase(Tv(aD(i, 1))) = UCase(doc) And UCase(Tv(aT(i, 1))) = UCase(tipo) Then
            ExisteMovimiento = True
            Exit Function
        End If
    Next i
End Function

' SKU de la fila seleccionada en INVENTARIO (o vacío)
Private Function SkuSeleccionado() As String
    On Error Resume Next
    If ActiveSheet.Name = SH_INV Then
        If ActiveCell.Row >= INV_FIRST Then SkuSeleccionado = Tv(Hoja(SH_INV).Cells(ActiveCell.Row, INV_SKU).Value)
    End If
    On Error GoTo 0
End Function

'-------------------------------------------------------------- botones
Public Sub NuevoProducto()
    Dim ws As Worksheet, r As Long
    Set ws = Hoja(SH_INV)
    ws.Activate
    QuitarFiltros ws
    r = UltimaFila(ws, INV_SKU, INV_FIRST) + 1
    If r > INV_LAST Then
        Msg "El inventario alcanzó su capacidad. Consulte el manual para ampliarlo.", vbExclamation
        Exit Sub
    End If
    Iniciar
    ws.Cells(r, INV_ACTIVO).Value = "SI"
    If Tv(ws.Cells(r, INV_STOCKINI).Value) = "" Then ws.Cells(r, INV_STOCKINI).Value = 0
    Terminar
    ws.Cells(r, INV_SKU).Select
    Msg "Complete los datos del nuevo producto en la fila " & r & ":" & vbLf & _
        "SKU único, nombre, categoría, talla, color, proveedor, costo, precio, stock inicial y stock mínimo." & vbLf & vbLf & _
        "Consejo: use un SKU por cada combinación de talla y color (ej. CAM-BAS-NEG-M)." & vbLf & _
        "Para mercancía nueva es mejor dejar stock inicial en 0 y registrar una COMPRA.", vbInformation
End Sub

Public Sub VerMovimientosSKU()
    Dim sku As String, ws As Worksheet
    sku = SkuSeleccionado()
    If sku = "" Then sku = Pedir("Escriba el SKU cuyo historial desea consultar:", "Movimientos de inventario")
    If sku = "" Then Exit Sub
    Set ws = Hoja(SH_MOV)
    ws.Activate
    On Error Resume Next
    ws.Unprotect Tv(CfgV(CFG_CLAVEHOJAS))
    QuitarFiltros ws
    ws.Range(ws.Cells(MOV_HDR, 1), ws.Cells(MOV_HDR, MOV_NCOLS)).AutoFilter Field:=MOV_SKU, Criteria1:=sku
    ProtegerTodo
    On Error GoTo 0
End Sub

Public Sub QuitarFiltrosHoja()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    On Error Resume Next
    ws.Unprotect Tv(CfgV(CFG_CLAVEHOJAS))
    QuitarFiltros ws
    ProtegerTodo
End Sub

' Ajuste manual de existencias (conteo físico, pérdida, daño). Requiere clave.
Public Sub AjusteInventario()
    Dim sku As String, s As String, cant As Double, motivo As String, st As Double, fi As Long
    sku = SkuSeleccionado()
    sku = Pedir("SKU a ajustar:", "Ajuste de inventario", sku)
    If sku = "" Then Exit Sub
    fi = BuscarFila(Hoja(SH_INV), INV_SKU, INV_FIRST, sku)
    If fi = 0 Then Msg "El SKU " & sku & " no existe.", vbExclamation: Exit Sub
    st = StockReal(sku)
    s = Pedir("Stock actual de " & sku & ": " & st & vbLf & vbLf & _
              "Escriba la cantidad a AJUSTAR (positiva para sumar, negativa para restar):", "Ajuste de inventario")
    If s = "" Then Exit Sub
    cant = NumTexto(s)
    If cant = 0 Or cant <> Int(cant) Then Msg "Cantidad no válida (debe ser un entero distinto de cero).", vbExclamation: Exit Sub
    motivo = Pedir("Motivo del ajuste (conteo físico, daño, pérdida, obsequio...):", "Ajuste de inventario")
    If motivo = "" Then Msg "El motivo es obligatorio.", vbExclamation: Exit Sub
    If st + cant < 0 Then
        If Msg("El ajuste dejaría el inventario en " & (st + cant) & " (NEGATIVO). ¿Desea continuar?", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    End If
    If Not Autorizar("Ajuste de inventario de " & sku & " por " & cant & " unidades") Then Exit Sub
    Iniciar
    On Error GoTo Fallo
    TxIniciar
    TxTabla SH_MOV, MOV_ID, MOV_FIRST, 1, MOV_NCOLS
    If cant > 0 Then
        RegistrarMovimiento "AJUSTE (+)", "AJUSTE", sku, cant, 0, motivo
    Else
        RegistrarMovimiento "AJUSTE (-)", "AJUSTE", sku, 0, -cant, motivo
    End If
    Auditar "INVENTARIO", "AJUSTE", sku, cant & " unidades. " & motivo
    TxConfirmar
    Terminar
    Msg "Ajuste registrado. Nuevo stock de " & sku & ": " & (st + cant), vbInformation
    Exit Sub
Fallo:
    TxRevertir
    Terminar
    Msg "No se pudo registrar el ajuste: " & Err.Description, vbCritical
End Sub

' Devolución de un cliente sobre una factura. Reingresa unidades, registra una línea negativa
' en ventas y descuenta el valor del saldo pendiente o lo registra como dinero devuelto.
Public Sub RegistrarDevolucion()
    Dim num As String, wh As Worksheet, wb As Worksheet, fh As Long, ult As Long, r As Long
    Dim sku As String, s As String, cant As Double, vendidas As Double, devueltas As Double
    Dim fLinea As Long, pu As Double, d As Double, cu As Double, valor As Double, motivo As String
    Dim saldoCxC As Double, nota As Double, dinero As Double, metodo As String, filaCar As Long, lista As String
    Dim a As Variant, wsC As Worksheet
    num = UCase(Pedir("Número de la factura sobre la que se hace la devolución (ej. FV-0012):", "Devolución de cliente"))
    If num = "" Then Exit Sub
    Set wh = Hoja(SH_HFA)
    fh = BuscarFila(wh, HFA_FACTURA, HFA_FIRST, num)
    If fh = 0 Then Msg "La factura " & num & " no existe.", vbExclamation: Exit Sub
    If Tv(wh.Cells(fh, HFA_ESTADO).Value) <> "ACTIVA" Then Msg "La factura " & num & " está ANULADA.", vbExclamation: Exit Sub
    Set wb = Hoja(SH_BD)
    ult = UltimaFila(wb, BD_FACTURA, BD_FIRST)
    For r = BD_FIRST To ult
        If UCase(Tv(wb.Cells(r, BD_FACTURA).Value)) = num And Nv(wb.Cells(r, BD_CANT).Value) > 0 Then
            lista = lista & vbLf & "  " & Tv(wb.Cells(r, BD_SKU).Value) & "  (" & wb.Cells(r, BD_CANT).Value & " und.)"
        End If
    Next r
    sku = UCase(Pedir("Productos de la factura " & num & ":" & lista & vbLf & vbLf & "Escriba el SKU devuelto:", "Devolución de cliente"))
    If sku = "" Then Exit Sub
    For r = BD_FIRST To ult
        If UCase(Tv(wb.Cells(r, BD_FACTURA).Value)) = num And UCase(Tv(wb.Cells(r, BD_SKU).Value)) = sku _
           And Tv(wb.Cells(r, BD_ESTADO).Value) = "ACTIVA" Then
            If Nv(wb.Cells(r, BD_CANT).Value) > 0 Then
                vendidas = vendidas + Nv(wb.Cells(r, BD_CANT).Value)
                If fLinea = 0 Then fLinea = r
            Else
                devueltas = devueltas - Nv(wb.Cells(r, BD_CANT).Value)
            End If
        End If
    Next r
    If fLinea = 0 Then Msg "El SKU " & sku & " no está en la factura " & num & ".", vbExclamation: Exit Sub
    s = Pedir("Vendidas: " & vendidas & "   Ya devueltas: " & devueltas & vbLf & "¿Cuántas unidades devuelve el cliente?", "Devolución de cliente")
    If s = "" Then Exit Sub
    cant = NumTexto(s)
    If cant <= 0 Or cant <> Int(cant) Or cant > vendidas - devueltas Then
        Msg "Cantidad no válida. Máximo permitido: " & (vendidas - devueltas), vbExclamation
        Exit Sub
    End If
    motivo = Pedir("Motivo de la devolución:", "Devolución de cliente", "Cambio / defecto")
    If motivo = "" Then Exit Sub
    pu = Nv(wb.Cells(fLinea, BD_PUNIT).Value)
    d = Nv(wb.Cells(fLinea, BD_DESCPCT).Value)
    cu = Nv(wb.Cells(fLinea, BD_COSTOUNIT).Value)
    valor = Round(cant * pu * (1 - d), 0)
    ' ¿cuánto se descuenta del saldo y cuánto se devuelve en dinero?
    Set wsC = Hoja(SH_CAR)
    filaCar = BuscarFila(wsC, CAR_DOC, CAR_FIRST, num)
    If filaCar > 0 Then
        If Tv(wsC.Cells(filaCar, CAR_REGESTADO).Value) = "ACTIVA" Then saldoCxC = Nv(wsC.Cells(filaCar, CAR_SALDO).Value)
    End If
    If saldoCxC >= valor Then
        nota = valor
    Else
        nota = saldoCxC
        dinero = valor - saldoCxC
    End If
    If dinero > 0 Then
        metodo = Pedir("Se devolverán " & Moneda(dinero) & " al cliente." & vbLf & _
                       "¿Por qué medio? (Efectivo, Transferencia, Nequi, Daviplata, Tarjeta, Otro)", "Devolución de dinero", "Efectivo")
        If Not EnLista(LST_METODOS, metodo) Then Msg "Método de pago no válido.", vbExclamation: Exit Sub
    End If
    If Msg("Confirmar devolución de " & cant & " x " & sku & " (factura " & num & ")" & vbLf & _
           "Valor: " & Moneda(valor) & vbLf & _
           IIf(nota > 0, "• Se descuenta del saldo pendiente: " & Moneda(nota) & vbLf, "") & _
           IIf(dinero > 0, "• Se devuelve en dinero (" & metodo & "): " & Moneda(dinero), ""), vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Iniciar
    On Error GoTo Fallo
    TxIniciar
    TxTabla SH_BD, BD_FACTURA, BD_FIRST, 1, BD_NCOLS
    TxTabla SH_MOV, MOV_ID, MOV_FIRST, 1, MOV_NCOLS
    TxTabla SH_ING, ING_ID, ING_FIRST, 1, ING_NCOLS
    ' línea negativa en ventas (copia de la original con cantidades negativas)
    r = NuevaFila(wb, BD_FACTURA, BD_FIRST, BD_LAST)
    a = wb.Range(wb.Cells(fLinea, 1), wb.Cells(fLinea, BD_NCOLS)).Value
    a(1, BD_FECHA) = Date
    a(1, BD_CANT) = -cant
    a(1, BD_DESCVAL) = -(cant * pu - valor)
    a(1, BD_SUBTOTAL) = -valor
    a(1, BD_COSTOTOTAL) = -Round(cant * cu, 0)
    a(1, BD_UTILIDAD) = -valor + Round(cant * cu, 0)
    a(1, BD_ESTADOPAGO) = "DEVOLUCIÓN"
    a(1, BD_ESTADO) = "ACTIVA"
    a(1, BD_USUARIO) = UsuarioActual()
    a(1, BD_REGISTRO) = Now
    EscribirFila wb, r, 1, a
    RegistrarMovimiento "DEVOLUCIÓN CLIENTE", num, sku, cant, 0, motivo, cu
    If nota > 0 Then RegistrarIngreso Date, "DEVOLUCIÓN", "NOTA CRÉDITO", "", num, Tv(wh.Cells(fh, HFA_CLIENTE).Value), _
                                       "Nota crédito por devolución " & sku, nota, motivo
    If dinero > 0 Then RegistrarIngreso Date, "DEVOLUCIÓN", "DEVOLUCIÓN", metodo, num, Tv(wh.Cells(fh, HFA_CLIENTE).Value), _
                                        "Dinero devuelto por " & sku, -dinero, motivo
    Auditar "INVENTARIO", "DEVOLUCIÓN CLIENTE", num, cant & " x " & sku & ". " & motivo
    TxConfirmar
    Terminar
    Msg "Devolución registrada. Nuevo stock de " & sku & ": " & StockReal(sku), vbInformation
    Exit Sub
Fallo:
    TxRevertir
    Terminar
    Msg "No se pudo registrar la devolución (no se guardó ningún cambio): " & Err.Description, vbCritical
End Sub
