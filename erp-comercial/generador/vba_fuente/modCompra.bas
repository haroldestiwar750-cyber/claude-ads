Attribute VB_Name = "modCompra"
'==============================================================================
' modCompra - Registro de compras a proveedores.
' Al GUARDAR: encabezado y lineas, entrada al kardex, costo promedio ponderado
' (con flete prorrateado), egreso por lo pagado y cuenta por pagar por el saldo.
' No permite registrar dos veces la misma factura del mismo proveedor.
'==============================================================================
Option Explicit

Private Const L1 As Long = 18
Private Const L2 As Long = 32
Private Const C_CLAVE As String = "M8"

Private Function HC() As Worksheet
    Set HC = WS("COMPRAS")
End Function

Public Sub NuevaCompra()
    If Not RequierePermiso("COMPRAS") Then Exit Sub
    LimpiarFormulario
    HC().Activate
    HC().Range("C11").Select
End Sub

Public Sub LimpiarCompra()
    If HC().Range("I8").Value = "NUEVO" Then
        If MsgBox("¿Borrar los datos del borrador actual?", vbQuestion + vbYesNo, APP_NOMBRE) = vbNo Then Exit Sub
    End If
    LimpiarFormulario
End Sub

Private Sub LimpiarFormulario()
    Dim sh As Worksheet
    Set sh = HC()
    Application.EnableEvents = False
    sh.Range("C11,C13,I11,I13,C34").ClearContents
    sh.Range("C" & L1 & ":C" & L2 & ",G" & L1 & ":G" & L2 & ",H" & L1 & ":H" & L2).ClearContents
    sh.Range("I35").Value = 0
    sh.Range("I36").Value = 0
    sh.Range("I12").Value = "PAGADA"
    sh.Range("I14").Formula = "=cfg_DiasCreditoProv"
    sh.Range("I6").Formula = "=cfg_PrefCompra&""-""&TEXT(cfg_NumCompra,""00000"")&"" (pendiente)"""
    sh.Range("I7").Value = Date
    sh.Range("I8").Value = "NUEVO"
    sh.Range(C_CLAVE).Value = ClaveUnica()
    Application.EnableEvents = True
End Sub

Public Sub GuardarCompra()
    Dim sh As Worksheet, hCo As Worksheet, hCd As Worksheet, hK As Worksheet, hInv As Worksheet
    Dim codPrv As String, factPrv As String, metodo As String, estPago As String, fecha As Date
    Dim total As Double, subtotal As Double, transp As Double, adic As Double, pagado As Double, saldo As Double
    Dim r As Long, n As Long, num As String, claveU As String, lineas As Collection, ln As Variant
    Dim filaInv As Long, ult As Long, existe As Object, sku As String, cant As Double, cu As Double
    Dim prorr As Double, costoReal As Double, exAnt As Double, costoAnt As Double, nuevoCosto As Double
    Dim prvNom As String, dias As Long

    If Not RequierePermiso("COMPRAS") Then Exit Sub
    Set sh = HC()
    Application.Calculate
    If sh.Range("I8").Value <> "NUEVO" Then
        MsgBox "Esta compra ya está registrada. Use NUEVA COMPRA.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    Set hCo = WS("BD_COMPRAS")
    claveU = CStr(sh.Range(C_CLAVE).Value)
    If claveU = "" Then claveU = ClaveUnica(): sh.Range(C_CLAVE).Value = claveU
    If BuscarFila(hCo, 17, claveU) > 0 Then
        MsgBox "Este borrador ya fue guardado. No se duplicará.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    codPrv = CodigoDe(sh.Range("C11").Value)
    If codPrv = "" Or BuscarFila(WS("PROVEEDORES"), 1, codPrv) = 0 Then
        MsgBox "Seleccione un proveedor válido.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    factPrv = Trim$(CStr(sh.Range("C13").Value))
    If factPrv = "" Then
        MsgBox "Escriba el número de factura del proveedor.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    ult = UltimaFila(hCo, 1)
    For r = FILA_DATOS To ult
        If hCo.Cells(r, 3).Value = codPrv And Clave(CStr(hCo.Cells(r, 5).Value)) = Clave(factPrv) And _
           hCo.Cells(r, 14).Value = "VIGENTE" Then
            MsgBox "La factura " & factPrv & " de este proveedor ya fue registrada en " & hCo.Cells(r, 1).Value & ".", _
                   vbExclamation, APP_NOMBRE
            Exit Sub
        End If
    Next r
    metodo = Trim$(CStr(sh.Range("I11").Value))
    estPago = Trim$(CStr(sh.Range("I12").Value))
    If estPago = "" Or (metodo = "" And Clave(estPago) <> "CREDITO") Then
        MsgBox "Indique forma y estado de pago.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    If Not IsDate(sh.Range("I7").Value) Then MsgBox "Fecha no válida.", vbExclamation, APP_NOMBRE: Exit Sub
    fecha = CDate(sh.Range("I7").Value)
    Set lineas = New Collection
    For r = L1 To L2
        If Not Vacio(sh.Range("C" & r).Value) Then
            If CStr(sh.Range("D" & r).Value) = "SKU NO EXISTE" Then
                MsgBox "Línea " & (r - L1 + 1) & ": el SKU no existe. Créelo en INVENTARIO > NUEVO PRODUCTO.", _
                       vbExclamation, APP_NOMBRE
                Exit Sub
            End If
            If NumV(sh.Range("G" & r).Value) <= 0 Or NumV(sh.Range("H" & r).Value) <= 0 Then
                MsgBox "Línea " & (r - L1 + 1) & ": cantidad y costo deben ser mayores que cero.", vbExclamation, APP_NOMBRE
                Exit Sub
            End If
            lineas.Add r
        End If
    Next r
    If lineas.Count = 0 Then MsgBox "Agregue al menos un producto.", vbExclamation, APP_NOMBRE: Exit Sub
    subtotal = NumV(sh.Range("I34").Value)
    transp = NumV(sh.Range("I35").Value)
    adic = NumV(sh.Range("I36").Value)
    total = NumV(sh.Range("I37").Value)
    Select Case Clave(estPago)
        Case "PAGADA": pagado = total
        Case "CREDITO": pagado = 0
        Case "ABONO PARCIAL"
            pagado = NumV(sh.Range("I13").Value)
            If pagado <= 0 Or pagado >= total Then
                MsgBox "Para ABONO PARCIAL el valor pagado debe ser mayor que 0 y menor que el total.", vbExclamation, APP_NOMBRE
                Exit Sub
            End If
        Case Else
            MsgBox "Estado de pago no válido.", vbExclamation, APP_NOMBRE: Exit Sub
    End Select
    saldo = total - pagado
    dias = CLng(NumV(sh.Range("I14").Value))
    If MsgBox("¿Registrar la compra por " & Format$(total, """$"" #,##0") & "?", vbQuestion + vbYesNo, APP_NOMBRE) = vbNo Then Exit Sub

    On Error GoTo FALLO
    InicioMacro
    IniciarTransaccion
    n = CLng(Cfg("cfg_NumCompra"))
    Do
        num = CStr(Cfg("cfg_PrefCompra")) & "-" & Format$(n, "00000")
        If BuscarFila(hCo, 1, num) = 0 Then Exit Do
        n = n + 1
    Loop
    EscribirCelda NR("cfg_NumCompra").Worksheet, NR("cfg_NumCompra").Row, NR("cfg_NumCompra").Column, n + 1
    prvNom = CStr(Buscar(WS("PROVEEDORES"), 1, codPrv, 2))
    Set hCd = WS("BD_COMPRA_DETALLE")
    Set hK = WS("HISTORIAL MOVIMIENTOS")
    Set hInv = WS("INVENTARIO")
    Set existe = CreateObject("Scripting.Dictionary")
    For Each ln In lineas
        r = CLng(ln)
        sku = CStr(sh.Range("M" & r).Value)
        cant = NumV(sh.Range("G" & r).Value)
        cu = NumV(sh.Range("H" & r).Value)
        prorr = NumV(sh.Range("J" & r).Value)
        costoReal = cu + prorr
        filaInv = BuscarFila(hInv, 1, sku)
        If Not existe.Exists(sku) Then existe(sku) = NumV(hInv.Cells(filaInv, 12).Value)
        exAnt = existe(sku)
        costoAnt = NumV(hInv.Cells(filaInv, 7).Value)
        If exAnt > 0 Then
            nuevoCosto = Round((exAnt * costoAnt + cant * costoReal) / (exAnt + cant), 0)
        Else
            nuevoCosto = Round(costoReal, 0)
        End If
        EscribirCelda hInv, filaInv, 7, nuevoCosto
        existe(sku) = exAnt + cant
        EscribirFila hCd, SiguienteFila(hCd), Array(num, fecha, codPrv, sku, sh.Range("D" & r).Value, cant, cu, _
            cant * cu, Round(prorr, 2), Round(costoReal, 2), "VIGENTE")
        EscribirFila hK, SiguienteFila(hK), Array(SiguienteID(hK, 1, "K", 6), fecha, Format$(Now, "hh:mm"), _
            UsuarioActual(), "COMPRA", sku, sh.Range("D" & r).Value & " " & sh.Range("E" & r).Value & " " & _
            sh.Range("F" & r).Value, cant, Empty, Round(costoReal, 0), num, "Compra a " & codPrv, "")
    Next ln
    EscribirFila hCo, SiguienteFila(hCo), Array(num, fecha, codPrv, prvNom, factPrv, subtotal, transp, adic, total, _
        metodo, estPago, pagado, saldo, "VIGENTE", UsuarioActual(), Now, claveU)
    If pagado > 0 Then
        Dim hE As Worksheet
        Set hE = WS("EGRESOS")
        EscribirFila hE, SiguienteFila(hE), Array(SiguienteID(hE, 1, "EGR", 5), fecha, "Compras", metodo, pagado, num, _
            codPrv, "Pago compra " & num & " (fact. " & factPrv & ")", UsuarioActual(), "VIGENTE")
    End If
    If saldo > 0 Then
        Dim hCa As Worksheet, rc As Long
        Set hCa = WS("CARTERA")
        rc = SiguienteFila(hCa)
        hCa.Range("A" & rc & ":G" & rc).Value = Array(SiguienteID(hCa, 1, "CAR", 5), "CxP", num, codPrv, prvNom, fecha, saldo)
        hCa.Range("J" & rc).Value = fecha + dias
        hCa.Range("M" & rc).Value = "NO"
        RegistrarEscritura hCa, rc, 1, 7
        RegistrarEscritura hCa, rc, 10, 10
        RegistrarEscritura hCa, rc, 13, 13
    End If
    Bitacora "COMPRAS", "GUARDAR", num, codPrv & " fact. " & factPrv & " · total " & Format$(total, "#,##0")
    CerrarTransaccion
    Application.EnableEvents = False
    sh.Range("I6").Value = num
    sh.Range("I8").Value = "VIGENTE"
    FinMacro
    GuardarLibro
    MsgBox "Compra " & num & " registrada. Inventario y costos actualizados." & _
           IIf(saldo > 0, vbCrLf & "Cuenta por pagar: " & Format$(saldo, """$"" #,##0"), ""), vbInformation, APP_NOMBRE
    Exit Sub
FALLO:
    DeshacerTransaccion
    FinMacro
    ErrorGeneral "GuardarCompra"
End Sub

Public Sub BuscarCompra()
    Dim num As String, hCo As Worksheet, hCd As Worksheet, sh As Worksheet, rf As Long, r As Long, fila As Long
    Dim lab As Variant, sku As String
    Set sh = HC()
    Set hCo = WS("BD_COMPRAS")
    num = UCase$(PedirTexto("Número de compra (ej. OC-00003):", CStr(hCo.Cells(UltimaFila(hCo, 1), 1).Value)))
    If num = "" Then Exit Sub
    rf = BuscarFila(hCo, 1, num)
    If rf = 0 Then MsgBox "No existe la compra " & num & ".", vbExclamation, APP_NOMBRE: Exit Sub
    InicioMacro
    LimpiarFormulario
    Application.EnableEvents = False
    sh.Range("I6").Value = num
    sh.Range("I7").Value = hCo.Cells(rf, 2).Value
    sh.Range("I8").Value = hCo.Cells(rf, 14).Value
    sh.Range("C11").Value = hCo.Cells(rf, 3).Value & " - " & hCo.Cells(rf, 4).Value
    sh.Range("C13").Value = hCo.Cells(rf, 5).Value
    sh.Range("I11").Value = hCo.Cells(rf, 10).Value
    sh.Range("I12").Value = hCo.Cells(rf, 11).Value
    sh.Range("I13").Value = hCo.Cells(rf, 12).Value
    sh.Range("I35").Value = hCo.Cells(rf, 7).Value
    sh.Range("I36").Value = hCo.Cells(rf, 8).Value
    sh.Range(C_CLAVE).Value = hCo.Cells(rf, 17).Value
    Set hCd = WS("BD_COMPRA_DETALLE")
    fila = L1
    For r = FILA_DATOS To UltimaFila(hCd, 1)
        If hCd.Cells(r, 1).Value = num And fila <= L2 Then
            sku = CStr(hCd.Cells(r, 4).Value)
            lab = Buscar(WS("INVENTARIO"), 1, sku, 16)
            If IsEmpty(lab) Then lab = sku & " - " & hCd.Cells(r, 5).Value
            sh.Range("C" & fila).Value = lab
            sh.Range("G" & fila).Value = hCd.Cells(r, 6).Value
            sh.Range("H" & fila).Value = hCd.Cells(r, 7).Value
            fila = fila + 1
        End If
    Next r
    FinMacro
    sh.Activate
End Sub

Public Sub ImprimirCompra()
    HC().PrintOut Copies:=1, Preview:=True
End Sub

Public Sub PDFCompra()
    Dim ruta As String
    If HC().Range("I8").Value = "NUEVO" Then MsgBox "Guarde la compra primero.", vbExclamation, APP_NOMBRE: Exit Sub
    ruta = modRespaldo.ExportarPDF(HC(), "COMPRAS", CStr(HC().Range("I6").Value), CDate(HC().Range("I7").Value))
    If ruta <> "" Then MsgBox "PDF generado:" & vbCrLf & ruta, vbInformation, APP_NOMBRE
End Sub
