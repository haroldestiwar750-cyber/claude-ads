Attribute VB_Name = "modFactura"
'==============================================================================
' modFactura - Comprobante de venta: NUEVA, GUARDAR, IMPRIMIR, PDF, BUSCAR,
' ANULAR / DEVOLUCION y LIMPIAR.
'
' Al GUARDAR (todo o nada): registra encabezado y lineas, descuenta inventario
' (kardex), registra el ingreso y/o la cuenta por cobrar, congela el costo para
' la utilidad, asocia vendedor y canal, genera el PDF y deja bitacora.
'==============================================================================
Option Explicit

Private Const L1 As Long = 18          ' primera linea de productos
Private Const L2 As Long = 32          ' ultima linea de productos
Private Const C_CLAVE As String = "M8" ' clave unica del borrador (anti duplicados)

Private Function HF() As Worksheet
    Set HF = WS("FACTURACION")
End Function

'------------------------------------------------------------------ NUEVA / LIMPIAR
Public Sub NuevaFactura()
    If Not RequierePermiso("FACTURAR") Then Exit Sub
    LimpiarFormulario
    HF().Activate
    HF().Range("C11").Select
End Sub

Public Sub LimpiarFactura()
    If HF().Range("I8").Value = "NUEVO" Then
        If MsgBox("¿Borrar los datos del borrador actual?", vbQuestion + vbYesNo, APP_NOMBRE) = vbNo Then Exit Sub
    End If
    LimpiarFormulario
End Sub

Private Sub LimpiarFormulario()
    Dim sh As Worksheet, r As Long
    Set sh = HF()
    Application.EnableEvents = False
    sh.Range("C11,I12,I13,I15,C34").ClearContents
    sh.Range("C" & L1 & ":C" & L2 & ",H" & L1 & ":H" & L2 & ",J" & L1 & ":J" & L2).ClearContents
    For r = L1 To L2
        sh.Range("I" & r).Formula = "=IF($M" & r & "="""","""",IFERROR(INDEX(inv_Precio,MATCH($M" & r & ",inv_SKU,0)),0))"
        sh.Range("N" & r).Formula = "=IF($M" & r & "="""","""",IFERROR(INDEX(inv_Costo,MATCH($M" & r & ",inv_SKU,0)),0))"
    Next r
    sh.Range("I6").Formula = "=cfg_PrefFactura&""-""&TEXT(cfg_NumFactura,""000000"")&"" (pendiente)"""
    sh.Range("I7").Value = Date
    sh.Range("I8").Value = "NUEVO"
    sh.Range("I14").Value = "PAGADA"
    sh.Range(C_CLAVE).Value = ClaveUnica()
    ' vendedor por defecto = el del usuario
    If VendedorActual() <> "" Then
        sh.Range("I11").Value = EtiquetaVendedor(VendedorActual())
    Else
        sh.Range("I11").ClearContents
    End If
    Application.EnableEvents = True
End Sub

Private Function EtiquetaVendedor(ByVal cod As String) As String
    Dim nom As Variant
    nom = Buscar(WS("VENDEDORES"), 1, cod, 2)
    If IsEmpty(nom) Then EtiquetaVendedor = cod Else EtiquetaVendedor = cod & " - " & nom
End Function

'------------------------------------------------------------------ GUARDAR
Public Sub GuardarFactura()
    Dim sh As Worksheet, hFac As Worksheet, hDet As Worksheet, hK As Worksheet, hInv As Worksheet
    Dim codCli As String, codVen As String, canal As String, metodo As String, estPago As String
    Dim fecha As Date, abono As Double, total As Double, bruto As Double, saldo As Double
    Dim costoTotal As Double, r As Long, n As Long, num As String, claveU As String
    Dim lineas As Collection, ln As Variant, stock As Object, k As Variant, autoriza As String
    Dim filaInv As Long, filaF As Long, filaD As Long, i As Long, seq As Long
    Dim cliNom As String, cliDoc As String, ruta As String, numCfg As Long

    If Not RequierePermiso("FACTURAR") Then Exit Sub
    Set sh = HF()
    Application.Calculate

    '---------------- validaciones (antes de escribir nada)
    If sh.Range("I8").Value <> "NUEVO" Then
        MsgBox "Este comprobante ya está registrado (" & sh.Range("I8").Value & "). Use NUEVA FACTURA.", _
               vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    claveU = CStr(sh.Range(C_CLAVE).Value)
    If claveU = "" Then claveU = ClaveUnica(): sh.Range(C_CLAVE).Value = claveU
    Set hFac = WS("BD_FACTURAS")
    If BuscarFila(hFac, 22, claveU) > 0 Then
        MsgBox "Este borrador ya fue guardado (clave " & claveU & "). No se duplicará.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    codCli = CodigoDe(sh.Range("C11").Value)
    If codCli = "" Or BuscarFila(WS("CLIENTES"), 1, codCli, FILA_DATOS_CLI) = 0 Then
        MsgBox "Seleccione un cliente válido de la lista (o créelo con + CLIENTE).", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    codVen = CodigoDe(sh.Range("I11").Value)
    If codVen = "" Or BuscarFila(WS("VENDEDORES"), 1, codVen) = 0 Then
        MsgBox "Seleccione el vendedor.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    If Clave(RolActual()) = "VENDEDOR" And VendedorActual() <> "" And codVen <> VendedorActual() Then
        MsgBox "Como vendedor solo puede facturar a su nombre (" & VendedorActual() & ").", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    canal = Trim$(CStr(sh.Range("I12").Value))
    If canal = "" Then
        MsgBox "El canal de venta es obligatorio.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    metodo = Trim$(CStr(sh.Range("I13").Value))
    estPago = Trim$(CStr(sh.Range("I14").Value))
    If metodo = "" Or estPago = "" Then
        MsgBox "Indique forma de pago y estado de pago.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    If Not IsDate(sh.Range("I7").Value) Then
        MsgBox "Fecha no válida.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    fecha = CDate(sh.Range("I7").Value)
    If fecha > Date Then
        MsgBox "La fecha no puede ser futura.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If

    Set lineas = New Collection
    Set stock = CreateObject("Scripting.Dictionary")
    For r = L1 To L2
        If Not Vacio(sh.Range("C" & r).Value) Then
            If CStr(sh.Range("D" & r).Value) = "SKU NO EXISTE" Or Vacio(sh.Range("M" & r).Value) Then
                MsgBox "Línea " & (r - L1 + 1) & ": el SKU no existe.", vbExclamation, APP_NOMBRE
                Exit Sub
            End If
            If NumV(sh.Range("H" & r).Value) <= 0 Or NumV(sh.Range("H" & r).Value) <> Int(NumV(sh.Range("H" & r).Value)) Then
                MsgBox "Línea " & (r - L1 + 1) & ": cantidad no válida.", vbExclamation, APP_NOMBRE
                Exit Sub
            End If
            If NumV(sh.Range("I" & r).Value) <= 0 Then
                MsgBox "Línea " & (r - L1 + 1) & ": el producto no tiene precio de venta.", vbExclamation, APP_NOMBRE
                Exit Sub
            End If
            lineas.Add r
            k = CStr(sh.Range("M" & r).Value)
            stock(k) = stock(k) + NumV(sh.Range("H" & r).Value)
        End If
    Next r
    If lineas.Count = 0 Then
        MsgBox "Agregue al menos un producto.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    ' existencias (suma por SKU si se repite en varias lineas)
    Set hInv = WS("INVENTARIO")
    For Each k In stock.Keys
        filaInv = BuscarFila(hInv, 1, k)
        If stock(k) > NumV(hInv.Cells(filaInv, 12).Value) Then
            If Clave(CStr(Cfg("cfg_PermitirNegativo"))) <> "SI" Then
                MsgBox "Stock insuficiente para " & k & ": disponible " & hInv.Cells(filaInv, 12).Value & _
                       ", solicitado " & stock(k) & "." & vbCrLf & _
                       "La configuración no permite inventario negativo.", vbExclamation, APP_NOMBRE
                Exit Sub
            End If
            autoriza = Autorizar("AJUSTE_INVENTARIO", "Venta con stock insuficiente de " & k)
            If autoriza = "" Then Exit Sub
        End If
    Next k
    total = NumV(sh.Range("K36").Value)
    bruto = NumV(sh.Range("K34").Value)
    If total <= 0 Then
        MsgBox "El total debe ser mayor que cero.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    Select Case Clave(estPago)
        Case "PAGADA": abono = total
        Case "CREDITO": abono = 0
        Case "ABONO PARCIAL"
            abono = NumV(sh.Range("I15").Value)
            If abono <= 0 Or abono >= total Then
                MsgBox "Para ABONO PARCIAL el abono inicial debe ser mayor que 0 y menor que el total.", _
                       vbExclamation, APP_NOMBRE
                Exit Sub
            End If
        Case Else
            MsgBox "Estado de pago no válido.", vbExclamation, APP_NOMBRE
            Exit Sub
    End Select
    saldo = total - abono
    If MsgBox("¿Guardar el comprobante por " & Format$(total, """$"" #,##0") & "?" & vbCrLf & _
              "Cliente: " & sh.Range("C11").Value & vbCrLf & "Canal: " & canal & " · " & estPago, _
              vbQuestion + vbYesNo, APP_NOMBRE) = vbNo Then Exit Sub

    '---------------- escritura (con deshacer si algo falla)
    On Error GoTo FALLO
    InicioMacro
    IniciarTransaccion
    numCfg = CLng(Cfg("cfg_NumFactura"))
    n = numCfg
    Do
        num = CStr(Cfg("cfg_PrefFactura")) & "-" & Format$(n, "000000")
        If BuscarFila(hFac, 1, num) = 0 Then Exit Do
        n = n + 1
    Loop
    EscribirCelda NR("cfg_NumFactura").Worksheet, NR("cfg_NumFactura").Row, NR("cfg_NumFactura").Column, n + 1

    Set hDet = WS("BD_FACT_DETALLE")
    Set hK = WS("HISTORIAL MOVIMIENTOS")
    cliNom = CStr(Buscar(WS("CLIENTES"), 1, codCli, 2, FILA_DATOS_CLI))
    cliDoc = CStr(Buscar(WS("CLIENTES"), 1, codCli, 4, FILA_DATOS_CLI))
    seq = Application.WorksheetFunction.CountIf(NR("fac_Cli"), codCli) + 1

    i = 0
    For Each ln In lineas
        i = i + 1
        r = CLng(ln)
        Dim sku As String, cant As Double, precio As Double, desc As Double, subt As Double, cu As Double
        sku = CStr(sh.Range("M" & r).Value)
        cant = NumV(sh.Range("H" & r).Value)
        precio = NumV(sh.Range("I" & r).Value)
        desc = NumV(sh.Range("J" & r).Value)
        subt = NumV(sh.Range("K" & r).Value)
        cu = NumV(sh.Range("N" & r).Value)
        costoTotal = costoTotal + cu * cant
        filaInv = BuscarFila(hInv, 1, sku)
        filaD = SiguienteFila(hDet)
        EscribirFila hDet, filaD, Array(num, fecha, canal, codVen, codCli, sku, sh.Range("D" & r).Value, _
            hInv.Cells(filaInv, 3).Value, sh.Range("E" & r).Value, sh.Range("F" & r).Value, cant, precio, desc, _
            subt, cu, cu * cant, subt - cu * cant, "VIGENTE", num & "|" & i)
        EscribirFila hK, SiguienteFila(hK), Array(SiguienteID(hK, 1, "K", 6), fecha, Format$(Now, "hh:mm"), _
            UsuarioActual(), "VENTA", sku, sh.Range("D" & r).Value & " " & sh.Range("E" & r).Value & " " & _
            sh.Range("F" & r).Value, Empty, cant, cu, num, "Venta " & canal, autoriza)
    Next ln

    filaF = SiguienteFila(hFac)
    EscribirFila hFac, filaF, Array(num, fecha, codCli, cliNom, cliDoc, codVen, canal, metodo, estPago, bruto, _
        bruto - total, total, abono, saldo, costoTotal, total - costoTotal, "VIGENTE", UsuarioActual(), Now, "", _
        "", claveU, codCli & "|" & seq, Empty)

    If abono > 0 Then
        Dim hI As Worksheet
        Set hI = WS("INGRESOS")
        EscribirFila hI, SiguienteFila(hI), Array(SiguienteID(hI, 1, "ING", 5), fecha, "Ventas", metodo, abono, num, _
            codCli, "Venta " & num, UsuarioActual(), "VIGENTE")
    End If
    If saldo > 0 Then
        Dim hC As Worksheet, rc As Long
        Set hC = WS("CARTERA")
        rc = SiguienteFila(hC)
        hC.Range("A" & rc & ":G" & rc).Value = Array(SiguienteID(hC, 1, "CAR", 5), "CxC", num, codCli, cliNom, fecha, saldo)
        hC.Range("J" & rc).Value = fecha + CLng(Cfg("cfg_DiasCredito"))
        hC.Range("M" & rc).Value = "NO"
        RegistrarEscritura hC, rc, 1, 7
        RegistrarEscritura hC, rc, 10, 10
        RegistrarEscritura hC, rc, 13, 13
    End If

    Bitacora "FACTURACIÓN", "GUARDAR", num, canal & " · total " & Format$(total, "#,##0") & " · " & estPago
    CerrarTransaccion
    Application.EnableEvents = False
    sh.Range("I6").Value = num
    sh.Range("I8").Value = "VIGENTE"
    FinMacro

    ' PDF y guardado (fuera de la transaccion: si falla el PDF, el comprobante ya quedo registrado)
    ruta = modRespaldo.ExportarPDF(sh, "FACTURAS", num, fecha)
    If ruta <> "" Then
        Application.EnableEvents = False
        hFac.Cells(filaF, 20).Value = modRespaldo.RutaRelativa(ruta)
        Application.EnableEvents = True
    End If
    GuardarLibro
    MsgBox "Comprobante " & num & " guardado." & vbCrLf & _
           "Total: " & Format$(total, """$"" #,##0") & IIf(saldo > 0, vbCrLf & "Saldo en cartera: " & _
           Format$(saldo, """$"" #,##0"), "") & IIf(ruta <> "", vbCrLf & "PDF: " & ruta, vbCrLf & _
           "(No se generó el PDF: revise la carpeta de OneDrive en CONFIGURACIÓN.)"), vbInformation, APP_NOMBRE
    Exit Sub

FALLO:
    DeshacerTransaccion
    FinMacro
    ErrorGeneral "GuardarFactura"
End Sub

'------------------------------------------------------------------ BUSCAR
Public Sub BuscarFactura()
    Dim num As String, hFac As Worksheet, hDet As Worksheet, sh As Worksheet, rf As Long, r As Long
    Dim fila As Long, ult As Long, sku As String, lab As Variant
    Set sh = HF()
    Set hFac = WS("BD_FACTURAS")
    num = UCase$(PedirTexto("Número del comprobante a buscar (ej. FV-000012):", CStr(hFac.Cells(UltimaFila(hFac, 1), 1).Value)))
    If num = "" Then Exit Sub
    If IsNumeric(num) Then num = CStr(Cfg("cfg_PrefFactura")) & "-" & Format$(CLng(num), "000000")
    rf = BuscarFila(hFac, 1, num)
    If rf = 0 Then
        MsgBox "No existe el comprobante " & num & ".", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    InicioMacro
    LimpiarFormulario
    Application.EnableEvents = False
    sh.Range("I6").Value = num
    sh.Range("I7").Value = hFac.Cells(rf, 2).Value
    sh.Range("I8").Value = hFac.Cells(rf, 17).Value
    sh.Range("C11").Value = hFac.Cells(rf, 3).Value & " - " & hFac.Cells(rf, 4).Value
    sh.Range("I11").Value = EtiquetaVendedor(CStr(hFac.Cells(rf, 6).Value))
    sh.Range("I12").Value = hFac.Cells(rf, 7).Value
    sh.Range("I13").Value = hFac.Cells(rf, 8).Value
    sh.Range("I14").Value = hFac.Cells(rf, 9).Value
    sh.Range("I15").Value = hFac.Cells(rf, 13).Value
    sh.Range(C_CLAVE).Value = hFac.Cells(rf, 22).Value
    If hFac.Cells(rf, 17).Value <> "VIGENTE" Then sh.Range("C34").Value = hFac.Cells(rf, 17).Value & ": " & hFac.Cells(rf, 21).Value
    Set hDet = WS("BD_FACT_DETALLE")
    ult = UltimaFila(hDet, 1)
    fila = L1
    For r = FILA_DATOS To ult
        If hDet.Cells(r, 1).Value = num And fila <= L2 Then
            sku = CStr(hDet.Cells(r, 6).Value)
            lab = Buscar(WS("INVENTARIO"), 1, sku, 16)
            If IsEmpty(lab) Then lab = sku & " - " & hDet.Cells(r, 7).Value
            sh.Range("C" & fila).Value = lab
            sh.Range("H" & fila).Value = hDet.Cells(r, 11).Value
            sh.Range("I" & fila).Value = hDet.Cells(r, 12).Value     ' precio historico
            If NumV(hDet.Cells(r, 13).Value) > 0 Then sh.Range("J" & fila).Value = hDet.Cells(r, 13).Value
            sh.Range("N" & fila).Value = hDet.Cells(r, 15).Value     ' costo historico
            fila = fila + 1
        End If
    Next r
    FinMacro
    sh.Activate
    Bitacora "FACTURACIÓN", "CONSULTAR", num, ""
End Sub

'------------------------------------------------------------------ ANULAR / DEVOLUCION
Public Sub AnularFactura()
    Dim num As String, hFac As Worksheet, rf As Long, autoriza As String, motivo As String
    Dim esDevol As Boolean, resp As VbMsgBoxResult, nuevoEstado As String
    Dim hDet As Worksheet, hK As Worksheet, hI As Worksheet, hE As Worksheet, hC As Worksheet, hA As Worksheet
    Dim r As Long, ult As Long, pagado As Double, abonos As Object, fecha As Date, k As Variant

    Set hFac = WS("BD_FACTURAS")
    If HF().Range("I8").Value = "VIGENTE" And ActiveSheet.Name = HF().Name Then num = CStr(HF().Range("I6").Value)
    num = UCase$(PedirTexto("Comprobante a anular o devolver:", num))
    If num = "" Then Exit Sub
    rf = BuscarFila(hFac, 1, num)
    If rf = 0 Then MsgBox "No existe " & num & ".", vbExclamation, APP_NOMBRE: Exit Sub
    If hFac.Cells(rf, 17).Value <> "VIGENTE" Then
        MsgBox num & " ya está " & hFac.Cells(rf, 17).Value & ".", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    autoriza = Autorizar("ANULAR_FACTURA", "Anulación / devolución de " & num)
    If autoriza = "" Then Exit Sub
    resp = MsgBox("¿Es una DEVOLUCIÓN del cliente?" & vbCrLf & vbCrLf & _
                  "SÍ = devolución: la mercancía regresa y se registra el reembolso de lo pagado." & vbCrLf & _
                  "NO = anulación: error o pedido cancelado; los ingresos del comprobante se anulan.", _
                  vbQuestion + vbYesNoCancel, APP_NOMBRE)
    If resp = vbCancel Then Exit Sub
    esDevol = (resp = vbYes)
    nuevoEstado = IIf(esDevol, "DEVUELTA", "ANULADA")
    motivo = PedirTexto("Motivo (obligatorio):")
    If motivo = "" Then Exit Sub
    fecha = Date

    On Error GoTo FALLO
    InicioMacro
    IniciarTransaccion
    EscribirCelda hFac, rf, 17, nuevoEstado
    EscribirCelda hFac, rf, 21, motivo
    EscribirCelda hFac, rf, 24, fecha
    ' lineas: estado y reingreso al inventario
    Set hDet = WS("BD_FACT_DETALLE")
    Set hK = WS("HISTORIAL MOVIMIENTOS")
    ult = UltimaFila(hDet, 1)
    For r = FILA_DATOS To ult
        If hDet.Cells(r, 1).Value = num Then
            EscribirCelda hDet, r, 18, nuevoEstado
            EscribirFila hK, SiguienteFila(hK), Array(SiguienteID(hK, 1, "K", 6), fecha, Format$(Now, "hh:mm"), _
                UsuarioActual(), IIf(esDevol, "DEVOLUCIÓN", "ANULACIÓN VENTA"), hDet.Cells(r, 6).Value, _
                hDet.Cells(r, 7).Value & " " & hDet.Cells(r, 9).Value & " " & hDet.Cells(r, 10).Value, _
                hDet.Cells(r, 11).Value, Empty, hDet.Cells(r, 15).Value, num, motivo, autoriza)
        End If
    Next r
    ' abonos de cartera ligados a este comprobante
    Set abonos = CreateObject("Scripting.Dictionary")
    Set hA = WS("BD_ABONOS")
    ult = UltimaFila(hA, 1)
    For r = FILA_DATOS To ult
        If hA.Cells(r, 5).Value = num Then abonos(CStr(hA.Cells(r, 1).Value)) = NumV(hA.Cells(r, 7).Value)
    Next r
    ' ingresos del comprobante y de sus abonos
    Set hI = WS("INGRESOS")
    ult = UltimaFila(hI, 1)
    For r = FILA_DATOS To ult
        If hI.Cells(r, 10).Value = "VIGENTE" Then
            If hI.Cells(r, 6).Value = num Or abonos.Exists(CStr(hI.Cells(r, 6).Value)) Then
                pagado = pagado + NumV(hI.Cells(r, 5).Value)
                If Not esDevol Then EscribirCelda hI, r, 10, "ANULADO"
            End If
        End If
    Next r
    ' cartera
    Set hC = WS("CARTERA")
    ult = UltimaFila(hC, 1)
    For r = FILA_DATOS To ult
        If hC.Cells(r, 3).Value = num And hC.Cells(r, 2).Value = "CxC" Then EscribirCelda hC, r, 13, "SI"
    Next r
    ' reembolso
    If esDevol And pagado > 0 Then
        Set hE = WS("EGRESOS")
        EscribirFila hE, SiguienteFila(hE), Array(SiguienteID(hE, 1, "EGR", 5), fecha, "Devoluciones a clientes", _
            hFac.Cells(rf, 8).Value, pagado, num, hFac.Cells(rf, 3).Value, "Reembolso devolución " & num, _
            UsuarioActual(), "VIGENTE")
    End If
    Bitacora "FACTURACIÓN", IIf(esDevol, "DEVOLUCIÓN", "ANULAR"), num, motivo & " · autoriza " & autoriza
    CerrarTransaccion
    Application.EnableEvents = False
    If CStr(HF().Range("I6").Value) = num Then HF().Range("I8").Value = nuevoEstado
    FinMacro
    GuardarLibro
    MsgBox num & " quedó " & nuevoEstado & ". El inventario fue restaurado." & _
           IIf(esDevol And pagado > 0, vbCrLf & "Reembolso registrado en EGRESOS: " & Format$(pagado, """$"" #,##0"), "") & _
           IIf(Not esDevol And pagado > 0, vbCrLf & "Ingresos anulados: " & Format$(pagado, """$"" #,##0"), ""), _
           vbInformation, APP_NOMBRE
    Exit Sub
FALLO:
    DeshacerTransaccion
    FinMacro
    ErrorGeneral "AnularFactura"
End Sub

'------------------------------------------------------------------ IMPRIMIR / PDF
Public Sub ImprimirFactura()
    If HF().Range("I8").Value = "NUEVO" Then
        If MsgBox("El comprobante aún no está guardado. ¿Imprimir como BORRADOR?", vbQuestion + vbYesNo, _
                  APP_NOMBRE) = vbNo Then Exit Sub
    End If
    HF().PrintOut Copies:=1, Preview:=True
End Sub

Public Sub PDFFactura()
    Dim ruta As String
    If HF().Range("I8").Value = "NUEVO" Then
        MsgBox "Guarde el comprobante antes de generar el PDF.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    ruta = modRespaldo.ExportarPDF(HF(), "FACTURAS", CStr(HF().Range("I6").Value), CDate(HF().Range("I7").Value))
    If ruta <> "" Then
        If MsgBox("PDF generado:" & vbCrLf & ruta & vbCrLf & vbCrLf & "¿Abrirlo?", vbQuestion + vbYesNo, APP_NOMBRE) = vbYes Then
            ThisWorkbook.FollowHyperlink ruta
        End If
    End If
End Sub

'------------------------------------------------------------------ cliente rapido
Public Sub ClienteRapido()
    Dim cod As String
    cod = modMaestros.NuevoCliente()
    If cod <> "" Then
        Application.EnableEvents = False
        HF().Range("C11").Value = cod & " - " & Buscar(WS("CLIENTES"), 1, cod, 2, FILA_DATOS_CLI)
        Application.EnableEvents = True
        HF().Activate
    End If
End Sub
