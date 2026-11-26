Attribute VB_Name = "modFinanzas"
'==============================================================================
' modFinanzas - Abonos de cartera, ingresos y egresos manuales, gastos fijos,
' cierres de caja y pago de comisiones. Todo sin duplicar movimientos.
'==============================================================================
Option Explicit

'------------------------------------------------------------------ ABONOS
Public Sub RegistrarAbono()
    Dim hC As Worksheet, r As Long, id As String, saldo As Double, valor As Variant, metodo As String
    Dim hA As Worksheet, idAb As String, tipo As String, doc As String, ter As String, fecha As Date
    Dim hFac As Worksheet, rf As Long

    If Not RequierePermiso("CARTERA") Then Exit Sub
    Set hC = WS("CARTERA")
    Application.Calculate
    If ActiveSheet.Name = hC.Name And ActiveCell.Row >= FILA_DATOS Then
        If Not Vacio(hC.Cells(ActiveCell.Row, 1).Value) Then id = CStr(hC.Cells(ActiveCell.Row, 1).Value)
    End If
    id = UCase$(PedirTexto("ID de cartera (CAR-xxxxx) o número de documento (FV-/OC-):", id))
    If id = "" Then Exit Sub
    r = BuscarFila(hC, 1, id)
    If r = 0 Then r = BuscarFila(hC, 3, id)
    If r = 0 Then MsgBox "No se encontró " & id & " en cartera.", vbExclamation, APP_NOMBRE: Exit Sub
    If hC.Cells(r, 13).Value = "SI" Then MsgBox "Ese documento está anulado.", vbExclamation, APP_NOMBRE: Exit Sub
    saldo = NumV(hC.Cells(r, 9).Value)
    If saldo <= 0 Then MsgBox "Ese documento no tiene saldo pendiente.", vbInformation, APP_NOMBRE: Exit Sub
    tipo = CStr(hC.Cells(r, 2).Value)
    doc = CStr(hC.Cells(r, 3).Value)
    ter = CStr(hC.Cells(r, 4).Value)
    valor = PedirNumero(IIf(tipo = "CxC", "Abono del cliente ", "Pago al proveedor ") & hC.Cells(r, 5).Value & _
                        vbCrLf & "Documento " & doc & " · saldo " & Format$(saldo, """$"" #,##0") & vbCrLf & vbCrLf & _
                        "Valor del abono:", saldo, 1)
    If IsEmpty(valor) Then Exit Sub
    If valor > saldo Then
        MsgBox "El abono no puede superar el saldo (" & Format$(saldo, """$"" #,##0") & ").", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    metodo = ElegirDeLista("Método de pago:", ListaCatalogo("Métodos de pago / cuentas"), "Transferencia")
    If metodo = "" Then Exit Sub
    fecha = Date

    On Error GoTo FALLO
    InicioMacro
    IniciarTransaccion
    Set hA = WS("BD_ABONOS")
    idAb = SiguienteID(hA, 1, "AB", 5)
    EscribirFila hA, SiguienteFila(hA), Array(idAb, fecha, hC.Cells(r, 1).Value, tipo, doc, ter, CDbl(valor), metodo, _
        UsuarioActual(), "Abono", Now)
    If tipo = "CxC" Then
        Dim hI As Worksheet
        Set hI = WS("INGRESOS")
        EscribirFila hI, SiguienteFila(hI), Array(SiguienteID(hI, 1, "ING", 5), fecha, "Abonos", metodo, CDbl(valor), _
            idAb, ter, "Abono a " & doc, UsuarioActual(), "VIGENTE")
        If valor >= saldo Then
            Set hFac = WS("BD_FACTURAS")
            rf = BuscarFila(hFac, 1, doc)
            If rf > 0 Then EscribirCelda hFac, rf, 9, "PAGADA"
        End If
    Else
        Dim hE As Worksheet
        Set hE = WS("EGRESOS")
        EscribirFila hE, SiguienteFila(hE), Array(SiguienteID(hE, 1, "EGR", 5), fecha, "Compras", metodo, CDbl(valor), _
            idAb, ter, "Abono a " & doc, UsuarioActual(), "VIGENTE")
    End If
    Bitacora "CARTERA", "ABONO", idAb, doc & " · " & Format$(valor, "#,##0") & " · " & metodo
    CerrarTransaccion
    FinMacro
    GuardarLibro
    MsgBox "Abono " & idAb & " registrado. Nuevo saldo: " & Format$(saldo - valor, """$"" #,##0"), vbInformation, APP_NOMBRE
    Exit Sub
FALLO:
    DeshacerTransaccion
    FinMacro
    ErrorGeneral "RegistrarAbono"
End Sub

'------------------------------------------------------------------ INGRESOS / EGRESOS MANUALES
Public Sub NuevoIngreso()
    Dim hI As Worksheet, metodo As String, valor As Variant, concepto As String, tercero As String, doc As String
    If Not RequierePermiso("FINANZAS") Then Exit Sub
    MsgBox "Use este botón solo para OTROS INGRESOS." & vbCrLf & _
           "Las ventas y los abonos se registran solos desde FACTURACIÓN y CARTERA (así no se duplican).", _
           vbInformation, APP_NOMBRE
    metodo = ElegirDeLista("Cuenta / método donde entra el dinero:", ListaCatalogo("Métodos de pago / cuentas"), "Efectivo")
    If metodo = "" Then Exit Sub
    valor = PedirNumero("Valor del ingreso:", "", 1)
    If IsEmpty(valor) Then Exit Sub
    concepto = PedirTexto("Concepto / detalle:")
    If concepto = "" Then Exit Sub
    tercero = PedirTexto("Tercero (quién paga) - opcional:", "", False)
    doc = PedirTexto("Documento soporte - opcional:", "", False)
    Set hI = WS("INGRESOS")
    InicioMacro
    EscribirFila hI, SiguienteFila(hI), Array(SiguienteID(hI, 1, "ING", 5), Date, "Otros ingresos", metodo, CDbl(valor), _
        doc, tercero, concepto, UsuarioActual(), "VIGENTE")
    Bitacora "INGRESOS", "NUEVO", "", concepto & " · " & Format$(valor, "#,##0")
    FinMacro
    GuardarLibro
End Sub

Public Sub NuevoEgreso()
    Dim hE As Worksheet, concepto As String, metodo As String, valor As Variant, benef As String, desc As String
    Dim doc As String
    If Not RequierePermiso("FINANZAS") Then Exit Sub
    concepto = ElegirDeLista("Concepto del egreso:" & vbCrLf & "(Compras de mercancía: use el módulo COMPRAS)", _
                             ListaCatalogo("Conceptos de egreso"), "Otros gastos")
    If concepto = "" Then Exit Sub
    If Clave(concepto) = "COMPRAS" Or Clave(concepto) = "COMISIONES" Then
        If MsgBox("Las compras y comisiones se registran desde sus módulos para no duplicar." & vbCrLf & _
                  "¿Registrar de todas formas como egreso manual?", vbQuestion + vbYesNo, APP_NOMBRE) = vbNo Then Exit Sub
    End If
    metodo = ElegirDeLista("Cuenta / método de pago:", ListaCatalogo("Métodos de pago / cuentas"), "Transferencia")
    If metodo = "" Then Exit Sub
    valor = PedirNumero("Valor:", "", 1)
    If IsEmpty(valor) Then Exit Sub
    benef = PedirTexto("Beneficiario:")
    If benef = "" Then Exit Sub
    desc = PedirTexto("Descripción:", concepto)
    doc = PedirTexto("Documento soporte (factura / recibo) - opcional:", "", False)
    If doc <> "" Then
        If BuscarFila(WS("EGRESOS"), 6, doc) > 0 Then
            If MsgBox("Ya existe un egreso con el documento " & doc & ". ¿Registrar otro?", vbQuestion + vbYesNo, _
                      APP_NOMBRE) = vbNo Then Exit Sub
        End If
    End If
    Set hE = WS("EGRESOS")
    InicioMacro
    EscribirFila hE, SiguienteFila(hE), Array(SiguienteID(hE, 1, "EGR", 5), Date, concepto, metodo, CDbl(valor), doc, _
        benef, desc, UsuarioActual(), "VIGENTE")
    Bitacora "EGRESOS", "NUEVO", doc, concepto & " · " & Format$(valor, "#,##0")
    FinMacro
    GuardarLibro
End Sub

Public Sub AnularEgreso()
    Dim hE As Worksheet, r As Long, motivo As String, doc As String
    If Not RequierePermiso("FINANZAS") Then Exit Sub
    Set hE = WS("EGRESOS")
    If ActiveSheet.Name <> hE.Name Or ActiveCell.Row < FILA_DATOS Then
        MsgBox "Seleccione una celda en la fila del egreso a anular.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    r = ActiveCell.Row
    If Vacio(hE.Cells(r, 1).Value) Or hE.Cells(r, 10).Value <> "VIGENTE" Then
        MsgBox "La fila seleccionada no es un egreso vigente.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    doc = CStr(hE.Cells(r, 6).Value)
    If Left$(doc, 3) = "OC-" Or Left$(doc, 3) = "AB-" Or Left$(doc, 3) = "FV-" Then
        MsgBox "Este egreso está ligado a " & doc & ". Corríjalo desde su módulo de origen.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    motivo = PedirTexto("Motivo de la anulación de " & hE.Cells(r, 1).Value & ":")
    If motivo = "" Then Exit Sub
    Application.EnableEvents = False
    hE.Cells(r, 10).Value = "ANULADO"
    hE.Cells(r, 8).Value = hE.Cells(r, 8).Value & " [ANULADO: " & motivo & "]"
    Application.EnableEvents = True
    Bitacora "EGRESOS", "ANULAR", CStr(hE.Cells(r, 1).Value), motivo
    GuardarLibro
End Sub

'------------------------------------------------------------------ GASTOS FIJOS
Public Sub GenerarGastosFijos()
    Dim hG As Worksheet, hE As Worksheet, r As Long, mes As Date, doc As String, n As Long, dia As Long
    Dim fecha As Date, lista As String
    If Not RequierePermiso("FINANZAS") Then Exit Sub
    Set hG = WS("GASTOS FIJOS")
    Set hE = WS("EGRESOS")
    mes = DateSerial(Year(Date), Month(Date), 1)
    If MsgBox("¿Generar los egresos de gastos fijos de " & Format$(mes, "mm/yyyy") & "?" & vbCrLf & _
              "Solo se crean los que no se hayan generado ese mes.", vbQuestion + vbYesNo, APP_NOMBRE) = vbNo Then Exit Sub
    On Error GoTo FALLO
    InicioMacro
    IniciarTransaccion
    For r = FILA_DATOS To UltimaFila(hG, 1)
        If Not Vacio(hG.Cells(r, 1).Value) And Clave(CStr(hG.Cells(r, 8).Value)) = "SI" Then
            doc = hG.Cells(r, 1).Value & "-" & Format$(mes, "yyyymm")
            If BuscarFila(hE, 6, doc) = 0 Then
                dia = Application.WorksheetFunction.Min(NumV(hG.Cells(r, 5).Value), Day(DateSerial(Year(mes), Month(mes) + 1, 0)))
                If dia < 1 Then dia = 1
                fecha = DateSerial(Year(mes), Month(mes), dia)
                EscribirFila hE, SiguienteFila(hE), Array(SiguienteID(hE, 1, "EGR", 5), fecha, hG.Cells(r, 2).Value, _
                    hG.Cells(r, 6).Value, NumV(hG.Cells(r, 4).Value), doc, hG.Cells(r, 7).Value, hG.Cells(r, 3).Value, _
                    UsuarioActual(), "VIGENTE")
                n = n + 1
                lista = lista & vbCrLf & "· " & hG.Cells(r, 3).Value & ": " & Format$(hG.Cells(r, 4).Value, """$"" #,##0")
            End If
            EscribirCelda hG, r, 9, mes
        End If
    Next r
    Bitacora "GASTOS FIJOS", "GENERAR", Format$(mes, "yyyymm"), n & " egresos"
    CerrarTransaccion
    FinMacro
    GuardarLibro
    MsgBox IIf(n = 0, "No había gastos pendientes por generar este mes.", n & " egresos generados:" & lista), _
           vbInformation, APP_NOMBRE
    Exit Sub
FALLO:
    DeshacerTransaccion
    FinMacro
    ErrorGeneral "GenerarGastosFijos"
End Sub

'------------------------------------------------------------------ CIERRES DE CAJA
Public Sub CierreDiario()
    HacerCierre "DIARIO"
End Sub

Public Sub CierreSemanal()
    HacerCierre "SEMANAL"
End Sub

Public Sub CierreMensual()
    HacerCierre "MENSUAL"
End Sub

Private Sub HacerCierre(ByVal tipo As String)
    Dim desde As Date, hasta As Date, hCi As Worksheet, cuentas As Range, i As Long, cuenta As String
    Dim sistema As Double, contado As Variant, r As Long, idC As String, obs As String, resumen As String
    Dim ing As Double, egr As Double, wf As WorksheetFunction

    If Not RequierePermiso("FINANZAS") Then Exit Sub
    hasta = Date
    Select Case tipo
        Case "DIARIO": desde = hasta
        Case "SEMANAL": desde = hasta - Weekday(hasta, vbMonday) + 1
        Case Else: desde = DateSerial(Year(hasta), Month(hasta), 1)
    End Select
    Set hCi = WS("BD_CIERRES")
    For r = FILA_DATOS To UltimaFila(hCi, 1)
        If hCi.Cells(r, 2).Value = tipo And hCi.Cells(r, 5).Value = hasta Then
            MsgBox "Ya existe un cierre " & tipo & " con corte " & Format$(hasta, "dd/mm/yyyy") & " (" & _
                   hCi.Cells(r, 1).Value & ").", vbExclamation, APP_NOMBRE
            Exit Sub
        End If
    Next r
    Set wf = Application.WorksheetFunction
    Set cuentas = NR("cfg_Cuentas")
    Application.Calculate
    On Error GoTo FALLO
    Application.EnableEvents = False
    IniciarTransaccion
    idC = SiguienteID(hCi, 1, "CIE", 4)
    For i = 1 To cuentas.Rows.Count
        cuenta = CStr(cuentas.Cells(i, 1).Value)
        If cuenta <> "" Then
            ing = wf.SumIfs(NR("ing_Valor"), NR("ing_Met"), cuenta, NR("ing_Estado"), "VIGENTE", NR("ing_Fecha"), "<=" & CLng(hasta))
            egr = wf.SumIfs(NR("egr_Valor"), NR("egr_Met"), cuenta, NR("egr_Estado"), "VIGENTE", NR("egr_Fecha"), "<=" & CLng(hasta))
            sistema = NumV(cuentas.Cells(i, 2).Value) + ing - egr
            contado = PedirNumero("Cierre " & tipo & " · " & cuenta & vbCrLf & "Saldo según el sistema: " & _
                                  Format$(sistema, """$"" #,##0") & vbCrLf & vbCrLf & "Saldo contado / extracto:", sistema, -1E+15)
            If IsEmpty(contado) Then GoTo CANCELA
            obs = ""
            If contado <> sistema Then obs = PedirTexto("Diferencia de " & Format$(contado - sistema, """$"" #,##0") & _
                                                       ". Explique la diferencia:", "", False)
            EscribirFila hCi, SiguienteFila(hCi), Array(idC, tipo, hasta, desde, hasta, cuenta, sistema, CDbl(contado), _
                CDbl(contado) - sistema, UsuarioActual(), obs)
            resumen = resumen & vbCrLf & cuenta & ": " & Format$(contado, """$"" #,##0") & _
                      IIf(contado <> sistema, "  (dif. " & Format$(contado - sistema, """$"" #,##0") & ")", "")
        End If
    Next i
    Bitacora "CAJA Y BANCOS", "CIERRE " & tipo, idC, Format$(desde, "dd/mm/yyyy") & " - " & Format$(hasta, "dd/mm/yyyy")
    CerrarTransaccion
    Application.EnableEvents = True
    GuardarLibro
    MsgBox "Cierre " & tipo & " " & idC & " registrado:" & resumen, vbInformation, APP_NOMBRE
    Exit Sub
CANCELA:
    DeshacerTransaccion
    Application.EnableEvents = True
    MsgBox "Cierre cancelado. No se guardó nada.", vbInformation, APP_NOMBRE
    Exit Sub
FALLO:
    DeshacerTransaccion
    Application.EnableEvents = True
    ErrorGeneral "Cierre de caja"
End Sub

'------------------------------------------------------------------ COMISIONES
Public Sub PagarComisiones()
    Dim hCo As Worksheet, hE As Worksheet, r As Long, mes As Date, doc As String, pend As Double, metodo As String
    Dim lista As String, total As Double, n As Long
    If Not RequierePermiso("COMISIONES") Then Exit Sub
    Set hCo = WS("COMISIONES")
    Set hE = WS("EGRESOS")
    Application.Calculate
    mes = CDate(hCo.Range("C4").Value)
    For r = 7 To 36
        If Not Vacio(hCo.Cells(r, 2).Value) And NumV(hCo.Cells(r, 11).Value) > 0 Then
            lista = lista & vbCrLf & hCo.Cells(r, 3).Value & ": " & Format$(hCo.Cells(r, 11).Value, """$"" #,##0")
            total = total + NumV(hCo.Cells(r, 11).Value)
        End If
    Next r
    If total = 0 Then MsgBox "No hay comisiones pendientes para " & Format$(mes, "mm/yyyy") & ".", vbInformation, APP_NOMBRE: Exit Sub
    If MsgBox("Pagar comisiones de " & Format$(mes, "mm/yyyy") & ":" & lista & vbCrLf & vbCrLf & "Total: " & _
              Format$(total, """$"" #,##0"), vbQuestion + vbYesNo, APP_NOMBRE) = vbNo Then Exit Sub
    metodo = ElegirDeLista("Método de pago:", ListaCatalogo("Métodos de pago / cuentas"), "Transferencia")
    If metodo = "" Then Exit Sub
    On Error GoTo FALLO
    InicioMacro
    IniciarTransaccion
    For r = 7 To 36
        pend = NumV(hCo.Cells(r, 11).Value)
        If Not Vacio(hCo.Cells(r, 2).Value) And pend > 0 Then
            doc = "COM-" & (Year(mes) * 100 + Month(mes)) & "-" & hCo.Cells(r, 2).Value
            EscribirFila hE, SiguienteFila(hE), Array(SiguienteID(hE, 1, "EGR", 5), Date, "Comisiones", metodo, pend, doc, _
                hCo.Cells(r, 2).Value, "Comisión " & Format$(mes, "mm/yyyy") & " " & hCo.Cells(r, 3).Value, _
                UsuarioActual(), "VIGENTE")
            n = n + 1
        End If
    Next r
    Bitacora "COMISIONES", "PAGAR", Format$(mes, "yyyymm"), n & " vendedores · " & Format$(total, "#,##0")
    CerrarTransaccion
    FinMacro
    GuardarLibro
    MsgBox n & " pagos de comisión registrados en EGRESOS.", vbInformation, APP_NOMBRE
    Exit Sub
FALLO:
    DeshacerTransaccion
    FinMacro
    ErrorGeneral "PagarComisiones"
End Sub
