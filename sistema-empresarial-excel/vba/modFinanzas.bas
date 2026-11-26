Attribute VB_Name = "modFinanzas"
'==============================================================
' modFinanzas - INGRESOS, EGRESOS, GASTOS FIJOS, CARTERA Y CAJA
'  Regla anti-duplicados: las ventas y abonos solo se registran
'  desde FACTURACIÓN / CARTERA; las compras desde COMPRAS. Los
'  formularios manuales no permiten esos tipos.
'==============================================================
Option Explicit

'-------------------------------------------------------------- núcleo
Public Function RegistrarIngreso(ByVal fecha As Date, ByVal origen As String, ByVal tipo As String, _
                                 ByVal metodo As String, ByVal doc As String, ByVal tercero As String, _
                                 ByVal concepto As String, ByVal valor As Double, ByVal obs As String) As String
    Dim ws As Worksheet, r As Long, a As Variant, id As String
    Set ws = Hoja(SH_ING)
    r = NuevaFila(ws, ING_ID, ING_FIRST, ING_LAST)
    id = SiguienteID(ws, ING_ID, ING_FIRST, "ING-", 5)
    a = NuevaFilaArr(ING_NCOLS)
    a(1, ING_ID) = id
    a(1, ING_FECHA) = fecha
    a(1, ING_ORIGEN) = origen
    a(1, ING_TIPO) = tipo
    If metodo <> "" Then
        a(1, ING_CLASIF) = tipo & " - " & metodo
    Else
        a(1, ING_CLASIF) = tipo
    End If
    a(1, ING_METODO) = metodo
    If tipo = "NOTA CRÉDITO" Then
        a(1, ING_CUENTA) = ""
    Else
        a(1, ING_CUENTA) = CuentaDeMetodo(metodo)
    End If
    a(1, ING_DOC) = doc
    a(1, ING_TERCERO) = tercero
    a(1, ING_CONCEPTO) = concepto
    a(1, ING_VALOR) = valor
    a(1, ING_USUARIO) = UsuarioActual()
    a(1, ING_REGISTRO) = Now
    a(1, ING_ESTADO) = "ACTIVO"
    a(1, ING_OBS) = obs
    EscribirFila ws, r, 1, a
    RegistrarIngreso = id
End Function

Public Function RegistrarEgreso(ByVal fecha As Date, ByVal concepto As String, ByVal categoria As String, _
                                ByVal tipo As String, ByVal valor As Double, ByVal metodo As String, _
                                ByVal responsable As String, ByVal comprobante As String, ByVal obs As String, _
                                ByVal origen As String, ByVal doc As String) As String
    Dim ws As Worksheet, r As Long, a As Variant, id As String
    Set ws = Hoja(SH_EGR)
    r = NuevaFila(ws, EGR_ID, EGR_FIRST, EGR_LAST)
    id = SiguienteID(ws, EGR_ID, EGR_FIRST, "EGR-", 5)
    a = NuevaFilaArr(EGR_NCOLS)
    a(1, EGR_ID) = id
    a(1, EGR_FECHA) = fecha
    a(1, EGR_CONCEPTO) = concepto
    a(1, EGR_CATEGORIA) = categoria
    a(1, EGR_TIPO) = tipo
    a(1, EGR_VALOR) = valor
    a(1, EGR_METODO) = metodo
    a(1, EGR_CUENTA) = CuentaDeMetodo(metodo)
    a(1, EGR_RESPONSABLE) = responsable
    a(1, EGR_COMPROBANTE) = comprobante
    a(1, EGR_OBS) = obs
    a(1, EGR_ORIGEN) = origen
    a(1, EGR_DOC) = doc
    a(1, EGR_USUARIO) = UsuarioActual()
    a(1, EGR_REGISTRO) = Now
    a(1, EGR_ESTADO) = "ACTIVO"
    EscribirFila ws, r, 1, a
    RegistrarEgreso = id
End Function

' Crea una cuenta por cobrar o por pagar (las columnas de saldo/mora/estado son fórmulas)
Public Function CrearObligacion(ByVal tipo As String, ByVal codTer As String, ByVal tercero As String, _
                                ByVal doc As String, ByVal fecha As Date, ByVal valorIni As Double, _
                                ByVal pagoIni As Double, ByVal vence As Date, ByVal obs As String) As String
    Dim ws As Worksheet, r As Long, a As Variant, id As String, pref As String
    Set ws = Hoja(SH_CAR)
    r = NuevaFila(ws, CAR_ID, CAR_FIRST, CAR_LAST)
    If tipo = "POR COBRAR" Then
        pref = "CXC-"
    Else
        pref = "CXP-"
    End If
    id = SiguienteID(ws, CAR_ID, CAR_FIRST, pref, 4)
    a = NuevaFilaArr(CAR_NCOLS)
    a(1, CAR_ID) = id
    a(1, CAR_TIPO) = tipo
    a(1, CAR_CODTER) = codTer
    a(1, CAR_TERCERO) = tercero
    a(1, CAR_DOC) = doc
    a(1, CAR_FECHA) = fecha
    a(1, CAR_VALORINI) = valorIni
    a(1, CAR_PAGOINI) = pagoIni
    a(1, CAR_VENCE) = vence
    a(1, CAR_REGESTADO) = "ACTIVA"
    If obs <> "" Then a(1, CAR_OBS) = obs
    EscribirValores ws, r, 1, a       ' respeta las columnas con fórmulas
    CrearObligacion = id
End Function

'-------------------------------------------------------------- INGRESOS manuales
Public Sub RegistrarIngresoManual()
    Dim ws As Worksheet, fecha As Variant, tipo As String, concepto As String, valor As Double, metodo As String
    Dim id As String
    Set ws = Hoja(SH_ING)
    Recalcular
    fecha = ws.Range(FING_FECHA).Value
    tipo = Tv(ws.Range(FING_TIPO).Value)
    concepto = Tv(ws.Range(FING_CONCEPTO).Value)
    valor = Nv(ws.Range(FING_VALOR).Value)
    metodo = Tv(ws.Range(FING_METODO).Value)
    If Not IsDate(fecha) Then Msg "Escriba una fecha válida.", vbExclamation: Exit Sub
    If Not EnLista(LST_TIPOSINGRESO, tipo) Then
        Msg "Seleccione un tipo de ingreso de la lista." & vbLf & _
            "Las VENTAS se registran en FACTURACIÓN y los ABONOS en CARTERA (así no se duplican).", vbExclamation
        Exit Sub
    End If
    If concepto = "" Then Msg "Escriba el concepto del ingreso.", vbExclamation: Exit Sub
    If valor <= 0 Then Msg "El valor debe ser mayor que cero.", vbExclamation: Exit Sub
    If Not EnLista(LST_METODOS, metodo) Then Msg "Seleccione un método de pago válido.", vbExclamation: Exit Sub
    If Msg("¿Registrar ingreso de " & Moneda(valor) & " (" & tipo & ", " & metodo & ")?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Iniciar
    On Error GoTo Fallo
    TxIniciar
    TxTabla SH_ING, ING_ID, ING_FIRST, 1, ING_NCOLS
    id = RegistrarIngreso(CDate(fecha), "MANUAL", tipo, metodo, Tv(ws.Range(FING_DOC).Value), Tv(ws.Range(FING_TERCERO).Value), _
                          concepto, valor, Tv(ws.Range(FING_OBS).Value))
    Auditar "INGRESOS", "REGISTRAR INGRESO", id, concepto & " " & Moneda(valor)
    TxConfirmar
    ws.Range(FING_CONCEPTO).Value = ""
    ws.Range(FING_VALOR).Value = ""
    ws.Range(FING_TERCERO).Value = ""
    ws.Range(FING_DOC).Value = ""
    ws.Range(FING_OBS).Value = ""
    Terminar
    Msg "Ingreso " & id & " registrado.", vbInformation
    Exit Sub
Fallo:
    TxRevertir
    Terminar
    Msg "No se pudo registrar: " & Err.Description, vbCritical
End Sub

Public Sub AnularIngreso()
    Dim ws As Worksheet, r As Long, motivo As String
    Set ws = Hoja(SH_ING)
    If ActiveSheet.Name <> SH_ING Or ActiveCell.Row < ING_FIRST Then Msg "Seleccione una celda de la fila del ingreso a anular.", vbExclamation: Exit Sub
    r = ActiveCell.Row
    If Tv(ws.Cells(r, ING_ID).Value) = "" Then Exit Sub
    If Tv(ws.Cells(r, ING_ESTADO).Value) <> "ACTIVO" Then Msg "Ese ingreso ya está anulado.", vbExclamation: Exit Sub
    If Tv(ws.Cells(r, ING_ORIGEN).Value) <> "MANUAL" Then
        Msg "Este ingreso proviene de " & Tv(ws.Cells(r, ING_ORIGEN).Value) & "." & vbLf & _
            "Para revertirlo anule el documento de origen (" & Tv(ws.Cells(r, ING_DOC).Value) & ").", vbExclamation
        Exit Sub
    End If
    If Not Autorizar("Anular el ingreso " & Tv(ws.Cells(r, ING_ID).Value)) Then Exit Sub
    motivo = Pedir("Motivo de la anulación:")
    If motivo = "" Then Exit Sub
    Iniciar
    ws.Cells(r, ING_ESTADO).Value = "ANULADO"
    ws.Cells(r, ING_OBS).Value = Trim(Tv(ws.Cells(r, ING_OBS).Value) & " [ANULADO: " & motivo & "]")
    Auditar "INGRESOS", "ANULAR INGRESO", Tv(ws.Cells(r, ING_ID).Value), motivo
    Terminar
    Msg "Ingreso anulado.", vbInformation
End Sub

'-------------------------------------------------------------- EGRESOS manuales
Public Sub RegistrarEgresoManual()
    Dim ws As Worksheet, fecha As Variant, concepto As String, cat As String, tipo As String
    Dim valor As Double, metodo As String, id As String
    Set ws = Hoja(SH_EGR)
    Recalcular
    fecha = ws.Range(FEGR_FECHA).Value
    concepto = Tv(ws.Range(FEGR_CONCEPTO).Value)
    cat = Tv(ws.Range(FEGR_CATEGORIA).Value)
    tipo = Tv(ws.Range(FEGR_TIPO).Value)
    valor = Nv(ws.Range(FEGR_VALOR).Value)
    metodo = Tv(ws.Range(FEGR_METODO).Value)
    If Not IsDate(fecha) Then Msg "Escriba una fecha válida.", vbExclamation: Exit Sub
    If concepto = "" Then Msg "Escriba el concepto del egreso.", vbExclamation: Exit Sub
    If Not EnLista(LST_CATGASTO, cat) Then Msg "Seleccione una categoría de la lista.", vbExclamation: Exit Sub
    If Not EnLista(LST_TIPOGASTO, tipo) Then Msg "Seleccione el tipo: FIJO, VARIABLE o EXTRAORDINARIO.", vbExclamation: Exit Sub
    If valor <= 0 Then Msg "El valor debe ser mayor que cero.", vbExclamation: Exit Sub
    If Not EnLista(LST_METODOS, metodo) Then Msg "Seleccione un método de pago válido.", vbExclamation: Exit Sub
    If UCase(cat) = UCase(Tv(CfgV(CFG_CATCOMPRAS))) Then
        If Msg("Las compras de MERCANCÍA deben registrarse en la hoja COMPRAS para que aumenten el inventario." & vbLf & _
               "¿Registrar de todos modos como egreso sin mercancía?", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    End If
    If Msg("¿Registrar egreso de " & Moneda(valor) & " (" & cat & ")?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Iniciar
    On Error GoTo Fallo
    TxIniciar
    TxTabla SH_EGR, EGR_ID, EGR_FIRST, 1, EGR_NCOLS
    id = RegistrarEgreso(CDate(fecha), concepto, cat, tipo, valor, metodo, Tv(ws.Range(FEGR_RESP).Value), _
                         Tv(ws.Range(FEGR_COMPROB).Value), Tv(ws.Range(FEGR_OBS).Value), "MANUAL", "")
    Auditar "EGRESOS", "REGISTRAR EGRESO", id, concepto & " " & Moneda(valor)
    TxConfirmar
    ws.Range(FEGR_CONCEPTO).Value = ""
    ws.Range(FEGR_VALOR).Value = ""
    ws.Range(FEGR_COMPROB).Value = ""
    ws.Range(FEGR_OBS).Value = ""
    Terminar
    Msg "Egreso " & id & " registrado.", vbInformation
    Exit Sub
Fallo:
    TxRevertir
    Terminar
    Msg "No se pudo registrar: " & Err.Description, vbCritical
End Sub

Public Sub AnularEgreso()
    Dim ws As Worksheet, r As Long, motivo As String, org As String
    Set ws = Hoja(SH_EGR)
    If ActiveSheet.Name <> SH_EGR Or ActiveCell.Row < EGR_FIRST Then Msg "Seleccione una celda de la fila del egreso a anular.", vbExclamation: Exit Sub
    r = ActiveCell.Row
    If Tv(ws.Cells(r, EGR_ID).Value) = "" Then Exit Sub
    If Tv(ws.Cells(r, EGR_ESTADO).Value) <> "ACTIVO" Then Msg "Ese egreso ya está anulado.", vbExclamation: Exit Sub
    org = Tv(ws.Cells(r, EGR_ORIGEN).Value)
    If org = "COMPRAS" Then
        Msg "Este egreso es el pago de una compra de mercancía. No se puede anular desde aquí.", vbExclamation
        Exit Sub
    End If
    If Not Autorizar("Anular el egreso " & Tv(ws.Cells(r, EGR_ID).Value)) Then Exit Sub
    motivo = Pedir("Motivo de la anulación:")
    If motivo = "" Then Exit Sub
    Iniciar
    ws.Cells(r, EGR_ESTADO).Value = "ANULADO"
    ws.Cells(r, EGR_OBS).Value = Trim(Tv(ws.Cells(r, EGR_OBS).Value) & " [ANULADO: " & motivo & "]")
    Auditar "EGRESOS", "ANULAR EGRESO", Tv(ws.Cells(r, EGR_ID).Value), motivo
    Terminar
    Msg "Egreso anulado.", vbInformation
End Sub

'-------------------------------------------------------------- GASTOS FIJOS
Public Sub PagarGastoFijo()
    Dim ws As Worksheet, r As Long, concepto As String, s As String, valor As Double, metodo As String, id As String
    Set ws = Hoja(SH_GFI)
    If gTest Then
        r = GFI_FIRST
    Else
        If ActiveSheet.Name <> SH_GFI Or ActiveCell.Row < GFI_FIRST Then Msg "Seleccione una celda de la fila del gasto fijo a pagar.", vbExclamation: Exit Sub
        r = ActiveCell.Row
    End If
    Recalcular
    concepto = Tv(ws.Cells(r, GFI_CONCEPTO).Value)
    If concepto = "" Then Msg "La fila seleccionada no tiene un gasto fijo.", vbExclamation: Exit Sub
    If UCase(Tv(ws.Cells(r, GFI_ACTIVO).Value)) = "NO" Then Msg "Ese gasto fijo está inactivo.", vbExclamation: Exit Sub
    If Tv(ws.Cells(r, GFI_ESTADO).Value) = "PAGADO" Then
        If Msg("El gasto '" & concepto & "' ya figura PAGADO este mes. ¿Registrar otro pago?", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    End If
    s = Pedir("Valor a pagar de '" & concepto & "':", "Pagar gasto fijo", CStr(Nv(ws.Cells(r, GFI_VALOR).Value)))
    If s = "" Then Exit Sub
    valor = NumTexto(s)
    If valor <= 0 Then Msg "Valor no válido.", vbExclamation: Exit Sub
    metodo = Tv(ws.Cells(r, GFI_METODO).Value)
    If Not EnLista(LST_METODOS, metodo) Then metodo = "Efectivo"
    Iniciar
    On Error GoTo Fallo
    TxIniciar
    TxTabla SH_EGR, EGR_ID, EGR_FIRST, 1, EGR_NCOLS
    id = RegistrarEgreso(Date, concepto & " " & Format(Date, "mm/yyyy"), Tv(ws.Cells(r, GFI_CATEGORIA).Value), "FIJO", valor, metodo, _
                         Tv(ws.Cells(r, GFI_RESPONSABLE).Value), "", "", "GASTO FIJO", "GF|" & concepto)
    Auditar "GASTOS FIJOS", "PAGAR", id, concepto & " " & Moneda(valor)
    TxConfirmar
    Terminar
    Msg "Pago registrado en EGRESOS (" & id & ").", vbInformation
    Exit Sub
Fallo:
    TxRevertir
    Terminar
    Msg "No se pudo registrar: " & Err.Description, vbCritical
End Sub

'-------------------------------------------------------------- CARTERA
Public Sub RegistrarAbono()
    Dim ws As Worksheet, combo As String, id As String, f As Long, saldo As Double, valor As Double
    Dim fecha As Variant, metodo As String, tipo As String, doc As String, tercero As String, nuevo As Double
    Dim wh As Worksheet, fh As Long, ref As String
    Set ws = Hoja(SH_CAR)
    Recalcular
    combo = Tv(ws.Range(FCAR_OBLIG).Value)
    If InStr(combo, " |") > 0 Then
        id = Trim(Left(combo, InStr(combo, " |") - 1))
    Else
        id = combo
    End If
    If id = "" Then Msg "Seleccione la obligación (cliente o proveedor) en la lista.", vbExclamation: Exit Sub
    f = BuscarFila(ws, CAR_ID, CAR_FIRST, id)
    If f = 0 Then Msg "No se encontró la obligación " & id, vbExclamation: Exit Sub
    If Tv(ws.Cells(f, CAR_REGESTADO).Value) <> "ACTIVA" Then Msg "La obligación está ANULADA.", vbExclamation: Exit Sub
    saldo = Nv(ws.Cells(f, CAR_SALDO).Value)
    valor = Nv(ws.Range(FCAR_VALOR).Value)
    fecha = ws.Range(FCAR_FECHA).Value
    metodo = Tv(ws.Range(FCAR_METODO).Value)
    If saldo <= 0 Then Msg "Esta obligación no tiene saldo pendiente.", vbExclamation: Exit Sub
    If valor <= 0 Then Msg "Escriba el valor del abono.", vbExclamation: Exit Sub
    If valor > saldo + 0.5 Then Msg "El abono (" & Moneda(valor) & ") supera el saldo pendiente (" & Moneda(saldo) & ").", vbExclamation: Exit Sub
    If Not IsDate(fecha) Then Msg "Escriba una fecha válida.", vbExclamation: Exit Sub
    If Not EnLista(LST_METODOS, metodo) Then Msg "Seleccione un método de pago válido.", vbExclamation: Exit Sub
    tipo = Tv(ws.Cells(f, CAR_TIPO).Value)
    doc = Tv(ws.Cells(f, CAR_DOC).Value)
    tercero = Tv(ws.Cells(f, CAR_TERCERO).Value)
    If Msg("¿Registrar abono de " & Moneda(valor) & " a " & doc & " (" & tercero & ")?" & vbLf & _
           "Saldo actual: " & Moneda(saldo) & "  ->  nuevo saldo: " & Moneda(saldo - valor), vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Iniciar
    On Error GoTo Fallo
    TxIniciar
    TxTabla SH_ING, ING_ID, ING_FIRST, 1, ING_NCOLS
    TxTabla SH_EGR, EGR_ID, EGR_FIRST, 1, EGR_NCOLS
    If tipo = "POR COBRAR" Then
        ref = RegistrarIngreso(CDate(fecha), "CARTERA", "ABONO", metodo, doc, tercero, "Abono a " & doc, valor, Tv(ws.Range(FCAR_OBS).Value))
        nuevo = saldo - valor
        If nuevo <= 0.5 Then
            Set wh = Hoja(SH_HFA)
            fh = BuscarFila(wh, HFA_FACTURA, HFA_FIRST, doc)
            If fh > 0 Then Modificar wh, fh, HFA_ESTADOPAGO, "PAGADA"
        End If
    Else
        ref = RegistrarEgreso(CDate(fecha), "Abono a " & doc & " - " & tercero, Tv(CfgV(CFG_CATCOMPRAS)), "VARIABLE", valor, metodo, _
                              UsuarioActual(), "", Tv(ws.Range(FCAR_OBS).Value), "CARTERA", doc)
    End If
    Auditar "CARTERA", "ABONO " & tipo, doc, Moneda(valor) & " (" & ref & ")"
    TxConfirmar
    ws.Range(FCAR_OBLIG).Value = ""
    ws.Range(FCAR_VALOR).Value = ""
    ws.Range(FCAR_OBS).Value = ""
    Terminar
    Msg "Abono registrado (" & ref & "). Nuevo saldo de " & doc & ": " & Moneda(saldo - valor), vbInformation
    Exit Sub
Fallo:
    TxRevertir
    Terminar
    Msg "No se pudo registrar el abono: " & Err.Description, vbCritical
End Sub

Public Sub FiltrarCarteraVencida()
    Dim ws As Worksheet
    Set ws = Hoja(SH_CAR)
    On Error Resume Next
    ws.Unprotect Tv(CfgV(CFG_CLAVEHOJAS))
    QuitarFiltros ws
    ws.Range(ws.Cells(CAR_HDR, 1), ws.Cells(CAR_HDR, CAR_NCOLS)).AutoFilter Field:=CAR_ESTADO, Criteria1:="VENCIDA"
    ProtegerTodo
End Sub

'-------------------------------------------------------------- CAJA: cierres
Private Sub FijarPeriodo(ByVal tipo As String, ByVal d0 As Date, ByVal d1 As Date)
    Dim ws As Worksheet
    Set ws = Hoja(SH_CAJ)
    Iniciar
    ws.Range(FCAJ_TIPO).Value = tipo
    ws.Range(FCAJ_DESDE).Value = d0
    ws.Range(FCAJ_HASTA).Value = d1
    Terminar
End Sub

Public Sub CierreDiario()
    FijarPeriodo "DIARIO", Date, Date
End Sub

Public Sub CierreSemanal()
    Dim d0 As Date
    d0 = Date - Weekday(Date, vbMonday) + 1
    FijarPeriodo "SEMANAL", d0, d0 + 6
End Sub

Public Sub CierreMensual()
    FijarPeriodo "MENSUAL", DateSerial(Year(Date), Month(Date), 1), DateSerial(Year(Date), Month(Date) + 1, 0)
End Sub

Public Sub GuardarCierre()
    Dim ws As Worksheet, r As Long, a As Variant, i As Long
    Set ws = Hoja(SH_CAJ)
    Recalcular
    If Not IsDate(ws.Range(FCAJ_DESDE).Value) Or Not IsDate(ws.Range(FCAJ_HASTA).Value) Then
        Msg "Defina las fechas del cierre (desde / hasta).", vbExclamation: Exit Sub
    End If
    If Msg("¿Guardar el cierre " & Tv(ws.Range(FCAJ_TIPO).Value) & " del " & FechaCorta(ws.Range(FCAJ_DESDE).Value) & _
           " al " & FechaCorta(ws.Range(FCAJ_HASTA).Value) & "?" & vbLf & _
           "Saldo final total: " & Moneda(Nv(ws.Cells(FCAJ_PERTOT, 7).Value)), vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Iniciar
    On Error GoTo Fallo
    r = NuevaFila(ws, CIE_FECHA, CIE_FIRST, CIE_LAST)
    a = NuevaFilaArr(CIE_NCOLS)
    a(1, CIE_FECHA - CIE_COL0 + 1) = Now
    a(1, CIE_TIPO - CIE_COL0 + 1) = Tv(ws.Range(FCAJ_TIPO).Value)
    a(1, CIE_DESDE - CIE_COL0 + 1) = ws.Range(FCAJ_DESDE).Value
    a(1, CIE_HASTA - CIE_COL0 + 1) = ws.Range(FCAJ_HASTA).Value
    a(1, CIE_ENTRADAS - CIE_COL0 + 1) = Nv(ws.Cells(FCAJ_PERTOT, 4).Value)
    a(1, CIE_SALIDAS - CIE_COL0 + 1) = Nv(ws.Cells(FCAJ_PERTOT, 5).Value)
    a(1, CIE_SALDO - CIE_COL0 + 1) = Nv(ws.Cells(FCAJ_PERTOT, 7).Value)
    For i = 0 To 4
        a(1, CIE_C1 - CIE_COL0 + 1 + i) = Nv(ws.Cells(FCAJ_PER1 + i, 7).Value)
    Next i
    a(1, CIE_USUARIO - CIE_COL0 + 1) = UsuarioActual()
    EscribirFila ws, r, CIE_COL0, a
    Auditar "CAJA", "CIERRE " & Tv(ws.Range(FCAJ_TIPO).Value), "", FechaCorta(ws.Range(FCAJ_DESDE).Value) & " - " & FechaCorta(ws.Range(FCAJ_HASTA).Value)
    Terminar
    Msg "Cierre guardado en el historial de cierres.", vbInformation
    Exit Sub
Fallo:
    Terminar
    Msg "No se pudo guardar el cierre: " & Err.Description, vbCritical
End Sub

Public Sub PDFCierre()
    Dim ws As Worksheet, ruta As String
    Set ws = Hoja(SH_CAJ)
    On Error GoTo Fallo
    ruta = Carpeta("REPORTES") & Application.PathSeparator & "Cierre de caja " & Format(Now, "yyyy-mm-dd hhmm") & ".pdf"
    ws.Range("B3:H" & FCAJ_PERTOT).ExportAsFixedFormat Type:=xlTypePDF, Filename:=ruta, OpenAfterPublish:=True
    Exit Sub
Fallo:
    Msg "No se pudo generar el PDF: " & Err.Description, vbExclamation
End Sub
