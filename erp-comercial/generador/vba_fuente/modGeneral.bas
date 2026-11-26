Attribute VB_Name = "modGeneral"
'==============================================================================
' modGeneral - Utilidades comunes del ERP Comercial
' Acceso a hojas, configuracion, consecutivos, escritura de filas, bitacora,
' proteccion de hojas y control de errores.
'==============================================================================
Option Explicit

Public Const APP_NOMBRE As String = "ERP Comercial"
Public Const FILA_DATOS As Long = 6          ' primera fila de datos en tablas BD y maestros
Public Const FILA_DATOS_CLI As Long = 17     ' CLIENTES tiene la ficha arriba

' Registro de filas escritas durante una operacion (para deshacer si hay error)
Private mDeshacer As Collection

'------------------------------------------------------------------ texto
Public Function SinTildes(ByVal s As String) As String
    Dim a As Variant, b As Variant, i As Long
    a = Array(193, 201, 205, 211, 218, 225, 233, 237, 243, 250, 209, 241, 220, 252)
    b = Array("A", "E", "I", "O", "U", "a", "e", "i", "o", "u", "N", "n", "U", "u")
    For i = 0 To UBound(a)
        s = Replace(s, ChrW(a(i)), b(i))
    Next i
    SinTildes = s
End Function

Public Function Clave(ByVal s As String) As String
    Clave = UCase$(Trim$(SinTildes(s)))
End Function

' Devuelve el codigo de una etiqueta "COD - Descripcion"
Public Function CodigoDe(ByVal etiqueta As Variant) As String
    Dim s As String, p As Long
    s = Trim$(CStr(etiqueta))
    p = InStr(1, s, " - ")
    If p > 0 Then CodigoDe = Left$(s, p - 1) Else CodigoDe = s
End Function

'------------------------------------------------------------------ hojas y nombres
Public Function WS(ByVal nombreHoja As String) As Worksheet
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If Clave(sh.Name) = Clave(nombreHoja) Then
            Set WS = sh
            Exit Function
        End If
    Next sh
    Err.Raise vbObjectError + 513, APP_NOMBRE, "No existe la hoja " & nombreHoja
End Function

Public Function NR(ByVal nombre As String) As Range
    Set NR = ThisWorkbook.Names(nombre).RefersToRange
End Function

Public Function Cfg(ByVal nombre As String) As Variant
    Cfg = ThisWorkbook.Names(nombre).RefersToRange.Value
End Function

Public Sub SetCfg(ByVal nombre As String, ByVal valor As Variant)
    ThisWorkbook.Names(nombre).RefersToRange.Value = valor
End Sub

Public Function UltimaFila(ByVal sh As Worksheet, ByVal col As Long) As Long
    UltimaFila = sh.Cells(sh.Rows.Count, col).End(xlUp).Row
End Function

Public Function SiguienteFila(ByVal sh As Worksheet, Optional ByVal col As Long = 1, _
                              Optional ByVal fila0 As Long = FILA_DATOS) As Long
    Dim r As Long
    r = UltimaFila(sh, col) + 1
    If r < fila0 Then r = fila0
    SiguienteFila = r
End Function

' Siguiente ID con prefijo (ej. ING-00012) buscando el maximo existente en la columna
Public Function SiguienteID(ByVal sh As Worksheet, ByVal col As Long, ByVal pref As String, _
                            ByVal ancho As Integer, Optional ByVal fila0 As Long = FILA_DATOS) As String
    Dim ult As Long, datos As Variant, i As Long, v As String, n As Long, mx As Long
    ult = UltimaFila(sh, col)
    If ult >= fila0 Then
        datos = sh.Range(sh.Cells(fila0, col), sh.Cells(ult, col)).Value
        If Not IsArray(datos) Then
            v = CStr(datos)
            If Left$(v, Len(pref) + 1) = pref & "-" Then mx = Val(Mid$(v, Len(pref) + 2))
        Else
            For i = 1 To UBound(datos, 1)
                v = CStr(datos(i, 1))
                If Left$(v, Len(pref) + 1) = pref & "-" Then
                    n = Val(Mid$(v, Len(pref) + 2))
                    If n > mx Then mx = n
                End If
            Next i
        End If
    End If
    SiguienteID = pref & "-" & Format$(mx + 1, String$(ancho, "0"))
End Function

' Fila donde aparece un valor en una columna (0 si no existe)
Public Function BuscarFila(ByVal sh As Worksheet, ByVal col As Long, ByVal valor As Variant, _
                           Optional ByVal fila0 As Long = FILA_DATOS) As Long
    Dim ult As Long, datos As Variant, i As Long
    ult = UltimaFila(sh, col)
    If ult < fila0 Then Exit Function
    datos = sh.Range(sh.Cells(fila0, col), sh.Cells(ult, col)).Value
    If Not IsArray(datos) Then
        If Clave(CStr(datos)) = Clave(CStr(valor)) Then BuscarFila = fila0
        Exit Function
    End If
    For i = 1 To UBound(datos, 1)
        If Clave(CStr(datos(i, 1))) = Clave(CStr(valor)) Then
            BuscarFila = fila0 + i - 1
            Exit Function
        End If
    Next i
End Function

' Valor de una columna en la fila donde otra columna = clave (Empty si no existe)
Public Function Buscar(ByVal sh As Worksheet, ByVal colClave As Long, ByVal valor As Variant, _
                       ByVal colDevolver As Long, Optional ByVal fila0 As Long = FILA_DATOS) As Variant
    Dim r As Long
    r = BuscarFila(sh, colClave, valor, fila0)
    If r > 0 Then Buscar = sh.Cells(r, colDevolver).Value Else Buscar = Empty
End Function

' Escribe una fila completa en una tabla BD copiando el formato de la fila anterior
Public Sub EscribirFila(ByVal sh As Worksheet, ByVal fila As Long, ByVal valores As Variant)
    Dim n As Long
    n = UBound(valores) - LBound(valores) + 1
    If fila > FILA_DATOS Then
        sh.Range(sh.Cells(fila - 1, 1), sh.Cells(fila - 1, n)).Copy Destination:=sh.Cells(fila, 1)
    Else
        With sh.Range(sh.Cells(fila, 1), sh.Cells(fila, n))
            .Font.Name = "Arial"
            .Font.Size = 9
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(191, 191, 191)
        End With
    End If
    sh.Range(sh.Cells(fila, 1), sh.Cells(fila, n)).Value = valores
    RegistrarEscritura sh, fila, 1, n
End Sub

'------------------------------------------------------------------ deshacer (pseudo transaccion)
Public Sub IniciarTransaccion()
    Set mDeshacer = New Collection
End Sub

Public Sub RegistrarEscritura(ByVal sh As Worksheet, ByVal fila As Long, ByVal col1 As Long, ByVal col2 As Long, _
                              Optional ByVal valorAnterior As Variant)
    If mDeshacer Is Nothing Then Exit Sub
    Dim it(0 To 4) As Variant
    it(0) = sh.Name: it(1) = fila: it(2) = col1: it(3) = col2
    If IsMissing(valorAnterior) Then it(4) = "#VACIAR#" Else it(4) = valorAnterior
    mDeshacer.Add it
End Sub

' Guarda el valor anterior de una celda para restaurarlo si la operacion falla
Public Sub EscribirCelda(ByVal sh As Worksheet, ByVal fila As Long, ByVal col As Long, ByVal valor As Variant)
    Dim ant As Variant
    If sh.Cells(fila, col).HasFormula Then ant = "=" & Mid$(sh.Cells(fila, col).Formula, 2) Else ant = sh.Cells(fila, col).Value
    If IsEmpty(ant) Then ant = ""
    RegistrarEscritura sh, fila, col, col, ant
    sh.Cells(fila, col).Value = valor
End Sub

Public Sub DeshacerTransaccion()
    Dim i As Long, it As Variant, sh As Worksheet
    If mDeshacer Is Nothing Then Exit Sub
    On Error Resume Next
    For i = mDeshacer.Count To 1 Step -1
        it = mDeshacer(i)
        Set sh = ThisWorkbook.Worksheets(CStr(it(0)))
        If CStr(it(4)) = "#VACIAR#" Then
            sh.Range(sh.Cells(it(1), it(2)), sh.Cells(it(1), it(3))).ClearContents
        Else
            sh.Cells(it(1), it(2)).Value = it(4)
        End If
    Next i
    Set mDeshacer = Nothing
    On Error GoTo 0
End Sub

Public Sub CerrarTransaccion()
    Set mDeshacer = Nothing
End Sub

'------------------------------------------------------------------ rendimiento
Public Sub InicioMacro()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait
End Sub

Public Sub FinMacro()
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Application.CutCopyMode = False
End Sub

'------------------------------------------------------------------ proteccion
Public Function ClaveHojas() As String
    On Error Resume Next
    ClaveHojas = CStr(Cfg("cfg_ClaveHojas"))
    If ClaveHojas = "" Then ClaveHojas = "ERP"
End Function

Public Sub ProtegerHoja(ByVal sh As Worksheet)
    On Error Resume Next
    sh.Unprotect ClaveHojas()
    sh.Protect Password:=ClaveHojas(), DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    sh.EnableSelection = xlNoRestrictions
End Sub

Public Sub DesprotegerHoja(ByVal sh As Worksheet)
    On Error Resume Next
    sh.Unprotect ClaveHojas()
End Sub

Public Sub ProtegerTodo()
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        ProtegerHoja sh
    Next sh
End Sub

Public Sub DesprotegerTodo()
    Dim sh As Worksheet
    If Not modSeguridad.RequierePermiso("CONFIGURACION") Then Exit Sub
    For Each sh In ThisWorkbook.Worksheets
        DesprotegerHoja sh
    Next sh
    MsgBox "Hojas desprotegidas para mantenimiento. Al volver a abrir el archivo se protegen de nuevo.", _
           vbExclamation, APP_NOMBRE
End Sub

'------------------------------------------------------------------ bitacora
Public Function Equipo() As String
    Equipo = Environ$("COMPUTERNAME")
    If Equipo = "" Then Equipo = "DESCONOCIDO"
End Function

Public Sub Bitacora(ByVal modulo As String, ByVal accion As String, ByVal documento As String, _
                    ByVal detalle As String)
    Dim sh As Worksheet, r As Long, ev As Boolean
    On Error GoTo Fin
    ev = Application.EnableEvents
    Application.EnableEvents = False
    Set sh = WS("BD_BITACORA")
    r = SiguienteFila(sh)
    sh.Range(sh.Cells(r, 1), sh.Cells(r, 8)).Value = Array(SiguienteID(sh, 1, "LOG", 6), Now, _
        modSeguridad.UsuarioActual(), modulo, accion, documento, detalle, Equipo())
    sh.Cells(r, 2).NumberFormat = "dd/mm/yyyy hh:mm"
    sh.Range(sh.Cells(r, 1), sh.Cells(r, 8)).Font.Size = 9
Fin:
    Application.EnableEvents = ev
End Sub

'------------------------------------------------------------------ utilidades varias
Public Function ClaveUnica() As String
    Randomize
    ClaveUnica = Format$(Now, "yyyymmddhhnnss") & "-" & Format$(Int(Rnd() * 1000000), "000000") & "-" & _
                 modSeguridad.UsuarioActual()
End Function

Public Function NumV(ByVal v As Variant) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then NumV = CDbl(v) Else NumV = 0
End Function

Public Function Vacio(ByVal v As Variant) As Boolean
    Vacio = (Trim$(CStr(v)) = "")
End Function

' Devuelve los valores no vacios de una columna del catalogo BASE DE DATOS
Public Function ListaCatalogo(ByVal nombreLista As String) As Variant
    Dim sh As Worksheet, c As Long, r As Long, col As New Collection, arr() As String, i As Long
    Set sh = WS("BASE DE DATOS")
    For c = 2 To 30
        If Clave(CStr(sh.Cells(5, c).Value)) = Clave(nombreLista) Then Exit For
    Next c
    If c > 30 Then ListaCatalogo = Array(): Exit Function
    r = 6
    Do While Trim$(CStr(sh.Cells(r, c).Value)) <> ""
        col.Add CStr(sh.Cells(r, c).Value)
        r = r + 1
    Loop
    If col.Count = 0 Then ListaCatalogo = Array(): Exit Function
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count
        arr(i) = col(i)
    Next i
    ListaCatalogo = arr
End Function

' Pide elegir un valor de una lista numerada. Devuelve "" si cancela.
Public Function ElegirDeLista(ByVal titulo As String, ByVal opciones As Variant, _
                              Optional ByVal porDefecto As String = "") As String
    Dim s As String, i As Long, resp As String, def As String
    If Not IsArray(opciones) Then Exit Function
    For i = LBound(opciones) To UBound(opciones)
        s = s & i & ". " & opciones(i) & vbCrLf
        If Clave(CStr(opciones(i))) = Clave(porDefecto) Then def = CStr(i)
    Next i
    resp = InputBox(titulo & vbCrLf & vbCrLf & s & vbCrLf & "Escriba el número:", APP_NOMBRE, def)
    If resp = "" Then Exit Function
    If Not IsNumeric(resp) Then Exit Function
    i = CLng(resp)
    If i < LBound(opciones) Or i > UBound(opciones) Then
        MsgBox "Opción no válida.", vbExclamation, APP_NOMBRE
        Exit Function
    End If
    ElegirDeLista = CStr(opciones(i))
End Function

Public Function PedirTexto(ByVal pregunta As String, Optional ByVal porDefecto As String = "", _
                           Optional ByVal obligatorio As Boolean = True) As String
    Dim s As String
    s = Trim$(InputBox(pregunta, APP_NOMBRE, porDefecto))
    If s = "" And obligatorio Then
        PedirTexto = ""
    Else
        PedirTexto = s
    End If
End Function

Public Function PedirNumero(ByVal pregunta As String, Optional ByVal porDefecto As Variant = "", _
                            Optional ByVal minimo As Double = 0) As Variant
    Dim s As String
    s = Trim$(InputBox(pregunta, APP_NOMBRE, CStr(porDefecto)))
    If s = "" Then PedirNumero = Empty: Exit Function
    s = Replace(s, "$", "")
    s = Replace(s, " ", "")
    If Not IsNumeric(s) Then
        MsgBox "Debe escribir un número.", vbExclamation, APP_NOMBRE
        PedirNumero = Empty
        Exit Function
    End If
    If CDbl(s) < minimo Then
        MsgBox "El valor debe ser mayor o igual a " & minimo & ".", vbExclamation, APP_NOMBRE
        PedirNumero = Empty
        Exit Function
    End If
    PedirNumero = CDbl(s)
End Function

Public Sub IrA(ByVal nombreHoja As String)
    On Error Resume Next
    WS(nombreHoja).Activate
    ActiveWindow.ScrollRow = 1
End Sub

Public Sub IrMenu()
    IrA "MENU"
End Sub

' Guarda el libro (el evento BeforeSave se encarga de ocultar hojas en el archivo)
Public Sub GuardarLibro()
    On Error Resume Next
    ThisWorkbook.Save
End Sub

Public Sub ErrorGeneral(ByVal origen As String)
    MsgBox "Ocurrió un error en " & origen & ":" & vbCrLf & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "La operación se deshizo; no quedaron registros a medias.", vbCritical, APP_NOMBRE
End Sub
