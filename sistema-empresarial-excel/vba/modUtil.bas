Attribute VB_Name = "modUtil"
'==============================================================
' modUtil - FUNCIONES COMUNES DEL SISTEMA
'  * Acceso a hojas y configuración
'  * Mensajes y preguntas (con modo de prueba automatizado)
'  * Transacciones: si una operación falla a mitad de camino,
'    se deshacen todas las filas y cambios que alcanzó a escribir
'  * Búsquedas, consecutivos, carpetas, PDF y auditoría
'==============================================================
Option Explicit

Public gTest As Boolean            ' True = modo prueba (sin cuadros de diálogo)
Public gRespuestas As Collection   ' respuestas en cola para Pedir() en modo prueba
Public gNoRespuestas As Long       ' cuántas preguntas Sí/No deben responderse NO en modo prueba
Public gLog As String              ' mensajes mostrados durante el modo prueba

' ---- transacción en curso
Private txHoja() As String, txCol() As Long, txFila() As Long, txCol0() As Long, txNCol() As Long, txN As Long
Private txVHoja() As String, txVAddr() As String, txVOld() As Variant, txVFormula() As Boolean, txVN As Long
Private txActiva As Boolean

'-------------------------------------------------------------- hojas
Public Function Hoja(ByVal nombre As String) As Worksheet
    Set Hoja = ThisWorkbook.Worksheets(nombre)
End Function

Public Function CfgV(ByVal direccion As String) As Variant
    CfgV = Hoja(SH_CFG).Range(direccion).Value
End Function

Public Sub CfgSet(ByVal direccion As String, ByVal valor As Variant)
    TxValor Hoja(SH_CFG), direccion
    Hoja(SH_CFG).Range(direccion).Value = valor
End Sub

Public Function UsuarioActual() As String
    UsuarioActual = Tv(Hoja(SH_INI).Range(FINI_USUARIO).Value)
    If UsuarioActual = "" Then UsuarioActual = "Administrador"
End Function

'-------------------------------------------------------------- conversiones seguras
Public Function Tv(ByVal v As Variant) As String
    If IsError(v) Then
        Tv = ""
    ElseIf IsEmpty(v) Or IsNull(v) Then
        Tv = ""
    Else
        Tv = Trim(CStr(v))
    End If
End Function

Public Function Nv(ByVal v As Variant) As Double
    If IsError(v) Then
        Nv = 0
    ElseIf IsEmpty(v) Then
        Nv = 0
    ElseIf VarType(v) = vbString Then
        If Trim(v) = "" Then
            Nv = 0
        ElseIf IsNumeric(v) Then
            Nv = CDbl(v)
        Else
            Nv = 0
        End If
    ElseIf IsNumeric(v) Or IsDate(v) Then
        Nv = CDbl(v)
    Else
        Nv = 0
    End If
End Function

' Convierte texto escrito por el usuario (ej. "150.000", "$ 25.500") a número.
Public Function NumTexto(ByVal s As String) As Double
    Dim i As Long, c As String, limpio As String
    s = Trim(s)
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c Like "#" Or c = "-" Or c = "," Then limpio = limpio & c
    Next i
    limpio = Replace(limpio, ",", ".")
    If limpio = "" Or limpio = "-" Then
        NumTexto = 0
    Else
        NumTexto = Val(limpio)
    End If
End Function

' Interpreta dd/mm/aaaa. Devuelve 0 si no es una fecha válida.
Public Function FechaTexto(ByVal s As String) As Date
    Dim p() As String
    s = Replace(Replace(Trim(s), "-", "/"), ".", "/")
    p = Split(s, "/")
    If UBound(p) = 2 Then
        If IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2)) Then
            If CLng(p(2)) < 100 Then p(2) = CStr(2000 + CLng(p(2)))
            If CLng(p(1)) >= 1 And CLng(p(1)) <= 12 And CLng(p(0)) >= 1 And CLng(p(0)) <= 31 Then
                FechaTexto = DateSerial(CLng(p(2)), CLng(p(1)), CLng(p(0)))
            End If
        End If
    End If
End Function

Public Function Moneda(ByVal v As Double) As String
    Moneda = "$ " & Format(v, "#,##0")
End Function

Public Function FechaCorta(ByVal d As Variant) As String
    If IsDate(d) Or IsNumeric(d) Then
        FechaCorta = Format(CDate(d), "dd/mm/yyyy")
    Else
        FechaCorta = Tv(d)
    End If
End Function

'-------------------------------------------------------------- diálogos
Public Function Msg(ByVal texto As String, Optional ByVal botones As Long = 64, _
                    Optional ByVal titulo As String = "Sistema Empresarial") As Long
    If gTest Then
        gLog = gLog & "[MSG] " & Replace(texto, vbLf, " / ") & vbLf
        If (botones And vbYesNo) = vbYesNo Or (botones And vbYesNoCancel) = vbYesNoCancel Then
            If gNoRespuestas > 0 Then
                gNoRespuestas = gNoRespuestas - 1
                Msg = vbNo
            Else
                Msg = vbYes
            End If
        Else
            Msg = vbOK
        End If
    Else
        Msg = MsgBox(texto, botones, titulo)
    End If
End Function

Public Function Pedir(ByVal texto As String, Optional ByVal titulo As String = "Sistema Empresarial", _
                      Optional ByVal defecto As String = "") As String
    If gTest Then
        If gRespuestas Is Nothing Then
            Pedir = defecto
        ElseIf gRespuestas.Count = 0 Then
            Pedir = defecto
        Else
            Pedir = CStr(gRespuestas.Item(1))
            gRespuestas.Remove 1
        End If
        gLog = gLog & "[PREGUNTA] " & Left(Replace(texto, vbLf, " / "), 80) & " => " & Pedir & vbLf
    Else
        Pedir = InputBox(texto, titulo, defecto)
    End If
End Function

' Solicita la clave de autorización configurada. Devuelve True si es correcta.
Public Function Autorizar(ByVal accion As String) As Boolean
    Dim clave As String
    clave = Pedir("Esta acción requiere autorización:" & vbLf & accion & vbLf & vbLf & _
                  "Escriba la clave de autorización:", "Autorización requerida")
    If clave = "" Then
        Autorizar = False
    ElseIf clave = Tv(CfgV(CFG_CLAVE)) Then
        Autorizar = True
        Auditar "SEGURIDAD", "AUTORIZACIÓN CONCEDIDA", "", accion
    Else
        Msg "Clave incorrecta. La acción fue cancelada.", vbCritical
        Auditar "SEGURIDAD", "AUTORIZACIÓN RECHAZADA", "", accion
        Autorizar = False
    End If
End Function

'-------------------------------------------------------------- entorno de ejecución
Public Sub Iniciar()
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculate
    Application.Calculation = xlCalculationManual
    DesprotegerTodo
    On Error GoTo 0
End Sub

Public Sub Terminar()
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    ProtegerTodo
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Public Sub Recalcular()
    On Error Resume Next
    Application.Calculate
    On Error GoTo 0
End Sub

Public Sub DesprotegerTodo()
    Dim ws As Worksheet, clave As String
    On Error Resume Next
    clave = Tv(ThisWorkbook.Worksheets(SH_CFG).Range(CFG_CLAVEHOJAS).Value)
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect clave
    Next ws
    On Error GoTo 0
End Sub

Public Sub ProtegerTodo()
    Dim ws As Worksheet, clave As String
    If gTest Then Exit Sub
    On Error Resume Next
    clave = Tv(ThisWorkbook.Worksheets(SH_CFG).Range(CFG_CLAVEHOJAS).Value)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SH_CALC Then
            Err.Clear
            ws.Protect Password:=clave, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                       UserInterfaceOnly:=True, AllowFiltering:=True, AllowFormattingColumns:=True, _
                       AllowFormattingRows:=True
            If Err.Number <> 0 Then Err.Clear: ws.Protect clave
        End If
    Next ws
    On Error GoTo 0
End Sub

'-------------------------------------------------------------- filas y búsquedas
Public Function UltimaFila(ByVal ws As Worksheet, ByVal col As Long, ByVal primera As Long) As Long
    ' Última fila con dato en la columna (ignora celdas vacías o con texto vacío)
    Dim tope As Long, arr As Variant, i As Long
    tope = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    UltimaFila = primera - 1
    If tope < primera Then Exit Function
    If tope = primera Then
        If Tv(ws.Cells(primera, col).Value) <> "" Then UltimaFila = primera
        Exit Function
    End If
    arr = ws.Range(ws.Cells(primera, col), ws.Cells(tope, col)).Value
    For i = UBound(arr, 1) To 1 Step -1
        If Tv(arr(i, 1)) <> "" Then
            UltimaFila = primera + i - 1
            Exit Function
        End If
    Next i
End Function

Public Function NuevaFila(ByVal ws As Worksheet, ByVal col As Long, ByVal primera As Long, ByVal ultimaCap As Long) As Long
    Dim r As Long
    r = UltimaFila(ws, col, primera) + 1
    If r > ultimaCap Then
        Err.Raise vbObjectError + 600, "NuevaFila", "Se alcanzó la capacidad máxima de la tabla en '" & ws.Name & _
                  "'. Amplíe la tabla copiando la última fila hacia abajo (ver manual)."
    End If
    NuevaFila = r
End Function

' Devuelve la primera fila cuyo valor (texto, sin mayúsculas/minúsculas) coincide; 0 si no existe.
Public Function BuscarFila(ByVal ws As Worksheet, ByVal col As Long, ByVal primera As Long, ByVal valor As String) As Long
    Dim ult As Long, arr As Variant, i As Long
    valor = UCase(Trim(valor))
    If valor = "" Then Exit Function
    ult = UltimaFila(ws, col, primera)
    If ult < primera Then Exit Function
    If ult = primera Then
        If UCase(Tv(ws.Cells(primera, col).Value)) = valor Then BuscarFila = primera
        Exit Function
    End If
    arr = ws.Range(ws.Cells(primera, col), ws.Cells(ult, col)).Value
    For i = 1 To UBound(arr, 1)
        If UCase(Tv(arr(i, 1))) = valor Then
            BuscarFila = primera + i - 1
            Exit Function
        End If
    Next i
End Function

' Lee una columna completa de una tabla como arreglo 1..n (vacío si no hay datos)
Public Function LeerColumna(ByVal ws As Worksheet, ByVal col As Long, ByVal primera As Long, ByVal ult As Long) As Variant
    Dim arr As Variant, out() As Variant
    If ult < primera Then
        ReDim out(1 To 1, 1 To 1)
        LeerColumna = out
        Exit Function
    End If
    If ult = primera Then
        ReDim out(1 To 1, 1 To 1)
        out(1, 1) = ws.Cells(primera, col).Value
        LeerColumna = out
    Else
        LeerColumna = ws.Range(ws.Cells(primera, col), ws.Cells(ult, col)).Value
    End If
End Function

' Siguiente identificador tipo PREFIJO-00001
Public Function SiguienteID(ByVal ws As Worksheet, ByVal col As Long, ByVal primera As Long, _
                            ByVal prefijo As String, ByVal digitos As Long) As String
    Dim ult As Long, arr As Variant, i As Long, mx As Long, s As String, n As Long
    ult = UltimaFila(ws, col, primera)
    If ult >= primera Then
        arr = LeerColumna(ws, col, primera, ult)
        For i = 1 To UBound(arr, 1)
            s = Tv(arr(i, 1))
            If UCase(Left(s, Len(prefijo))) = UCase(prefijo) Then
                n = CLng(Val(Mid(s, Len(prefijo) + 1)))
                If n > mx Then mx = n
            End If
        Next i
    End If
    SiguienteID = prefijo & Format(mx + 1, String(digitos, "0"))
End Function

Public Function NuevaFilaArr(ByVal ncols As Long) As Variant
    Dim a() As Variant
    ReDim a(1 To 1, 1 To ncols)
    NuevaFilaArr = a
End Function

Public Sub EscribirFila(ByVal ws As Worksheet, ByVal fila As Long, ByVal col0 As Long, ByVal arr As Variant)
    ws.Range(ws.Cells(fila, col0), ws.Cells(fila, col0 + UBound(arr, 2) - 1)).Value = arr
End Sub

' Escribe solo las posiciones no vacías del arreglo (para tablas con columnas de fórmula)
Public Sub EscribirValores(ByVal ws As Worksheet, ByVal fila As Long, ByVal col0 As Long, ByVal arr As Variant)
    Dim i As Long
    For i = 1 To UBound(arr, 2)
        If Not IsEmpty(arr(1, i)) Then ws.Cells(fila, col0 + i - 1).Value = arr(1, i)
    Next i
End Sub

' Cambia una celda existente dejando registro para poder revertir
Public Sub Modificar(ByVal ws As Worksheet, ByVal fila As Long, ByVal col As Long, ByVal valor As Variant)
    TxValor ws, ws.Cells(fila, col).Address(False, False)
    ws.Cells(fila, col).Value = valor
End Sub

' Borra el contenido (no las fórmulas) de las filas desde 'desde' hasta la última usada
Public Sub LimpiarFilasDesde(ByVal ws As Worksheet, ByVal keyCol As Long, ByVal desde As Long, _
                             ByVal col0 As Long, ByVal ncols As Long)
    Dim ult As Long, r As Long, c As Long
    ult = UltimaFila(ws, keyCol, desde)
    For r = desde To ult
        For c = col0 To col0 + ncols - 1
            If Left(CStr(ws.Cells(r, c).Formula), 1) <> "=" Then ws.Cells(r, c).ClearContents
        Next c
    Next r
End Sub

'-------------------------------------------------------------- transacciones
Public Sub TxIniciar()
    txN = 0: txVN = 0: txActiva = True
    ReDim txHoja(1 To 20): ReDim txCol(1 To 20): ReDim txFila(1 To 20): ReDim txCol0(1 To 20): ReDim txNCol(1 To 20)
    ReDim txVHoja(1 To 50): ReDim txVAddr(1 To 50): ReDim txVOld(1 To 50): ReDim txVFormula(1 To 50)
End Sub

Public Sub TxTabla(ByVal hojaNombre As String, ByVal keyCol As Long, ByVal primera As Long, _
                   ByVal col0 As Long, ByVal ncols As Long)
    If Not txActiva Then Exit Sub
    txN = txN + 1
    txHoja(txN) = hojaNombre: txCol(txN) = keyCol: txCol0(txN) = col0: txNCol(txN) = ncols
    txFila(txN) = UltimaFila(Hoja(hojaNombre), keyCol, primera) + 1
End Sub

Public Sub TxValor(ByVal ws As Worksheet, ByVal direccion As String)
    If Not txActiva Then Exit Sub
    txVN = txVN + 1
    If txVN > UBound(txVHoja) Then
        ReDim Preserve txVHoja(1 To txVN + 50): ReDim Preserve txVAddr(1 To txVN + 50)
        ReDim Preserve txVOld(1 To txVN + 50): ReDim Preserve txVFormula(1 To txVN + 50)
    End If
    txVHoja(txVN) = ws.Name
    txVAddr(txVN) = direccion
    If Left(CStr(ws.Range(direccion).Formula), 1) = "=" Then
        txVFormula(txVN) = True
        txVOld(txVN) = ws.Range(direccion).Formula
    Else
        txVFormula(txVN) = False
        txVOld(txVN) = ws.Range(direccion).Value
    End If
End Sub

Public Sub TxConfirmar()
    txActiva = False: txN = 0: txVN = 0
End Sub

Public Sub TxRevertir()
    Dim i As Long, ws As Worksheet
    On Error Resume Next
    If Not txActiva Then Exit Sub
    For i = txVN To 1 Step -1
        Set ws = Hoja(txVHoja(i))
        If txVFormula(i) Then
            ws.Range(txVAddr(i)).Formula = txVOld(i)
        Else
            ws.Range(txVAddr(i)).Value = txVOld(i)
        End If
    Next i
    For i = 1 To txN
        LimpiarFilasDesde Hoja(txHoja(i)), txCol(i), txFila(i), txCol0(i), txNCol(i)
    Next i
    txActiva = False: txN = 0: txVN = 0
    On Error GoTo 0
End Sub

'-------------------------------------------------------------- inventario (stock real)
' Stock calculado directamente desde los datos (no depende del recálculo de fórmulas)
Public Function StockReal(ByVal sku As String) As Double
    Dim wsI As Worksheet, wsM As Worksheet, fi As Long, ult As Long, i As Long
    Dim arrSku As Variant, arrEnt As Variant, arrSal As Variant, total As Double
    Set wsI = Hoja(SH_INV)
    fi = BuscarFila(wsI, INV_SKU, INV_FIRST, sku)
    If fi = 0 Then Exit Function
    total = Nv(wsI.Cells(fi, INV_STOCKINI).Value)
    Set wsM = Hoja(SH_MOV)
    ult = UltimaFila(wsM, MOV_ID, MOV_FIRST)
    If ult >= MOV_FIRST Then
        arrSku = LeerColumna(wsM, MOV_SKU, MOV_FIRST, ult)
        arrEnt = LeerColumna(wsM, MOV_ENTRADA, MOV_FIRST, ult)
        arrSal = LeerColumna(wsM, MOV_SALIDA, MOV_FIRST, ult)
        sku = UCase(sku)
        For i = 1 To UBound(arrSku, 1)
            If UCase(Tv(arrSku(i, 1))) = sku Then total = total + Nv(arrEnt(i, 1)) - Nv(arrSal(i, 1))
        Next i
    End If
    StockReal = total
End Function

Public Function CuentaDeMetodo(ByVal metodo As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = Hoja(SH_CFG)
    For r = CFG_LIST_FIRST To CFG_LIST_FIRST + CFG_LIST_N - 1
        If UCase(Tv(ws.Range(LST_METODOS & r).Value)) = UCase(Trim(metodo)) Then
            CuentaDeMetodo = Tv(ws.Range(LST_METODOCUENTA & r).Value)
            Exit For
        End If
    Next r
    If CuentaDeMetodo = "" Then CuentaDeMetodo = "OTROS"
End Function

Public Function EnLista(ByVal colLista As String, ByVal valor As String) As Boolean
    Dim ws As Worksheet, r As Long
    Set ws = Hoja(SH_CFG)
    For r = CFG_LIST_FIRST To CFG_LIST_FIRST + CFG_LIST_N - 1
        If UCase(Tv(ws.Range(colLista & r).Value)) = UCase(Trim(valor)) And Trim(valor) <> "" Then
            EnLista = True
            Exit Function
        End If
    Next r
End Function

'-------------------------------------------------------------- carpetas y PDF
Public Function CarpetaEscritorio() As String
    Dim p As String
    On Error Resume Next
    p = CreateObject("WScript.Shell").SpecialFolders("Desktop")   ' detecta OneDrive y cualquier usuario
    If p = "" Then p = Environ("USERPROFILE") & "\Desktop"
    If Dir(p, vbDirectory) = "" Then p = ThisWorkbook.Path
    On Error GoTo 0
    CarpetaEscritorio = p
End Function

' Devuelve (y crea si falta) ESCRITORIO\SISTEMA EMPRESARIAL\<sub>
Public Function Carpeta(Optional ByVal subTxt As String = "") As String
    Dim raiz As String, sep As String, s As Variant
    sep = Application.PathSeparator
    raiz = CarpetaEscritorio() & sep & "SISTEMA EMPRESARIAL"
    On Error Resume Next
    If Dir(raiz, vbDirectory) = "" Then MkDir raiz
    For Each s In Array("FACTURAS", "GUIAS", "REPORTES", "COPIAS DE SEGURIDAD")
        If Dir(raiz & sep & s, vbDirectory) = "" Then MkDir raiz & sep & s
    Next s
    On Error GoTo 0
    If subTxt = "" Then
        Carpeta = raiz
    Else
        Carpeta = raiz & sep & subTxt
    End If
End Function

Public Function NombreArchivo(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, c, "-")
    Next c
    NombreArchivo = Left(Trim(s), 120)
End Function

' Exporta el área de impresión de una hoja a PDF. Devuelve la ruta o "" si no fue posible.
Public Function ExportarPDF(ByVal ws As Worksheet, ByVal subcarpeta As String, ByVal nombre As String) As String
    Dim ruta As String
    If gTest Then Exit Function
    On Error GoTo Fallo
    ruta = Carpeta(subcarpeta) & Application.PathSeparator & NombreArchivo(nombre) & ".pdf"
    PrepararImpresion ws, True
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ruta, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    PrepararImpresion ws, False
    ExportarPDF = ruta
    Exit Function
Fallo:
    PrepararImpresion ws, False
    ExportarPDF = ""
End Function

' Quita (True) o restaura (False) el color amarillo de captura para imprimir un documento limpio
Public Sub PrepararImpresion(ByVal ws As Worksheet, ByVal imprimir As Boolean)
    Dim rangos As String, color As Long
    On Error Resume Next
    If ws.Name = SH_FAC Then
        rangos = "I3,C8:F8,C11:F11,B14:B28,F14:F28,H14:H28,C30:D32,A36:I37"
    ElseIf ws.Name = SH_GUI Then
        rangos = "I3:I5,G8:I14,C17:D18,G17:I18,B21:B32,F21:F32,A36:I37"
    Else
        Exit Sub
    End If
    If imprimir Then
        color = RGB(255, 255, 255)
    Else
        color = RGB(255, 248, 225)
    End If
    ws.Range(rangos).Interior.Color = color
    On Error GoTo 0
End Sub

'-------------------------------------------------------------- auditoría
Public Sub Auditar(ByVal modulo As String, ByVal accion As String, ByVal doc As String, ByVal detalle As String)
    Dim ws As Worksheet, r As Long, a As Variant
    On Error Resume Next
    Set ws = Hoja(SH_AUD)
    r = UltimaFila(ws, AUD_FECHA, AUD_FIRST) + 1
    a = NuevaFilaArr(AUD_NCOLS)
    a(1, AUD_FECHA) = Now
    a(1, AUD_USUARIO) = UsuarioActual()
    a(1, AUD_MODULO) = modulo
    a(1, AUD_ACCION) = accion
    a(1, AUD_DOC) = doc
    a(1, AUD_DETALLE) = detalle
    EscribirFila ws, r, 1, a
    On Error GoTo 0
End Sub

Public Sub QuitarFiltros(ByVal ws As Worksheet)
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo 0
End Sub
