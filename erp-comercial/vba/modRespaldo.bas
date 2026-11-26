Attribute VB_Name = "modRespaldo"
'==============================================================================
' modRespaldo - OneDrive (deteccion automatica de la ruta, sin rutas fijas),
' estructura de carpetas, exportacion a PDF, copias de seguridad programadas y
' recuperacion de copias.
'
' Estructura:  <OneDrive>\SISTEMA EMPRESARIAL\
'                 FACTURAS\AAAA\MM  GUIAS\AAAA\MM  COMPRAS\AAAA\MM  REPORTES
'                 CLIENTES  COPIAS DE SEGURIDAD  MANUALES
'==============================================================================
Option Explicit

Private mProximoRespaldo As Date

Private Function Subcarpetas() As Variant
    Subcarpetas = Array("FACTURAS", "GUIAS", "COMPRAS", "REPORTES", "CLIENTES", "COPIAS DE SEGURIDAD", "MANUALES")
End Function

Private Function ExisteCarpeta(ByVal ruta As String) As Boolean
    On Error Resume Next
    ExisteCarpeta = ((GetAttr(ruta) And vbDirectory) = vbDirectory)
End Function

Private Sub CrearCarpeta(ByVal ruta As String)
    Dim partes As Variant, i As Long, acum As String
    If ExisteCarpeta(ruta) Then Exit Sub
    partes = Split(ruta, "\")
    acum = partes(0)
    For i = 1 To UBound(partes)
        acum = acum & "\" & partes(i)
        If Not ExisteCarpeta(acum) Then MkDir acum
    Next i
End Sub

Private Function LeerRegistro(ByVal regKey As String) As String
    On Error Resume Next
    LeerRegistro = CreateObject("WScript.Shell").RegRead(regKey)
End Function

' Raices de OneDrive presentes en el equipo (empresa primero)
Public Function RaicesOneDrive() As Collection
    Dim c As New Collection, cand As Variant, v As Variant, i As Integer
    cand = Array(Environ$("OneDriveCommercial"), _
                 LeerRegistro("HKCU\Software\Microsoft\OneDrive\Accounts\Business1\UserFolder"), _
                 LeerRegistro("HKCU\Software\Microsoft\OneDrive\Accounts\Business2\UserFolder"), _
                 Environ$("OneDrive"), Environ$("OneDriveConsumer"), _
                 LeerRegistro("HKCU\Software\Microsoft\OneDrive\Accounts\Personal\UserFolder"))
    On Error Resume Next
    For Each v In cand
        If CStr(v) <> "" Then
            If ExisteCarpeta(CStr(v)) Then c.Add CStr(v), LCase$(CStr(v))   ' la clave evita repetidos
        End If
    Next v
    Set RaicesOneDrive = c
End Function

' Carpeta raiz del sistema. Orden: ruta guardada > carpeta existente en OneDrive > crearla > elegir manualmente
Public Function CarpetaSistema(Optional ByVal preguntar As Boolean = True) As String
    Dim guardada As String, raiz As Variant, nombre As String, raices As Collection, fd As FileDialog, s As Variant
    nombre = CStr(Cfg("cfg_CarpetaSistema"))
    guardada = CStr(Cfg("cfg_RutaOneDrive"))
    If guardada <> "" Then
        If ExisteCarpeta(guardada) Then CarpetaSistema = guardada: Exit Function
    End If
    Set raices = RaicesOneDrive()
    For Each raiz In raices
        If ExisteCarpeta(raiz & "\" & nombre) Then CarpetaSistema = raiz & "\" & nombre: GoTo Guardar
    Next raiz
    If raices.Count > 0 Then
        If Not preguntar Then CarpetaSistema = raices(1) & "\" & nombre: GoTo Crear
        If MsgBox("No se encontró la carpeta """ & nombre & """ en OneDrive." & vbCrLf & vbCrLf & _
                  "¿Crearla en " & raices(1) & "?" & vbCrLf & vbCrLf & _
                  "Si el administrador ya la compartió con usted, responda NO, abra OneDrive en el navegador, " & _
                  "use «Agregar acceso directo a Mis archivos» y luego DETECTAR ONEDRIVE.", _
                  vbQuestion + vbYesNo, APP_NOMBRE) = vbYes Then
            CarpetaSistema = raices(1) & "\" & nombre
            GoTo Crear
        End If
    End If
    If Not preguntar Then Exit Function
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Seleccione la carpeta " & nombre & " (o la carpeta donde crearla)"
    If fd.Show <> -1 Then Exit Function
    s = fd.SelectedItems(1)
    If UCase$(Right$(s, Len(nombre))) <> UCase$(nombre) Then s = s & "\" & nombre
    CarpetaSistema = s
Crear:
    CrearCarpeta CarpetaSistema
Guardar:
    Dim ev As Boolean
    ev = Application.EnableEvents
    Application.EnableEvents = False
    SetCfg "cfg_RutaOneDrive", CarpetaSistema
    Application.EnableEvents = ev
    CrearEstructura CarpetaSistema
End Function

Private Sub CrearEstructura(ByVal raiz As String)
    Dim s As Variant
    On Error Resume Next
    For Each s In Subcarpetas()
        CrearCarpeta raiz & "\" & s
    Next s
End Sub

Public Sub DetectarOneDrive()
    Dim ruta As String, ev As Boolean
    ev = Application.EnableEvents
    Application.EnableEvents = False
    SetCfg "cfg_RutaOneDrive", ""
    Application.EnableEvents = ev
    ruta = CarpetaSistema(True)
    If ruta <> "" Then
        MsgBox "Carpeta del sistema:" & vbCrLf & ruta & vbCrLf & vbCrLf & "Subcarpetas verificadas.", vbInformation, APP_NOMBRE
        Bitacora "CONFIGURACIÓN", "ONEDRIVE", "", ruta
    End If
End Sub

Public Sub CrearCarpetas()
    Dim ruta As String
    ruta = CarpetaSistema(True)
    If ruta <> "" Then CrearEstructura ruta: MsgBox "Estructura creada en:" & vbCrLf & ruta, vbInformation, APP_NOMBRE
End Sub

Public Function RutaRelativa(ByVal ruta As String) As String
    Dim raiz As String
    raiz = CStr(Cfg("cfg_RutaOneDrive"))
    If raiz <> "" And LCase$(Left$(ruta, Len(raiz))) = LCase$(raiz) Then
        RutaRelativa = Replace(Mid$(ruta, Len(raiz) + 2), "\", "/")
    Else
        RutaRelativa = ruta
    End If
End Function

'------------------------------------------------------------------ PDF
Public Function ExportarPDF(ByVal sh As Worksheet, ByVal tipo As String, ByVal nombreArchivo As String, _
                            ByVal fecha As Date) As String
    Dim raiz As String, carpeta As String, ruta As String
    On Error GoTo FALLO
    raiz = CarpetaSistema(True)
    If raiz = "" Then Exit Function
    carpeta = raiz & "\" & tipo & "\" & Format$(fecha, "yyyy") & "\" & Format$(fecha, "mm")
    CrearCarpeta carpeta
    ruta = carpeta & "\" & nombreArchivo & ".pdf"
    sh.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ruta, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    ExportarPDF = ruta
    Exit Function
FALLO:
    MsgBox "No se pudo generar el PDF: " & Err.Description, vbExclamation, APP_NOMBRE
End Function

'------------------------------------------------------------------ COPIAS DE SEGURIDAD
Public Sub RespaldoAhora()
    If Not RequierePermiso("RESPALDOS") Then Exit Sub
    CrearRespaldo False
End Sub

' Copia exacta del libro (SaveCopyAs no altera el archivo abierto)
Public Function CrearRespaldo(Optional ByVal silencioso As Boolean = False) As String
    Dim raiz As String, carpeta As String, ruta As String, ext As String
    On Error GoTo FALLO
    raiz = CarpetaSistema(Not silencioso)
    If raiz = "" Then Exit Function
    carpeta = raiz & "\COPIAS DE SEGURIDAD"
    CrearCarpeta carpeta
    ext = Mid$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "."))
    ruta = carpeta & "\ERP_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & UsuarioActual() & ext
    ThisWorkbook.SaveCopyAs ruta
    DepurarRespaldos carpeta, CLng(NumV(Cfg("cfg_CopiasConservar")))
    Bitacora "RESPALDOS", "COPIA", "", ruta
    CrearRespaldo = ruta
    If Not silencioso Then MsgBox "Copia de seguridad creada:" & vbCrLf & ruta, vbInformation, APP_NOMBRE
    Exit Function
FALLO:
    If Not silencioso Then MsgBox "No se pudo crear la copia: " & Err.Description, vbExclamation, APP_NOMBRE
End Function

' Conserva solo las N copias mas recientes (los nombres llevan fecha, asi que el orden alfabetico es cronologico)
Private Sub DepurarRespaldos(ByVal carpeta As String, ByVal conservar As Long)
    Dim f As String, nombres As New Collection, arr() As String, i As Long, j As Long, tmp As String
    If conservar < 5 Then conservar = 5
    f = Dir(carpeta & "\ERP_*.xls*")
    Do While f <> ""
        nombres.Add f
        f = Dir()
    Loop
    If nombres.Count <= conservar Then Exit Sub
    ReDim arr(1 To nombres.Count)
    For i = 1 To nombres.Count
        arr(i) = nombres(i)
    Next i
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
        Next j
    Next i
    On Error Resume Next
    For i = 1 To UBound(arr) - conservar
        Kill carpeta & "\" & arr(i)
    Next i
End Sub

Public Sub ProgramarRespaldo()
    Dim horas As Double
    On Error Resume Next
    horas = NumV(Cfg("cfg_HorasRespaldo"))
    If horas <= 0 Then Exit Sub
    mProximoRespaldo = Now + horas / 24
    Application.OnTime mProximoRespaldo, "EjecutarRespaldoProgramado"
End Sub

Public Sub EjecutarRespaldoProgramado()
    If SesionActiva() Then CrearRespaldo True
    ProgramarRespaldo
End Sub

Public Sub CancelarRespaldoProgramado()
    On Error Resume Next
    If mProximoRespaldo > 0 Then Application.OnTime mProximoRespaldo, "EjecutarRespaldoProgramado", , False
End Sub

' Recuperacion: abre una copia elegida para verificarla y, si se confirma, la deja como nuevo archivo de trabajo
Public Sub RestaurarCopia()
    Dim fd As FileDialog, origen As String, destino As String, raiz As String, wb As Workbook
    If Not RequierePermiso("RESPALDOS") Then Exit Sub
    raiz = CarpetaSistema(True)
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Elija la copia de seguridad a recuperar"
    fd.Filters.Clear
    fd.Filters.Add "Libros de Excel", "*.xlsm;*.xlsx"
    If raiz <> "" Then fd.InitialFileName = raiz & "\COPIAS DE SEGURIDAD\"
    If fd.Show <> -1 Then Exit Sub
    origen = fd.SelectedItems(1)
    If MsgBox("Se hará primero una copia del archivo actual y luego se abrirá:" & vbCrLf & origen & vbCrLf & vbCrLf & _
              "La copia recuperada se guardará junto al archivo actual con el sufijo _RECUPERADO para que la " & _
              "revise. Si está correcta, ciérrelos y renombre el recuperado como archivo principal." & vbCrLf & vbCrLf & _
              "¿Continuar?", vbQuestion + vbYesNo, APP_NOMBRE) = vbNo Then Exit Sub
    CrearRespaldo True
    destino = ThisWorkbook.Path & "\" & Left$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & _
              "_RECUPERADO_" & Format$(Now, "yyyymmdd_hhnn") & Mid$(origen, InStrRev(origen, "."))
    FileCopy origen, destino
    Bitacora "RESPALDOS", "RECUPERAR", "", origen & " -> " & destino
    GuardarLibro
    Set wb = Workbooks.Open(destino)
    MsgBox "Copia recuperada abierta: " & destino, vbInformation, APP_NOMBRE
End Sub
