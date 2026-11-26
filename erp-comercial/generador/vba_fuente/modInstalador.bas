Attribute VB_Name = "modInstalador"
'==============================================================================
' modInstalador - Convierte ERP_Comercial (guardado como .xlsm) en el sistema
' completo: importa todos los modulos .bas de la carpeta "vba", inserta el
' codigo de ThisWorkbook, crea los formularios de inicio de sesion y dibuja
' los botones.
'
' Requisito (una sola vez): Archivo > Opciones > Centro de confianza >
' Configuracion > Configuracion de macros > marcar "Confiar en el acceso al
' modelo de objetos de proyectos de VBA".
'
' Uso: importe ESTE archivo (Alt+F11 > Archivo > Importar) y ejecute InstalarERP.
'==============================================================================
Option Explicit

Public Sub InstalarERP()
    Dim proj As Object, carpeta As String, f As String, n As Long, nombreMod As String, comp As Object
    Dim lista As String

    If LCase$(Right$(ThisWorkbook.Name, 5)) <> ".xlsm" Then
        MsgBox "Primero guarde el archivo como «Libro de Excel habilitado para macros (*.xlsm)» y vuelva a ejecutar.", _
               vbExclamation, "Instalador ERP"
        Exit Sub
    End If
    On Error Resume Next
    Set proj = ThisWorkbook.VBProject
    n = proj.VBComponents.Count
    On Error GoTo 0
    If n = 0 Then
        MsgBox "Excel no permite que el instalador modifique el proyecto VBA." & vbCrLf & vbCrLf & _
               "Active: Archivo > Opciones > Centro de confianza > Configuración del Centro de confianza > " & _
               "Configuración de macros > «Confiar en el acceso al modelo de objetos de proyectos de VBA». " & _
               "Luego ejecute InstalarERP de nuevo (puede desactivarlo al terminar).", vbExclamation, "Instalador ERP"
        Exit Sub
    End If

    carpeta = ThisWorkbook.Path & "\vba"
    If Dir(carpeta & "\modGeneral.bas") = "" Then
        With Application.FileDialog(4)          ' msoFileDialogFolderPicker
            .Title = "Seleccione la carpeta 'vba' con los módulos del ERP"
            If .Show <> -1 Then Exit Sub
            carpeta = .SelectedItems(1)
        End With
    End If

    ' 1. importar modulos
    f = Dir(carpeta & "\*.bas")
    Do While f <> ""
        nombreMod = Left$(f, Len(f) - 4)
        If LCase$(nombreMod) <> "modinstalador" Then
            On Error Resume Next
            Set comp = Nothing
            Set comp = proj.VBComponents(nombreMod)
            If Not comp Is Nothing Then
                comp.Name = nombreMod & "_viejo" & Format$(Timer * 100, "0")   ' Remove es diferido: liberar el nombre
                proj.VBComponents.Remove comp
            End If
            On Error GoTo 0
            proj.VBComponents.Import carpeta & "\" & f
            lista = lista & " " & nombreMod
        End If
        f = Dir()
    Loop

    ' 2. codigo de ThisWorkbook (su nombre cambia segun el idioma: ThisWorkbook / EsteLibro)
    InsertarCodigo proj.VBComponents(ThisWorkbook.CodeName), carpeta & "\ThisWorkbook.txt"

    ' 3. formularios con clave oculta
    CrearFormLogin proj
    CrearFormClave proj

    ' 4. botones y proteccion
    Application.Run "CrearBotones"
    Application.Run "ProtegerTodo"

    MsgBox "Instalación terminada." & vbCrLf & vbCrLf & "Módulos:" & lista & vbCrLf & vbCrLf & _
           "1. Guarde (Ctrl+G) y cierre el archivo." & vbCrLf & "2. Ábralo de nuevo y habilite el contenido." & vbCrLf & _
           "3. Ingrese con usuario admin y clave Admin2026 (se le pedirá cambiarla).", vbInformation, "Instalador ERP"
End Sub

Private Sub InsertarCodigo(ByVal comp As Object, ByVal archivo As String)
    Dim cm As Object, txt As String, linea As String, ff As Integer
    If Dir(archivo) = "" Then
        MsgBox "No se encontró " & archivo, vbExclamation, "Instalador ERP"
        Exit Sub
    End If
    ff = FreeFile
    Open archivo For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, linea
        txt = txt & linea & vbCrLf
    Loop
    Close #ff
    Set cm = comp.CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString txt
End Sub

Private Function NuevoForm(ByVal proj As Object, ByVal nombre As String, ByVal titulo As String, _
                           ByVal ancho As Double, ByVal alto As Double) As Object
    Dim comp As Object
    On Error Resume Next
    Set comp = proj.VBComponents(nombre)
    If Not comp Is Nothing Then
        comp.Name = nombre & "_viejo" & Format$(Timer * 100, "0")
        proj.VBComponents.Remove comp
    End If
    Set comp = Nothing
    On Error GoTo 0
    Set comp = proj.VBComponents.Add(3)          ' vbext_ct_MSForm
    comp.Name = nombre
    comp.Properties("Caption") = titulo
    comp.Properties("Width") = ancho
    comp.Properties("Height") = alto
    Set NuevoForm = comp
End Function

Private Function Ctl(ByVal frm As Object, ByVal tipo As String, ByVal nombre As String, ByVal x As Double, _
                     ByVal y As Double, ByVal w As Double, ByVal h As Double) As Object
    Set Ctl = frm.Designer.Controls.Add("Forms." & tipo & ".1", nombre)
    Ctl.Left = x
    Ctl.Top = y
    Ctl.Width = w
    Ctl.Height = h
    Ctl.Font.Name = "Arial"
    Ctl.Font.Size = 10
End Function

Private Function CodigoBotones() As String
    CodigoBotones = "Public Aceptado As Boolean" & vbCrLf & _
        "Private Sub btnOK_Click()" & vbCrLf & "    Aceptado = True" & vbCrLf & "    Me.Hide" & vbCrLf & "End Sub" & vbCrLf & _
        "Private Sub btnCancel_Click()" & vbCrLf & "    Aceptado = False" & vbCrLf & "    Me.Hide" & vbCrLf & "End Sub" & vbCrLf & _
        "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf & _
        "    If CloseMode = 0 Then" & vbCrLf & "        Cancel = True" & vbCrLf & "        Aceptado = False" & vbCrLf & _
        "        Me.Hide" & vbCrLf & "    End If" & vbCrLf & "End Sub" & vbCrLf
End Function

Private Sub CrearFormLogin(ByVal proj As Object)
    Dim frm As Object, c As Object
    Set frm = NuevoForm(proj, "frmLogin", "Iniciar sesión", 270, 165)
    Set c = Ctl(frm, "Label", "lblU", 12, 16, 70, 16): c.Caption = "Usuario"
    Set c = Ctl(frm, "TextBox", "txtUsuario", 85, 12, 160, 20)
    Set c = Ctl(frm, "Label", "lblC", 12, 46, 70, 16): c.Caption = "Clave"
    Set c = Ctl(frm, "TextBox", "txtClave", 85, 42, 160, 20): c.PasswordChar = "*"
    Set c = Ctl(frm, "CommandButton", "btnOK", 85, 80, 75, 26): c.Caption = "Entrar": c.Default = True
    Set c = Ctl(frm, "CommandButton", "btnCancel", 170, 80, 75, 26): c.Caption = "Cancelar": c.Cancel = True
    frm.CodeModule.AddFromString CodigoBotones() & _
        "Private Sub UserForm_Activate()" & vbCrLf & "    If Me.txtUsuario.Value = """" Then Me.txtUsuario.SetFocus Else Me.txtClave.SetFocus" & _
        vbCrLf & "End Sub" & vbCrLf
End Sub

Private Sub CrearFormClave(ByVal proj As Object)
    Dim frm As Object, c As Object
    Set frm = NuevoForm(proj, "frmClave", "Clave", 300, 140)
    Set c = Ctl(frm, "Label", "lblMensaje", 12, 10, 270, 30): c.Caption = "Escriba la clave": c.WordWrap = True
    Set c = Ctl(frm, "TextBox", "txtClave", 12, 44, 270, 20): c.PasswordChar = "*"
    Set c = Ctl(frm, "CommandButton", "btnOK", 120, 74, 75, 26): c.Caption = "Aceptar": c.Default = True
    Set c = Ctl(frm, "CommandButton", "btnCancel", 207, 74, 75, 26): c.Caption = "Cancelar": c.Cancel = True
    frm.CodeModule.AddFromString CodigoBotones() & _
        "Private Sub UserForm_Activate()" & vbCrLf & "    Me.txtClave.SetFocus" & vbCrLf & "End Sub" & vbCrLf
End Sub
