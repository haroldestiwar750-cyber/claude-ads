Attribute VB_Name = "modSeguridad"
'==============================================================================
' modSeguridad - Inicio de sesion, roles, permisos, visibilidad de hojas,
' control de edicion y ciclo de apertura / guardado / cierre del libro.
'
' IMPORTANTE: es una capa de control operativo. La proteccion de hojas y el
' VBA de Excel NO son seguridad fuerte; la seguridad real llega con la base de
' datos central (usuarios y permisos en el servidor).
'==============================================================================
Option Explicit

Private mGuardando As Boolean
Private mSesion As Boolean
Private mIntentos As Integer

'------------------------------------------------------------------ hash de clave
' Mismo algoritmo que hash_clave() en generador/datos_prueba.py
Public Function HashClave(ByVal usuario As String, ByVal claveTexto As String) As String
    Dim s As String, h As Double, i As Long, r As Long
    s = LCase$(usuario) & "|" & claveTexto & "|ERP"
    h = 5381
    For r = 1 To 500
        For i = 1 To Len(s)
            h = h * 33 + AscW(Mid$(s, i, 1)) + r
            h = h - Int(h / 2147483647#) * 2147483647#
        Next i
    Next r
    HashClave = Format$(h, "0")
End Function

'------------------------------------------------------------------ sesion actual
Public Function UsuarioActual() As String
    On Error Resume Next
    UsuarioActual = CStr(Cfg("app_Usuario"))
    If UsuarioActual = "" Then UsuarioActual = "(sin sesión)"
End Function

Public Function RolActual() As String
    On Error Resume Next
    RolActual = CStr(Cfg("app_Rol"))
End Function

Public Function VendedorActual() As String
    On Error Resume Next
    VendedorActual = CStr(Cfg("app_Vendedor"))
End Function

Public Function SesionActiva() As Boolean
    SesionActiva = mSesion And RolActual() <> ""
End Function

Public Function EsAdmin() As Boolean
    EsAdmin = (Clave(RolActual()) = "ADMINISTRADOR")
End Function

'------------------------------------------------------------------ permisos
Public Function TienePermiso(ByVal codigo As String, Optional ByVal rol As String = "") As Boolean
    Dim sh As Worksheet, r As Long, c As Long
    If rol = "" Then rol = RolActual()
    If Clave(rol) = "ADMINISTRADOR" Then TienePermiso = True: Exit Function
    If rol = "" Then Exit Function
    Set sh = WS("USUARIOS Y PERMISOS")
    For c = 12 To 20
        If Clave(CStr(sh.Cells(5, c).Value)) = Clave(rol) Then Exit For
    Next c
    If c > 20 Then Exit Function
    r = BuscarFila(sh, 10, codigo)
    If r = 0 Then Exit Function
    TienePermiso = (Clave(CStr(sh.Cells(r, c).Value)) = "SI")
End Function

Public Function RequierePermiso(ByVal codigo As String) As Boolean
    If Not SesionActiva() Then
        MsgBox "Debe iniciar sesión.", vbExclamation, APP_NOMBRE
        Exit Function
    End If
    RequierePermiso = TienePermiso(codigo)
    If Not RequierePermiso Then
        MsgBox "Su rol (" & RolActual() & ") no tiene el permiso " & codigo & "." & vbCrLf & _
               "Solicítelo al administrador.", vbExclamation, APP_NOMBRE
    End If
End Function

' Permiso necesario para VER cada hoja ("" = todos los usuarios con sesion)
Public Function PermisoVer(ByVal nombreHoja As String) As String
    Select Case Clave(nombreHoja)
        Case "MENU", "INVENTARIO", "CLIENTES", "VENDEDORES": PermisoVer = ""
        Case "DASHBOARD", "CANALES DE VENTA": PermisoVer = "VER_DASHBOARD"
        Case "FACTURACION": PermisoVer = "FACTURAR"
        Case "GUIAS", "BD_GUIAS": PermisoVer = "GUIAS"
        Case "COMPRAS": PermisoVer = "COMPRAS"
        Case "PROVEEDORES": PermisoVer = "PROVEEDORES"
        Case "COSTOS Y RENTABILIDAD": PermisoVer = "VER_COSTOS"
        Case "INGRESOS", "EGRESOS", "GASTOS FIJOS", "CAJA Y BANCOS", "UTILIDADES", "BD_CIERRES": PermisoVer = "FINANZAS"
        Case "CARTERA", "BD_ABONOS": PermisoVer = "CARTERA"
        Case "COMISIONES": PermisoVer = "COMISIONES"
        Case "REPORTES": PermisoVer = "REPORTES"
        Case "HISTORIAL MOVIMIENTOS": PermisoVer = "INVENTARIO"
        Case "CONFIGURACION", "BASE DE DATOS": PermisoVer = "CONFIGURACION"
        Case "USUARIOS Y PERMISOS", "BD_BITACORA", "BD_FACTURAS", "BD_FACT_DETALLE", "BD_COMPRAS", _
             "BD_COMPRA_DETALLE": PermisoVer = "USUARIOS"
        Case Else: PermisoVer = "USUARIOS"
    End Select
End Function

' Permiso necesario para EDITAR celdas desbloqueadas de cada hoja
Public Function PermisoEditar(ByVal nombreHoja As String) As String
    Select Case Clave(nombreHoja)
        Case "DASHBOARD", "CANALES DE VENTA", "CAJA Y BANCOS", "UTILIDADES", "REPORTES", "COMISIONES", _
             "FACTURACION", "GUIAS", "COMPRAS", "COSTOS Y RENTABILIDAD", "CARTERA"
            PermisoEditar = PermisoVer(nombreHoja)          ' solo filtros o formularios
        Case "INVENTARIO": PermisoEditar = "INVENTARIO"
        Case "CLIENTES": PermisoEditar = "CLIENTES"
        Case "VENDEDORES", "USUARIOS Y PERMISOS": PermisoEditar = "USUARIOS"
        Case "INGRESOS", "EGRESOS", "GASTOS FIJOS": PermisoEditar = "FINANZAS"
        Case "BD_GUIAS": PermisoEditar = "GUIAS"
        Case "CONFIGURACION", "BASE DE DATOS": PermisoEditar = "CONFIGURACION"
        Case Else: PermisoEditar = "#"                       ' solo las macros
    End Select
End Function

Private Function PuedeVer(ByVal sh As Worksheet) As Boolean
    Dim p As String
    p = PermisoVer(sh.Name)
    If p = "" Then PuedeVer = True Else PuedeVer = TienePermiso(p)
End Function

Public Sub AplicarPermisos()
    Dim sh As Worksheet, verCostos As Boolean
    Application.ScreenUpdating = False
    WS("MENU").Visible = xlSheetVisible
    For Each sh In ThisWorkbook.Worksheets
        If Clave(sh.Name) <> "MENU" Then
            If SesionActiva() And PuedeVer(sh) Then
                sh.Visible = xlSheetVisible
            Else
                sh.Visible = xlSheetVeryHidden
            End If
        End If
    Next sh
    ' columnas de costos y utilidad
    verCostos = TienePermiso("VER_COSTOS")
    On Error Resume Next
    WS("INVENTARIO").Columns("G").Hidden = Not verCostos
    WS("INVENTARIO").Columns("N").Hidden = Not verCostos
    WS("VENDEDORES").Columns("N").Hidden = Not verCostos
    WS("FACTURACION").Columns("M:P").Hidden = Not verCostos
    WS("CLIENTES").Columns("T").Hidden = True
    WS("DASHBOARD").Columns("Q:R").Hidden = True
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

Public Sub OcultarTodo()
    Dim sh As Worksheet
    WS("MENU").Visible = xlSheetVisible
    For Each sh In ThisWorkbook.Worksheets
        If Clave(sh.Name) <> "MENU" Then sh.Visible = xlSheetVeryHidden
    Next sh
End Sub

'------------------------------------------------------------------ inicio de sesion
' Devuelve True si el usuario ingreso. Usa frmLogin (clave oculta) si existe.
Private Function PedirCredenciales(ByRef usuario As String, ByRef claveTexto As String, _
                                   ByVal titulo As String) As Boolean
    Dim f As Object
    On Error GoTo SinFormulario
    Set f = VBA.UserForms.Add("frmLogin")
    f.Caption = titulo
    f.Controls("txtUsuario").Value = usuario
    f.Controls("txtClave").Value = ""
    f.Show
    If f.Aceptado Then
        usuario = Trim$(f.Controls("txtUsuario").Value)
        claveTexto = f.Controls("txtClave").Value
        PedirCredenciales = True
    End If
    Unload f
    Exit Function
SinFormulario:
    On Error GoTo 0
    usuario = InputBox(titulo & vbCrLf & vbCrLf & "Usuario:", APP_NOMBRE, usuario)
    If usuario = "" Then Exit Function
    claveTexto = InputBox("Clave para " & usuario & ":" & vbCrLf & _
                          "(Instale frmLogin con InstalarERP para ocultar la clave)", APP_NOMBRE)
    PedirCredenciales = (claveTexto <> "")
End Function

Public Function PedirClaveOculta(ByVal titulo As String) As String
    Dim f As Object
    On Error GoTo SinFormulario
    Set f = VBA.UserForms.Add("frmClave")
    f.Caption = titulo
    f.Controls("lblMensaje").Caption = titulo
    f.Show
    If f.Aceptado Then PedirClaveOculta = f.Controls("txtClave").Value
    Unload f
    Exit Function
SinFormulario:
    PedirClaveOculta = InputBox(titulo, APP_NOMBRE)
End Function

' Valida usuario y clave contra USUARIOS Y PERMISOS. Devuelve la fila o 0.
Public Function ValidarCredenciales(ByVal usuario As String, ByVal claveTexto As String) As Long
    Dim sh As Worksheet, r As Long
    Set sh = WS("USUARIOS Y PERMISOS")
    r = BuscarFila(sh, 1, usuario)
    If r = 0 Then Exit Function
    If Clave(CStr(sh.Cells(r, 5).Value)) <> "ACTIVO" Then Exit Function
    If CStr(sh.Cells(r, 6).Value) <> HashClave(CStr(sh.Cells(r, 1).Value), claveTexto) Then Exit Function
    ValidarCredenciales = r
End Function

Public Sub IniciarSesion()
    Dim usuario As String, claveTexto As String, r As Long, sh As Worksheet
    Set sh = WS("USUARIOS Y PERMISOS")
    mSesion = False
    For mIntentos = 1 To 3
        If Not PedirCredenciales(usuario, claveTexto, "Iniciar sesión - " & CStr(Cfg("cfg_Empresa"))) Then Exit For
        r = ValidarCredenciales(usuario, claveTexto)
        If r > 0 Then
            Application.EnableEvents = False
            SetCfg "app_Usuario", CStr(sh.Cells(r, 1).Value)
            SetCfg "app_Rol", CStr(sh.Cells(r, 3).Value)
            SetCfg "app_Vendedor", CStr(sh.Cells(r, 4).Value)
            sh.Cells(r, 7).Value = Now
            sh.Cells(r, 7).NumberFormat = "dd/mm/yyyy hh:mm"
            Application.EnableEvents = True
            mSesion = True
            Bitacora "SEGURIDAD", "INGRESO", "", "Inicio de sesión en " & Equipo()
            If Clave(CStr(sh.Cells(r, 8).Value)) = "SI" Then
                MsgBox "Por seguridad debe cambiar su clave ahora.", vbInformation, APP_NOMBRE
                If Not CambiarClaveFila(r) Then
                    MsgBox "No se cambió la clave. Se cerrará la sesión.", vbExclamation, APP_NOMBRE
                    CerrarSesionSilenciosa
                    Exit Sub
                End If
            End If
            AplicarPermisos
            IrA "MENU"
            Exit Sub
        Else
            Bitacora "SEGURIDAD", "INGRESO FALLIDO", "", "Usuario: " & usuario
            MsgBox "Usuario o clave incorrectos, o usuario inactivo. Intento " & mIntentos & " de 3.", _
                   vbExclamation, APP_NOMBRE
        End If
    Next mIntentos
    ' sin sesion: el libro queda solo con el MENU visible
    CerrarSesionSilenciosa
    MsgBox "No hay sesión activa. Use el botón INICIAR SESIÓN del menú.", vbInformation, APP_NOMBRE
End Sub

Private Sub CerrarSesionSilenciosa()
    Application.EnableEvents = False
    mSesion = False
    SetCfg "app_Usuario", "(sin sesión)"
    SetCfg "app_Rol", ""
    SetCfg "app_Vendedor", ""
    OcultarTodo
    Application.EnableEvents = True
End Sub

Public Sub CerrarSesion()
    If SesionActiva() Then Bitacora "SEGURIDAD", "SALIDA", "", "Cierre de sesión"
    CerrarSesionSilenciosa
    IniciarSesion
End Sub

'------------------------------------------------------------------ claves
Private Function CambiarClaveFila(ByVal r As Long) As Boolean
    Dim sh As Worksheet, c1 As String, c2 As String
    Set sh = WS("USUARIOS Y PERMISOS")
    c1 = PedirClaveOculta("Nueva clave para " & sh.Cells(r, 1).Value & " (mínimo 8 caracteres):")
    If Len(c1) < 8 Then
        If c1 <> "" Then MsgBox "La clave debe tener al menos 8 caracteres.", vbExclamation, APP_NOMBRE
        Exit Function
    End If
    c2 = PedirClaveOculta("Repita la nueva clave:")
    If c1 <> c2 Then
        MsgBox "Las claves no coinciden.", vbExclamation, APP_NOMBRE
        Exit Function
    End If
    Application.EnableEvents = False
    sh.Cells(r, 6).Value = HashClave(CStr(sh.Cells(r, 1).Value), c1)
    sh.Cells(r, 8).Value = "NO"
    Application.EnableEvents = True
    Bitacora "SEGURIDAD", "CAMBIO CLAVE", CStr(sh.Cells(r, 1).Value), ""
    GuardarLibro
    CambiarClaveFila = True
End Function

Public Sub CambiarMiClave()
    Dim r As Long, actual As String
    If Not SesionActiva() Then Exit Sub
    actual = PedirClaveOculta("Clave actual de " & UsuarioActual() & ":")
    If actual = "" Then Exit Sub
    r = ValidarCredenciales(UsuarioActual(), actual)
    If r = 0 Then
        MsgBox "La clave actual no es correcta.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    If CambiarClaveFila(r) Then MsgBox "Clave actualizada.", vbInformation, APP_NOMBRE
End Sub

' Administrador: asigna una clave temporal al usuario de la fila seleccionada
Public Sub RestablecerClave()
    Dim sh As Worksheet, r As Long, temporal As String
    If Not RequierePermiso("USUARIOS") Then Exit Sub
    Set sh = WS("USUARIOS Y PERMISOS")
    If ActiveSheet.Name <> sh.Name Then sh.Activate
    r = ActiveCell.Row
    If r < FILA_DATOS Or Vacio(sh.Cells(r, 1).Value) Then
        MsgBox "Seleccione una celda en la fila del usuario.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    temporal = PedirTexto("Clave temporal para " & sh.Cells(r, 1).Value & " (el usuario deberá cambiarla):", _
                          "Temporal" & Format$(Now, "nnss"))
    If Len(temporal) < 6 Then Exit Sub
    Application.EnableEvents = False
    sh.Cells(r, 6).Value = HashClave(CStr(sh.Cells(r, 1).Value), temporal)
    sh.Cells(r, 8).Value = "SI"
    Application.EnableEvents = True
    Bitacora "SEGURIDAD", "RESTABLECER CLAVE", CStr(sh.Cells(r, 1).Value), ""
    GuardarLibro
    MsgBox "Clave temporal asignada. Entréguela al usuario por un medio seguro.", vbInformation, APP_NOMBRE
End Sub

' Pide autorizacion de un usuario con el permiso indicado. Devuelve el usuario que autoriza o "".
Public Function Autorizar(ByVal permiso As String, ByVal motivo As String) As String
    Dim usuario As String, claveTexto As String, r As Long, sh As Worksheet
    If TienePermiso(permiso) Then Autorizar = UsuarioActual(): Exit Function
    If MsgBox(motivo & vbCrLf & vbCrLf & "Se requiere autorización de un usuario con el permiso " & permiso & _
              ". ¿Desea solicitarla ahora?", vbQuestion + vbYesNo, APP_NOMBRE) = vbNo Then Exit Function
    If Not PedirCredenciales(usuario, claveTexto, "Autorización: " & permiso) Then Exit Function
    r = ValidarCredenciales(usuario, claveTexto)
    Set sh = WS("USUARIOS Y PERMISOS")
    If r = 0 Then
        MsgBox "Credenciales de autorización no válidas.", vbExclamation, APP_NOMBRE
        Exit Function
    End If
    If Not TienePermiso(permiso, CStr(sh.Cells(r, 3).Value)) Then
        MsgBox "Ese usuario no tiene el permiso " & permiso & ".", vbExclamation, APP_NOMBRE
        Exit Function
    End If
    Autorizar = CStr(sh.Cells(r, 1).Value)
    Bitacora "SEGURIDAD", "AUTORIZACIÓN", permiso, "Autoriza " & Autorizar & ": " & motivo
End Function

'------------------------------------------------------------------ eventos del libro
Public Sub AlAbrir()
    On Error Resume Next
    Application.EnableEvents = False
    mSesion = False
    SetCfg "app_Usuario", "(sin sesión)"
    SetCfg "app_Rol", ""
    SetCfg "app_Vendedor", ""
    OcultarTodo
    ProtegerTodo
    Application.EnableEvents = True
    On Error GoTo 0
    ThisWorkbook.Saved = True
    IniciarSesion
    modRespaldo.ProgramarRespaldo
End Sub

Public Sub AlCerrar()
    On Error Resume Next
    modRespaldo.CancelarRespaldoProgramado
    If SesionActiva() Then
        Bitacora "SEGURIDAD", "SALIDA", "", "Cierre del libro"
        modRespaldo.CrearRespaldo True
    End If
    Application.EnableEvents = False
    mSesion = False
    SetCfg "app_Usuario", "(sin sesión)"
    SetCfg "app_Rol", ""
    SetCfg "app_Vendedor", ""
    OcultarTodo
    ThisWorkbook.Save
    Application.EnableEvents = True
End Sub

' Antes de guardar: el archivo en disco queda con todas las hojas ocultas salvo el MENU,
' para que sin macros no se vea ni se modifique la informacion.
Public Sub AntesDeGuardar(ByVal SaveAsUI As Boolean, ByRef Cancel As Boolean)
    Dim activa As String
    If mGuardando Then Exit Sub
    If SaveAsUI Then Exit Sub
    On Error GoTo Fin
    mGuardando = True
    Cancel = True
    activa = ActiveSheet.Name
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    OcultarTodo
    ThisWorkbook.Save
    If SesionActiva() Then
        AplicarPermisos
        If ThisWorkbook.Worksheets(activa).Visible = xlSheetVisible Then ThisWorkbook.Worksheets(activa).Activate
    End If
    ThisWorkbook.Saved = True
Fin:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    mGuardando = False
End Sub

' Control de edicion manual: deshace cambios en hojas sin permiso de edicion
Public Sub ValidarEdicion(ByVal sh As Object, ByVal Target As Range)
    Dim p As String
    If mGuardando Then Exit Sub
    If Clave(sh.Name) = "CLIENTES" And Target.Address = "$E$4" Then Exit Sub    ' selector de ficha
    If SesionActiva() Then
        p = PermisoEditar(sh.Name)
        If p <> "#" Then
            If p = "" Then Exit Sub
            If TienePermiso(p) Then Exit Sub
        End If
    End If
    Application.EnableEvents = False
    On Error Resume Next
    Application.Undo
    On Error GoTo 0
    Application.EnableEvents = True
    MsgBox "No tiene permiso para modificar esta hoja (" & sh.Name & ").", vbExclamation, APP_NOMBRE
End Sub
