Attribute VB_Name = "modMaestros"
'==============================================================================
' modMaestros - Alta de clientes, proveedores, vendedores, usuarios y productos
' (un SKU por cada combinacion producto + talla + color) y ajustes de
' inventario autorizados. Evita codigos y documentos duplicados.
'==============================================================================
Option Explicit

'------------------------------------------------------------------ CLIENTES
Public Function NuevoCliente() As String
    Dim sh As Worksheet, r As Long, cod As String, nombre As String, tdoc As String, documento As String
    Dim tel As String, ciudad As String, canal As String, correo As String, direccion As String, wa As String
    If Not RequierePermiso("CLIENTES") Then Exit Function
    Set sh = WS("CLIENTES")
    nombre = PedirTexto("Nombre completo o razón social del cliente:")
    If nombre = "" Then Exit Function
    tdoc = ElegirDeLista("Tipo de documento:", ListaCatalogo("Tipos de documento"), "CC")
    If tdoc = "" Then Exit Function
    documento = PedirTexto("Número de " & tdoc & " (sin puntos):")
    If documento = "" Then Exit Function
    r = BuscarFila(sh, 4, documento, FILA_DATOS_CLI)
    If r > 0 Then
        MsgBox "Ya existe el cliente " & sh.Cells(r, 1).Value & " - " & sh.Cells(r, 2).Value & " con ese documento.", _
               vbExclamation, APP_NOMBRE
        NuevoCliente = CStr(sh.Cells(r, 1).Value)
        Exit Function
    End If
    tel = PedirTexto("Teléfono:")
    If tel = "" Then Exit Function
    wa = PedirTexto("WhatsApp:", tel, False)
    correo = PedirTexto("Correo (opcional):", "", False)
    direccion = PedirTexto("Dirección de entrega:")
    ciudad = ElegirDeLista("Ciudad (si no está, escriba 0 y agréguela en BASE DE DATOS):", ListaCatalogo("Ciudades"), "Bogotá")
    If ciudad = "" Then ciudad = PedirTexto("Escriba la ciudad:")
    canal = ElegirDeLista("¿Por qué canal llegó el cliente?", ListaCatalogo("Canales de venta"), "WhatsApp")
    If canal = "" Then Exit Function
    cod = SiguienteID(sh, 1, "CLI", 4, FILA_DATOS_CLI)
    r = SiguienteFila(sh, 1, FILA_DATOS_CLI)
    Application.EnableEvents = False
    sh.Range("A" & r & ":K" & r).Value = Array(cod, nombre, tdoc, documento, tel, wa, correo, direccion, ciudad, Date, canal)
    Application.EnableEvents = True
    Bitacora "CLIENTES", "NUEVO", cod, nombre & " · " & tdoc & " " & documento
    GuardarLibro
    MsgBox "Cliente " & cod & " creado.", vbInformation, APP_NOMBRE
    NuevoCliente = cod
End Function

Public Sub BotonNuevoCliente()
    Dim cod As String
    cod = NuevoCliente()
End Sub

'------------------------------------------------------------------ PROVEEDORES
Public Sub NuevoProveedor()
    Dim sh As Worksheet, r As Long, cod As String, nombre As String, nit As String
    If Not RequierePermiso("PROVEEDORES") Then Exit Sub
    Set sh = WS("PROVEEDORES")
    nombre = PedirTexto("Razón social del proveedor:")
    If nombre = "" Then Exit Sub
    nit = PedirTexto("NIT (con dígito de verificación):")
    If nit = "" Then Exit Sub
    r = BuscarFila(sh, 3, nit)
    If r > 0 Then MsgBox "Ya existe " & sh.Cells(r, 1).Value & " con ese NIT.", vbExclamation, APP_NOMBRE: Exit Sub
    cod = SiguienteID(sh, 1, "PRV", 3)
    r = SiguienteFila(sh)
    Application.EnableEvents = False
    sh.Range("A" & r & ":C" & r).Value = Array(cod, nombre, nit)
    sh.Range("D" & r).Value = PedirTexto("Nombre del contacto:", "", False)
    sh.Range("E" & r).Value = PedirTexto("Teléfono:", "", False)
    sh.Range("F" & r).Value = PedirTexto("WhatsApp:", CStr(sh.Range("E" & r).Value), False)
    sh.Range("G" & r).Value = PedirTexto("Correo:", "", False)
    sh.Range("I" & r).Value = PedirTexto("Ciudad:", "", False)
    sh.Range("J" & r).Value = PedirTexto("Productos que suministra:", "", False)
    Application.EnableEvents = True
    Bitacora "PROVEEDORES", "NUEVO", cod, nombre & " · NIT " & nit
    GuardarLibro
    sh.Activate
    sh.Range("H" & r).Select
    MsgBox "Proveedor " & cod & " creado. Complete dirección, precios y datos bancarios en su fila.", vbInformation, APP_NOMBRE
End Sub

'------------------------------------------------------------------ VENDEDORES
Public Sub NuevoVendedor()
    Dim sh As Worksheet, r As Long, cod As String, nombre As String, pct As Variant, meta As Variant
    If Not RequierePermiso("USUARIOS") Then Exit Sub
    Set sh = WS("VENDEDORES")
    nombre = PedirTexto("Nombre del vendedor:")
    If nombre = "" Then Exit Sub
    pct = PedirNumero("% de comisión (ej. 5 para 5%):", 5, 0)
    If IsEmpty(pct) Then Exit Sub
    meta = PedirNumero("Meta mensual de ventas (COP):", 3000000, 0)
    If IsEmpty(meta) Then Exit Sub
    cod = SiguienteID(sh, 1, "VEN", 2)
    r = SiguienteFila(sh)
    Application.EnableEvents = False
    sh.Range("A" & r & ":I" & r).Value = Array(cod, nombre, PedirTexto("Identificación:", "", False), _
        LCase$(PedirTexto("Usuario del sistema (si tiene):", "", False)), "Vendedor", "ACTIVO", _
        ElegirDeLista("Canal principal:", ListaCatalogo("Canales de venta")), CDbl(pct) / 100, CDbl(meta))
    Application.EnableEvents = True
    Bitacora "VENDEDORES", "NUEVO", cod, nombre
    GuardarLibro
    MsgBox "Vendedor " & cod & " creado. Si usará el sistema, créele un usuario con rol Vendedor y asocie " & cod & ".", _
           vbInformation, APP_NOMBRE
End Sub

'------------------------------------------------------------------ USUARIOS
Public Sub NuevoUsuario()
    Dim sh As Worksheet, r As Long, usuario As String, rol As String, temporal As String, ven As String
    If Not RequierePermiso("USUARIOS") Then Exit Sub
    Set sh = WS("USUARIOS Y PERMISOS")
    usuario = LCase$(PedirTexto("Nombre de usuario (sin espacios):"))
    If usuario = "" Then Exit Sub
    If InStr(usuario, " ") > 0 Then MsgBox "El usuario no puede tener espacios.", vbExclamation, APP_NOMBRE: Exit Sub
    If BuscarFila(sh, 1, usuario) > 0 Then MsgBox "Ese usuario ya existe.", vbExclamation, APP_NOMBRE: Exit Sub
    rol = ElegirDeLista("Rol:", ListaCatalogo("Roles"), "Vendedor")
    If rol = "" Then Exit Sub
    If Clave(rol) = "VENDEDOR" Then ven = UCase$(PedirTexto("Código de vendedor asociado (ej. VEN-01):"))
    temporal = PedirTexto("Clave temporal (mínimo 8 caracteres; deberá cambiarla al entrar):")
    If Len(temporal) < 8 Then MsgBox "Clave demasiado corta.", vbExclamation, APP_NOMBRE: Exit Sub
    r = SiguienteFila(sh)
    Application.EnableEvents = False
    sh.Range("A" & r & ":H" & r).Value = Array(usuario, PedirTexto("Nombre completo:", usuario), rol, ven, "ACTIVO", _
        HashClave(usuario, temporal), Empty, "SI")
    Application.EnableEvents = True
    Bitacora "SEGURIDAD", "NUEVO USUARIO", usuario, rol
    GuardarLibro
    MsgBox "Usuario " & usuario & " creado con rol " & rol & ".", vbInformation, APP_NOMBRE
End Sub

'------------------------------------------------------------------ PRODUCTOS / SKU
Private Function Abrev(ByVal s As String, ByVal n As Integer) As String
    s = UCase$(SinTildes(Replace(Trim$(s), " ", "")))
    If Clave(s) = "UNICA" Then s = "U"
    Abrev = Left$(s, n)
End Function

Public Sub NuevoProducto()
    Dim sh As Worksheet, codBase As String, nombre As String, cat As String, prv As String, tallas As Variant
    Dim colores As Variant, costo As Variant, precio As Variant, minimo As Variant, inicial As Variant
    Dim t As Variant, c As Variant, sku As String, r As Long, creados As String, n As Long, existentes As String
    Dim hK As Worksheet

    If Not RequierePermiso("INVENTARIO") Then Exit Sub
    Set sh = WS("INVENTARIO")
    nombre = PedirTexto("Nombre del producto (ej. Camiseta oversize):")
    If nombre = "" Then Exit Sub
    codBase = UCase$(PedirTexto("Código base de 3-5 letras para los SKU (ej. CAMO):", Abrev(nombre, 4)))
    If codBase = "" Then Exit Sub
    cat = ElegirDeLista("Categoría (para una nueva, agréguela en BASE DE DATOS):", ListaCatalogo("Categorías"))
    If cat = "" Then Exit Sub
    prv = UCase$(PedirTexto("Código del proveedor principal (ej. PRV-001):", "", False))
    tallas = Split(PedirTexto("Tallas separadas por coma (ej. S,M,L o 38,39,40 o Única):", "Única"), ",")
    colores = Split(PedirTexto("Colores separados por coma (ej. Negro,Blanco):", "Negro"), ",")
    If UBound(tallas) < 0 Or UBound(colores) < 0 Then Exit Sub
    costo = PedirNumero("Costo de compra por unidad (COP):", "", 0)
    If IsEmpty(costo) Then Exit Sub
    precio = PedirNumero("Precio de venta (COP):", "", 1)
    If IsEmpty(precio) Then Exit Sub
    If precio <= costo Then
        If MsgBox("El precio de venta es menor o igual al costo. ¿Continuar?", vbExclamation + vbYesNo, APP_NOMBRE) = vbNo Then Exit Sub
    End If
    minimo = PedirNumero("Stock mínimo por SKU:", 2, 0)
    If IsEmpty(minimo) Then Exit Sub
    inicial = PedirNumero("Stock inicial por SKU (0 si entrará por compras):", 0, 0)
    If IsEmpty(inicial) Then Exit Sub

    InicioMacro
    Set hK = WS("HISTORIAL MOVIMIENTOS")
    For Each t In tallas
        For Each c In colores
            sku = codBase & "-" & Abrev(CStr(t), 4) & "-" & Abrev(CStr(c), 3)
            If BuscarFila(sh, 1, sku) > 0 Then
                existentes = existentes & " " & sku
            Else
                r = SiguienteFila(sh)
                sh.Range("A" & r & ":I" & r).Value = Array(sku, nombre, cat, Trim$(CStr(t)), Trim$(CStr(c)), prv, _
                    CDbl(costo), CDbl(precio), CDbl(inicial))
                sh.Range("M" & r).Value = CDbl(minimo)
                If inicial > 0 Then
                    EscribirFila hK, SiguienteFila(hK), Array(SiguienteID(hK, 1, "K", 6), Date, Format$(Now, "hh:mm"), _
                        UsuarioActual(), "INICIAL", sku, nombre & " " & Trim$(CStr(t)) & " " & Trim$(CStr(c)), _
                        CDbl(inicial), Empty, CDbl(costo), "INICIAL", "Stock inicial al crear el SKU", UsuarioActual())
                End If
                creados = creados & vbCrLf & sku
                n = n + 1
            End If
        Next c
    Next t
    Bitacora "INVENTARIO", "NUEVO PRODUCTO", codBase, n & " SKU · " & nombre
    FinMacro
    GuardarLibro
    MsgBox n & " SKU creados:" & creados & IIf(existentes <> "", vbCrLf & vbCrLf & "Ya existían (no se duplicaron):" & _
           existentes, ""), vbInformation, APP_NOMBRE
End Sub

Public Sub AjusteInventario()
    Dim sh As Worksheet, sku As String, r As Long, tipo As String, cant As Variant, motivo As String
    Dim autoriza As String, exist As Double, hK As Worksheet
    If Not RequierePermiso("INVENTARIO") Then Exit Sub
    Set sh = WS("INVENTARIO")
    Application.Calculate
    If ActiveSheet.Name = sh.Name And ActiveCell.Row >= FILA_DATOS Then sku = CStr(sh.Cells(ActiveCell.Row, 1).Value)
    sku = UCase$(PedirTexto("SKU a ajustar:", sku))
    If sku = "" Then Exit Sub
    r = BuscarFila(sh, 1, sku)
    If r = 0 Then MsgBox "No existe el SKU " & sku & ".", vbExclamation, APP_NOMBRE: Exit Sub
    exist = NumV(sh.Cells(r, 12).Value)
    tipo = ElegirDeLista("Tipo de ajuste para " & sku & " (existencias: " & exist & "):", _
                         Array("AJUSTE ENTRADA", "AJUSTE SALIDA"), "AJUSTE SALIDA")
    If tipo = "" Then Exit Sub
    cant = PedirNumero("Cantidad:", 1, 1)
    If IsEmpty(cant) Then Exit Sub
    If tipo = "AJUSTE SALIDA" And cant > exist And Clave(CStr(Cfg("cfg_PermitirNegativo"))) <> "SI" Then
        MsgBox "El ajuste dejaría el inventario negativo y la configuración no lo permite.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    motivo = PedirTexto("Motivo del ajuste (avería, conteo físico, obsequio...):")
    If motivo = "" Then Exit Sub
    autoriza = Autorizar("AJUSTE_INVENTARIO", "Ajuste de inventario " & sku & " · " & tipo & " " & cant)
    If autoriza = "" Then Exit Sub
    Set hK = WS("HISTORIAL MOVIMIENTOS")
    InicioMacro
    EscribirFila hK, SiguienteFila(hK), Array(SiguienteID(hK, 1, "K", 6), Date, Format$(Now, "hh:mm"), UsuarioActual(), _
        tipo, sku, sh.Cells(r, 2).Value & " " & sh.Cells(r, 4).Value & " " & sh.Cells(r, 5).Value, _
        IIf(tipo = "AJUSTE ENTRADA", CDbl(cant), Empty), IIf(tipo = "AJUSTE SALIDA", CDbl(cant), Empty), _
        sh.Cells(r, 7).Value, SiguienteID(hK, 11, "AJ", 4), motivo, autoriza)
    Bitacora "INVENTARIO", "AJUSTE", sku, tipo & " " & cant & " · " & motivo & " · autoriza " & autoriza
    FinMacro
    GuardarLibro
    MsgBox "Ajuste registrado en el historial de movimientos.", vbInformation, APP_NOMBRE
End Sub
