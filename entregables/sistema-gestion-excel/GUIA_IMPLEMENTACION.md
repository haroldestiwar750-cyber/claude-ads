# Sistema Integral de Gestión Empresarial en Excel — Guía de implementación

**Archivos del paquete**

| Archivo | Para qué sirve |
|---|---|
| `Sistema_Gestion_Empresarial.xlsx` | Libro completo: 18 hojas, unas 33.700 fórmulas, gráficos, listas desplegables, formatos condicionales y datos de ejemplo. |
| `modGestion.bas` | Módulo VBA completo y comentado, listo para **importar** (no hace falta copiar y pegar). |
| `GUIA_IMPLEMENTACION.md` | Este documento: estructura de cada hoja, fórmulas en español y código VBA. |

> El libro va como `.xlsx` y el código por separado. El `.xlsm` con macros se arma en su propio Excel en unos 3 minutos (paso 1), y así usted revisa todo el código antes de activarlo.

---

## 1. Implementación paso a paso (Windows, Excel 2016, 2019, 2021 o Microsoft 365)

### 1.1 Convertir el libro en `.xlsm`
1. Abra `Sistema_Gestion_Empresarial.xlsx`.
2. **Archivo → Guardar como → Examinar**. En *Tipo* elija **Libro de Excel habilitado para macros (*.xlsm)** → Guardar.
3. Cierre y vuelva a abrir el `.xlsm`. Si aparece la barra amarilla *"Advertencia de seguridad: las macros se han deshabilitado"*, pulse **Habilitar contenido**.
   - Si no aparece: **Archivo → Opciones → Centro de confianza → Configuración del Centro de confianza → Configuración de macros → "Deshabilitar macros de VBA con notificación"**.
   - Si Windows bloquea el archivo descargado: clic derecho sobre el archivo → **Propiedades → marque "Desbloquear"** → Aceptar.

### 1.2 Pegar el código VBA
**Opción A, recomendada: importar el módulo**
1. Pulse **Alt + F11** (abre el editor de Visual Basic).
2. Menú **Archivo → Importar archivo…** → seleccione `modGestion.bas` → Abrir.
3. En el panel izquierdo (*Proyecto – VBAProject*) aparecerá **Módulos → modGestion**.

**Opción B: copiar y pegar**
1. Alt + F11 → menú **Insertar → Módulo**.
2. Copie el código de la sección 6 de esta guía **sin la primera línea** `Attribute VB_Name = "modGestion"` (esa línea solo sirve al importar).
3. Péguelo en la ventana en blanco. Opcional: en la ventana *Propiedades* (F4) cambie el nombre `Módulo1` por `modGestion`.

**Opcional: código de ThisWorkbook.** Abre el DASHBOARD al iniciar y vuelve a aplicar la protección. En el panel izquierdo haga doble clic en **ThisWorkbook** y pegue:

```vba
Private Sub Workbook_Open()
    'Quite el apóstrofo de la siguiente línea si usa ProtegerHojas:
    'ProtegerHojas
    ThisWorkbook.Worksheets("DASHBOARD").Activate
End Sub
```

4. Menú **Depuración → Compilar VBAProject**. Si no sale ningún mensaje, el código está correcto.
5. Cierre el editor (Alt + Q) y **guarde** (Ctrl + G / Ctrl + S).

### 1.3 Crear los botones
**Opción automática (recomendada):** pulse **Alt + F8**, seleccione **CrearBotones** y luego **Ejecutar**. Se crean los 4 botones de FACTURACION y los 5 de GUIAS en la columna K, ya asignados, y marcados para **no imprimirse**.

**Opción manual (si prefiere sus propias formas o colores):**
1. Active la pestaña **Programador**: Archivo → Opciones → Personalizar cinta → marque **Programador**.
2. **Programador → Insertar → Controles de formulario → Botón** (el primer ícono). Dibuje el botón en la columna K.
3. Se abre *Asignar macro*: elija la macro correspondiente → Aceptar.
4. Clic derecho sobre el botón → *Modificar texto*. Clic derecho → *Formato de control → Propiedades* → **desmarque "Imprimir objeto"**.
5. También puede usar una forma (Insertar → Formas): clic derecho → **Asignar macro**.

| Hoja | Texto del botón | Macro a asignar |
|---|---|---|
| FACTURACION | IMPRIMIR FACTURA | `ImprimirFactura` |
| FACTURACION | LIMPIAR FACTURA | `LimpiarFactura` |
| FACTURACION | GUARDAR (Historial + PDF) | `GuardarFactura` |
| FACTURACION | NUEVA FACTURA | `NuevaFactura` |
| GUIAS | CARGAR DESDE FACTURA | `CargarFacturaEnGuia` |
| GUIAS | IMPRIMIR GUÍA | `ImprimirGuia` |
| GUIAS | LIMPIAR GUÍA | `LimpiarGuia` |
| GUIAS | GUARDAR GUÍA (PDF) | `GuardarGuia` |
| GUIAS | NUEVA GUÍA | `NuevaGuia` |

### 1.4 Dejar el libro listo para su negocio
1. **CONFIG**: nombre, NIT, dirección, IVA, prefijos, consecutivo inicial (por ejemplo, si su última factura fue la 350, escriba 351), nombre de la carpeta del Escritorio y formato (PDF / EXCEL / AMBOS).
2. Borre los **datos de ejemplo**. Use la tecla *Supr* y **solo en las celdas amarillas**: PRODUCTOS, CLIENTES, PROVEEDORES, COMPRAS, COSTOS (col. A y D:H), INGRESOS_EGRESOS (A, B, D:G), GASTOS_FIJOS (A:D). En HIST_FACTURAS, HIST_DETALLE y HIST_GUIAS borre las filas 5 en adelante. En CARTERA borre **solo A:F**, porque G:K son fórmulas.
3. En FACTURACION pulse **NUEVA FACTURA** y en GUIAS **NUEVA GUÍA** para que tomen el consecutivo inicial.
4. Opcional: **Alt+F8 → ProtegerHojas** bloquea todas las fórmulas, de modo que solo se editan las celdas amarillas. Para modificar la estructura use **DesprotegerHojas**. Si define una clave, escríbala en la constante `CLAVE` del módulo.

### 1.5 Qué hace cada botón de Facturación

| Botón | Qué hace exactamente |
|---|---|
| **Imprimir factura** | Fija el área de impresión en `A1:H44` (sin la columna de stock ni los botones), 1 página, centrada, en vertical. Pregunta: **Sí** = imprimir, **No** = vista previa. |
| **Limpiar factura** | Borra **solo constantes** de B10, B14:B33, D14:D33 y A36. Nunca borra una celda con fórmula. Además **reescribe** las fórmulas de la tabla por si alguien las sobrescribió. Pone la fecha de hoy y "Contado". No cambia el número. |
| **Guardar** | 1) Valida la fecha, que el cliente exista, que los SKU y las cantidades sean válidos, que no haya un consecutivo duplicado, y avisa si falta stock. 2) Exporta el PDF y/o el Excel a `Escritorio\<Carpeta CONFIG B17>\Facturas\AAAA-MM\FE-1007_Cliente.pdf` y crea las carpetas que falten. 3) Escribe HIST_DETALLE (descuenta inventario) y HIST_FACTURAS. 4) Si es **Crédito**, la crea en CARTERA. Si es **Contado**, registra el ingreso en INGRESOS_EGRESOS. 5) Marca la factura como GUARDADA, guarda el libro y ofrece crear una nueva. |
| **Nueva factura** | Si hay una factura sin guardar, pide confirmación. Limpia y pone como número **el mayor consecutivo guardado + 1** (o el inicial de CONFIG). Así no quedan huecos en la numeración aunque se pulse varias veces. |

**Ruta dinámica del Escritorio.** El código usa `CreateObject("WScript.Shell").SpecialFolders("Desktop")`. Esa ruta corresponde al Escritorio real del usuario que tiene la sesión abierta, **incluso si está redirigido a OneDrive** (`C:\Users\<usuario>\OneDrive\Escritorio`). Si eso falla, usa `%USERPROFILE%\Desktop`. Funciona en cualquier PC con Windows sin cambiar nada.

---

## 2. Separador de argumentos y nombres de funciones

- **Dentro del archivo no tiene que hacer nada.** Excel guarda las fórmulas internamente en inglés y las muestra traducidas al idioma de su instalación, con el separador de su configuración regional.
- **Para copiar y pegar a mano**, esta guía muestra todas las fórmulas **en español con punto y coma `;`**, que es el formato de Excel en español en Colombia, México, España y la mayoría de Latinoamérica.
- Si su Excel usa **coma** como separador de argumentos (configuración regional con punto decimal, por ejemplo EE. UU.), reemplace `;` por `,`. Si su Excel está **en inglés**, use los nombres en inglés: SUMAR.SI.CONJUNTO = SUMIFS, SI.ERROR = IFERROR, INDICE = INDEX, COINCIDIR = MATCH, BUSCARV = VLOOKUP, BUSCAR = LOOKUP, FIN.MES = EOMONTH, K.ESIMO.MAYOR = LARGE, ELEGIR = CHOOSE, CONTAR.SI.CONJUNTO = COUNTIFS, REDONDEAR.MAS = ROUNDUP, FILA = ROW, ESNUMERO = ISNUMBER.
- Para saber cuál usa: escriba `=SUMA(1;2)` en una celda. Si da 3, su separador es `;`.
- **Copiar una fórmula de tabla:** pegue la de la fila 5 (o la indicada) y **arrástrela hacia abajo**. Las referencias con `$` quedan fijas y las demás se ajustan solas.
- Solo se usan funciones compatibles con **Excel 2010 o posterior**. No se usan BUSCARX, FILTRAR ni UNICOS, así que funciona también en Excel 2016 y 2019.

---

## 3. Arquitectura y flujo de datos

```
CONFIG ───────────────► (todas las hojas leen parámetros, listas y el mes de análisis)
PRODUCTOS ──► INVENTARIO ◄── COMPRAS (entradas)          ◄── HIST_DETALLE (salidas)
CLIENTES ───► FACTURACION ──[GUARDAR]──► HIST_FACTURAS + HIST_DETALLE
                                     ├─► CARTERA            (si es Crédito)
                                     ├─► INGRESOS_EGRESOS   (si es Contado)
                                     └─► PDF / Excel en  Escritorio\Carpeta\Facturas\AAAA-MM
HIST_FACTURAS ──[CARGAR]──► GUIAS ──[GUARDAR GUÍA]──► HIST_GUIAS + PDF (…\Guias\AAAA-MM)
INGRESOS_EGRESOS ──► CARTERA (abonos) · GASTOS_FIJOS (real) · PROVEEDORES (pagos)
INVENTARIO ──► COSTOS (costo directo) ──► PUNTO_EQUILIBRIO ◄── GASTOS_FIJOS
Todo ──► DASHBOARD
```

**Reglas de oro para que nada falle**
1. **No inserte ni elimine filas o columnas** en FACTURACION, GUIAS ni CONFIG: las macros usan direcciones fijas (constantes al inicio del módulo). Si necesita hacerlo, actualice las constantes.
2. En PRODUCTOS **no ordene ni elimine filas** con movimientos: INVENTARIO lee la misma fila. Para retirar un producto déjelo con stock 0. Los productos nuevos van al final.
3. Los SKU y NIT se escriben como **texto** (las columnas ya tienen formato Texto). No los cambie después de usarlos.
4. Los textos de las listas (Contado/Crédito, categorías del flujo de caja) deben coincidir **exactamente**. Por eso se eligen desde menús desplegables.
5. **Abono de un cliente**: en INGRESOS_EGRESOS use la categoría `Abono cartera` y escriba el **número de factura** (por ejemplo `FE-1002`) en *Documento ref.*
6. **Pago a proveedor**: categoría `Pago proveedores` y el **nombre exacto** del proveedor en *Tercero*.
7. **Gasto fijo**: categoría `Gasto fijo` y una descripción que **empiece** con el concepto de GASTOS_FIJOS (por ejemplo "Arriendo noviembre").
8. **Compra de contado**: regístrela en COMPRAS (inventario) **y** en INGRESOS_EGRESOS como `Compra inventario` (caja). La compra a crédito va solo en COMPRAS, y su pago después como `Pago proveedores`.
9. **Anular una factura**: no la borre del historial. Registre una devolución en COMPRAS (cantidad positiva, proveedor "DEVOLUCIÓN CLIENTE") y, si era a crédito, un abono por el total con la descripción "Nota crédito".

---

## 4. Estructura de cada hoja, celda por celda

Convención: **Entrada** = usted digita (celda amarilla, texto azul) · **Fórmula** = no tocar · **Macro** = la escribe el código VBA.
En las tablas, la fórmula se muestra para la **primera fila de datos**. Cópiela hacia abajo.


### CONFIG

Parámetros del negocio y listas desplegables. **Todo el sistema lee de aquí**: no mueva estas celdas.

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| B3 | Nombre del negocio | Entrada | — |
| B4 | NIT | Entrada | — |
| B5 | Dirección | Entrada | — |
| B6 | Ciudad | Entrada | — |
| B7 | Teléfono / WhatsApp | Entrada | — |
| B8 | Correo electrónico | Entrada | — |
| B9 | Texto legal (resolución, régimen) | Entrada | — |
| B10 | Mensaje al pie de la factura | Entrada | — |
| B12 | Tarifa de IVA | Entrada | — |
| B13 | Prefijo de factura | Entrada | — |
| B14 | Consecutivo inicial de factura | Entrada | — |
| B15 | Prefijo de guía | Entrada | — |
| B16 | Consecutivo inicial de guía | Entrada | — |
| B17 | Carpeta en el Escritorio | Entrada | — |
| B18 | Formato de copia al guardar (PDF / EXCEL / AMBOS) | Entrada | — |
| B19 | Margen mínimo para ser RENTABLE (SÍ) | Entrada | — |
| B20 | Margen mínimo aceptable (REVISAR) | Entrada | — |
| B21 | Plazo de crédito por defecto (días) | Entrada | — |
| B22 | Días de alerta 'Por vencer' | Entrada | — |
| B23 | Saldo inicial de caja | Entrada | — |
| B24 | Mes de análisis (primer día del mes) | Entrada | — |
| B26 | Costos variables adicionales (% sobre ventas) | Entrada | — |
| B27 | Método de valoración de inventario (PROMEDIO / ULTIMO) | Entrada | — |
| B25 | Fin del mes de análisis | Fórmula | `=FIN.MES(B24;0)` |
| D4:E30 | Tabla Categoría → Tipo (Ingreso/Egreso) del flujo de caja | Entrada | — |
| G4:G20 | Lista de transportadoras | Entrada | — |
| I4:I5 | Formas de pago (no cambie el texto 'Contado' / 'Crédito') | Valor fijo | — |
| K4:K20 | Categorías de gasto fijo | Entrada | — |

### PRODUCTOS

Catálogo maestro. Encabezados en fila 4, datos de la fila 5 a la 504.

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| A5 | SKU / Código | Entrada | — |
| B5 | Producto | Entrada | — |
| C5 | Categoría | Entrada | — |
| D5 | Unidad | Entrada | — |
| E5 | Costo unitario (compra) | Entrada | — |
| F5 | Precio de venta (sin IVA) | Entrada | — |
| G5 | Stock inicial | Entrada | — |
| H5 | Stock mínimo | Entrada | — |
| I5 | ¿Grava IVA? (SI/NO) | Entrada | — |
| J5 | Proveedor principal | Entrada | — |

### CLIENTES

Catálogo de clientes. Encabezados en fila 4, datos desde la fila 5.

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| A5 | NIT / CC | Entrada | — |
| B5 | Nombre / Razón social | Entrada | — |
| C5 | Teléfono | Entrada | — |
| D5 | Email | Entrada | — |
| E5 | Dirección | Entrada | — |
| F5 | Ciudad / Municipio | Entrada | — |
| G5 | Departamento | Entrada | — |
| H5 | Plazo crédito (días) | Entrada | — |

### PROVEEDORES

Catálogo de proveedores + cuentas por pagar. Encabezados en fila 4, datos desde fila 5.

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| A5 | NIT | Entrada | — |
| B5 | Razón social | Entrada | — |
| C5 | Contacto | Entrada | — |
| D5 | Teléfono | Entrada | — |
| E5 | Email | Entrada | — |
| F5 | Ciudad | Entrada | — |
| G5 | Plazo de pago (días) | Entrada | — |
| H5 | Total comprado | Fórmula | `=SI(B5="";"";SUMAR.SI.CONJUNTO(COMPRAS!$H:$H;COMPRAS!$C:$C;B5))` |
| I5 | Compras a crédito | Fórmula | `=SI(B5="";"";SUMAR.SI.CONJUNTO(COMPRAS!$H:$H;COMPRAS!$C:$C;B5;COMPRAS!$I:$I;"Crédito"))` |
| J5 | Pagos realizados | Fórmula | `=SI(B5="";"";SUMAR.SI.CONJUNTO(INGRESOS_EGRESOS!$I:$I;INGRESOS_EGRESOS!$B:$B;"Pago proveedores";INGRESOS_EGRESOS!$F:$F;B5))` |
| K5 | Saldo por pagar | Fórmula | `=SI(B5="";"";I5-J5)` |

### COMPRAS

Registro de entradas de mercancía. Encabezados en fila 4, datos desde fila 5 (fórmulas precargadas hasta la 3004).

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| A5 | Fecha | Entrada | — |
| B5 | No. factura proveedor | Entrada | — |
| C5 | Proveedor | Entrada | — |
| D5 | SKU | Entrada | — |
| E5 | Producto | Fórmula | `=SI(D5="";"";SI.ERROR(INDICE(PRODUCTOS!$B:$B;COINCIDIR(D5;PRODUCTOS!$A:$A;0));"SKU NO EXISTE"))` |
| F5 | Cantidad | Entrada | — |
| G5 | Costo unitario | Entrada | — |
| H5 | Total compra | Fórmula | `=SI(D5="";"";F5*G5)` |
| I5 | Forma de pago | Entrada | — |
| J5 | Observaciones | Entrada | — |

### INVENTARIO

100 % fórmulas. La fila *n* de INVENTARIO corresponde a la fila *n* de PRODUCTOS. Fórmulas precargadas de la fila 5 a la 504 (500 productos).

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| A5 | SKU | Fórmula | `=SI(PRODUCTOS!A5="";"";PRODUCTOS!A5)` |
| B5 | Producto | Fórmula | `=SI(A5="";"";PRODUCTOS!B5)` |
| C5 | Categoría | Fórmula | `=SI(A5="";"";PRODUCTOS!C5)` |
| D5 | Stock inicial | Fórmula | `=SI(A5="";"";PRODUCTOS!G5)` |
| E5 | Entradas (compras) | Fórmula | `=SI(A5="";"";SUMAR.SI.CONJUNTO(COMPRAS!$F:$F;COMPRAS!$D:$D;A5))` |
| F5 | Salidas (facturado) | Fórmula | `=SI(A5="";"";SUMAR.SI.CONJUNTO(HIST_DETALLE!$E:$E;HIST_DETALLE!$C:$C;A5))` |
| G5 | Stock actual | Fórmula | `=SI(A5="";"";D5+E5-F5)` |
| H5 | Costo promedio ponderado | Fórmula | `=SI(A5="";"";SI.ERROR((D5*PRODUCTOS!E5+SUMAR.SI.CONJUNTO(COMPRAS!$H:$H;COMPRAS!$D:$D;A5))/(D5+E5);PRODUCTOS!E5))` |
| I5 | Último costo | Fórmula | `=SI(A5="";"";SI.ERROR(BUSCAR(2;1/(COMPRAS!$D$5:$D$3004=A5);COMPRAS!$G$5:$G$3004);PRODUCTOS!E5))` |
| J5 | Costo de valoración | Fórmula | `=SI(A5="";"";SI(CONFIG!$B$27="ULTIMO";I5;H5))` |
| K5 | Valor total inventario | Fórmula | `=SI(A5="";"";G5*J5)` |
| L5 | Stock mínimo | Fórmula | `=SI(A5="";"";PRODUCTOS!H5)` |
| M5 | Alerta | Fórmula | `=SI(A5="";"";SI(G5<=0;"AGOTADO";SI(G5<=L5;"REPONER";"OK")))` |
| N5 | Unid. vendidas (mes análisis) | Fórmula | `=SI(A5="";"";SUMAR.SI.CONJUNTO(HIST_DETALLE!$E:$E;HIST_DETALLE!$C:$C;A5;HIST_DETALLE!$B:$B;">="&CONFIG!$B$24;HIST_DETALLE!$B:$B;"<="&CONFIG!$B$25))` |
| O5 | Ventas $ mes (sin IVA) | Fórmula | `=SI(A5="";"";SUMAR.SI.CONJUNTO(HIST_DETALLE!$G:$G;HIST_DETALLE!$C:$C;A5;HIST_DETALLE!$B:$B;">="&CONFIG!$B$24;HIST_DETALLE!$B:$B;"<="&CONFIG!$B$25))` |
| P5 | Aux. ranking | Fórmula | `=SI(A5="";"";N5+FILA()/1000000)` |

**Costo promedio ponderado** = (Stock inicial × Costo de catálogo + Σ Total de compras) / (Stock inicial + Σ Cantidad comprada). **Último costo** usa `BUSCAR(2;1/(rango=SKU);costos)`, que devuelve el costo de la última compra registrada. El método que se usa para valorizar se elige en CONFIG!B27.

### FACTURACION

Formato imprimible en **A1:H44** (área de impresión). La columna I (stock disponible) y la columna K (estado + botones) quedan **fuera** del área de impresión. Solo se digitan las celdas amarillas: B7, D7, B10, B14:B33, D14:D33 y A36.

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| A1 | Nombre del negocio | Fórmula | `=CONFIG!B3` |
| A2 | NIT del negocio | Fórmula | `="NIT: "&CONFIG!B4` |
| H2 | Número de factura visible (prefijo + consecutivo) | Fórmula | `=CONFIG!$B$13&H3` |
| A3 | Dirección y ciudad del negocio | Fórmula | `=CONFIG!B5&" - "&CONFIG!B6` |
| H3 | Consecutivo numérico (lo escribe la macro NUEVA FACTURA) | Lo escribe la macro | — |
| A4 | Teléfono y correo | Fórmula | `="Tel: "&CONFIG!B7&"   |   "&CONFIG!B8` |
| A5 | Texto legal | Fórmula | `=CONFIG!B9` |
| B7 | Fecha de la factura | Entrada | — |
| D7 | Forma de pago (Contado / Crédito) | Entrada | — |
| G7 | Fecha de vencimiento | Fórmula | `=SI(B7="";"";SI(D7="Crédito";B7+SI(SI.ERROR(INDICE(CLIENTES!$G:$G;COINCIDIR(B10;CLIENTES!$A:$A;0));0)>0;INDICE(CLIENTES!$G:$G;COINCIDIR(B10;CLIENTES!$A:$A;0));CONFIG!$B$21);B7))` |
| B10 | NIT/CC del cliente (lista de CLIENTES) | Entrada | — |
| C10 | Nombre del cliente | Fórmula | `=SI(B10="";"";SI.ERROR(INDICE(CLIENTES!$B:$B;COINCIDIR(B10;CLIENTES!$A:$A;0));"CLIENTE NO REGISTRADO"))` |
| E10 | Ciudad del cliente | Fórmula | `=SI(B10="";"";SI.ERROR(INDICE(CLIENTES!$F:$F;COINCIDIR(B10;CLIENTES!$A:$A;0));""))` |
| H10 | Teléfono del cliente | Fórmula | `=SI(B10="";"";SI.ERROR(INDICE(CLIENTES!$C:$C;COINCIDIR(B10;CLIENTES!$A:$A;0));""))` |
| B11 | Dirección del cliente | Fórmula | `=SI(B10="";"";SI.ERROR(INDICE(CLIENTES!$E:$E;COINCIDIR(B10;CLIENTES!$A:$A;0));""))` |
| E11 | Email del cliente | Fórmula | `=SI(B10="";"";SI.ERROR(INDICE(CLIENTES!$D:$D;COINCIDIR(B10;CLIENTES!$A:$A;0));""))` |
| A14:A33 | Número de ítem (1 a 20) | Valor fijo | — |
| B14:B33 | SKU del producto (lista de PRODUCTOS) | Entrada | — |
| C14 | Producto (copiar hasta C33) | Fórmula | `=SI(B14="";"";SI.ERROR(INDICE(PRODUCTOS!$B:$B;COINCIDIR(B14;PRODUCTOS!$A:$A;0));"SKU NO EXISTE"))` |
| D14:D33 | Cantidad | Entrada | — |
| E14 | Precio unitario sin IVA (copiar hasta fila 33) | Fórmula | `=SI(B14="";"";SI.ERROR(INDICE(PRODUCTOS!$F:$F;COINCIDIR(B14;PRODUCTOS!$A:$A;0));0))` |
| F14 | Subtotal (copiar hasta fila 33) | Fórmula | `=SI(B14="";"";D14*E14)` |
| G14 | IVA del ítem (copiar hasta fila 33) | Fórmula | `=SI(B14="";"";SI(SI.ERROR(INDICE(PRODUCTOS!$I:$I;COINCIDIR(B14;PRODUCTOS!$A:$A;0));"NO")="SI";REDONDEAR(F14*CONFIG!$B$12;0);0))` |
| H14 | Total del ítem (copiar hasta fila 33) | Fórmula | `=SI(B14="";"";F14+G14)` |
| I14 | Stock disponible (ayuda, no se imprime) (copiar hasta fila 33) | Fórmula | `=SI(B14="";"";SI.ERROR(INDICE(INVENTARIO!$G:$G;COINCIDIR(B14;INVENTARIO!$A:$A;0));""))` |
| A36 | Observaciones (celda combinada A36:E38) | Entrada | — |
| H35 | Subtotal | Fórmula | `=SUMA(F14:F33)` |
| H36 | IVA | Fórmula | `=SUMA(G14:G33)` |
| H37 | TOTAL A PAGAR | Fórmula | `=SUMA(H14:H33)` |
| H38 | Total unidades | Fórmula | `=SUMA(D14:D33)` |
| A40 | Estado de pago impreso | Fórmula | `=SI(D7="Crédito";"ESTADO DE PAGO: CRÉDITO — pendiente de pago hasta la fecha de vencimiento";"ESTADO DE PAGO: CONTADO — PAGADA")` |
| A44 | Mensaje al pie | Fórmula | `=CONFIG!B10` |
| K3 | Estado del documento (BORRADOR / GUARDADA) — lo escribe la macro | Lo escribe la macro | — |

### GUIAS

Misma estructura visual que la factura (encabezado A1:H5, bloque de datos, tabla de ítems de 20 filas, totales, firmas). Área de impresión **A1:H44**. Se llena a mano o con el botón CARGAR DESDE FACTURA.

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| A1 | Nombre del negocio (remitente) | Fórmula | `=CONFIG!B3` |
| A2 | NIT | Fórmula | `="Remitente — NIT: "&CONFIG!B4` |
| H2 | Número de guía visible | Fórmula | `=CONFIG!$B$15&H3` |
| A3 | Dirección remitente | Fórmula | `=CONFIG!B5&" - "&CONFIG!B6` |
| H3 | Consecutivo numérico de guía (lo escribe la macro NUEVA GUÍA) | Lo escribe la macro | — |
| A4 | Teléfono remitente | Fórmula | `="Tel: "&CONFIG!B7&"   |   "&CONFIG!B8` |
| B7 | Fecha de despacho | Entrada | — |
| D7 | Factura asociada (ej. FE-1006) | Entrada | — |
| G7 | Transportadora (lista CONFIG) | Entrada | — |
| B8 | Valor declarado para el seguro | Entrada | — |
| G8 | Número de rastreo de la transportadora | Entrada | — |
| B11 | Nombre destinatario | Entrada | — |
| E11 | NIT/CC destinatario | Entrada | — |
| H11 | Teléfono destinatario | Entrada | — |
| B12 | Dirección destinatario | Entrada | — |
| E12 | Ciudad / municipio destinatario | Entrada | — |
| H12 | Departamento destinatario | Entrada | — |
| B15:B34 | SKU (la macro CARGAR DESDE FACTURA lo llena) | Entrada | — |
| C15 | Producto (copiar hasta C34) | Fórmula | `=SI(B15="";"";SI.ERROR(INDICE(PRODUCTOS!$B:$B;COINCIDIR(B15;PRODUCTOS!$A:$A;0));"SKU NO EXISTE"))` |
| D15:D34 | Cantidad despachada | Entrada | — |
| E15 | Unidad (copiar hasta E34) | Fórmula | `=SI(B15="";"";SI.ERROR(INDICE(PRODUCTOS!$D:$D;COINCIDIR(B15;PRODUCTOS!$A:$A;0));""))` |
| F15:H34 | Bultos, peso y observación por ítem | Entrada | — |
| D35 | Total unidades | Fórmula | `=SUMA(D15:D34)` |
| F35 | Total bultos | Fórmula | `=SUMA(F15:F34)` |
| G35 | Peso total (kg) | Fórmula | `=SUMA(G15:G34)` |
| A38 | Observaciones (celda combinada A38:H40) | Entrada | — |
| K3 | Estado (BORRADOR / GUARDADA) — lo escribe la macro | Lo escribe la macro | — |

### COSTOS

Calculadora de precios por producto. Encabezados en fila 5, datos de la fila 6 a la 105.

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| A6 | SKU | Entrada | — |
| B6 | Producto | Fórmula | `=SI(A6="";"";SI.ERROR(INDICE(PRODUCTOS!$B:$B;COINCIDIR(A6;PRODUCTOS!$A:$A;0));"SKU NO EXISTE"))` |
| C6 | Costo directo unitario | Fórmula | `=SI(A6="";"";SI.ERROR(INDICE(INVENTARIO!$J:$J;COINCIDIR(A6;INVENTARIO!$A:$A;0));0))` |
| D6 | Empaque $ | Entrada | — |
| E6 | Envío / logística $ | Entrada | — |
| F6 | Otros costos variables $ | Entrada | — |
| G6 | Comisión % (pasarela, marketplace, vendedor) | Entrada | — |
| H6 | Margen deseado % | Entrada | — |
| I6 | Costo variable total $ | Fórmula | `=SI(A6="";"";C6+D6+E6+F6)` |
| J6 | Precio sugerido (sin IVA) | Fórmula | `=SI(A6="";"";SI(G6+H6>=1;"REVISAR %";I6/(1-G6-H6)))` |
| K6 | Precio sugerido (con IVA) | Fórmula | `=SI(A6="";"";SI(ESNUMERO(J6);J6*(1+SI(SI.ERROR(INDICE(PRODUCTOS!$I:$I;COINCIDIR(A6;PRODUCTOS!$A:$A;0));"NO")="SI";CONFIG!$B$12;0));""))` |
| L6 | Precio actual de venta (catálogo) | Fórmula | `=SI(A6="";"";SI.ERROR(INDICE(PRODUCTOS!$F:$F;COINCIDIR(A6;PRODUCTOS!$A:$A;0));0))` |
| M6 | Utilidad unitaria $ | Fórmula | `=SI(A6="";"";L6-I6-L6*G6)` |
| N6 | Utilidad unitaria % | Fórmula | `=SI(A6="";"";SI.ERROR(M6/L6;0))` |
| O6 | ¿Es rentable? | Fórmula | `=SI(A6="";"";SI(N6>=CONFIG!$B$19;"SÍ";SI(N6>=CONFIG!$B$20;"REVISAR";"NO")))` |
| P6 | Precio actual − sugerido | Fórmula | `=SI(A6="";"";SI(ESNUMERO(J6);L6-J6;""))` |

**Lógica del precio sugerido:** si el precio debe cubrir el costo variable (C+D+E+F), pagar la comisión (G % del precio) y dejar el margen deseado (H % del precio), entonces `Precio = Costo variable / (1 − Comisión% − Margen%)`. **¿Es rentable?** compara la utilidad % *al precio actual* con los umbrales de CONFIG: ≥ B19 → **SÍ**; entre B20 y B19 → **REVISAR**; menor que B20 → **NO**.

### INGRESOS_EGRESOS

Libro de caja. Encabezados en fila 4, datos de la fila 5 a la 3004. Resumen mensual en L4:P17. **Registre en orden cronológico** para que el saldo acumulado sea correcto.

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| A5 | Fecha | Entrada | — |
| B5 | Categoría | Entrada | — |
| C5 | Tipo | Fórmula | `=SI(B5="";"";SI.ERROR(BUSCARV(B5;CONFIG!$D$4:$E$30;2;0);"REVISAR"))` |
| D5 | Descripción / Concepto | Entrada | — |
| E5 | Documento ref. | Entrada | — |
| F5 | Tercero (cliente / proveedor) | Entrada | — |
| G5 | Valor | Entrada | — |
| H5 | Ingreso | Fórmula | `=SI(A5="";"";SI(C5="Ingreso";G5;0))` |
| I5 | Egreso | Fórmula | `=SI(A5="";"";SI(C5="Egreso";G5;0))` |
| J5 | Saldo acumulado | Fórmula | `=SI(A5="";"";CONFIG!$B$23+SUMA($H$5:H5)-SUMA($I$5:I5))` |
| L5 | Mes (enero del año de análisis; L6..L16 cambian el 1 por 2..12) | Fórmula | `=FECHA(AÑO(CONFIG!$B$24);1;1)` |
| M5 | Ingresos del mes | Fórmula | `=SUMAR.SI.CONJUNTO($H:$H;$A:$A;">="&L5;$A:$A;"<="&FIN.MES(L5;0))` |
| N5 | Egresos del mes | Fórmula | `=SUMAR.SI.CONJUNTO($I:$I;$A:$A;">="&L5;$A:$A;"<="&FIN.MES(L5;0))` |
| O5 | Flujo neto | Fórmula | `=M5-N5` |
| P5 | Saldo final de caja | Fórmula | `=CONFIG!$B$23+SUMAR.SI.CONJUNTO($H:$H;$A:$A;"<="&FIN.MES(L5;0))-SUMAR.SI.CONJUNTO($I:$I;$A:$A;"<="&FIN.MES(L5;0))` |

### CARTERA

Columnas A:F las escribe la macro GUARDAR cuando la factura es a Crédito; G:K son fórmulas (precargadas hasta la fila 1004). Resumen por estado en M4:O9.

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| A5 | No. factura | Lo escribe la macro | — |
| B5 | NIT/CC | Lo escribe la macro | — |
| C5 | Cliente | Lo escribe la macro | — |
| D5 | Fecha emisión | Lo escribe la macro | — |
| E5 | Fecha vencimiento | Lo escribe la macro | — |
| F5 | Valor total | Lo escribe la macro | — |
| G5 | Abonos realizados | Fórmula | `=SI(A5="";"";SUMAR.SI.CONJUNTO(INGRESOS_EGRESOS!$H:$H;INGRESOS_EGRESOS!$B:$B;"Abono cartera";INGRESOS_EGRESOS!$E:$E;A5))` |
| H5 | Saldo pendiente | Fórmula | `=SI(A5="";"";F5-G5)` |
| I5 | Días vencido | Fórmula | `=SI(A5="";"";SI(H5<=0;0;MAX(0;HOY()-E5)))` |
| J5 | Estado | Fórmula | `=SI(A5="";"";SI(H5<=0;"Pagado";SI(HOY()>E5;"Vencido";SI(E5-HOY()<=CONFIG!$B$22;"Por vencer";"Al día"))))` |
| K5 | Edad de cartera | Fórmula | `=SI(A5="";"";SI(H5<=0;"Pagada";SI(I5=0;"Corriente";SI(I5<=30;"1-30 días";SI(I5<=60;"31-60 días";SI(I5<=90;"61-90 días";"Más de 90"))))))` |
| N5 | Resumen: saldo por estado (N8 'Pagado' usa la columna F) | Fórmula | `=SUMAR.SI.CONJUNTO($H$5:$H$1004;$J$5:$J$1004;M5)` |

### GASTOS_FIJOS

Filas 5 a 24 para 20 conceptos; totales en fila 25; impacto en el punto de equilibrio en filas 28-32.

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| A5 | Concepto (inicio de la descripción en el flujo de caja) | Entrada | — |
| B5 | Categoría | Entrada | — |
| C5 | Día de pago | Entrada | — |
| D5 | Presupuesto mensual | Entrada | — |
| E5 | Real ejecutado (mes análisis) | Fórmula | `=SI(A5="";"";SUMAR.SI.CONJUNTO(INGRESOS_EGRESOS!$I:$I;INGRESOS_EGRESOS!$B:$B;"Gasto fijo";INGRESOS_EGRESOS!$D:$D;A5&"*";INGRESOS_EGRESOS!$A:$A;">="&CONFIG!$B$24;INGRESOS_EGRESOS!$A:$A;"<="&CONFIG!$B$25))` |
| F5 | Diferencia (Ppto − Real) | Fórmula | `=SI(A5="";"";D5-E5)` |
| G5 | % ejecución | Fórmula | `=SI(A5="";"";SI.ERROR(E5/D5;0))` |
| H5 | Estado | Fórmula | `=SI(A5="";"";SI(E5=0;"PENDIENTE";SI(E5>D5;"SOBRE PPTO";"OK")))` |
| I5 | % del total de gastos fijos | Fórmula | `=SI(A5="";"";SI.ERROR(D5/$D$25;0))` |
| D25 | Total presupuesto (E25 y F25 igual con su columna) | Fórmula | `=SUMA(D5:D24)` |
| D28 | Margen de contribución promedio | Fórmula | `=PUNTO_EQUILIBRIO!B10` |
| D29 | Ventas mínimas para cubrir estos gastos fijos (PE) | Fórmula | `=PUNTO_EQUILIBRIO!B11` |
| D30 | Ventas netas del mes | Fórmula | `=PUNTO_EQUILIBRIO!B7` |
| D31 | Gastos fijos como % de las ventas del mes | Fórmula | `=SI.ERROR(MAX(D25;E25)/PUNTO_EQUILIBRIO!B7;0)` |
| D32 | Cada $100.000 de gasto fijo adicional exige vender | Fórmula | `=SI.ERROR(100000/PUNTO_EQUILIBRIO!B10;0)` |

### PUNTO_EQUILIBRIO

Columna A = concepto, columna B = cálculo. Tabla de sensibilidad en D4:F9.

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| B4 | Gastos fijos presupuestados | Fórmula | `=GASTOS_FIJOS!D25` |
| B5 | Gastos fijos reales del mes | Fórmula | `=GASTOS_FIJOS!E25` |
| B6 | Gastos fijos usados en el cálculo (el mayor) | Fórmula | `=MAX(B4;B5)` |
| B7 | Ventas netas del mes (sin IVA) | Fórmula | `=SUMAR.SI.CONJUNTO(HIST_FACTURAS!$G:$G;HIST_FACTURAS!$C:$C;">="&CONFIG!$B$24;HIST_FACTURAS!$C:$C;"<="&CONFIG!$B$25)` |
| B8 | Costo de ventas del mes | Fórmula | `=SUMAR.SI.CONJUNTO(HIST_FACTURAS!$L:$L;HIST_FACTURAS!$C:$C;">="&CONFIG!$B$24;HIST_FACTURAS!$C:$C;"<="&CONFIG!$B$25)` |
| B9 | Costos variables adicionales (% sobre ventas) | Fórmula | `=CONFIG!B26` |
| B10 | Margen de contribución % | Fórmula | `=SI(B7>0;(B7-B8)/B7-B9;SI.ERROR(PROMEDIO(COSTOS!$N$6:$N$105);0))` |
| B11 | PUNTO DE EQUILIBRIO EN VENTAS ($ sin IVA) | Fórmula | `=SI(B10<=0;0;B6/B10)` |
| B12 | Ventas diarias necesarias | Fórmula | `=SI.ERROR(B11/DIA(CONFIG!B25);0)` |
| B13 | Ticket promedio por factura | Fórmula | `=SI.ERROR(B7/CONTAR.SI.CONJUNTO(HIST_FACTURAS!$C:$C;">="&CONFIG!$B$24;HIST_FACTURAS!$C:$C;"<="&CONFIG!$B$25);0)` |
| B14 | Facturas necesarias en el mes | Fórmula | `=SI.ERROR(REDONDEAR.MAS(B11/B13;0);0)` |
| B15 | % de cumplimiento del punto de equilibrio | Fórmula | `=SI.ERROR(B7/B11;0)` |
| B16 | Ventas que faltan para llegar al PE | Fórmula | `=MAX(0;B11-B7)` |
| B17 | Margen de seguridad | Fórmula | `=SI.ERROR((B7-B11)/B7;0)` |
| B18 | Diagnóstico | Fórmula | `=SI(B11=0;"Sin datos suficientes";SI(B7>=B11;"Sobre el punto de equilibrio: el mes genera utilidad";"Por debajo del punto de equilibrio: faltan ventas"))` |
| E5 | Sensibilidad: ventas | Fórmula | `=$B$11*D5` |
| F5 | Sensibilidad: utilidad operativa | Fórmula | `=E5*$B$10-$B$6` |

### DASHBOARD

Resumen ejecutivo del **mes de análisis** (CONFIG!B24). Fila 4-6 y 8-10: 8 tarjetas KPI (título, valor grande, indicador secundario). Filas 12-44: 3 gráficos + tabla de alertas. Columnas R:U: datos de soporte de los gráficos y cálculo de la utilidad neta (no editar).

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| F2 | Mes de análisis | Fórmula | `=CONFIG!B24` |
| B5 | VENTAS DEL MES (sin IVA) | Fórmula | `=SUMAR.SI.CONJUNTO(HIST_FACTURAS!$G:$G;HIST_FACTURAS!$C:$C;">="&CONFIG!$B$24;HIST_FACTURAS!$C:$C;"<="&CONFIG!$B$25)` |
| D6 | VENTAS DEL MES (sin IVA) — Facturas emitidas: | Fórmula | `=CONTAR.SI.CONJUNTO(HIST_FACTURAS!$C:$C;">="&CONFIG!$B$24;HIST_FACTURAS!$C:$C;"<="&CONFIG!$B$25)` |
| F5 | INGRESOS DEL MES (caja) | Fórmula | `=SUMAR.SI.CONJUNTO(INGRESOS_EGRESOS!$H:$H;INGRESOS_EGRESOS!$A:$A;">="&CONFIG!$B$24;INGRESOS_EGRESOS!$A:$A;"<="&CONFIG!$B$25)` |
| H6 | INGRESOS DEL MES (caja) — Flujo neto: | Fórmula | `=F5-J5` |
| J5 | EGRESOS DEL MES (caja) | Fórmula | `=SUMAR.SI.CONJUNTO(INGRESOS_EGRESOS!$I:$I;INGRESOS_EGRESOS!$A:$A;">="&CONFIG!$B$24;INGRESOS_EGRESOS!$A:$A;"<="&CONFIG!$B$25)` |
| L6 | EGRESOS DEL MES (caja) — Cobertura (Ingr./Egr.): | Fórmula | `=SI.ERROR(F5/J5;0)` |
| N5 | UTILIDAD NETA DEL MES | Fórmula | `=S41` |
| P6 | UTILIDAD NETA DEL MES — Margen neto: | Fórmula | `=SI.ERROR(N5/B5;0)` |
| B9 | VALOR DEL INVENTARIO | Fórmula | `=SUMA(INVENTARIO!$K$5:$K$504)` |
| D10 | VALOR DEL INVENTARIO — Productos a reponer: | Fórmula | `=CONTAR.SI(INVENTARIO!$M$5:$M$504;"REPONER")+CONTAR.SI(INVENTARIO!$M$5:$M$504;"AGOTADO")` |
| F9 | CARTERA POR COBRAR | Fórmula | `=SUMA(CARTERA!$H$5:$H$1004)` |
| H10 | CARTERA POR COBRAR — Cartera vencida: | Fórmula | `=SUMAR.SI.CONJUNTO(CARTERA!$H:$H;CARTERA!$J:$J;"Vencido")` |
| J9 | GASTOS FIJOS DEL MES (real) | Fórmula | `=GASTOS_FIJOS!E25` |
| L10 | GASTOS FIJOS DEL MES (real) — Presupuesto: | Fórmula | `=GASTOS_FIJOS!D25` |
| N9 | CUMPLIMIENTO PUNTO DE EQUILIBRIO | Fórmula | `=PUNTO_EQUILIBRIO!B15` |
| P10 | CUMPLIMIENTO PUNTO DE EQUILIBRIO — PE en ventas: | Fórmula | `=PUNTO_EQUILIBRIO!B11` |
| R5:U16 | Serie mensual Ingresos/Egresos (de INGRESOS_EGRESOS L:O) | Fórmula | `=ELEGIR(MES(INGRESOS_EGRESOS!L5);"Ene";"Feb";"Mar";"Abr";"May";"Jun";"Jul";"Ago";"Sep";"Oct";"Nov";"Dic")` |
| R21 | Top 1: nombre del producto (k = 1..5 en filas 21..25) | Fórmula | `=SI.ERROR(SI(K.ESIMO.MAYOR(INVENTARIO!$P$5:$P$504;1)<1;"";INDICE(INVENTARIO!$B$5:$B$504;COINCIDIR(K.ESIMO.MAYOR(INVENTARIO!$P$5:$P$504;1);INVENTARIO!$P$5:$P$504;0)));"")` |
| S21 | Top 1: unidades | Fórmula | `=SI(R21="";0;INDICE(INVENTARIO!$N$5:$N$504;COINCIDIR(K.ESIMO.MAYOR(INVENTARIO!$P$5:$P$504;1);INVENTARIO!$P$5:$P$504;0)))` |
| T21 | Top 1: ventas $ | Fórmula | `=SI(R21="";0;INDICE(INVENTARIO!$O$5:$O$504;COINCIDIR(K.ESIMO.MAYOR(INVENTARIO!$P$5:$P$504;1);INVENTARIO!$P$5:$P$504;0)))` |
| S30 | Saldo de cartera por estado | Fórmula | `=SUMAR.SI.CONJUNTO(CARTERA!$H:$H;CARTERA!$J:$J;R30)` |
| S37 | Ventas netas (sin IVA) | Fórmula | `=B5` |
| S38 | (−) Costo de ventas | Fórmula | `=SUMAR.SI.CONJUNTO(HIST_FACTURAS!$L:$L;HIST_FACTURAS!$C:$C;">="&CONFIG!$B$24;HIST_FACTURAS!$C:$C;"<="&CONFIG!$B$25)` |
| S39 | (−) Gastos operativos (egresos del mes) | Fórmula | `=J5` |
| S40 | (+) Excluir compras, pagos a proveedores e IVA | Fórmula | `=SUMAR.SI.CONJUNTO(INGRESOS_EGRESOS!$I:$I;INGRESOS_EGRESOS!$B:$B;"Compra inventario";INGRESOS_EGRESOS!$A:$A;">="&CONFIG!$B$24;INGRESOS_EGRESOS!$A:$A;"<="&CONFIG!$B$25)+SUMAR.SI.CONJUNTO(INGRESOS_EGRESOS!$I:$I;INGRESOS_EGRESOS!$B:$B;"Pago proveedores";INGRESOS_EGRESOS!$A:$A;">="&CONFIG!$B$24;INGRESOS_EGRESOS!$A:$A;"<="&CONFIG!$B$25)+SUMAR.SI.CONJUNTO(INGRESOS_EGRESOS!$I:$I;INGRESOS_EGRESOS!$B:$B;"Pago IVA (DIAN)";INGRESOS_EGRESOS!$A:$A;">="&CONFIG!$B$24;INGRESOS_EGRESOS!$A:$A;"<="&CONFIG!$B$25)` |
| S41 | UTILIDAD NETA DEL MES | Fórmula | `=S37-S38-S39+S40` |
| O29 | Productos agotados | Fórmula | `=CONTAR.SI(INVENTARIO!$M$5:$M$504;"AGOTADO")` |
| O30 | Productos por reponer | Fórmula | `=CONTAR.SI(INVENTARIO!$M$5:$M$504;"REPONER")` |
| O31 | Facturas vencidas | Fórmula | `=CONTAR.SI(CARTERA!$J$5:$J$1004;"Vencido")` |
| O32 | Facturas por vencer | Fórmula | `=CONTAR.SI(CARTERA!$J$5:$J$1004;"Por vencer")` |
| O33 | Productos NO rentables / a revisar | Fórmula | `=CONTAR.SI(COSTOS!$O$6:$O$105;"NO")+CONTAR.SI(COSTOS!$O$6:$O$105;"REVISAR")` |
| O34 | Gastos fijos sobre presupuesto | Fórmula | `=CONTAR.SI(GASTOS_FIJOS!$H$5:$H$24;"SOBRE PPTO")` |
| O35 | Saldo por pagar a proveedores | Fórmula | `=SUMA(PROVEEDORES!$K$5:$K$304)` |
| O36 | Saldo actual de caja | Fórmula | `=CONFIG!B23+SUMA(INGRESOS_EGRESOS!$H$5:$H$3004)-SUMA(INGRESOS_EGRESOS!$I$5:$I$3004)` |

**Gráficos** (ya creados en el archivo):
1. *Ingresos vs. Egresos por mes*: columnas agrupadas con datos de R4:T16 (12 meses del año de análisis).
2. *Top 5 productos más vendidos*: barras horizontales con R20:S25. El ranking usa `K.ESIMO.MAYOR` sobre la columna auxiliar oculta INVENTARIO!P (unidades + FILA()/1.000.000 para desempatar).
3. *Estado de la cartera*: columnas con R29:S32 (Al día, Por vencer, Vencido).

Para recrear un gráfico: seleccione el rango de datos → **Insertar → Gráfico de columnas o barras**.

**Utilidad neta del mes (S41)** = Ventas sin IVA − Costo de ventas − Egresos del mes + (Compras de inventario + Pagos a proveedores + Pago de IVA). Esas tres partidas se excluyen de los gastos porque el costo de la mercancía ya entra como *costo de ventas* y el IVA no es un gasto.

### HIST_FACTURAS

Una fila por factura guardada (la escribe la macro). Encabezados en fila 4.

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| A5 | Consecutivo | Lo escribe la macro | — |
| B5 | No. factura | Lo escribe la macro | — |
| C5 | Fecha | Lo escribe la macro | — |
| D5 | NIT/CC cliente | Lo escribe la macro | — |
| E5 | Cliente | Lo escribe la macro | — |
| F5 | Ciudad | Lo escribe la macro | — |
| G5 | Subtotal (sin IVA) | Lo escribe la macro | — |
| H5 | IVA | Lo escribe la macro | — |
| I5 | Total | Lo escribe la macro | — |
| J5 | Forma de pago | Lo escribe la macro | — |
| K5 | Vencimiento | Lo escribe la macro | — |
| L5 | Costo de ventas | Lo escribe la macro | — |
| M5 | Utilidad bruta | Lo escribe la macro | — |
| N5 | Archivo guardado | Lo escribe la macro | — |
| O5 | Registrado el | Lo escribe la macro | — |

### HIST_DETALLE

Una fila por ítem facturado (la escribe la macro). Es la fuente de las SALIDAS del inventario.

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| A5 | No. factura | Lo escribe la macro | — |
| B5 | Fecha | Lo escribe la macro | — |
| C5 | SKU | Lo escribe la macro | — |
| D5 | Producto | Lo escribe la macro | — |
| E5 | Cantidad | Lo escribe la macro | — |
| F5 | Precio unit. | Lo escribe la macro | — |
| G5 | Subtotal | Lo escribe la macro | — |
| H5 | IVA | Lo escribe la macro | — |
| I5 | Total | Lo escribe la macro | — |
| J5 | Costo unit. | Lo escribe la macro | — |
| K5 | Costo total | Lo escribe la macro | — |
| L5 | Utilidad bruta | Lo escribe la macro | — |
| M5 | NIT/CC cliente | Lo escribe la macro | — |
| N5 | Consecutivo | Lo escribe la macro | — |

### HIST_GUIAS

Una fila por guía guardada (la escribe la macro).

| Celda / Col. | Contenido | Tipo | Fórmula (español, separador ;) |
|---|---|---|---|
| A5 | Consecutivo | Lo escribe la macro | — |
| B5 | No. guía | Lo escribe la macro | — |
| C5 | Fecha | Lo escribe la macro | — |
| D5 | Factura asociada | Lo escribe la macro | — |
| E5 | Destinatario | Lo escribe la macro | — |
| F5 | Ciudad | Lo escribe la macro | — |
| G5 | Transportadora | Lo escribe la macro | — |
| H5 | No. rastreo | Lo escribe la macro | — |
| I5 | Bultos | Lo escribe la macro | — |
| J5 | Archivo guardado | Lo escribe la macro | — |
| K5 | Registrado el | Lo escribe la macro | — |

---

## 5. Hojas adicionales que agregué y por qué

| Hoja | Por qué es clave para el control real |
|---|---|
| **COMPRAS** | Sin un registro de entradas, el inventario solo puede bajar. Aquí entra la mercancía con su costo real, que alimenta el **costo promedio ponderado** y, por lo tanto, la utilidad bruta de cada factura. |
| **PROVEEDORES** | Da las **cuentas por pagar** (compras a crédito − pagos). Sin esta hoja la caja parece sana mientras se acumulan deudas con proveedores. |
| **PUNTO_EQUILIBRIO** | Traduce los gastos fijos a una **meta de ventas mensual y diaria** y al número de facturas necesarias. Incluye una tabla de sensibilidad. Es el indicador que más ayuda a decidir precios y gastos. |
| **HIST_DETALLE** | Guarda cada producto vendido con su **costo en el momento de la venta**. Permite descontar inventario, rankear productos y calcular la utilidad real sin depender de precios que cambian. |
| **HIST_GUIAS** | Trazabilidad de despachos: qué factura salió, con qué transportadora y con qué número de rastreo. |
| **CONFIG** | Un solo lugar para los parámetros (IVA, prefijos, umbrales, carpeta, mes de análisis). Cambia el comportamiento del sistema sin tocar fórmulas ni código. |
| **INSTRUCCIONES** | Manual de uso dentro del archivo para cualquier persona del equipo. |

**Siguientes mejoras sugeridas** (cuando el negocio crezca): una hoja de **Nómina** detallada, **Cotizaciones** que se conviertan en factura, **Presupuesto anual** por mes, y conexión con la **facturación electrónica DIAN**. Este formato es un soporte interno: si usted es responsable de facturar electrónicamente, la factura legal debe emitirse en su proveedor tecnológico autorizado.

---

## 6. Código VBA completo (`modGestion.bas`)

```vba

Attribute VB_Name = "modGestion"
'==============================================================================
'  SISTEMA INTEGRAL DE GESTIÓN EMPRESARIAL - MÓDULO DE MACROS
'------------------------------------------------------------------------------
'  Botones de FACTURACION : ImprimirFactura, LimpiarFactura, GuardarFactura,
'                           NuevaFactura
'  Botones de GUIAS       : CargarFacturaEnGuia, ImprimirGuia, LimpiarGuia,
'                           GuardarGuia, NuevaGuia
'  Utilidades             : CrearBotones (crea y asigna todos los botones),
'                           ProtegerHojas / DesprotegerHojas
'
'  IMPORTANTE: las direcciones de celda de este módulo coinciden EXACTAMENTE
'  con la plantilla Sistema_Gestion_Empresarial.xlsx. Si mueve filas o
'  columnas en FACTURACION o GUIAS, actualice las constantes de abajo.
'  Las fórmulas se escriben con .Formula (en inglés y con coma), que Excel
'  traduce solo al idioma de su instalación.
'==============================================================================
Option Explicit

'----------------------------- NOMBRES DE HOJAS -------------------------------
Private Const H_FAC As String = "FACTURACION"
Private Const H_GUIA As String = "GUIAS"
Private Const H_CFG As String = "CONFIG"
Private Const H_HFAC As String = "HIST_FACTURAS"
Private Const H_HDET As String = "HIST_DETALLE"
Private Const H_HGUIA As String = "HIST_GUIAS"
Private Const H_CART As String = "CARTERA"
Private Const H_CAJA As String = "INGRESOS_EGRESOS"
Private Const H_INV As String = "INVENTARIO"
Private Const H_CLI As String = "CLIENTES"

Private Const FILA_DATOS As Long = 5          'primera fila de datos en tablas

'----------------------------- FACTURA ----------------------------------------
Private Const FAC_AREA As String = "$A$1:$H$44"   'área de impresión
Private Const FAC_NUM As String = "H3"            'consecutivo numérico
Private Const FAC_NUMTXT As String = "H2"         'prefijo + consecutivo
Private Const FAC_FECHA As String = "B7"
Private Const FAC_PAGO As String = "D7"
Private Const FAC_VENCE As String = "G7"
Private Const FAC_CLI As String = "B10"
Private Const FAC_CLINOM As String = "C10"
Private Const FAC_CIUDAD As String = "E10"
Private Const FAC_OBS As String = "A36"
Private Const FAC_SUB As String = "H35"
Private Const FAC_IVA As String = "H36"
Private Const FAC_TOT As String = "H37"
Private Const FAC_ESTADO As String = "K3"
Private Const FAC_F1 As Long = 14                 'primera fila de ítems
Private Const FAC_F2 As Long = 33                 'última fila de ítems

'----------------------------- GUÍA -------------------------------------------
Private Const GUI_AREA As String = "$A$1:$H$44"
Private Const GUI_NUM As String = "H3"
Private Const GUI_NUMTXT As String = "H2"
Private Const GUI_FECHA As String = "B7"
Private Const GUI_FACT As String = "D7"
Private Const GUI_TRANSP As String = "G7"
Private Const GUI_VALOR As String = "B8"
Private Const GUI_RASTREO As String = "G8"
Private Const GUI_NOMBRE As String = "B11"
Private Const GUI_ID As String = "E11"
Private Const GUI_TEL As String = "H11"
Private Const GUI_DIR As String = "B12"
Private Const GUI_CIUDAD As String = "E12"
Private Const GUI_DEPTO As String = "H12"
Private Const GUI_OBS As String = "A38"
Private Const GUI_BULTOS As String = "F35"
Private Const GUI_ESTADO As String = "K3"
Private Const GUI_F1 As Long = 15
Private Const GUI_F2 As Long = 34

'----------------------------- PROTECCIÓN -------------------------------------
Private Const CLAVE As String = ""        'escriba una clave si desea, ej. "1234"

'==============================================================================
'                               FACTURACIÓN
'==============================================================================

'--- BOTÓN: IMPRIMIR SOLO FACTURA ---------------------------------------------
Public Sub ImprimirFactura()
    Dim ws As Worksheet, resp As VbMsgBoxResult
    Set ws = ThisWorkbook.Worksheets(H_FAC)

    If Not TieneItems(ws, FAC_F1, FAC_F2) Then
        MsgBox "La factura no tiene productos.", vbExclamation, "Imprimir factura"
        Exit Sub
    End If

    ConfigurarPagina ws, FAC_AREA

    resp = MsgBox("¿Cómo desea continuar?" & vbCrLf & vbCrLf & _
                  "SÍ  = Imprimir directamente" & vbCrLf & _
                  "NO  = Ver vista previa" & vbCrLf & _
                  "CANCELAR = Salir", vbYesNoCancel + vbQuestion, "Imprimir factura " & ws.Range(FAC_NUMTXT).Value)
    Select Case resp
        Case vbYes: ws.PrintOut Copies:=1, Collate:=True
        Case vbNo:  ws.PrintPreview
    End Select
End Sub

'--- BOTÓN: LIMPIAR FACTURA ---------------------------------------------------
Public Sub LimpiarFactura()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(H_FAC)
    If MsgBox("¿Borrar los datos de la factura actual?" & vbCrLf & _
              "(No se borran fórmulas, encabezado ni el número de factura)", _
              vbYesNo + vbQuestion, "Limpiar factura") = vbNo Then Exit Sub
    LimpiarFacturaInterno ws
End Sub

'--- BOTÓN: NUEVA FACTURA -----------------------------------------------------
Public Sub NuevaFactura()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(H_FAC)

    If ws.Range(FAC_ESTADO).Value <> "GUARDADA" And TieneItems(ws, FAC_F1, FAC_F2) Then
        If MsgBox("La factura actual NO ha sido guardada." & vbCrLf & _
                  "¿Desea descartarla y crear una nueva?", vbYesNo + vbExclamation, _
                  "Nueva factura") = vbNo Then Exit Sub
    End If

    LimpiarFacturaInterno ws
    ws.Range(FAC_NUM).Value = SiguienteConsecutivo(H_HFAC, ThisWorkbook.Worksheets(H_CFG).Range("B14").Value)
    ws.Activate
    ws.Range(FAC_CLI).Select
End Sub

'--- BOTÓN: GUARDAR FACTURA (historial + cartera/caja + PDF/Excel) ------------
Public Sub GuardarFactura()
    Dim ws As Worksheet, wsH As Worksheet, wsD As Worksheet, wsCfg As Worksheet
    Dim wsInv As Worksheet, wsCart As Worksheet, wsCaja As Worksheet
    Dim r As Long, fh As Long, fd As Long, n As Long
    Dim numero As Long, numTxt As String, fecha As Date, cliente As String
    Dim sku As String, cant As Double, costoU As Double, costoTot As Double
    Dim m As Variant, avisos As String, carpeta As String, archivo As String
    Dim formato As String, rutas As String

    On Error GoTo Fallo
    Set ws = ThisWorkbook.Worksheets(H_FAC)
    Set wsH = ThisWorkbook.Worksheets(H_HFAC)
    Set wsD = ThisWorkbook.Worksheets(H_HDET)
    Set wsCfg = ThisWorkbook.Worksheets(H_CFG)
    Set wsInv = ThisWorkbook.Worksheets(H_INV)
    Set wsCart = ThisWorkbook.Worksheets(H_CART)
    Set wsCaja = ThisWorkbook.Worksheets(H_CAJA)
    Application.Calculate

    '---------- 1. VALIDACIONES ----------
    numero = Val(ws.Range(FAC_NUM).Value)
    numTxt = CStr(ws.Range(FAC_NUMTXT).Value)
    If numero <= 0 Then Err.Raise vbObjectError + 1, , "No hay número de factura. Use NUEVA FACTURA."
    If Not IsError(Application.Match(numero, wsH.Columns(1), 0)) Then
        Err.Raise vbObjectError + 2, , "La factura " & numTxt & " YA fue guardada." & vbCrLf & _
                  "Use NUEVA FACTURA para generar el siguiente consecutivo."
    End If
    If Not IsDate(ws.Range(FAC_FECHA).Value) Then Err.Raise vbObjectError + 3, , "Digite una fecha válida en " & FAC_FECHA & "."
    fecha = CDate(ws.Range(FAC_FECHA).Value)
    cliente = CStr(ws.Range(FAC_CLINOM).Value)
    If Trim$(CStr(ws.Range(FAC_CLI).Value)) = "" Then Err.Raise vbObjectError + 4, , "Seleccione el NIT/CC del cliente."
    If cliente = "" Or cliente = "CLIENTE NO REGISTRADO" Then _
        Err.Raise vbObjectError + 5, , "El cliente no existe en la hoja CLIENTES. Regístrelo primero."
    If Trim$(CStr(ws.Range(FAC_PAGO).Value)) = "" Then Err.Raise vbObjectError + 6, , "Seleccione la forma de pago (Contado / Crédito)."
    If Not TieneItems(ws, FAC_F1, FAC_F2) Then Err.Raise vbObjectError + 7, , "La factura no tiene productos."

    For r = FAC_F1 To FAC_F2
        sku = Trim$(CStr(ws.Cells(r, "B").Value))
        If sku = "" And Not IsEmpty(ws.Cells(r, "D").Value) Then _
            Err.Raise vbObjectError + 8, , "Fila " & r & ": hay cantidad pero falta el SKU."
        If sku <> "" Then
            If ws.Cells(r, "C").Value = "SKU NO EXISTE" Then _
                Err.Raise vbObjectError + 9, , "Fila " & r & ": el SKU " & sku & " no existe en PRODUCTOS."
            If Not IsNumeric(ws.Cells(r, "D").Value) Or Val(ws.Cells(r, "D").Value) <= 0 Then _
                Err.Raise vbObjectError + 10, , "Fila " & r & ": la cantidad debe ser mayor que cero."
            If IsNumeric(ws.Cells(r, "I").Value) Then
                If CDbl(ws.Cells(r, "D").Value) > CDbl(ws.Cells(r, "I").Value) Then _
                    avisos = avisos & "  - " & sku & ": pide " & ws.Cells(r, "D").Value & _
                             ", disponible " & ws.Cells(r, "I").Value & vbCrLf
            End If
        End If
    Next r
    If avisos <> "" Then
        If MsgBox("Stock insuficiente:" & vbCrLf & avisos & vbCrLf & _
                  "¿Desea guardar de todas formas? (el inventario quedará negativo)", _
                  vbYesNo + vbExclamation, "Inventario") = vbNo Then Exit Sub
    End If

    Application.ScreenUpdating = False

    '---------- 2. COPIA EN EL ESCRITORIO (primero: si falla, no se registra nada) ----------
    carpeta = CarpetaDestino("Facturas", fecha)
    archivo = carpeta & "\" & LimpiarNombre(numTxt & "_" & cliente)
    formato = UCase$(Trim$(CStr(wsCfg.Range("B18").Value)))
    If formato = "" Then formato = "PDF"

    ConfigurarPagina ws, FAC_AREA
    If formato = "PDF" Or formato = "AMBOS" Then
        ExportarPDF ws, archivo & ".pdf"
        rutas = archivo & ".pdf"
    End If
    If formato = "EXCEL" Or formato = "AMBOS" Then
        ExportarExcel ws, archivo & ".xlsx", "I:K"
        rutas = rutas & IIf(rutas = "", "", " | ") & archivo & ".xlsx"
    End If

    '---------- 3. DETALLE (HIST_DETALLE) ----------
    fd = SiguienteFila(wsD, 1)
    For r = FAC_F1 To FAC_F2
        sku = Trim$(CStr(ws.Cells(r, "B").Value))
        If sku <> "" Then
            cant = CDbl(ws.Cells(r, "D").Value)
            m = Application.Match(sku, wsInv.Columns(1), 0)
            If IsError(m) Then costoU = 0 Else costoU = Val(wsInv.Cells(m, "J").Value)
            wsD.Cells(fd, 1).Value = numTxt
            wsD.Cells(fd, 2).Value = fecha
            wsD.Cells(fd, 3).NumberFormat = "@"
            wsD.Cells(fd, 3).Value = sku
            wsD.Cells(fd, 4).Value = ws.Cells(r, "C").Value
            wsD.Cells(fd, 5).Value = cant
            wsD.Cells(fd, 6).Value = ws.Cells(r, "E").Value
            wsD.Cells(fd, 7).Value = ws.Cells(r, "F").Value
            wsD.Cells(fd, 8).Value = ws.Cells(r, "G").Value
            wsD.Cells(fd, 9).Value = ws.Cells(r, "H").Value
            wsD.Cells(fd, 10).Value = costoU
            wsD.Cells(fd, 11).Value = costoU * cant
            wsD.Cells(fd, 12).Value = ws.Cells(r, "F").Value - costoU * cant
            wsD.Cells(fd, 13).NumberFormat = "@"
            wsD.Cells(fd, 13).Value = CStr(ws.Range(FAC_CLI).Value)
            wsD.Cells(fd, 14).Value = numero
            wsD.Range(wsD.Cells(fd, 1), wsD.Cells(fd, 14)).Font.Name = "Arial"
            costoTot = costoTot + costoU * cant
            fd = fd + 1
        End If
    Next r
    FormatearFilasHist wsD

    '---------- 4. ENCABEZADO (HIST_FACTURAS) ----------
    fh = SiguienteFila(wsH, 1)
    wsH.Cells(fh, 1).Value = numero
    wsH.Cells(fh, 2).Value = numTxt
    wsH.Cells(fh, 3).Value = fecha
    wsH.Cells(fh, 4).NumberFormat = "@"
    wsH.Cells(fh, 4).Value = CStr(ws.Range(FAC_CLI).Value)
    wsH.Cells(fh, 5).Value = cliente
    wsH.Cells(fh, 6).Value = ws.Range(FAC_CIUDAD).Value
    wsH.Cells(fh, 7).Value = ws.Range(FAC_SUB).Value
    wsH.Cells(fh, 8).Value = ws.Range(FAC_IVA).Value
    wsH.Cells(fh, 9).Value = ws.Range(FAC_TOT).Value
    wsH.Cells(fh, 10).Value = ws.Range(FAC_PAGO).Value
    wsH.Cells(fh, 11).Value = ws.Range(FAC_VENCE).Value
    wsH.Cells(fh, 12).Value = costoTot
    wsH.Cells(fh, 13).Value = ws.Range(FAC_SUB).Value - costoTot
    wsH.Cells(fh, 14).Value = rutas
    wsH.Cells(fh, 15).Value = Now
    FormatearFilasHist wsH

    '---------- 5. CARTERA (crédito) o CAJA (contado) ----------
    If EsCredito(ws.Range(FAC_PAGO).Value) Then
        n = SiguienteFila(wsCart, 1)
        wsCart.Cells(n, 1).Value = numTxt
        wsCart.Cells(n, 2).NumberFormat = "@"
        wsCart.Cells(n, 2).Value = CStr(ws.Range(FAC_CLI).Value)
        wsCart.Cells(n, 3).Value = cliente
        wsCart.Cells(n, 4).Value = fecha
        wsCart.Cells(n, 5).Value = ws.Range(FAC_VENCE).Value
        wsCart.Cells(n, 6).Value = ws.Range(FAC_TOT).Value
        AsegurarFormulasCartera wsCart, n
    Else
        n = SiguienteFila(wsCaja, 1)
        wsCaja.Cells(n, 1).Value = fecha
        wsCaja.Cells(n, 2).Value = "Venta contado"
        wsCaja.Cells(n, 4).Value = "Venta factura " & numTxt
        wsCaja.Cells(n, 5).Value = numTxt
        wsCaja.Cells(n, 6).Value = cliente
        wsCaja.Cells(n, 7).Value = ws.Range(FAC_TOT).Value
        AsegurarFormulasCaja wsCaja, n
    End If

    '---------- 6. CIERRE ----------
    ws.Range(FAC_ESTADO).Value = "GUARDADA"
    Application.ScreenUpdating = True
    ThisWorkbook.Save

    If MsgBox("Factura " & numTxt & " guardada correctamente." & vbCrLf & vbCrLf & _
              "Copia en: " & rutas & vbCrLf & vbCrLf & _
              "¿Desea crear una NUEVA factura?", vbYesNo + vbInformation, "Guardar factura") = vbYes Then
        NuevaFactura
    End If
    Exit Sub

Fallo:
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbCritical, "No se pudo guardar la factura"
End Sub

'--- Limpieza interna (sin preguntar) -----------------------------------------
Private Sub LimpiarFacturaInterno(ws As Worksheet)
    Application.ScreenUpdating = False
    BorrarSoloValores ws.Range(FAC_CLI)
    BorrarSoloValores ws.Range("B" & FAC_F1 & ":B" & FAC_F2)
    BorrarSoloValores ws.Range("D" & FAC_F1 & ":D" & FAC_F2)
    ws.Range(FAC_OBS).MergeArea.ClearContents
    ws.Range(FAC_FECHA).Value = Date
    ws.Range(FAC_PAGO).Value = "Contado"
    RestaurarFormulasFactura ws
    ws.Range(FAC_ESTADO).Value = "BORRADOR"
    Application.ScreenUpdating = True
End Sub

'--- Reescribe las fórmulas de la factura por si alguien las sobrescribió ------
Private Sub RestaurarFormulasFactura(ws As Worksheet)
    Dim a As String, b As String
    a = FAC_F1: b = FAC_F2
    ws.Range("C" & a & ":C" & b).Formula = "=IF(B" & a & "="""","""",IFERROR(INDEX(PRODUCTOS!$B:$B,MATCH(B" & a & ",PRODUCTOS!$A:$A,0)),""SKU NO EXISTE""))"
    ws.Range("E" & a & ":E" & b).Formula = "=IF(B" & a & "="""","""",IFERROR(INDEX(PRODUCTOS!$F:$F,MATCH(B" & a & ",PRODUCTOS!$A:$A,0)),0))"
    ws.Range("F" & a & ":F" & b).Formula = "=IF(B" & a & "="""","""",D" & a & "*E" & a & ")"
    ws.Range("G" & a & ":G" & b).Formula = "=IF(B" & a & "="""","""",IF(IFERROR(INDEX(PRODUCTOS!$I:$I,MATCH(B" & a & ",PRODUCTOS!$A:$A,0)),""NO"")=""SI"",ROUND(F" & a & "*CONFIG!$B$12,0),0))"
    ws.Range("H" & a & ":H" & b).Formula = "=IF(B" & a & "="""","""",F" & a & "+G" & a & ")"
    ws.Range("I" & a & ":I" & b).Formula = "=IF(B" & a & "="""","""",IFERROR(INDEX(INVENTARIO!$G:$G,MATCH(B" & a & ",INVENTARIO!$A:$A,0)),""""))"
    ws.Range(FAC_SUB).Formula = "=SUM(F" & a & ":F" & b & ")"
    ws.Range(FAC_IVA).Formula = "=SUM(G" & a & ":G" & b & ")"
    ws.Range(FAC_TOT).Formula = "=SUM(H" & a & ":H" & b & ")"
    ws.Range(FAC_NUMTXT).Formula = "=CONFIG!$B$13&H3"
End Sub

'==============================================================================
'                               GUÍAS DE ENVÍO
'==============================================================================

'--- BOTÓN: CARGAR DATOS DESDE LA FACTURA ASOCIADA ----------------------------
Public Sub CargarFacturaEnGuia()
    Dim ws As Worksheet, wsH As Worksheet, wsD As Worksheet, wsC As Worksheet
    Dim fac As String, m As Variant, mc As Variant, r As Long, fila As Long, idCli As String

    Set ws = ThisWorkbook.Worksheets(H_GUIA)
    Set wsH = ThisWorkbook.Worksheets(H_HFAC)
    Set wsD = ThisWorkbook.Worksheets(H_HDET)
    Set wsC = ThisWorkbook.Worksheets(H_CLI)

    fac = Trim$(CStr(ws.Range(GUI_FACT).Value))
    If fac = "" Then MsgBox "Escriba o seleccione la factura asociada en " & GUI_FACT & ".", vbExclamation: Exit Sub
    m = Application.Match(fac, wsH.Columns(2), 0)
    If IsError(m) Then MsgBox "La factura " & fac & " no está en HIST_FACTURAS (¿la guardó?).", vbExclamation: Exit Sub

    Application.ScreenUpdating = False
    idCli = CStr(wsH.Cells(m, 4).Value)
    ws.Range(GUI_NOMBRE).Value = wsH.Cells(m, 5).Value
    ws.Range(GUI_ID).NumberFormat = "@"
    ws.Range(GUI_ID).Value = idCli
    ws.Range(GUI_CIUDAD).Value = wsH.Cells(m, 6).Value
    ws.Range(GUI_VALOR).Value = wsH.Cells(m, 7).Value
    mc = Application.Match(idCli, wsC.Columns(1), 0)
    If Not IsError(mc) Then
        ws.Range(GUI_TEL).NumberFormat = "@"
        ws.Range(GUI_TEL).Value = CStr(wsC.Cells(mc, 3).Value)
        ws.Range(GUI_DIR).Value = wsC.Cells(mc, 5).Value
        ws.Range(GUI_DEPTO).Value = wsC.Cells(mc, 7).Value
    End If

    BorrarSoloValores ws.Range("B" & GUI_F1 & ":B" & GUI_F2)
    BorrarSoloValores ws.Range("D" & GUI_F1 & ":D" & GUI_F2)
    BorrarSoloValores ws.Range("F" & GUI_F1 & ":H" & GUI_F2)
    fila = GUI_F1
    For r = FILA_DATOS To SiguienteFila(wsD, 1) - 1
        If CStr(wsD.Cells(r, 1).Value) = fac Then
            If fila > GUI_F2 Then
                MsgBox "La factura tiene más de " & (GUI_F2 - GUI_F1 + 1) & " ítems; solo se cargaron los primeros.", vbInformation
                Exit For
            End If
            ws.Cells(fila, "B").NumberFormat = "@"
            ws.Cells(fila, "B").Value = CStr(wsD.Cells(r, 3).Value)
            ws.Cells(fila, "D").Value = wsD.Cells(r, 5).Value
            fila = fila + 1
        End If
    Next r
    Application.ScreenUpdating = True
    MsgBox "Datos de la factura " & fac & " cargados. Complete transportadora, bultos y rastreo.", vbInformation
End Sub

'--- BOTÓN: IMPRIMIR GUÍA -----------------------------------------------------
Public Sub ImprimirGuia()
    Dim ws As Worksheet, resp As VbMsgBoxResult
    Set ws = ThisWorkbook.Worksheets(H_GUIA)
    ConfigurarPagina ws, GUI_AREA
    resp = MsgBox("SÍ = Imprimir   |   NO = Vista previa   |   CANCELAR = Salir", _
                  vbYesNoCancel + vbQuestion, "Imprimir guía " & ws.Range(GUI_NUMTXT).Value)
    Select Case resp
        Case vbYes: ws.PrintOut Copies:=1
        Case vbNo:  ws.PrintPreview
    End Select
End Sub

'--- BOTÓN: LIMPIAR GUÍA ------------------------------------------------------
Public Sub LimpiarGuia()
    If MsgBox("¿Borrar los datos de la guía actual?", vbYesNo + vbQuestion, "Limpiar guía") = vbNo Then Exit Sub
    LimpiarGuiaInterno ThisWorkbook.Worksheets(H_GUIA)
End Sub

'--- BOTÓN: NUEVA GUÍA --------------------------------------------------------
Public Sub NuevaGuia()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(H_GUIA)
    If ws.Range(GUI_ESTADO).Value <> "GUARDADA" And TieneItems(ws, GUI_F1, GUI_F2) Then
        If MsgBox("La guía actual no se ha guardado. ¿Descartarla?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If
    LimpiarGuiaInterno ws
    ws.Range(GUI_NUM).Value = SiguienteConsecutivo(H_HGUIA, ThisWorkbook.Worksheets(H_CFG).Range("B16").Value)
    ws.Activate
    ws.Range(GUI_FACT).Select
End Sub

'--- BOTÓN: GUARDAR GUÍA (historial + PDF en el Escritorio) -------------------
Public Sub GuardarGuia()
    Dim ws As Worksheet, wsH As Worksheet, n As Long, numero As Long
    Dim numTxt As String, archivo As String, fecha As Date

    On Error GoTo Fallo
    Set ws = ThisWorkbook.Worksheets(H_GUIA)
    Set wsH = ThisWorkbook.Worksheets(H_HGUIA)
    Application.Calculate

    numero = Val(ws.Range(GUI_NUM).Value)
    numTxt = CStr(ws.Range(GUI_NUMTXT).Value)
    If numero <= 0 Then Err.Raise vbObjectError + 20, , "No hay número de guía. Use NUEVA GUÍA."
    If Not IsError(Application.Match(numero, wsH.Columns(1), 0)) Then _
        Err.Raise vbObjectError + 21, , "La guía " & numTxt & " ya fue guardada. Use NUEVA GUÍA."
    If Not IsDate(ws.Range(GUI_FECHA).Value) Then Err.Raise vbObjectError + 22, , "Digite la fecha de despacho."
    If Trim$(CStr(ws.Range(GUI_NOMBRE).Value)) = "" Or Trim$(CStr(ws.Range(GUI_DIR).Value)) = "" _
       Or Trim$(CStr(ws.Range(GUI_CIUDAD).Value)) = "" Then _
        Err.Raise vbObjectError + 23, , "Faltan datos del destinatario (nombre, dirección o ciudad)."
    If Trim$(CStr(ws.Range(GUI_TRANSP).Value)) = "" Then Err.Raise vbObjectError + 24, , "Seleccione la transportadora."
    If Not TieneItems(ws, GUI_F1, GUI_F2) Then Err.Raise vbObjectError + 25, , "La guía no tiene productos."
    fecha = CDate(ws.Range(GUI_FECHA).Value)

    Application.ScreenUpdating = False
    archivo = CarpetaDestino("Guias", fecha) & "\" & LimpiarNombre(numTxt & "_" & ws.Range(GUI_NOMBRE).Value) & ".pdf"
    ConfigurarPagina ws, GUI_AREA
    ExportarPDF ws, archivo

    n = SiguienteFila(wsH, 1)
    wsH.Cells(n, 1).Value = numero
    wsH.Cells(n, 2).Value = numTxt
    wsH.Cells(n, 3).Value = fecha
    wsH.Cells(n, 4).Value = ws.Range(GUI_FACT).Value
    wsH.Cells(n, 5).Value = ws.Range(GUI_NOMBRE).Value
    wsH.Cells(n, 6).Value = ws.Range(GUI_CIUDAD).Value
    wsH.Cells(n, 7).Value = ws.Range(GUI_TRANSP).Value
    wsH.Cells(n, 8).Value = ws.Range(GUI_RASTREO).Value
    wsH.Cells(n, 9).Value = ws.Range(GUI_BULTOS).Value
    wsH.Cells(n, 10).Value = archivo
    wsH.Cells(n, 11).Value = Now
    FormatearFilasHist wsH

    ws.Range(GUI_ESTADO).Value = "GUARDADA"
    Application.ScreenUpdating = True
    ThisWorkbook.Save
    MsgBox "Guía " & numTxt & " guardada." & vbCrLf & "PDF: " & archivo, vbInformation, "Guardar guía"
    Exit Sub
Fallo:
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbCritical, "No se pudo guardar la guía"
End Sub

Private Sub LimpiarGuiaInterno(ws As Worksheet)
    Application.ScreenUpdating = False
    BorrarSoloValores ws.Range(GUI_FACT & "," & GUI_VALOR & "," & GUI_TEL & "," & GUI_DEPTO)
    ws.Range(GUI_TRANSP).MergeArea.ClearContents
    ws.Range(GUI_RASTREO).MergeArea.ClearContents
    ws.Range(GUI_NOMBRE).MergeArea.ClearContents
    ws.Range(GUI_ID).MergeArea.ClearContents
    ws.Range(GUI_DIR).MergeArea.ClearContents
    ws.Range(GUI_CIUDAD).MergeArea.ClearContents
    ws.Range(GUI_OBS).MergeArea.ClearContents
    BorrarSoloValores ws.Range("B" & GUI_F1 & ":B" & GUI_F2)
    BorrarSoloValores ws.Range("D" & GUI_F1 & ":D" & GUI_F2)
    BorrarSoloValores ws.Range("F" & GUI_F1 & ":H" & GUI_F2)
    ws.Range(GUI_FECHA).Value = Date
    ws.Range(GUI_ESTADO).Value = "BORRADOR"
    Application.ScreenUpdating = True
End Sub

'==============================================================================
'                     FUNCIONES DE APOYO (no se asignan a botones)
'==============================================================================

'Ruta: <Escritorio del usuario>\<carpeta CONFIG B17>\<subcarpeta>\<AAAA-MM>
'Usa la carpeta especial "Desktop" de Windows, que funciona aunque el Escritorio
'esté redirigido a OneDrive; si falla, usa %USERPROFILE%\Desktop.
Private Function CarpetaDestino(ByVal subcarpeta As String, ByVal fecha As Date) As String
    Dim escritorio As String, base As String
    On Error Resume Next
    escritorio = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    On Error GoTo 0
    If escritorio = "" Then escritorio = Environ$("USERPROFILE") & "\Desktop"
    base = Trim$(CStr(ThisWorkbook.Worksheets(H_CFG).Range("B17").Value))
    If base = "" Then base = "Gestion_Empresa"
    CarpetaDestino = escritorio & "\" & LimpiarNombre(base) & "\" & subcarpeta & "\" & Format$(fecha, "yyyy-mm")
    CrearCarpeta CarpetaDestino
End Function

'Crea la carpeta y todas las intermedias que no existan
Private Sub CrearCarpeta(ByVal ruta As String)
    Dim fso As Object, padre As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(ruta) Then Exit Sub
    padre = fso.GetParentFolderName(ruta)
    If padre <> "" And Not fso.FolderExists(padre) Then CrearCarpeta padre
    fso.CreateFolder ruta
End Sub

Private Sub ExportarPDF(ws As Worksheet, ByVal archivo As String)
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=archivo, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End Sub

'Copia la hoja a un libro nuevo, la convierte en valores, quita botones,
'validaciones, formatos condicionales y vínculos, y la guarda como .xlsx
Private Sub ExportarExcel(ws As Worksheet, ByVal archivo As String, ByVal colsAux As String)
    Dim wb As Workbook, sh As Worksheet, l As Variant
    ws.Copy
    Set wb = ActiveWorkbook
    Set sh = wb.Worksheets(1)
    On Error Resume Next
    sh.Unprotect Password:=CLAVE
    On Error GoTo 0
    sh.UsedRange.Copy
    sh.UsedRange.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    On Error Resume Next
    Dim k As Long
    For k = sh.Shapes.Count To 1 Step -1: sh.Shapes(k).Delete: Next k
    sh.Cells.Validation.Delete
    sh.Cells.FormatConditions.Delete
    sh.Columns(colsAux).Delete
    If Not IsEmpty(wb.LinkSources(xlExcelLinks)) Then
        For Each l In wb.LinkSources(xlExcelLinks)
            wb.BreakLink Name:=l, Type:=xlLinkTypeExcelLinks
        Next l
    End If
    On Error GoTo 0
    sh.Range("A1").Select
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=archivo, FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    ThisWorkbook.Activate
End Sub

Private Sub ConfigurarPagina(ws As Worksheet, ByVal area As String)
    On Error Resume Next             'por si no hay impresora instalada
    With ws.PageSetup
        .PrintArea = area
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .PrintGridlines = False
    End With
    On Error GoTo 0
End Sub

'Siguiente consecutivo = mayor número guardado + 1 (o el inicial de CONFIG)
Private Function SiguienteConsecutivo(ByVal hojaHist As String, ByVal inicial As Variant) As Long
    Dim mx As Double
    mx = Application.WorksheetFunction.Max(ThisWorkbook.Worksheets(hojaHist).Columns(1))
    If mx < Val(inicial) - 1 Then mx = Val(inicial) - 1
    SiguienteConsecutivo = CLng(mx) + 1
End Function

'Primera fila vacía de una tabla según la columna indicada (mínimo FILA_DATOS)
Private Function SiguienteFila(ws As Worksheet, ByVal col As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, col).End(xlUp).Row + 1
    If r < FILA_DATOS Then r = FILA_DATOS
    SiguienteFila = r
End Function

Private Function TieneItems(ws As Worksheet, ByVal f1 As Long, ByVal f2 As Long) As Boolean
    TieneItems = Application.WorksheetFunction.CountA(ws.Range("B" & f1 & ":B" & f2)) > 0
End Function

Private Function EsCredito(ByVal v As Variant) As Boolean
    EsCredito = (LCase$(Left$(Trim$(CStr(v)), 2)) = "cr")
End Function

'Borra solo constantes: nunca toca una celda que tenga fórmula
Private Sub BorrarSoloValores(rng As Range)
    Dim c As Range
    For Each c In rng.Cells
        If Not c.HasFormula Then
            If c.MergeCells Then c.MergeArea.ClearContents Else c.ClearContents
        End If
    Next c
End Sub

Private Function LimpiarNombre(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, ch, "_")
    Next ch
    LimpiarNombre = Left$(Trim$(s), 80)
End Function

'Copia el formato de la fila 5 a las filas nuevas del historial
Private Sub FormatearFilasHist(ws As Worksheet)
    Dim ultima As Long, ncol As Long
    ultima = SiguienteFila(ws, 1) - 1
    If ultima <= FILA_DATOS Then Exit Sub
    ncol = ws.Cells(FILA_DATOS - 1, ws.Columns.Count).End(xlToLeft).Column
    ws.Range(ws.Cells(FILA_DATOS, 1), ws.Cells(FILA_DATOS, ncol)).Copy
    ws.Range(ws.Cells(FILA_DATOS + 1, 1), ws.Cells(ultima, ncol)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub

'Si la fila supera las fórmulas precargadas, las copia de la fila de arriba
Private Sub AsegurarFormulasCartera(ws As Worksheet, ByVal fila As Long)
    If Not ws.Cells(fila, "G").HasFormula Then ws.Range("G" & fila - 1 & ":K" & fila - 1).Copy ws.Range("G" & fila)
End Sub

Private Sub AsegurarFormulasCaja(ws As Worksheet, ByVal fila As Long)
    If Not ws.Cells(fila, "C").HasFormula Then
        ws.Range("C" & fila - 1).Copy ws.Range("C" & fila)
        ws.Range("H" & fila - 1 & ":J" & fila - 1).Copy ws.Range("H" & fila)
    End If
End Sub

'==============================================================================
'           CREAR BOTONES (ejecutar UNA vez con Alt+F8 > CrearBotones)
'==============================================================================
Public Sub CrearBotones()
    CrearBotonesHoja ThisWorkbook.Worksheets(H_FAC), _
        Array("IMPRIMIR FACTURA", "LIMPIAR FACTURA", "GUARDAR (Historial + PDF)", "NUEVA FACTURA"), _
        Array("ImprimirFactura", "LimpiarFactura", "GuardarFactura", "NuevaFactura")
    CrearBotonesHoja ThisWorkbook.Worksheets(H_GUIA), _
        Array("CARGAR DESDE FACTURA", "IMPRIMIR GUÍA", "LIMPIAR GUÍA", "GUARDAR GUÍA (PDF)", "NUEVA GUÍA"), _
        Array("CargarFacturaEnGuia", "ImprimirGuia", "LimpiarGuia", "GuardarGuia", "NuevaGuia")
    MsgBox "Botones creados en FACTURACION y GUIAS.", vbInformation
End Sub

Private Sub CrearBotonesHoja(ws As Worksheet, textos As Variant, macros As Variant)
    Dim i As Long, b As Object, celda As Range
    For i = ws.Shapes.Count To 1 Step -1            'elimina botones anteriores
        If Left$(ws.Shapes(i).Name, 4) = "btn_" Then ws.Shapes(i).Delete
    Next i
    For i = LBound(textos) To UBound(textos)
        Set celda = ws.Range("K" & (5 + i * 3))
        Set b = ws.Buttons.Add(celda.Left + 2, celda.Top, celda.Width - 4, 34)
        b.Name = "btn_" & macros(i)
        b.Caption = textos(i)
        b.OnAction = macros(i)
        b.Font.Bold = True
        b.Font.Size = 10
        b.PrintObject = False                       'el botón nunca se imprime
        b.Placement = xlFreeFloating
    Next i
End Sub

'==============================================================================
'      PROTECCIÓN OPCIONAL: bloquea fórmulas; solo se editan celdas amarillas
'==============================================================================
Public Sub ProtegerHojas()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect Password:=CLAVE, UserInterfaceOnly:=True, AllowFiltering:=True, _
                   AllowFormattingColumns:=True
    Next ws
End Sub

Public Sub DesprotegerHojas()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:=CLAVE
    Next ws
End Sub
```

---

## 7. Solución de problemas

| Síntoma | Causa y solución |
|---|---|
| "No se puede ejecutar la macro…" | Las macros están deshabilitadas, o el archivo no es `.xlsm`. Revise el paso 1.1. |
| "Error de compilación: Sub o Function no definida" | El código se pegó incompleto. Vuelva a importar `modGestion.bas`. |
| Caracteres raros (Ã©) en los mensajes | Se pegó desde un editor con otra codificación. Use **Importar archivo** con el `.bas`, que viene en codificación Windows-1252. |
| "Acceso denegado" al guardar el PDF | El antivirus o la protección de carpetas de OneDrive bloquean el Escritorio. Permita Excel en *Acceso controlado a carpetas* o cambie la ruta en la función `CarpetaDestino`. |
| La factura dice "SKU NO EXISTE" | El SKU no está en PRODUCTOS, o tiene espacios. Elíjalo de la lista desplegable. |
| "CLIENTE NO REGISTRADO" | Cree primero el cliente en CLIENTES (NIT como texto, sin puntos). |
| El inventario no descuenta | La factura no se guardó con el botón GUARDAR. Solo las facturas en HIST_DETALLE descuentan stock. |
| El DASHBOARD sale en ceros | El mes de análisis (CONFIG!B24) no corresponde a sus datos. |
| Mac | El código usa rutas y objetos de Windows (`WScript.Shell`, `FileSystemObject`). En Mac funcionan imprimir, limpiar y nueva factura, pero para guardar hay que adaptar `CarpetaDestino` y `CrearCarpeta`. |
