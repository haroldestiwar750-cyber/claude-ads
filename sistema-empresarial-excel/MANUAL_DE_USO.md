---
title: "Sistema Integral de Gestión Empresarial en Excel"
subtitle: "Manual de instalación y uso paso a paso"
---

# 1. Qué es y qué incluye

Un pequeño ERP construido en Microsoft Excel con macros VBA para una comercializadora de ropa, calzado y accesorios (adaptable a cualquier categoría). Controla facturación, guías de envío, inventario con kardex, compras, proveedores, clientes, costos y rentabilidad, ingresos, egresos, gastos fijos, cartera, caja y bancos, utilidades, reportes y un panel gerencial.

**Contenido del paquete `SistemaEmpresarial.zip`:**

| Archivo | Para qué sirve |
|---|---|
| `INSTALAR.bat` | Instalador con doble clic (ejecuta `instalar.ps1`). |
| `instalar.ps1` | Script que arma el archivo final `.xlsm`. |
| `SistemaEmpresarial_base.xlsx` | Libro con las 22 hojas, fórmulas, gráficos y datos de demostración. |
| `vba\*.bas`, `vba\ThisWorkbook.txt` | Código VBA completo y documentado (10 módulos). |
| `MANUAL_DE_USO.pdf` | Este manual. |

**Requisitos:** Windows 10/11 con Microsoft Excel de escritorio 2013, 2016, 2019, 2021 o Microsoft 365. No funciona en Excel para la web, Mac, móviles ni Google Sheets (las macros VBA y las rutas de Windows no existen allí).

# 2. Instalación

## 2.1 Instalación automática (recomendada)

1. Cierre Excel.
2. Descomprima `SistemaEmpresarial.zip` en cualquier carpeta (por ejemplo, Descargas). **No lo ejecute desde dentro del ZIP.**
3. Haga doble clic en **`INSTALAR.bat`** y siga las instrucciones de la ventana.

El instalador:

- Detecta automáticamente el escritorio del usuario (funciona con cualquier nombre de usuario y con escritorio en OneDrive) y crea:

```
ESCRITORIO\
   SISTEMA EMPRESARIAL\
      SISTEMA EMPRESARIAL.xlsm      <- el sistema
      FACTURAS\                     <- PDF de cada factura
      GUIAS\                        <- PDF de cada guía
      REPORTES\                     <- reportes PDF y Excel
      COPIAS DE SEGURIDAD\          <- respaldos automáticos y manuales
```

- Activa por unos segundos el permiso "Confiar en el acceso al modelo de objetos de proyectos de VBA", importa los módulos, crea los botones, guarda el `.xlsm` y **restaura el permiso a su valor original**.
- Ofrece (opcional) agregar la carpeta como **ubicación de confianza** para que Excel abra el sistema con macros activas sin la barra amarilla.

## 2.2 Habilitar macros

Si al abrir aparece la barra amarilla **"ADVERTENCIA DE SEGURIDAD – Las macros se han deshabilitado"**, pulse **Habilitar contenido**. Si aparece la barra roja **"RIESGO DE SEGURIDAD"**, el archivo fue copiado desde internet o correo: cierre Excel, clic derecho sobre el archivo → **Propiedades** → marque **Desbloquear** → Aceptar. (O ejecute de nuevo el instalador y acepte la ubicación de confianza.)

## 2.3 Instalación manual (si el instalador no puede ejecutarse)

1. Abra `SistemaEmpresarial_base.xlsx` y guárdelo como **Libro de Excel habilitado para macros (\*.xlsm)** en `Escritorio\SISTEMA EMPRESARIAL\SISTEMA EMPRESARIAL.xlsm`.
2. Pulse **Alt + F11** (editor de VBA).
3. Menú **Archivo → Importar archivo…** e importe, en este orden: `modEstructura.bas`, `modUtil.bas`, `modInventario.bas`, `modFinanzas.bas`, `modFactura.bas`, `modGuia.bas`, `modCompras.bas`, `modReportes.bas`, `modSistema.bas`, `modPruebas.bas`.
4. En el panel izquierdo haga doble clic en **ThisWorkbook** (en Excel en español: **EsteLibro**), abra `ThisWorkbook.txt` con el Bloc de notas, copie todo y péguelo en la ventana de código.
5. Cierre el editor, pulse **Alt + F8**, seleccione **CrearInterfaz** y pulse **Ejecutar**. Aparecerán todos los botones.
6. Guarde (Ctrl + G / Ctrl + S).

# 3. Primeros pasos (una sola vez)

1. **Pruebe el sistema con los datos de demostración.** Trae 6 meses de operación simulada de una tienda ficticia (164 facturas, compras, abonos, gastos, guías y una factura anulada) para que vea todo funcionando.
2. Vaya a **CONFIGURACIÓN → EJECUTAR PRUEBAS** para comprobar en su propio Excel que todo funciona (ver sección 9).
3. Cuando esté listo para operar de verdad: **CONFIGURACIÓN → BORRAR DATOS DEMO** (pide clave; antes crea una copia de seguridad). Conserva la configuración, las listas y el cliente "CONSUMIDOR FINAL".
4. **CONFIGURACIÓN:** escriba nombre, NIT, dirección, teléfono, correo, datos bancarios, Nequi/Daviplata, márgenes y textos. Pulse **CARGAR LOGO** para poner su logo en la factura y la guía.
5. **CAMBIAR CLAVES:** cambie las claves iniciales. Clave de autorización: `admin123`. Clave de protección de hojas: `1234`.
6. **CAMBIAR NUMERACIÓN** (si necesita continuar una numeración existente de facturas, guías o compras).
7. **CAJA Y BANCOS:** fecha y saldo inicial de cada cuenta (Efectivo, Bancos, Nequi, Daviplata, Otros).
8. Registre **PROVEEDORES**, **CLIENTES** y los **productos** en INVENTARIO (o cree el producto con stock inicial 0 y cárguelo con una COMPRA, que es lo recomendado porque deja el costo y el kardex completos).
9. En **INICIO**, elija el **Usuario** que está operando: queda registrado en cada movimiento.

# 4. Cómo está organizado

**Colores:** las celdas **amarillas** son de captura; las blancas/grises tienen fórmulas y están protegidas. En los encabezados de las tablas: azul = captura, verde = fórmula automática, azul oscuro = registrado por el sistema.

| Hoja | Función |
|---|---|
| INICIO | Panel gerencial: 15 indicadores, 8 gráficos, alertas, filtros y menú de navegación. |
| FACTURACIÓN | Formato de factura imprimible y sus botones. |
| GUÍAS DE ENVÍO | Formato de guía imprimible y sus botones. |
| INVENTARIO | Maestro de productos (SKU por talla/color) con stock en tiempo real y consulta de existencias. |
| COMPRAS | Formulario de compra + historial. |
| PROVEEDORES / CLIENTES | Bases de datos con totales y saldos automáticos. |
| COSTOS Y RENTABILIDAD | Calculadora de precios y rentabilidad de cada producto. |
| INGRESOS / EGRESOS | Registro de entradas y salidas de dinero, con resumen por período. |
| GASTOS FIJOS | Gastos recurrentes con vencimiento y estado del mes. |
| CARTERA | Cuentas por cobrar y por pagar, abonos, mora y alertas. |
| CAJA Y BANCOS | Saldo por cuenta, traslados, cierres diarios/semanales/mensuales. |
| UTILIDADES | Estado de resultados mensual, por categoría y por producto. |
| REPORTES | 11 reportes imprimibles y exportables. |
| CONFIGURACIÓN | Datos de la empresa, numeración, márgenes, claves y todas las listas desplegables. |
| BASE DE DATOS | Detalle de ventas (una fila por producto facturado). |
| HISTORIAL FACTURAS / HISTORIAL GUÍAS | Encabezado de cada documento guardado. |
| MOVIMIENTOS INV | Kardex: quién, cuándo, por qué cambió cada existencia. |
| AUDITORÍA | Bitácora de acciones sensibles (anulaciones, ajustes, claves, respaldos). |

Cada hoja tiene un botón **⌂ INICIO** y el panel tiene un menú lateral para ir a cualquier módulo.

# 5. Operación diaria paso a paso

## 5.1 Vender (FACTURACIÓN)

1. Pulse **NUEVA FACTURA** (pone el siguiente número y la fecha de hoy).
2. En **Cliente** elija de la lista (para uno nuevo: **NUEVO CLIENTE**, complete la fila en CLIENTES y vuelva).
3. En cada línea elija el **Código (SKU)**, escriba la **Cantidad** y, si aplica, el **Desc. %**. Descripción, talla, color, precio y subtotal se llenan solos.
4. Revise la columna **Validación** (no se imprime): avisa si no hay stock suficiente, si el producto está inactivo o si el precio quedó por debajo del costo real.
5. Elija **Método de pago** y **Estado de pago**: *PAGADA*, *PENDIENTE* (crédito) o *ABONO PARCIAL* (escriba el **Valor pagado**).
6. Pulse **GUARDAR FACTURA**. En una sola operación el sistema:
   - registra la venta en BASE DE DATOS y en HISTORIAL FACTURAS,
   - descuenta el inventario (kardex),
   - registra el dinero recibido en INGRESOS (si hubo pago),
   - crea la cuenta por cobrar en CARTERA (si quedó saldo),
   - avanza el consecutivo y guarda el **PDF** en `SISTEMA EMPRESARIAL\FACTURAS`.
7. **IMPRIMIR** imprime solo el formato de la factura (sin colores de captura ni columnas auxiliares).

**Protecciones:** una factura guardada no se puede guardar de nuevo (no hay doble descuento de inventario ni doble ingreso); el número solo avanza al guardar (no quedan huecos por borradores); si algo falla a mitad de camino se revierte todo; vender sin stock suficiente exige clave; vender bajo costo exige confirmación.

**BUSCAR FACTURA:** escriba el número, una fecha (dd/mm/aaaa) o parte del nombre/código del cliente. La factura se muestra con sus **precios históricos** para reimprimirla.

**ANULAR FACTURA:** pide clave y motivo, crea una copia de seguridad, devuelve las unidades al inventario, anula los ingresos y la cartera asociados y conserva la factura en el historial marcada como ANULADA (con quién, cuándo y por qué). No se puede anular dos veces.

**LIMPIAR** borra solo los datos capturados (nunca fórmulas ni formato).

## 5.2 Guías de envío

- Desde FACTURACIÓN: **GENERAR GUÍA** (con la factura guardada en pantalla), o en GUÍAS DE ENVÍO: **DESDE FACTURA**. Se llenan destinatario, documento, teléfono, ciudad, departamento, dirección, valor declarado y productos.
- Complete transportadora y número de seguimiento y pulse **GUARDAR GUÍA** (PDF en `SISTEMA EMPRESARIAL\GUIAS`).
- Para actualizar el **estado del envío** de una guía existente: **BUSCAR GUÍA**, cambie el estado/seguimiento y pulse **GUARDAR GUÍA**.
- La guía se imprime independiente de la factura (**IMPRIMIR GUÍA**).

## 5.3 Compras

En COMPRAS diligencie proveedor, forma y estado de pago, transporte/otros costos y hasta 8 productos con cantidad y costo unitario. **REGISTRAR COMPRA**:

- aumenta el inventario y deja el movimiento en el kardex,
- recalcula el **costo promedio ponderado** del producto,
- reparte transporte y otros costos entre los productos según su valor (costo real con fletes),
- registra el pago en EGRESOS (categoría *Compras*) y el saldo pendiente en CARTERA (por pagar).

## 5.4 Inventario

- **NUEVO PRODUCTO** lleva a la primera fila libre. Use un SKU por combinación de talla y color.
- **Nunca escriba en Entradas/Salidas/Stock**: se calculan desde el kardex.
- **DEVOLUCIÓN CLIENTE:** indique factura, SKU, cantidad y motivo. Reingresa las unidades, registra la venta negativa (corrige ventas y utilidad) y descuenta el valor del saldo pendiente del cliente o lo registra como dinero devuelto.
- **AJUSTE (CLAVE):** para conteos físicos, daños o pérdidas. Exige motivo y clave.
- **VER MOVIMIENTOS:** filtra el kardex del producto seleccionado.
- **Consulta de existencias** (arriba): por categoría, talla, color o parte del SKU.
- Estados automáticos: DISPONIBLE, BAJO STOCK (≤ mínimo), AGOTADO, INACTIVO. Rotación: ALTA, MEDIA, BAJA, SIN VENTAS.

## 5.5 Cartera

Seleccione la obligación en **Obligación** (solo aparecen las que tienen saldo), escriba fecha, valor y método, y pulse **REGISTRAR ABONO**. Para clientes genera un ingreso tipo ABONO; para proveedores un egreso. Saldo, días de mora y estado (PENDIENTE, ABONO PARCIAL, VENCIDA, PAGADA, ANULADA) se recalculan solos. **VER SOLO VENCIDA** filtra la cartera vencida.

## 5.6 Ingresos, egresos y gastos fijos

- **INGRESOS:** las ventas y los abonos llegan solos (no se registran a mano, así no se duplican). El formulario es solo para ingresos adicionales (aportes, préstamos, rendimientos…).
- **EGRESOS:** registre fecha, concepto, categoría (12), tipo (FIJO/VARIABLE/EXTRAORDINARIO), valor, método, responsable, comprobante y observaciones.
- **GASTOS FIJOS:** defina cada gasto recurrente una vez. Seleccione la fila y pulse **PAGAR GASTO SELECCIONADO**; el estado del mes cambia a PAGADO. Se marcan VENCIDO o PRÓXIMO A VENCER (5 días).
- Ingresos y egresos **manuales** se pueden anular (con clave) seleccionando la fila y pulsando **ANULAR SELECCIONADO**; los automáticos se revierten anulando el documento de origen.

## 5.7 Caja y bancos

Saldo disponible por cuenta = saldo inicial + entradas − salidas ± traslados. Cada método de pago entra a la cuenta definida en CONFIGURACIÓN (Transferencia y Tarjeta → BANCOS, etc.). Registre consignaciones o retiros entre cuentas en **TRASLADOS**. Para cierres pulse **DIARIO**, **SEMANAL** o **MENSUAL** (o escriba las fechas), revise y pulse **GUARDAR CIERRE**; **PDF CIERRE** lo exporta.

## 5.8 Panel gerencial (INICIO) y reportes

Filtros: **año**, **mes** (o TODOS), **categoría** y **día**. Los indicadores y gráficos se recalculan solos con cada operación. **HOY** vuelve al período actual.

En REPORTES elija el tipo y las fechas y pulse **GENERAR REPORTE**; luego **EXPORTAR PDF**, **EXPORTAR EXCEL** (guardados en `SISTEMA EMPRESARIAL\REPORTES`) o **IMPRIMIR**. Reportes: ventas diario/mensual/anual, inventario, productos vendidos, utilidades, gastos, cartera, compras, clientes y flujo de caja.

# 6. Cómo se calculan las cifras

- **Stock actual** = stock inicial + entradas − salidas, tomado del kardex (MOVIMIENTOS INV).
- **Costo de compra** = promedio ponderado: (stock × costo anterior + unidades compradas × costo nuevo) ÷ (stock + unidades compradas).
- **Costo real unitario** = costo de compra + fletes y otros costos de compra por unidad. Es el costo que se congela en cada venta.
- **Utilidad bruta** = ventas − costo real de lo vendido. **Utilidad neta** = utilidad bruta − gastos operativos (todos los egresos **excepto** las compras de mercancía, que son inventario y no gasto).
- **Ventas facturadas ≠ dinero recibido ≠ ganancia.** UTILIDADES muestra las tres cifras por separado: lo vendido (aunque sea a crédito), lo que realmente entró a caja y lo que realmente se ganó.
- **Margen sobre ventas** = (precio − costo) ÷ precio → precio = costo ÷ (1 − margen).
  **Recargo sobre costo** = (precio − costo) ÷ costo → precio = costo × (1 + recargo).
  Ejemplo con costo $50.000 y 40 %: margen → $83.334; recargo → $70.000 (solo 28,6 % de margen).
- **Precio mínimo** = costo real ÷ (1 − comisión %): por debajo se pierde dinero.
- **Escenarios:** competitivo, recomendado (precio sugerido) y premium, con los márgenes de CONFIGURACIÓN, redondeados a la centena.
- Empaques y comisiones se usan en la calculadora para decidir precios; lo que realmente se paga por ellos se registra en EGRESOS. Así no se cuentan dos veces.

# 7. Seguridad, integridad y respaldo

- **Hojas protegidas:** solo las celdas amarillas se pueden editar; las fórmulas no se pueden borrar. Las macros trabajan con protección "solo interfaz".
- **Validaciones:** listas desplegables, SKU y códigos únicos, cantidades enteras positivas, descuentos 0–100 %, fechas válidas, registros completos.
- **Claves:** la de autorización se pide para anular, ajustar inventario, vender con stock negativo, cambiar numeración, restaurar y borrar datos.
- **Consecutivos** bloqueados: solo avanzan al guardar y solo se cambian con clave, nunca a un número ya usado.
- **Auditoría:** cada acción sensible queda en la hoja AUDITORÍA con fecha, usuario y detalle.
- **Copias de seguridad:** automática una vez al día al cerrar Excel, antes de cada anulación y antes de borrar datos; manual con **COPIA DE SEGURIDAD** (INICIO o CONFIGURACIÓN). Se guardan en `SISTEMA EMPRESARIAL\COPIAS DE SEGURIDAD`.
- **RESTAURAR COPIA:** seleccione un respaldo; el sistema copia sus datos dentro del archivo actual (antes respalda el estado presente).

# 8. Datos de demostración

La tienda ficticia "MODA & ESTILO S.A.S." trae 24 productos, 12 clientes, 5 proveedores, 7 compras, 164 facturas (una anulada), abonos, gastos fijos y variables, traslados y 6 guías. Sirven para conocer el sistema. Bórrelos con **BORRAR DATOS DEMO** antes de operar.

# 9. Pruebas automáticas

**CONFIGURACIÓN → EJECUTAR PRUEBAS** realiza, usando exactamente las mismas macros de los botones y sin cuadros de diálogo:

1. Registrar un producto. 2. Comprarle 20 unidades con flete (costo de compra). 3. Calcular su precio sugerido. 4. Crear una factura por 3 unidades. 5. Guardarla. 6. Comprobar que el inventario pasa de 20 a 17 (y que un segundo guardado no descuenta otra vez). 7. Comprobar el registro de la venta, el ingreso y la cartera. 8. Comprobar la utilidad. 9. Generar la guía de envío. 10. Registrar un abono del cliente. 11. Registrar un gasto. 12. Comprobar que el panel refleja los cambios. 13. Anular una factura de prueba y comprobar que el inventario se restaura (y que no se puede anular dos veces ni con clave errada).

Son 37 verificaciones. El informe (Esperado / Obtenido / Resultado) queda en REPORTES y **todos los registros de prueba se eliminan al final**, dejando el archivo exactamente como estaba.

# 10. Mantenimiento

- **Agregar categorías, métodos de pago, transportadoras, usuarios, tallas, colores, categorías de gasto:** escriba al final de la columna correspondiente en CONFIGURACIÓN. Las listas desplegables se actualizan solas.
- **No cambie el nombre** de la categoría de gasto indicada en "Categoría de gasto para compras de mercancía" sin actualizar también esa celda.
- **Capacidad:** 1.000 productos, 2.000 clientes, 500 proveedores, 100 gastos fijos, 3.000 obligaciones de cartera y 100.000 registros por tabla de movimientos. Si se acerca al límite, el sistema avisa; la capacidad se amplía regenerando el libro con el script `build/construir_libro.py` (parámetros en `build/estructura.py`) y restaurando sus datos con **RESTAURAR COPIA**.
- **Desproteger una hoja** (para cambios de diseño): Revisar → Desproteger hoja → clave de hojas.
- **Volver a crear los botones:** Alt + F8 → **CrearInterfaz**.

# 11. Solución de problemas

| Problema | Solución |
|---|---|
| Los botones no hacen nada | Las macros están deshabilitadas: vea la sección 2.2. |
| "Clave incorrecta" | Use la clave de autorización vigente (inicial `admin123`). |
| No se generó el PDF | Verifique que tiene permiso de escritura en el escritorio; la venta igual quedó guardada y puede reimprimirla con BUSCAR + IMPRIMIR. |
| "El número no corresponde al consecutivo" | Pulse NUEVA FACTURA / NUEVA GUÍA. |
| Un desplegable no muestra un dato nuevo | Verifique que lo escribió en la columna correcta, sin filas vacías intermedias. |
| Excel muestra la hoja lenta | Los indicadores se recalculan con cada cambio; en archivos muy grandes cambie a cálculo manual (Fórmulas → Opciones de cálculo) y pulse ACTUALIZAR. |

# 12. Alcance y limitaciones

- Es una **factura comercial** (documento interno). **No reemplaza la facturación electrónica de la DIAN**: si su empresa está obligada, debe emitirla además a través de un proveedor autorizado.
- El sistema no liquida IVA (está pensado para no responsables de IVA). Los impuestos pagados se registran como egreso en la categoría *Impuestos*.
- Requiere Excel de escritorio para Windows.

# 13. Arquitectura técnica (para quien mantenga el sistema)

- **Una sola fuente de verdad:** `build/estructura.py` define hojas, columnas y celdas; de ahí se generan el libro y el módulo `modEstructura.bas` (constantes), por lo que macros y hojas nunca quedan desalineadas.
- **Registro único:** cada operación se escribe una sola vez (venta, movimiento, ingreso, obligación); todo lo demás son fórmulas (stock, saldos, utilidades, panel).
- **Transacciones:** las operaciones registran el punto de partida de cada tabla y, si ocurre un error, borran lo escrito y restauran valores (`TxIniciar` / `TxRevertir`).

| Módulo | Contenido |
|---|---|
| modEstructura | Constantes generadas (hojas, filas, columnas, celdas). |
| modUtil | Diálogos, transacciones, búsquedas, carpetas, PDF, auditoría, protección. |
| modInventario | Kardex, ajustes, devoluciones, nuevo producto. |
| modFactura | Nueva, guardar, imprimir, limpiar, buscar, anular. |
| modGuia | Guías: nueva, guardar/actualizar, imprimir, buscar, desde factura. |
| modCompras | Registro de compras, costo promedio, proveedores. |
| modFinanzas | Ingresos, egresos, gastos fijos, abonos, cierres de caja. |
| modReportes | 11 reportes, PDF, Excel, impresión. |
| modSistema | Botones, respaldo/restauración, claves, numeración, logo, borrado de demo. |
| modPruebas | Pruebas automáticas obligatorias. |
| ThisWorkbook | Protección al abrir y respaldo diario al cerrar. |
