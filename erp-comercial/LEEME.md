# ERP Comercial: interfaz en Excel (versión 1.0, para revisión)

## Qué contiene

| Archivo | Qué es |
|---|---|
| `ERP_Comercial.xlsm` | El libro con los 22 módulos, más 8 tablas `BD_*`, fórmulas, dashboard, gráficos y datos de prueba. Todavía no trae macros: se instalan en el paso 2. |
| `vba/` | Código VBA: 11 módulos `.bas` y `ThisWorkbook.txt`, en codificación Windows-1252 lista para importar. |
| `generador/` | Scripts en Python que construyen el libro y los datos de prueba. Solo sirven para regenerar el archivo. |

## 1. Solo para ver cómo queda (sin macros)

Abra `ERP_Comercial.xlsm`. Si Excel pregunta, **no** habilite las macros todavía.
Todo se ve y calcula con fórmulas: menú, dashboard, facturación, inventario, cartera, utilidades, etc.
Las celdas **amarillas** son editables. Las **grises** están calculadas.

## 2. Instalar las macros (botones funcionales)

1. Deje la carpeta `vba` junto al archivo `ERP_Comercial.xlsm`.
2. En Excel, vaya a **Archivo › Opciones › Centro de confianza › Configuración del Centro de confianza › Configuración de macros** y marque **«Confiar en el acceso al modelo de objetos de proyectos de VBA»**. Puede desmarcarlo cuando termine la instalación.
3. Abra `ERP_Comercial.xlsm` y presione **Alt+F11**.
4. En el editor, vaya a **Archivo › Importar archivo…** y elija `vba\modInstalador.bas`.
5. Ponga el cursor dentro de `InstalarERP` y presione **F5**. El instalador hace esto:
   - importa los módulos;
   - crea la ventana de inicio de sesión con la clave oculta;
   - dibuja los botones de cada módulo;
   - protege las hojas.
6. Guarde, cierre y vuelva a abrir el archivo. Presione **Habilitar contenido** e inicie sesión.

**Usuarios de prueba.** Cada uno debe cambiar su clave en el primer ingreso.

| Usuario | Clave | Rol |
|---|---|---|
| admin | Admin2026 | Administrador |
| aux | Aux2026 | Auxiliar administrativo |
| laura / andres / camila | Laura2026 / Andres2026 / Camila2026 | Vendedor |

Desactive el **Autoguardado** de OneDrive para este archivo, porque el sistema guarda solo después de cada operación.

## Qué hace cada botón

- **Facturación:** nueva factura, guardar, imprimir, PDF, buscar, anular/devolución, limpiar y crear cliente.
  Al guardar se hace todo esto (si algo falla, se deshace todo):
  - registra la factura y descuenta el inventario;
  - registra el ingreso o la cuenta por cobrar;
  - congela el costo para calcular la utilidad;
  - asocia el vendedor y el canal;
  - genera el PDF en `OneDrive\SISTEMA EMPRESARIAL\FACTURAS\AAAA\MM`;
  - deja el registro en la bitácora.

  Una factura no se puede guardar dos veces. Las anuladas se conservan.
- **Compras:** actualiza el inventario, calcula el costo promedio ponderado (con fletes prorrateados), registra el egreso y la cuenta por pagar. Bloquea la misma factura del mismo proveedor si se intenta registrar dos veces.
- **Guías:** se crean a partir de una factura y cargan cliente, dirección y productos.
- **Inventario:** crea un SKU por cada talla y color, y registra ajustes. Los ajustes requieren la autorización de un usuario con el permiso `AJUSTE_INVENTARIO`.
- **Cartera:** abonos parciales. **Gastos fijos:** se generan una vez por mes. **Caja:** cierres diario, semanal y mensual. **Comisiones:** pago sin duplicar.
- **Reportes:** 13 reportes en PDF o Excel, que se guardan en `REPORTES`.
- **Configuración:**
  - detecta la ruta de OneDrive (no hay rutas fijas) y crea las carpetas;
  - hace respaldos manuales;
  - restaura copias.

Además, el sistema hace un respaldo automático cada 4 horas y al cerrar el archivo.

## Limitaciones de esta versión

1. **El VBA no se ha ejecutado en Excel real.** Se construyó en un entorno sin Excel.
   - Las fórmulas sí se verificaron: 25.465 fórmulas, 0 errores, y los resultados cuadran con una simulación independiente.
   - El código VBA pasó una revisión estática.
   - Haga las pruebas de la Etapa 9 con datos de prueba antes de usarlo con datos reales.
2. **Los datos viven dentro de este archivo.** Úselo en un solo equipo, o como demostración. Repartir copias a varios empleados duplicaría y desincronizaría la información. El trabajo multiusuario requiere la base de datos central (Etapa 1).
3. **La protección no es seguridad fuerte.** La protección de hojas, las claves de Excel y el VBA evitan errores, pero no protegen la información de un usuario con conocimientos.
4. **El comprobante no es factura electrónica DIAN.** Si su empresa está obligada a facturar electrónicamente, debe integrarse con un proveedor tecnológico autorizado.

## Regenerar el libro (opcional)

```bash
cd generador
python generar_erp.py && python a_xlsm.py && python exportar_vba.py
```
