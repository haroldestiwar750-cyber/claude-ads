# Sistema Integral de Gestión Empresarial en Excel (VBA)

ERP en Excel para una comercializadora de ropa, calzado y accesorios: facturación, guías de envío,
inventario con kardex, compras, proveedores, clientes, costos y rentabilidad, ingresos, egresos,
gastos fijos, cartera, caja y bancos, utilidades, 11 reportes y panel gerencial con 8 gráficos.

## Entrega

| Ruta | Contenido |
|---|---|
| `dist/SistemaEmpresarial.zip` | **Paquete para Windows**: descomprimir y ejecutar `INSTALAR.bat` → crea `Escritorio\SISTEMA EMPRESARIAL\SISTEMA EMPRESARIAL.xlsm`. |
| `MANUAL_DE_USO.pdf` / `.md` | Manual de instalación y uso paso a paso. |
| `SistemaEmpresarial_base.xlsx` | Libro base (22 hojas, fórmulas, gráficos, datos de demostración). |
| `vba/` | Código VBA documentado (10 módulos + código del libro). |
| `instalador/` | `INSTALAR.bat` + `instalar.ps1`. |
| `build/` | Generadores: `estructura.py` (fuente única de hojas/columnas → también genera `modEstructura.bas`), `construir_libro.py`, `demo.py`, `empaquetar.py`. |
| `pruebas/` | Verificación automática con LibreOffice: fórmulas vs. simulación, macros VBA reales (13 pasos obligatorios + pruebas extendidas). |

## Regenerar y probar

```bash
python3 build/empaquetar.py                                   # libro + módulo de constantes + ZIP
cd pruebas
python3 lo_util.py errores ../SistemaEmpresarial_base.xlsx      # celdas con error tras recalcular
python3 verificar_formulas.py ../SistemaEmpresarial_base.xlsx   # stock, costos, ventas, cartera vs. simulación
python3 ejecutar_vba_lo.py ../SistemaEmpresarial_base.xlsx      # 13 pruebas obligatorias (37 verificaciones)
python3 pruebas_extendidas.py ../SistemaEmpresarial_base.xlsx   # reversión exacta, devoluciones, reportes, etc.
```

Requiere Python 3 con `xlsxwriter`, `openpyxl`, `Pillow`, y LibreOffice (solo para las pruebas).
