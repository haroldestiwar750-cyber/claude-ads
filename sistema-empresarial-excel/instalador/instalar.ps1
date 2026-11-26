# =====================================================================
#  INSTALADOR DEL SISTEMA EMPRESARIAL (Excel + VBA)
#  1. Crea ESCRITORIO\SISTEMA EMPRESARIAL\ con sus subcarpetas
#     (detecta el escritorio real del usuario, incluso con OneDrive)
#  2. Abre el libro base en Excel, importa los módulos VBA y el código
#     del libro, crea los botones y guarda "SISTEMA EMPRESARIAL.xlsm"
#  3. Restaura la configuración de seguridad de Excel que haya tocado
# =====================================================================
$ErrorActionPreference = "Stop"
$here   = Split-Path -Parent $MyInvocation.MyCommand.Path
$base   = Join-Path $here "SistemaEmpresarial_base.xlsx"
$vbaDir = Join-Path $here "vba"
$orden  = @("modEstructura","modUtil","modInventario","modFinanzas","modFactura","modGuia",
            "modCompras","modReportes","modSistema","modPruebas")

function Titulo($t) { Write-Host ""; Write-Host "== $t" -ForegroundColor Cyan }

Titulo "Verificando archivos"
if (-not (Test-Path $base)) { throw "No se encontró $base. Descomprima todo el paquete antes de instalar." }
foreach ($m in $orden) { if (-not (Test-Path (Join-Path $vbaDir "$m.bas"))) { throw "Falta vba\$m.bas" } }
Get-ChildItem -Path $here -Recurse -File | Unblock-File -ErrorAction SilentlyContinue

Titulo "Creando carpetas en el escritorio"
$desk = [Environment]::GetFolderPath("Desktop")
$dest = Join-Path $desk "SISTEMA EMPRESARIAL"
foreach ($s in @("FACTURAS","GUIAS","REPORTES","COPIAS DE SEGURIDAD")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $dest $s) | Out-Null
}
$out = Join-Path $dest "SISTEMA EMPRESARIAL.xlsm"
Write-Host "Destino: $out"
if (Test-Path $out) {
    $r = Read-Host "Ya existe un SISTEMA EMPRESARIAL.xlsm. Se guardará una copia y se reemplazará. ¿Continuar? (S/N)"
    if ($r -notmatch '^[sS]') { Write-Host "Instalación cancelada."; exit 1 }
    $bk = Join-Path (Join-Path $dest "COPIAS DE SEGURIDAD") ("Antes de reinstalar " + (Get-Date -Format "yyyy-MM-dd HHmmss") + ".xlsm")
    Copy-Item $out $bk
    Write-Host "Copia del archivo anterior: $bk"
}

Titulo "Habilitando temporalmente el acceso al proyecto VBA"
$cambios = @()
$office = "HKCU:\Software\Microsoft\Office"
$versiones = @()
if (Test-Path $office) { $versiones = Get-ChildItem $office | Where-Object { $_.PSChildName -match '^\d+\.\d+$' } | ForEach-Object { $_.PSChildName } }
if ($versiones.Count -eq 0) { $versiones = @("16.0") }
foreach ($v in $versiones) {
    $k = "$office\$v\Excel\Security"
    if (-not (Test-Path $k)) { New-Item -Path $k -Force | Out-Null }
    $old = (Get-ItemProperty -Path $k -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
    $cambios += [pscustomobject]@{ Key = $k; Old = $old }
    Set-ItemProperty -Path $k -Name AccessVBOM -Value 1 -Type DWord
}

$excel = $null; $wb = $null
try {
    Titulo "Abriendo Excel"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1   # permite ejecutar las macros recién importadas
    $wb = $excel.Workbooks.Open($base)
    $proj = $wb.VBProject

    Titulo "Importando módulos VBA"
    foreach ($m in $orden) {
        $null = $proj.VBComponents.Import((Join-Path $vbaDir "$m.bas"))
        Write-Host "  + $m"
    }
    # módulo del libro (ThisWorkbook / EsteLibro, según el idioma de Office)
    $codigos = @(); foreach ($ws in $wb.Worksheets) { $codigos += $ws.CodeName }
    $doc = $null
    if ($wb.CodeName) { try { $doc = $proj.VBComponents.Item($wb.CodeName) } catch { $doc = $null } }
    if (-not $doc) { foreach ($c in $proj.VBComponents) { if ($c.Type -eq 100 -and ($codigos -notcontains $c.Name)) { $doc = $c } } }
    $codigoLibro = [System.IO.File]::ReadAllText((Join-Path $vbaDir "ThisWorkbook.txt"), [System.Text.Encoding]::GetEncoding(1252))
    $doc.CodeModule.AddFromString($codigoLibro)
    Write-Host "  + Código del libro ($($doc.Name))"

    Titulo "Guardando como libro habilitado para macros"
    $wb.SaveAs($out, 52)

    Titulo "Creando botones e interfaz"
    $excel.Run("CrearInterfaz")
    $wb.Save()
    $wb.Close($false)
    $wb = $null
    Write-Host ""
    Write-Host "INSTALACIÓN COMPLETADA" -ForegroundColor Green
    Write-Host "Archivo: $out"
}
catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Si el error menciona 'acceso al proyecto de Visual Basic', siga la instalación manual del MANUAL (sección 2.3)." -ForegroundColor Yellow
    if ($wb) { try { $wb.Close($false) } catch {} }
    throw
}
finally {
    if ($excel) { try { $excel.Quit() } catch {}; [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    foreach ($c in $cambios) {
        if ($null -eq $c.Old) { Remove-ItemProperty -Path $c.Key -Name AccessVBOM -ErrorAction SilentlyContinue }
        else { Set-ItemProperty -Path $c.Key -Name AccessVBOM -Value $c.Old -Type DWord }
    }
}

Titulo "Ubicación de confianza (opcional)"
Write-Host "Si agrega la carpeta del sistema como 'ubicación de confianza', Excel abrirá el archivo"
Write-Host "con las macros activas sin mostrar la barra amarilla de advertencia."
$r = Read-Host "¿Agregar '$dest' como ubicación de confianza de Excel? (S/N)"
if ($r -match '^[sS]') {
    foreach ($v in $versiones) {
        $k = "$office\$v\Excel\Security\Trusted Locations\SistemaEmpresarial"
        New-Item -Path $k -Force | Out-Null
        Set-ItemProperty -Path $k -Name Path -Value ($dest + "\")
        Set-ItemProperty -Path $k -Name AllowSubfolders -Value 1 -Type DWord
        Set-ItemProperty -Path $k -Name Description -Value "Sistema Empresarial"
    }
    Write-Host "Ubicación de confianza agregada." -ForegroundColor Green
}

Start-Process $out
