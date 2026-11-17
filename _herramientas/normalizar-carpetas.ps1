param(
    [string]$Raiz = '',
    [switch]$Aplicar    # sin este switch solo genera inventario y plan (no renombra nada)
)
$ErrorActionPreference = 'Stop'
$tools = $PSScriptRoot
if (-not $Raiz) { $Raiz = Join-Path (Split-Path $tools -Parent) 'CERTIFICADOS DICIEMBRE 2026' }
$Raiz = (Get-Item -LiteralPath $Raiz).FullName.TrimEnd('\')
$invPath = Join-Path $tools 'inventario-carpetas.csv'
$logPath = Join-Path $tools 'log-carpetas.csv'

# Reglas AUTOMATICAS (seguras): recortar y colapsar espacios. Nada mas.
function Get-NombreSeguro([string]$n) {
    $x = $n.Trim()
    $x = [regex]::Replace($x, '\s{2,}', ' ')
    return $x
}
# SUGERENCIAS (no se aplican): requieren criterio humano.
function Get-Sugerencias([string]$n) {
    $s = @()
    if ($n -match '(^|\s)-\s*$' -or $n -match '^\s*-') { $s += 'guion-colgante' }
    if ($n -cmatch '[a-z]' -and $n -cmatch '[A-Z]') { $s += 'mayus-minus-mezcladas' }
    if ($n -match '\s-\S|\S-\s') { $s += 'espaciado-guion-irregular' }
    return ($s -join ';')
}
function Get-Tipo([string]$name, [int]$depth) {
    if ($name -match '^(19|20)\d\d$') { return 'Anio' }
    if ($name -match '^CERTIFICADOS?\b') { return 'Certificados' }
    if ($depth -eq 1) { return 'Empresa' }
    return 'Sucursal/Otro'
}

# ---------- Etapa 1: inventario ----------
$dirs = @(Get-ChildItem -LiteralPath $Raiz -Recurse -Directory | Sort-Object { $_.FullName.Length } -Descending)
$inv = @()
foreach ($d in $dirs) {
    $rel = $d.FullName.Substring($Raiz.Length + 1)
    $depth = ($rel -split '\\').Count
    $prop = Get-NombreSeguro $d.Name
    $inv += [pscustomobject]@{
        RutaOriginal = $d.FullName
        NombreActual = $d.Name
        NombrePropuesto = $prop
        Tipo = Get-Tipo $d.Name $depth
        CambioSeguro = ($prop -cne $d.Name)
        Sugerencias = Get-Sugerencias $d.Name
    }
}
$inv | Export-Csv -LiteralPath $invPath -NoTypeInformation -Encoding UTF8
Write-Host ("Carpetas inspeccionadas: {0}" -f $inv.Count)

# ---------- Etapa 2: plan con deteccion de colisiones ----------
$cambios = @($inv | Where-Object CambioSeguro)
$log = @()
foreach ($c in $cambios) {
    $parent = Split-Path $c.RutaOriginal -Parent
    $dest = Join-Path $parent $c.NombrePropuesto
    $estado = 'PENDIENTE'; $nota = ''
    if (-not $c.NombrePropuesto) { $estado = 'OMITIDO'; $nota = 'nombre-vacio' }
    elseif (Test-Path -LiteralPath $dest) {
        # colision (Windows ignora mayusculas): no se combina ni sobrescribe
        $estado = 'CONFLICTO'; $nota = 'destino-ya-existe'
    }
    else {
        $otros = @($cambios | Where-Object { $_.RutaOriginal -ne $c.RutaOriginal -and (Split-Path $_.RutaOriginal -Parent) -eq $parent -and $_.NombrePropuesto -ieq $c.NombrePropuesto })
        if ($otros.Count -gt 0) { $estado = 'CONFLICTO'; $nota = 'dos-carpetas-misma-propuesta' }
    }
    $log += [pscustomobject]@{ RutaOriginal = $c.RutaOriginal; NombreActual = $c.NombreActual; NombreNuevo = $c.NombrePropuesto; RutaNueva = $dest; Estado = $estado; Nota = $nota }
}

Write-Host ("Cambios seguros propuestos: {0} (conflictos: {1})" -f $log.Count, @($log | Where-Object Estado -eq 'CONFLICTO').Count)
$sug = @($inv | Where-Object { $_.Sugerencias })
Write-Host ("Carpetas con sugerencias NO aplicadas (revision manual): {0}" -f $sug.Count)
$sug | ForEach-Object { Write-Host ("  [{0}] {1}" -f $_.Sugerencias, $_.RutaOriginal.Substring($Raiz.Length)) }

# ---------- Etapa 3: aplicar (de lo mas profundo a lo mas superficial) ----------
if ($Aplicar) {
    foreach ($l in ($log | Where-Object Estado -eq 'PENDIENTE')) {
        try {
            Rename-Item -LiteralPath $l.RutaOriginal -NewName $l.NombreNuevo
            if (-not (Test-Path -LiteralPath $l.RutaNueva)) { throw 'ruta nueva no valida tras renombrar' }
            $l.Estado = 'RENOMBRADO'
        } catch { $l.Estado = 'ERROR'; $l.Nota = $_.Exception.Message }
    }
    Write-Host 'NOTA: las rutas cambiaron; vuelva a ejecutar extraer-*.ps1 para regenerar el inventario de certificados.'
}
$log | Export-Csv -LiteralPath $logPath -NoTypeInformation -Encoding UTF8
$log | Group-Object Estado | ForEach-Object { "{0}: {1}" -f $_.Name, $_.Count }
if (-not $Aplicar) { Write-Host 'Modo plan: no se renombro nada. Use -Aplicar para ejecutar.' }
