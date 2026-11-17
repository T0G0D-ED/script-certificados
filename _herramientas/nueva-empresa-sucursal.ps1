param(
    [ValidateSet('Empresa','Sucursal')]
    [string]$Modo = 'Empresa',
    [string]$CarpetaEmpresa = '',     # Para Modo=Empresa: ej. "42 NUEVA EMPRESA" | Para Modo=Sucursal: carpeta existente
    [string]$NombreSucursal = '',     # Opcional en Modo=Empresa; obligatorio en Modo=Sucursal
    [string]$Inmueble = '',           # Texto para INMUEBLE TRATADO
    [string]$Direccion = '',          # Texto para DIRECCION
    [string]$Comuna = '',             # Texto para COMUNA
    [string]$MesBase = 'ENERO',       # Mes en el que se crea el certificado inicial
    [int]$AnioBase = 2027,
    [string]$FolioInicial = '',       # Si vacio, usa (max folio del mes + 1)
    [string]$Raiz = '',
    [switch]$AbrirEnWord
)
$ErrorActionPreference = 'Stop'
$tools = $PSScriptRoot
$proj = Split-Path $tools -Parent
if (-not $Raiz) {
    $cand = Get-ChildItem -LiteralPath $proj -Directory | Where-Object { $_.Name -match '^CERTIFICADOS' } | Select-Object -First 1
    $Raiz = if ($cand) { $cand.FullName } else { Join-Path $proj 'CERTIFICADOS DICIEMBRE 2026' }
}

$meses = @('ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE')
function Format-Folio([int]$n) {
    $s = $n.ToString()
    if ($s.Length -le 3) { return $s }
    return $s.Insert($s.Length - 3, '.')
}
function Get-NextMonth([string]$mes, [int]$anio) {
    $idx = [array]::IndexOf($meses, $mes.ToUpperInvariant())
    if ($idx -eq 11) { return @{ Mes = 'ENERO'; Anio = ($anio + 1) } }
    return @{ Mes = $meses[$idx + 1]; Anio = $anio }
}

$MesBase = $MesBase.ToUpperInvariant()
$nextM = Get-NextMonth $MesBase $AnioBase
$HastaMes = $nextM.Mes
$HastaAnio = $nextM.Anio

# 1. Leer inventario actual del mes base para conocer el ultimo folio
$invPath = Join-Path $tools ("inventario-{0}-{1}.csv" -f $MesBase.ToLowerInvariant(), $AnioBase)
if (-not (Test-Path -LiteralPath $invPath)) {
    & (Join-Path $tools 'extraer-mes.ps1') -Mes $MesBase -Anio $AnioBase -Raiz $Raiz | Out-Null
}
$invRows = if (Test-Path -LiteralPath $invPath) { @(Import-Csv -LiteralPath $invPath) } else { @() }
$nums = @($invRows | Where-Object { $_.FolioNum -match '^\d+$' } | ForEach-Object { [int]$_.FolioNum })
$maxFolio = if ($nums.Count -gt 0) { ($nums | Measure-Object -Maximum).Maximum } else { 37800 }

if (-not $FolioInicial) {
    $FolioInicial = Format-Folio ($maxFolio + 1)
}
$folioNumClean = ($FolioInicial -replace '\.', '')

# 2. Determinar carpeta destino y plantilla base (.doc)
$empDir = Join-Path $Raiz $CarpetaEmpresa.Trim()
$targetBaseDir = $empDir

if ($Modo -eq 'Empresa') {
    if (Test-Path -LiteralPath $empDir) {
        throw "Ya existe la carpeta de empresa: $CarpetaEmpresa"
    }
    New-Item -ItemType Directory -Path $empDir -Force | Out-Null
    if ($NombreSucursal.Trim()) {
        $targetBaseDir = Join-Path $empDir $NombreSucursal.Trim().ToUpperInvariant()
        New-Item -ItemType Directory -Path $targetBaseDir -Force | Out-Null
    }
} else {
    if (-not (Test-Path -LiteralPath $empDir)) {
        throw "No existe la empresa seleccionada: $CarpetaEmpresa"
    }
    if (-not $NombreSucursal.Trim()) {
        throw "Debe indicar el nombre de la nueva sucursal."
    }
    # Detectar si la empresa usa subcarpeta intermedia (ej. CERTIFICADOS o CERTIFICADOS - SUCURSALES)
    $subCert = Get-ChildItem -LiteralPath $empDir -Directory | Where-Object { $_.Name -match '^CERTIFICADOS' } | Select-Object -First 1
    $parentForBranch = if ($subCert) { $subCert.FullName } else { $empDir }
    $targetBaseDir = Join-Path $parentForBranch $NombreSucursal.Trim().ToUpperInvariant()
    if (Test-Path -LiteralPath $targetBaseDir) {
        throw "Ya existe la sucursal '$NombreSucursal' en '$CarpetaEmpresa'."
    }
    New-Item -ItemType Directory -Path $targetBaseDir -Force | Out-Null
}

# Crear carpetas de anio (2026 y AnioBase)
foreach ($y in (@(2026, $AnioBase) | Select-Object -Unique)) {
    $yd = Join-Path $targetBaseDir "$y"
    if (-not (Test-Path -LiteralPath $yd)) { New-Item -ItemType Directory -Path $yd -Force | Out-Null }
}
$yearDir = Join-Path $targetBaseDir "$AnioBase"

# Elegir documento plantilla: primero buscar un .doc del mismo mes dentro de la misma empresa, si no usar 41 AGR EL BOSQUE o el ultimo del inventario
$templateRow = $null
if ($Modo -eq 'Sucursal') {
    $empPrefix = '\' + $CarpetaEmpresa.Trim() + '\'
    $templateRow = $invRows | Where-Object { $_.Rel.StartsWith($empPrefix, [System.StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
}
if (-not $templateRow) {
    $templateRow = $invRows | Where-Object { $_.Rel -match '41 AGR EL BOSQUE' } | Select-Object -First 1
}
if (-not $templateRow -and $invRows.Count -gt 0) {
    $templateRow = $invRows[-1]
}
if (-not $templateRow -or -not (Test-Path -LiteralPath $templateRow.Path)) {
    throw "No se encontro un documento base en $MesBase $AnioBase para usar como plantilla."
}

# Nombre del archivo nuevo
$baseLabel = if ($NombreSucursal.Trim()) {
    $NombreSucursal.Trim().ToUpperInvariant()
} else {
    ($CarpetaEmpresa.Trim() -replace '^\d+[\-\s]*', '').Trim().ToUpperInvariant()
}
$newFileName = "{0} - {1} {2}.doc" -f $baseLabel, $AnioBase, $MesBase
$newFilePath = Join-Path $yearDir $newFileName

Copy-Item -LiteralPath $templateRow.Path -Destination $newFilePath -Force

# 3. Actualizar Folio, Inmueble, Direccion, Comuna y Fechas en Word
$rxFolio = 'FOLIO\s*N[^\d\s]{0,2}\s*[:\-]?\s*(?<f>[0-9][0-9\.]*)'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($newFilePath, $false, $false, $false)
    try {
        foreach ($para in $doc.Paragraphs) {
            $txt = $para.Range.Text
            # Actualizar Folio
            if ($txt -match $rxFolio) {
                $oldF = $Matches['f']
                $f = $para.Range.Find
                $f.ClearFormatting() | Out-Null; $f.Replacement.ClearFormatting() | Out-Null
                [void]$f.Execute($oldF, $true, $false, $false, $false, $false, $true, 0, $false, $FolioInicial, 1)
            }
            # Actualizar INMUEBLE TRATADO conservando la etiqueta MES / ANIO
            if ($Inmueble.Trim() -and $txt -match 'INMUEBLE\s+TRATADO\s*:\s*(?<val>.+?)(?<tail>\s+' + $MesBase + '\s*/\s*' + $AnioBase + '|$)') {
                $oldVal = $Matches['val'].Trim()
                if ($oldVal) {
                    $f = $para.Range.Find
                    $f.ClearFormatting() | Out-Null; $f.Replacement.ClearFormatting() | Out-Null
                    [void]$f.Execute($oldVal, $true, $false, $false, $false, $false, $true, 0, $false, $Inmueble.Trim().ToUpperInvariant(), 1)
                }
            }
            # Actualizar DIRECCION y COMUNA en la linea de DIRECCION:
            if ($txt -match '^\s*DIRECCION\s*:') {
                if ($Direccion.Trim() -and $txt -match 'DIRECCION\s*:\s*(?<dir>.+?)(\s+COMUNA\s*:|$)') {
                    $oldDir = $Matches['dir'].Trim()
                    if ($oldDir) {
                        $f = $para.Range.Find
                        $f.ClearFormatting() | Out-Null; $f.Replacement.ClearFormatting() | Out-Null
                        [void]$f.Execute($oldDir, $true, $false, $false, $false, $false, $true, 0, $false, $Direccion.Trim().ToUpperInvariant(), 1)
                    }
                }
                if ($Comuna.Trim() -and $txt -match 'COMUNA\s*:\s*(?<com>[^\r\n\a]+)') {
                    $oldCom = $Matches['com'].Trim()
                    if ($oldCom) {
                        $f = $para.Range.Find
                        $f.ClearFormatting() | Out-Null; $f.Replacement.ClearFormatting() | Out-Null
                        [void]$f.Execute($oldCom, $true, $false, $false, $false, $false, $true, 0, $false, $Comuna.Trim().ToUpperInvariant(), 1)
                    }
                }
            }
        }
        $doc.Save()
    } finally {
        $doc.Close([ref]$false)
    }
} finally {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# 4. Agregar al inventario del mes base y reordenar por ruta
$newRel = $newFilePath.Substring($Raiz.Length)
$newRow = [pscustomobject]@{
    Path = $newFilePath
    Rel = $newRel
    FileName = $newFileName
    FolioRaw = $FolioInicial
    FolioNum = $folioNumClean
    Desde = "Desde el 01 de $MesBase $AnioBase"
    Hasta = "Hasta el 15 de $HastaMes $HastaAnio"
    MesLabel = "$MesBase / $AnioBase"
    Error = ''
}
$invRows = @($invRows | Where-Object { $_.Rel -ne $newRel }) + @($newRow) | Sort-Object Path
$invRows | Export-Csv -LiteralPath $invPath -NoTypeInformation -Encoding UTF8

if ($AbrirEnWord) {
    Start-Process -FilePath $newFilePath
}

[pscustomobject]@{
    CarpetaCreada = $targetBaseDir
    ArchivoCreado = $newFilePath
    Rel = $newRel
    FolioAsignado = $FolioInicial
    TotalEnInventario = $invRows.Count
}
