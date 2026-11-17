param(
    [string]$Mes = 'ENERO',
    [int]$Anio = 2027,
    [string]$Raiz = ''
)
$ErrorActionPreference = 'Stop'
$outDir = $PSScriptRoot
if (-not $Raiz) {
    $parent = Split-Path $outDir -Parent
    $cand = Get-ChildItem -LiteralPath $parent -Directory | Where-Object { $_.Name -match '^CERTIFICADOS' } | Select-Object -First 1
    if ($cand) { $Raiz = $cand.FullName } else { throw "No se encontro carpeta CERTIFICADOS en $parent" }
}
$Mes = $Mes.ToUpperInvariant()
$csvPath = Join-Path $outDir ("inventario-{0}-{1}.csv" -f $Mes.ToLowerInvariant(), $Anio)

function Get-MonthFromName([string]$name) {
    $u = $name.ToUpperInvariant()
    foreach ($m in @('DICIEMBRE','NOVIEMBRE','SEPTIEMBRE','OCTUBRE','FEBRERO','ENERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO')) {
        if ($u -match $m) { return $m }
    }
    return 'UNKNOWN'
}

$files = @(Get-ChildItem -LiteralPath $Raiz -Recurse -File |
    Where-Object {
        $_.Directory.Name -eq "$Anio" -and
        $_.Extension -match '^\.docx?$' -and
        $_.Name -notmatch '^~' -and
        (Get-MonthFromName $_.Name) -eq $Mes
    } | Sort-Object FullName)

Write-Host ("Archivos de {0} {1} encontrados: {2}" -f $Mes, $Anio, $files.Count)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$rows = @()
$folioRx = 'FOLIO\s*N[^\d\s]{0,2}\s*[:\-]?\s*([0-9][0-9\.]*)'

try {
    $i = 0
    foreach ($f in $files) {
        $i++
        $rel = $f.FullName.Substring($Raiz.Length)
        Write-Host ("[{0}/{1}] {2}" -f $i, $files.Count, $rel)
        $doc = $null
        $folioRaw = ''; $folioNum = ''; $desde = ''; $hasta = ''; $mesLabel = ''; $err = ''
        try {
            $doc = $word.Documents.Open($f.FullName, $false, $true, $false)
            $t = $doc.Content.Text -replace "`r", ' ' -replace "`n", ' ' -replace "`t", ' '
            $t = [regex]::Replace($t, '\s+', ' ')

            $extra = ''
            try { foreach ($sr in $doc.StoryRanges) { $r = $sr; while ($r) { $extra += ' ' + $r.Text; $r = $r.NextStoryRange } } } catch {}
            try { foreach ($sh in $doc.Shapes) { try { $extra += ' ' + $sh.TextFrame.TextRange.Text } catch {} } } catch {}
            $extra = [regex]::Replace($extra, '\s+', ' ')

            if ($t -match $folioRx) { $folioRaw = $Matches[1].TrimEnd('.') }
            elseif ($extra -match $folioRx) { $folioRaw = $Matches[1].TrimEnd('.') }
            if ($folioRaw) { $folioNum = ($folioRaw -replace '\.', '') }

            if ($t -match 'Desde\s+el(.{0,80}?202[0-9])') { $desde = $Matches[0].Trim() }
            if ($t -match 'Hasta\s+el(.{0,80}?202[0-9])') { $hasta = $Matches[0].Trim() }
            if ($t -match ("({0}\s*/\s*{1})" -f $Mes, $Anio)) { $mesLabel = $Matches[1] }
        } catch {
            $err = $_.Exception.Message
        } finally {
            if ($doc) { $doc.Close([ref]$false) }
        }

        $rows += [pscustomobject]@{
            Path = $f.FullName
            Rel = $rel
            FileName = $f.Name
            FolioRaw = $folioRaw
            FolioNum = $folioNum
            Desde = $desde
            Hasta = $hasta
            MesLabel = $mesLabel
            Error = $err
        }
    }
} finally {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# Overrides manuales: _herramientas\folios-manuales.csv (columnas: Rel,Folio)
$manualPath = Join-Path $outDir 'folios-manuales.csv'
if (Test-Path -LiteralPath $manualPath) {
    $manual = @(Import-Csv -LiteralPath $manualPath)
    foreach ($r in $rows) {
        $m = $manual | Where-Object { $_.Rel -eq $r.Rel -and $_.Folio } | Select-Object -First 1
        if ($m) {
            $r.FolioRaw = $m.Folio
            $r.FolioNum = ($m.Folio -replace '\.', '')
            $r.Error = 'folio-manual'
        }
    }
}

$faltan = @($rows | Where-Object { -not $_.FolioNum } | ForEach-Object { [pscustomobject]@{ Rel = $_.Rel; Folio = '' } })
$faltantesPath = Join-Path $outDir 'folios-faltantes.csv'
if ($faltan.Count -gt 0) {
    $faltan | Export-Csv -LiteralPath $faltantesPath -NoTypeInformation -Encoding UTF8
    Write-Host ("Sin folio: {0}. Complete folios-faltantes.csv o ingreselos en la interfaz." -f $faltan.Count)
} elseif (Test-Path -LiteralPath $faltantesPath) {
    Remove-Item -LiteralPath $faltantesPath -Force
}

$rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "Wrote $csvPath"

$nums = @($rows | Where-Object { $_.FolioNum -match '^\d+$' } | ForEach-Object { [int]$_.FolioNum })
Write-Host ("Parsed folios: {0}" -f $nums.Count)
if ($nums.Count -gt 0) {
    Write-Host ("Min={0} Max={1}" -f ($nums | Measure-Object -Minimum).Minimum, ($nums | Measure-Object -Maximum).Maximum)
    $dup = @($rows | Where-Object { $_.FolioNum } | Group-Object FolioNum | Where-Object { $_.Count -gt 1 })
    Write-Host ("Duplicates: {0}" -f $dup.Count)
    $missing = @($rows | Where-Object { -not $_.FolioNum })
    Write-Host ("Missing folio: {0}" -f $missing.Count)
}
