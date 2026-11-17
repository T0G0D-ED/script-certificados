$ErrorActionPreference = 'Stop'

# 1. Rutas apuntando a la nueva carpeta de Diciembre
$root = 'C:\Users\Maxi-\OneDrive\Escritorio\proyecto-certificados\CERTIFICADOS DICIEMBRE 2026'
$outDir = 'C:\Users\Maxi-\OneDrive\Escritorio\proyecto-certificados\_herramientas'
$csvPath = Join-Path $outDir 'inventario-diciembre-2026.csv'

function Get-MonthFromName([string]$name) {
    $u = $name.ToUpperInvariant()
    foreach ($m in @('DICIEMBRE','NOVIEMBRE','SEPTIEMBRE','OCTUBRE','FEBRERO','ENERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO')) {
        if ($u -match $m) { return $m }
    }
    return 'UNKNOWN'
}

# 2. Filtrar solo los archivos que digan DICIEMBRE
$dicFiles = Get-ChildItem -LiteralPath $root -Recurse -File |
    Where-Object {
        $_.Directory.Name -eq '2026' -and
        $_.Extension -match '^\.docx?$' -and
        $_.Name -notmatch '^~' -and
        (Get-MonthFromName $_.Name) -eq 'DICIEMBRE'
    } | Sort-Object FullName

Write-Host ("December files: {0}" -f $dicFiles.Count)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$rows = @()

try {
    $i = 0
    foreach ($f in $dicFiles) {
        $i++
        Write-Host ("[{0}/{1}] {2}" -f $i, $dicFiles.Count, $f.FullName.Substring($root.Length))
        $doc = $null
        $folioRaw = ''
        $folioNum = ''
        $desde = ''
        $hasta = ''
        $mesLabel = ''
        $err = ''
        try {
            $doc = $word.Documents.Open($f.FullName, $false, $true, $false)
            $t = $doc.Content.Text -replace "`r", ' ' -replace "`n", ' ' -replace "`t", ' '
            $t = [regex]::Replace($t, '\s+', ' ')
            
            # Texto de respaldo: encabezados/pies + cuadros de texto
            $extra = ''
            try { foreach ($sr in $doc.StoryRanges) { $r = $sr; while ($r) { $extra += ' ' + $r.Text; $r = $r.NextStoryRange } } } catch {}
            try { foreach ($sh in $doc.Shapes) { try { $extra += ' ' + $sh.TextFrame.TextRange.Text } catch {} } } catch {}
            $extra = [regex]::Replace($extra, '\s+', ' ')
            # Acepta FOLIO N + (o, grado, punto, etc.) + numero con puntos
            $folioRx = 'FOLIO\s*N[^\d\s]{0,2}\s*[:\-]?\s*([0-9][0-9\.]*)'
            if ($t -match $folioRx) { $folioRaw = $Matches[1].TrimEnd('.') }
            elseif ($extra -match $folioRx) { $folioRaw = $Matches[1].TrimEnd('.') }
            if ($folioRaw) { $folioNum = ($folioRaw -replace '\.', '') }
            if ($t -match 'Desde el(.{0,80}?2026)') { $desde = $Matches[0].Trim() }
            if ($t -match 'Hasta el(.{0,80}?202[0-9])') { $hasta = $Matches[0].Trim() }
            
            # 3. Buscar la etiqueta interna de DICIEMBRE
            if ($t -match '(DICIEMBRE\s*/\s*2026)') { $mesLabel = $Matches[1] }
            
        } catch {
            $err = $_.Exception.Message
        } finally {
            if ($doc) { $doc.Close([ref]$false) }
        }
        
        $rows += [pscustomobject]@{
            Path = $f.FullName
            Rel = $f.FullName.Substring($root.Length)
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
        if ($m) { $r.FolioRaw = $m.Folio; $r.FolioNum = ($m.Folio -replace '\.', ''); $r.Error = 'folio-manual' }
    }
}
# Plantilla con los que siguen sin folio, para completar a mano
$faltan = @($rows | Where-Object { -not $_.FolioNum } | ForEach-Object { [pscustomobject]@{ Rel = $_.Rel; Folio = '' } })
if ($faltan.Count -gt 0) {
    $faltan | Export-Csv -LiteralPath (Join-Path $outDir 'folios-faltantes.csv') -NoTypeInformation -Encoding UTF8
    Write-Host ('Sin folio: {0}. Complete la columna Folio en folios-faltantes.csv, guardelo como folios-manuales.csv y vuelva a ejecutar.' -f $faltan.Count)
}$rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "Wrote $csvPath"

$nums = $rows | Where-Object { $_.FolioNum -match '^\d+$' } | ForEach-Object { [int]$_.FolioNum }
Write-Host ("Parsed folios: {0}" -f $nums.Count)

if ($nums.Count -gt 0) {
    Write-Host ("Min={0} Max={1}" -f ($nums | Measure-Object -Minimum).Minimum, ($nums | Measure-Object -Maximum).Maximum)
    $dup = $rows | Where-Object { $_.FolioNum } | Group-Object FolioNum | Where-Object { $_.Count -gt 1 }
    Write-Host ("Duplicates: {0}" -f $dup.Count)
    $dup | ForEach-Object { Write-Host ("  {0} x{1}" -f $_.Name, $_.Count) }
    $missing = $rows | Where-Object { -not $_.FolioNum }
    Write-Host ("Missing folio: {0}" -f $missing.Count)
    $missing | ForEach-Object { Write-Host $_.Rel }
}