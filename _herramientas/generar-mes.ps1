param(
    [string]$Origen = 'DICIEMBRE',     # mes del documento fuente
    [int]$OrigenAnio = 2026,
    [string]$Destino = 'ENERO',        # mes a generar
    [int]$DestinoAnio = 2027,
    [string]$HastaMes = 'FEBRERO',     # mes del "Hasta el"
    [int]$HastaAnio = 2027,
    [string]$DiaDesde = '01',
    [string]$DiaHasta = '15',
    [string]$Inventario = '',          # CSV de extraer-*.ps1 (por defecto inventario-<origen>-<anio>.csv)
    [string]$Raiz = '',                # carpeta raiz de certificados
    [string]$OnlyRel = '',
    [int]$MaxFiles = 0,
    [switch]$DryRun                    # solo muestra el plan, no crea nada
)
$ErrorActionPreference = 'Stop'
$tools = $PSScriptRoot
$logPath = Join-Path $tools ("log-{0}-{1}.csv" -f $Destino.ToLower(), $DestinoAnio)
if (-not $Inventario) { $Inventario = Join-Path $tools ("inventario-{0}-{1}.csv" -f $Origen.ToLower(), $OrigenAnio) }
$inventory = @(Import-Csv -LiteralPath $Inventario)

function Format-Folio([int]$n) {
    $s = $n.ToString()
    if ($s.Length -le 3) { return $s }
    return $s.Insert($s.Length - 3, '.')
}

# Reemplaza un grupo de un regex dentro de un parrafo, tocando solo esos caracteres
# (el resto del texto y su formato quedan intactos).
function Set-GroupInParagraph($doc, $para, [string]$pattern, [hashtable]$newByGroup) {
    $txt = $para.Range.Text
    $m = [regex]::Match($txt, $pattern, 'IgnoreCase')
    if (-not $m.Success) { return $false }
    $start = $para.Range.Start
    foreach ($g in ($newByGroup.Keys | Sort-Object { $m.Groups[$_].Index } -Descending)) {
        $grp = $m.Groups[$g]
        if (-not $grp.Success) { return $false }
        if ($grp.Value -ceq $newByGroup[$g]) { continue }
        $r = $doc.Range($start + $grp.Index, $start + $grp.Index + $grp.Length)
        $r.Text = $newByGroup[$g]
    }
    return $true
}

$rxFolio = 'FOLIO\s*N[^\d\s]{0,2}\s*[:\-]?\s*(?<f>[0-9][0-9\.]*)'
$rxDesde = 'Desde\s+el\s+(?<d>\d{1,2})\s+de\s+(?<m>[A-Za-z]+)\s+(?:de\s+)?(?<a>\d{4})'
$rxHasta = 'Hasta\s+el\s+(?<d>\d{1,2})\s+de\s+(?<m>[A-Za-z]+)\s+(?:de\s+)?(?<a>\d{4})'
$rxLabel = '(?<m>' + $Origen + ')\s*/\s*(?<a>' + $OrigenAnio + ')'

# Folio fuera del cuerpo (encabezados, pies, cuadros de texto)
function Update-FolioStories($doc, [string]$old, [string]$new) {
    $cnt = 0
    foreach ($sr in $doc.StoryRanges) {
        $r = $sr
        while ($r) {
            if ($r.StoryType -ne 1) {
                $m = [regex]::Match($r.Text, $rxFolio)
                if ($m.Success -and $m.Groups['f'].Value -eq $old) {
                    # Busqueda de texto exacto dentro de esta seccion (no depende de offsets,
                    # que se desfasan con caracteres especiales como imagenes inline)
                    $rr = $r.Duplicate
                    $f = $rr.Find
                    $f.ClearFormatting() | Out-Null
                    $f.Replacement.ClearFormatting() | Out-Null
                    if ($f.Execute($old, $true, $false, $false, $false, $false, $true, 0, $false, $new, 1)) { $cnt++ }
                }
            }
            $r = $r.NextStoryRange
        }
    }
    return $cnt
}

function Update-Certificate($doc, [string]$oldFolio, [string]$newFolio) {
    $n = @{ folio = 0; desde = 0; hasta = 0; label = 0 }
    foreach ($para in $doc.Paragraphs) {
        $txt = $para.Range.Text
        if ($txt -match $rxFolio -and $Matches['f'] -eq $oldFolio) {
            if (Set-GroupInParagraph $doc $para $rxFolio @{ f = $newFolio }) { $n.folio++ }
        }
        if ($txt -match 'Desde\s+el') {
            if (Set-GroupInParagraph $doc $para $rxDesde @{ d = $DiaDesde; m = $Destino; a = "$DestinoAnio" }) { $n.desde++ }
        }
        if ($txt -match 'Hasta\s+el') {
            if (Set-GroupInParagraph $doc $para $rxHasta @{ d = $DiaHasta; m = $HastaMes; a = "$HastaAnio" }) { $n.hasta++ }
        }
        if ($txt -match $rxLabel) {
            if (Set-GroupInParagraph $doc $para $rxLabel @{ m = $Destino; a = "$DestinoAnio" }) { $n.label++ }
        }
    }
    $n.folio += (Update-FolioStories $doc $oldFolio $newFolio)
    return $n
}

# Verificacion sobre el documento ya guardado y reabierto
function Test-Certificate($word, [string]$path, [string]$newFolio, [string]$oldFolio) {
    $errs = @()
    $doc = $word.Documents.Open($path, $false, $true, $false)
    try {
        $all = $doc.Content.Text
        foreach ($sr in $doc.StoryRanges) { $r = $sr; while ($r) { if ($r.StoryType -ne 1) { $all += ' ' + $r.Text }; $r = $r.NextStoryRange } }
        $t = [regex]::Replace(($all -replace "[`r`n`a`t]", ' '), '\s+', ' ')
    } finally { $doc.Close([ref]$false) }
    if ($t -notmatch $rxFolio) { $errs += 'sin-folio' }
    elseif ($Matches['f'] -ne $newFolio) { $errs += "folio-incorrecto($($Matches['f']))" }
    if ($t -notmatch $rxDesde) { $errs += 'sin-desde' }
    elseif ($Matches['d'] -ne $DiaDesde -or $Matches['m'] -cne $Destino -or $Matches['a'] -ne "$DestinoAnio") { $errs += "desde-incorrecto($($Matches[0]))" }
    if ($t -notmatch $rxHasta) { $errs += 'sin-hasta' }
    elseif ($Matches['d'] -ne $DiaHasta -or $Matches['m'] -cne $HastaMes -or $Matches['a'] -ne "$HastaAnio") { $errs += "hasta-incorrecto($($Matches[0]))" }
    if ($t -match ('(?i)' + $Origen + '\s*/\s*' + $OrigenAnio)) { $errs += 'etiqueta-mes-vieja' }
    return $errs
}

# ---------- Plan ----------
$nums = @()
foreach ($row in $inventory) { if ($row.FolioNum -match '^\d+$') { $nums += [int]$row.FolioNum } }
$sinFolio = @($inventory | Where-Object { $_.FolioNum -notmatch '^\d+$' })
if ($sinFolio.Count -gt 0) {
    Write-Host ("ABORTADO: {0} documentos sin folio en el inventario. Complete folios-manuales.csv y re-extraiga:" -f $sinFolio.Count) -ForegroundColor Red
    $sinFolio | ForEach-Object { Write-Host "  $($_.Rel)" }
    exit 1
}
$dup = $nums | Group-Object | Where-Object { $_.Count -gt 1 }
if ($dup) { Write-Host ("ABORTADO: folios duplicados en el inventario: {0}" -f (($dup | ForEach-Object Name) -join ', ')) -ForegroundColor Red; exit 1 }

$next = ($nums | Measure-Object -Maximum).Maximum + 1
Write-Host "Folio maximo $Origen = $($next - 1); la secuencia de $Destino empieza en $next"

function Get-TargetDir([string]$srcFile, [int]$destYear) {
    $srcDir = Split-Path $srcFile -Parent
    $dirName = Split-Path $srcDir -Leaf
    if ($dirName -match '^\d{4}$' -and $dirName -ne "$destYear") {
        $parent = Split-Path $srcDir -Parent
        return (Join-Path $parent "$destYear")
    }
    return $srcDir
}

$plan = @()
foreach ($row in $inventory) {
    $newName = [regex]::Replace($row.FileName, $Origen, $Destino, 'IgnoreCase')
    $newName = $newName.Replace("$OrigenAnio", "$DestinoAnio")
    $dstDir = Get-TargetDir $row.Path $DestinoAnio
    $plan += [pscustomobject]@{ Row = $row; NewName = $newName; DestDir = $dstDir; NewFolio = (Format-Folio $next); OldFolio = $row.FolioRaw }
    $next++
}
$targets = @($plan)
if ($OnlyRel) { $targets = @($targets | Where-Object { $_.Row.Rel -eq $OnlyRel }) }
if ($MaxFiles -gt 0) { $targets = @($targets | Select-Object -First $MaxFiles) }

if ($DryRun) {
    Write-Host '--- DRY RUN: no se crea nada ---'
    foreach ($t in $targets) {
        $dst = Join-Path $t.DestDir $t.NewName
        $estado = if (-not (Test-Path -LiteralPath $t.Row.Path)) { 'FALTA-FUENTE' } elseif ($t.NewName -ceq $t.Row.FileName -and $t.DestDir -eq (Split-Path $t.Row.Path -Parent)) { 'NOMBRE-IGUAL' } elseif (Test-Path -LiteralPath $dst) { 'YA-EXISTE' } else { 'CREARIA' }
        Write-Host ("{0,-13} {1} -> {2}  folio {3} -> {4}" -f $estado, $t.Row.FileName, $t.NewName, $t.OldFolio, $t.NewFolio)
    }
    exit 0
}

# ---------- Ejecucion ----------
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$logs = @()
try {
    $i = 0
    foreach ($t in $targets) {
        $i++
        $row = $t.Row
        $src = $row.Path
        if (-not (Test-Path -LiteralPath $t.DestDir)) { New-Item -ItemType Directory -Path $t.DestDir -Force | Out-Null }
        $dst = Join-Path $t.DestDir $t.NewName
        $status = ''; $notes = ''
        Write-Host ("[{0}/{1}] {2}" -f $i, $targets.Count, $row.Rel)

        if (-not (Test-Path -LiteralPath $src)) { $status = 'SKIP-missing-source' }
        elseif ($t.NewName -ceq $row.FileName -and $t.DestDir -eq (Split-Path $src -Parent)) { $status = 'SKIP-name-sin-mes' }
        elseif (Test-Path -LiteralPath $dst) { $status = 'SKIP-ya-existe' }
        else {
            $hashAntes = (Get-FileHash -LiteralPath $src -Algorithm SHA256).Hash
            Copy-Item -LiteralPath $src -Destination $dst
            $doc = $null
            try {
                $doc = $word.Documents.Open($dst, $false, $false, $false)
                $n = Update-Certificate $doc $row.FolioRaw $t.NewFolio
                $doc.Save()
                $doc.Close([ref]$false); $doc = $null
                $errs = @(Test-Certificate $word $dst $t.NewFolio $row.FolioRaw)
                if ((Get-FileHash -LiteralPath $src -Algorithm SHA256).Hash -ne $hashAntes) { $errs += 'ORIGINAL-MODIFICADO' }
                if ($errs.Count -gt 0) { throw ('verificacion: ' + ($errs -join ';')) }
                $status = 'CREATED-OK'
                $notes = "folio=$($n.folio) desde=$($n.desde) hasta=$($n.hasta) etiqueta=$($n.label)"
            } catch {
                $status = 'ERROR'
                $notes = $_.Exception.Message
                if ($doc) { try { $doc.Close([ref]$false) } catch {} }
                # Solo borramos la copia que acabamos de crear nosotros
                if (Test-Path -LiteralPath $dst) { Remove-Item -LiteralPath $dst -Force }
            }
        }
        $logs += [pscustomobject]@{
            Rel = $row.Rel; Source = $row.FileName; Dest = $t.NewName
            OldFolio = $row.FolioRaw; NewFolio = $t.NewFolio
            Status = $status; Notes = $notes; DestPath = $dst
        }
        Write-Host ("  -> {0} folio {1} => {2} {3}" -f $status, $row.FolioRaw, $t.NewFolio, $notes)
    }
} finally {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

if (Test-Path -LiteralPath $logPath) {
    $existing = @(Import-Csv -LiteralPath $logPath)
    $keep = $existing | Where-Object { $rel = $_.Rel; -not ($logs | Where-Object { $_.Rel -eq $rel }) }
    $logs = @($keep) + @($logs)
}
$logs | Export-Csv -LiteralPath $logPath -NoTypeInformation -Encoding UTF8
Write-Host "Log: $logPath"
$logs | Group-Object Status | ForEach-Object { "{0}: {1}" -f $_.Name, $_.Count }
