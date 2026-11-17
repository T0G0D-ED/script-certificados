param(
    [int]$MaxFiles = 0,
    [string]$OnlyRel = ''
)

$ErrorActionPreference = 'Stop'
$root = 'C:\Users\Maxi-\OneDrive\Escritorio\proyecto-certificados\CERTIFICADOS NOVIEMBRE 2026'
$tools = 'C:\Users\Maxi-\OneDrive\Escritorio\proyecto-certificados\_herramientas'
$inventory = Import-Csv -LiteralPath (Join-Path $tools 'inventario-noviembre-2026.csv')
$logPath = Join-Path $tools 'log-diciembre-2026.csv'

function Format-Folio([int]$n) {
    $s = $n.ToString()
    if ($s.Length -le 3) { return $s }
    return $s.Insert($s.Length - 3, '.')
}

function Invoke-WordReplace($range, [string]$findText, [string]$replaceText, [bool]$wildcards = $false) {
    if ([string]::IsNullOrEmpty($findText)) { return $false }
    $find = $range.Find
    $find.ClearFormatting() | Out-Null
    $find.Replacement.ClearFormatting() | Out-Null
    return [bool]$find.Execute($findText, $true, $false, $wildcards, $false, $false, $true, 0, $false, $replaceText, 2)
}

function Update-Certificate($doc, [string]$oldFolioRaw, [string]$newFolioRaw) {
    $notes = @()
    if ($oldFolioRaw -and $newFolioRaw) {
        $ok = Invoke-WordReplace $doc.Content $oldFolioRaw $newFolioRaw
        if ($ok) { $notes += "folio:$oldFolioRaw->$newFolioRaw" } else { $notes += "folio-NOT-FOUND:$oldFolioRaw" }
    } else {
        $notes += 'folio-skipped'
    }

    foreach ($para in $doc.Paragraphs) {
        $txt = $para.Range.Text
        if ($txt -match 'Hasta el') {
            Invoke-WordReplace $para.Range 'DICIEMBRE' 'ENERO' | Out-Null
            Invoke-WordReplace $para.Range 'DIIEMBRE' 'ENERO' | Out-Null
            Invoke-WordReplace $para.Range '2026' '2027' | Out-Null
            foreach ($d in @('31','30','28','04')) {
                Invoke-WordReplace $para.Range $d '15' | Out-Null
            }
            if ($txt -match 'Hasta el\s+de') {
                Invoke-WordReplace $para.Range 'Hasta el( {1,})de' 'Hasta el\1 15 de' $true | Out-Null
            }
            $notes += 'hasta-updated'
        }
        elseif ($txt -match 'Desde el') {
            Invoke-WordReplace $para.Range 'NOVIEMBRE' 'DICIEMBRE' | Out-Null
            foreach ($d in @('04','31','30','28')) {
                Invoke-WordReplace $para.Range $d '01' | Out-Null
            }
            if ($txt -match 'Desde el\s+de') {
                Invoke-WordReplace $para.Range 'Desde el( {1,})de' 'Desde el\1 01 de' $true | Out-Null
            }
            $notes += 'desde-updated'
        }
        elseif ($txt -match 'NOVIEMBRE\s*/') {
            Invoke-WordReplace $para.Range 'NOVIEMBRE' 'DICIEMBRE' | Out-Null
            $notes += 'mes-label-updated'
        }
    }
    return ($notes -join '|')
}

$parsed = @()
foreach ($row in $inventory) {
    if ($row.FolioNum -match '^\d+$') { $parsed += [int]$row.FolioNum }
}
$maxFolio = ($parsed | Measure-Object -Maximum).Maximum
$nextFolio = $maxFolio + 1
Write-Host "November max folio=$maxFolio; December sequence starts at $nextFolio"

# Pre-assign numbers in inventory order so a test file does not steal the first number.
$plan = @()
foreach ($row in $inventory) {
    $item = [pscustomobject]@{
        Row = $row
        NewFolio = ''
        NewFolioRaw = ''
    }
    if ($row.FolioNum -match '^\d+$') {
        $item.NewFolio = $nextFolio.ToString()
        $item.NewFolioRaw = Format-Folio $nextFolio
        $nextFolio++
    }
    $plan += $item
}

$targets = @($plan)
if ($OnlyRel) {
    $targets = @($plan | Where-Object { $_.Row.Rel -eq $OnlyRel })
}
if ($MaxFiles -gt 0) {
    $targets = @($targets | Select-Object -First $MaxFiles)
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$logs = @()

try {
    $i = 0
    foreach ($item in $targets) {
        $i++
        $row = $item.Row
        $src = $row.Path
        $dir = Split-Path $src -Parent
        $newName = [regex]::Replace($row.FileName, 'NOVIEMBRE', 'DICIEMBRE', 'IgnoreCase')
        $dst = Join-Path $dir $newName
        $status = ''
        $notes = ''

        Write-Host ("[{0}/{1}] {2}" -f $i, $targets.Count, $row.Rel)

        if (-not (Test-Path -LiteralPath $src)) {
            $status = 'SKIP-missing-source'
        }
        elseif ($newName -eq $row.FileName) {
            $status = 'SKIP-name-no-noviembre'
        }
        elseif (Test-Path -LiteralPath $dst) {
            $status = 'SKIP-diciembre-exists'
        }
        else {
            Copy-Item -LiteralPath $src -Destination $dst
            $doc = $null
            try {
                $doc = $word.Documents.Open($dst, $false, $false, $false)
                $notes = Update-Certificate $doc $row.FolioRaw $item.NewFolioRaw
                $doc.Save()
                $status = 'CREATED'
            } catch {
                $status = 'ERROR'
                $notes = $_.Exception.Message
                if (Test-Path -LiteralPath $dst) {
                    # Keep the copy for inspection unless Word never created useful output.
                }
            } finally {
                if ($doc) { $doc.Close() }
            }
        }

        $logs += [pscustomobject]@{
            Rel = $row.Rel
            Source = $row.FileName
            Dest = $newName
            OldFolio = $row.FolioRaw
            NewFolio = $item.NewFolioRaw
            Status = $status
            Notes = $notes
            DestPath = $dst
        }
        Write-Host ("  -> {0} folio {1} => {2}" -f $status, $row.FolioRaw, $item.NewFolioRaw)
    }
} finally {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$existing = @()
if (Test-Path -LiteralPath $logPath) {
    $existing = @(Import-Csv -LiteralPath $logPath)
    $keep = $existing | Where-Object { $rel = $_.Rel; -not ($logs | Where-Object { $_.Rel -eq $rel }) }
    $logs = @($keep) + @($logs)
}
$logs | Export-Csv -LiteralPath $logPath -NoTypeInformation -Encoding UTF8
Write-Host "Log: $logPath"
$logs | Group-Object Status | ForEach-Object { "{0}: {1}" -f $_.Name, $_.Count }
