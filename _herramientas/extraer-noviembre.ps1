$ErrorActionPreference = 'Stop'
$root = 'C:\Users\Maxi-\OneDrive\Escritorio\proyecto-certificados\CERTIFICADOS NOVIEMBRE 2026'
$outDir = 'C:\Users\Maxi-\OneDrive\Escritorio\proyecto-certificados\_herramientas'
$csvPath = Join-Path $outDir 'inventario-noviembre-2026.csv'

function Get-MonthFromName([string]$name) {
    $u = $name.ToUpperInvariant()
    foreach ($m in @('DICIEMBRE','NOVIEMBRE','SEPTIEMBRE','OCTUBRE','FEBRERO','ENERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO')) {
        if ($u -match $m) { return $m }
    }
    return 'UNKNOWN'
}

$novFiles = Get-ChildItem -LiteralPath $root -Recurse -File |
    Where-Object {
        $_.Directory.Name -eq '2026' -and
        $_.Extension -match '^\.docx?$' -and
        $_.Name -notmatch '^~' -and
        (Get-MonthFromName $_.Name) -eq 'NOVIEMBRE'
    } | Sort-Object FullName

Write-Host ("November files: {0}" -f $novFiles.Count)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$rows = @()
try {
    $i = 0
    foreach ($f in $novFiles) {
        $i++
        Write-Host ("[{0}/{1}] {2}" -f $i, $novFiles.Count, $f.FullName.Substring($root.Length))
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
            if ($t -match 'FOLIO\s*N[ºo°ª]?\s*([0-9\.]+)') {
                $folioRaw = $Matches[1]
                $folioNum = ($folioRaw -replace '\.', '')
            }
            if ($t -match 'Desde el(.{0,80}?2026)') { $desde = $Matches[0].Trim() }
            if ($t -match 'Hasta el(.{0,80}?202[0-9])') { $hasta = $Matches[0].Trim() }
            if ($t -match '(NOVIEMBRE\s*/\s*2026)') { $mesLabel = $Matches[1] }
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

$rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
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
