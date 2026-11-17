Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$ErrorActionPreference = 'Stop'
$tools = $PSScriptRoot
$proj = Split-Path $tools -Parent
$cand = Get-ChildItem -LiteralPath $proj -Directory | Where-Object { $_.Name -match '^CERTIFICADOS' } | Select-Object -First 1
$raiz = if ($cand) { $cand.FullName } else { Join-Path $proj 'CERTIFICADOS DICIEMBRE 2026' }

$meses = @('ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE')

function Format-Folio([int]$n) {
    $s = $n.ToString()
    if ($s.Length -le 3) { return $s }
    return $s.Insert($s.Length - 3, '.')
}

function Get-NextMonth([string]$mes, [int]$anio) {
    $idx = [array]::IndexOf($meses, $mes.ToUpperInvariant())
    if ($idx -lt 0) { $idx = 0 }
    if ($idx -eq 11) {
        return @{ Mes = 'ENERO'; Anio = ($anio + 1) }
    } else {
        return @{ Mes = $meses[$idx + 1]; Anio = $anio }
    }
}

function Get-MonthFromName([string]$name) {
    $u = $name.ToUpperInvariant()
    foreach ($m in @('DICIEMBRE','NOVIEMBRE','SEPTIEMBRE','OCTUBRE','FEBRERO','ENERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO')) {
        if ($u -match $m) { return $m }
    }
    return 'UNKNOWN'
}

# Detecta automaticamente el ultimo (Anio, Mes) que tenga al menos 50 certificados
function Find-LatestBaseMonth() {
    $yearDirs = @(Get-ChildItem -LiteralPath $raiz -Recurse -Directory | Where-Object { $_.Name -match '^\d{4}$' })
    $years = @($yearDirs | ForEach-Object { [int]$_.Name } | Sort-Object -Unique -Descending)
    foreach ($y in $years) {
        $files = @(Get-ChildItem -LiteralPath $raiz -Recurse -File | Where-Object {
            $_.Directory.Name -eq "$y" -and $_.Extension -match '^\.docx?$' -and $_.Name -notmatch '^~'
        })
        for ($mIdx = 11; $mIdx -ge 0; $mIdx--) {
            $mName = $meses[$mIdx]
            $cnt = @($files | Where-Object { (Get-MonthFromName $_.Name) -eq $mName }).Count
            if ($cnt -ge 50) {
                return @{ Mes = $mName; Anio = $y; Count = $cnt }
            }
        }
    }
    return @{ Mes = 'ENERO'; Anio = 2027; Count = 0 }
}

# Ventana para pedir folio manual si algun documento no lo tiene
function Request-ManualFolio([string]$relPath, [string]$fileName) {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Ingresar Folio Manualmente'
    $dlg.Size = New-Object System.Drawing.Size(640, 290)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 12)

    $lbl1 = New-Object System.Windows.Forms.Label
    $lbl1.Text = "No se pudo detectar automaticamente el folio en este documento:`r`n$relPath"
    $lbl1.Location = New-Object System.Drawing.Point(20, 20)
    $lbl1.Size = New-Object System.Drawing.Size(590, 70)
    $dlg.Controls.Add($lbl1)

    $lbl2 = New-Object System.Windows.Forms.Label
    $lbl2.Text = 'Escriba el numero de folio (por ejemplo 37.850):'
    $lbl2.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
    $lbl2.Location = New-Object System.Drawing.Point(20, 100)
    $lbl2.Size = New-Object System.Drawing.Size(580, 30)
    $dlg.Controls.Add($lbl2)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Font = New-Object System.Drawing.Font('Segoe UI', 16, [System.Drawing.FontStyle]::Bold)
    $txt.Location = New-Object System.Drawing.Point(20, 135)
    $txt.Size = New-Object System.Drawing.Size(260, 40)
    $dlg.Controls.Add($txt)

    $btnOpen = New-Object System.Windows.Forms.Button
    $btnOpen.Text = 'Abrir documento en Word'
    $btnOpen.Location = New-Object System.Drawing.Point(300, 135)
    $btnOpen.Size = New-Object System.Drawing.Size(290, 40)
    $btnOpen.Add_Click({
        $full = Join-Path $raiz $relPath.TrimStart('\')
        if (Test-Path -LiteralPath $full) { Start-Process -FilePath $full }
    })
    $dlg.Controls.Add($btnOpen)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Guardar Folio'
    $btnOk.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.FlatStyle = 'Flat'
    $btnOk.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
    $btnOk.Location = New-Object System.Drawing.Point(20, 195)
    $btnOk.Size = New-Object System.Drawing.Size(260, 44)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $txt.Text.Trim()
    }
    return ''
}

# Ventana de DOBLE VERIFICACION para Deshacer
function Confirm-DoubleVerification([string]$mesDest, [int]$anioDest, [int]$cantidad) {
    $msg1 = "ATENCION - PASO 1 DE 2:`r`n`r`n" +
            "Esta a punto de BORRAR los $cantidad certificados generados del mes de $mesDest $anioDest.`r`n" +
            "(Los certificados de meses anteriores NO se tocaran).`r`n`r`n" +
            "¿Esta seguro de que desea continuar al paso de confirmacion final?"
    $r1 = [System.Windows.Forms.MessageBox]::Show($msg1, "Verificacion 1 de 2 - Deshacer $mesDest $anioDest", 'YesNo', 'Warning', 'Button2')
    if ($r1 -ne [System.Windows.Forms.DialogResult]::Yes) { return $false }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Verificacion Final (2 de 2) - Confirmar Borrado de $mesDest $anioDest"
    $dlg.Size = New-Object System.Drawing.Size(580, 310)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 245)
    $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 12)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "PASO 2 DE 2 (SEGURIDAD):`r`nPara confirmar que desea eliminar los $cantidad certificados de $mesDest $anioDest, escriba la palabra BORRAR en el recuadro:"
    $lbl.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
    $lbl.ForeColor = [System.Drawing.Color]::FromArgb(160, 20, 20)
    $lbl.Location = New-Object System.Drawing.Point(20, 20)
    $lbl.Size = New-Object System.Drawing.Size(530, 85)
    $dlg.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Font = New-Object System.Drawing.Font('Segoe UI', 18, [System.Drawing.FontStyle]::Bold)
    $txt.Location = New-Object System.Drawing.Point(20, 115)
    $txt.Size = New-Object System.Drawing.Size(260, 45)
    $dlg.Controls.Add($txt)

    $btnConfirm = New-Object System.Windows.Forms.Button
    $btnConfirm.Text = 'CONFIRMAR BORRADO'
    $btnConfirm.BackColor = [System.Drawing.Color]::FromArgb(190, 30, 30)
    $btnConfirm.ForeColor = [System.Drawing.Color]::White
    $btnConfirm.FlatStyle = 'Flat'
    $btnConfirm.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
    $btnConfirm.Location = New-Object System.Drawing.Point(20, 190)
    $btnConfirm.Size = New-Object System.Drawing.Size(260, 50)
    $btnConfirm.Enabled = $false
    $dlg.Controls.Add($btnConfirm)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancelar (No borrar nada)'
    $btnCancel.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
    $btnCancel.Location = New-Object System.Drawing.Point(300, 190)
    $btnCancel.Size = New-Object System.Drawing.Size(245, 50)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)

    $txt.Add_TextChanged({
        $btnConfirm.Enabled = ($txt.Text.Trim().ToUpperInvariant() -eq 'BORRAR')
    })
    $btnConfirm.Add_Click({
        if ($txt.Text.Trim().ToUpperInvariant() -eq 'BORRAR') {
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $dlg.Close()
        }
    })

    return ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
}

# ==================== VENTANA PRINCIPAL ====================
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Generador Mensual de Certificados - Parra & Mauro Ltda.'
$form.Size = New-Object System.Drawing.Size(1120, 780)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 12)

# Panel Superior: Seleccion y Resumen Claro
$pnlTop = New-Object System.Windows.Forms.Panel
$pnlTop.Location = New-Object System.Drawing.Point(16, 14)
$pnlTop.Size = New-Object System.Drawing.Size(1070, 150)
$pnlTop.BackColor = [System.Drawing.Color]::White
$pnlTop.BorderStyle = 'FixedSingle'
$form.Controls.Add($pnlTop)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = '1. Mes de origen (mes anterior ya listo):'
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
$lblTitle.Location = New-Object System.Drawing.Point(18, 16)
$lblTitle.Size = New-Object System.Drawing.Size(360, 30)
$pnlTop.Controls.Add($lblTitle)

$cmbMes = New-Object System.Windows.Forms.ComboBox
$cmbMes.DropDownStyle = 'DropDownList'
$cmbMes.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
$cmbMes.Location = New-Object System.Drawing.Point(380, 12)
$cmbMes.Size = New-Object System.Drawing.Size(190, 36)
$meses | ForEach-Object { [void]$cmbMes.Items.Add($_) }
$pnlTop.Controls.Add($cmbMes)

$numAnio = New-Object System.Windows.Forms.NumericUpDown
$numAnio.Minimum = 2024
$numAnio.Maximum = 2040
$numAnio.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
$numAnio.Location = New-Object System.Drawing.Point(585, 12)
$numAnio.Size = New-Object System.Drawing.Size(105, 36)
$pnlTop.Controls.Add($numAnio)

$btnAuto = New-Object System.Windows.Forms.Button
$btnAuto.Text = 'Detectar ultimo mes automaticamente'
$btnAuto.Font = New-Object System.Drawing.Font('Segoe UI', 11)
$btnAuto.Location = New-Object System.Drawing.Point(710, 11)
$btnAuto.Size = New-Object System.Drawing.Size(335, 36)
$pnlTop.Controls.Add($btnAuto)

$lblTargetInfo = New-Object System.Windows.Forms.Label
$lblTargetInfo.Font = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
$lblTargetInfo.ForeColor = [System.Drawing.Color]::FromArgb(15, 82, 186)
$lblTargetInfo.Location = New-Object System.Drawing.Point(18, 58)
$lblTargetInfo.Size = New-Object System.Drawing.Size(1030, 36)
$pnlTop.Controls.Add($lblTargetInfo)

$lblSummaryInfo = New-Object System.Windows.Forms.Label
$lblSummaryInfo.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Regular)
$lblSummaryInfo.ForeColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
$lblSummaryInfo.Location = New-Object System.Drawing.Point(18, 100)
$lblSummaryInfo.Size = New-Object System.Drawing.Size(1030, 36)
$pnlTop.Controls.Add($lblSummaryInfo)

# Panel de Botones Grandes
$btnSimular = New-Object System.Windows.Forms.Button
$btnSimular.Text = 'PASO 1: Revisar y Simular'
$btnSimular.Font = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
$btnSimular.BackColor = [System.Drawing.Color]::FromArgb(30, 100, 200)
$btnSimular.ForeColor = [System.Drawing.Color]::White
$btnSimular.FlatStyle = 'Flat'
$btnSimular.Location = New-Object System.Drawing.Point(16, 178)
$btnSimular.Size = New-Object System.Drawing.Size(275, 58)
$form.Controls.Add($btnSimular)

$btnGenerar = New-Object System.Windows.Forms.Button
$btnGenerar.Text = 'PASO 2: Crear Certificados'
$btnGenerar.Font = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
$btnGenerar.BackColor = [System.Drawing.Color]::FromArgb(25, 135, 60)
$btnGenerar.ForeColor = [System.Drawing.Color]::White
$btnGenerar.FlatStyle = 'Flat'
$btnGenerar.Location = New-Object System.Drawing.Point(305, 178)
$btnGenerar.Size = New-Object System.Drawing.Size(295, 58)
$form.Controls.Add($btnGenerar)

$btnCarpeta = New-Object System.Windows.Forms.Button
$btnCarpeta.Text = 'Abrir Carpeta Certificados'
$btnCarpeta.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$btnCarpeta.BackColor = [System.Drawing.Color]::FromArgb(230, 235, 242)
$btnCarpeta.Location = New-Object System.Drawing.Point(614, 178)
$btnCarpeta.Size = New-Object System.Drawing.Size(245, 58)
$form.Controls.Add($btnCarpeta)

$btnDeshacer = New-Object System.Windows.Forms.Button
$btnDeshacer.Text = 'Deshacer Mes Creado'
$btnDeshacer.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
$btnDeshacer.BackColor = [System.Drawing.Color]::FromArgb(250, 225, 225)
$btnDeshacer.ForeColor = [System.Drawing.Color]::FromArgb(160, 20, 20)
$btnDeshacer.FlatStyle = 'Flat'
$btnDeshacer.Location = New-Object System.Drawing.Point(873, 178)
$btnDeshacer.Size = New-Object System.Drawing.Size(213, 58)
$form.Controls.Add($btnDeshacer)

# Barra de progreso y estado
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Listo. Pulse "PASO 1: Revisar y Simular" para verificar los documentos antes de crear.'
$lblStatus.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$lblStatus.Location = New-Object System.Drawing.Point(18, 246)
$lblStatus.Size = New-Object System.Drawing.Size(1065, 28)
$form.Controls.Add($lblStatus)

$prog = New-Object System.Windows.Forms.ProgressBar
$prog.Location = New-Object System.Drawing.Point(16, 278)
$prog.Size = New-Object System.Drawing.Size(1070, 24)
$prog.Style = 'Continuous'
$form.Controls.Add($prog)

# Tabla de vista previa y resultados
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(16, 312)
$grid.Size = New-Object System.Drawing.Size(1070, 410)
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.RowHeadersVisible = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.AutoSizeColumnsMode = 'Fill'
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.Font = New-Object System.Drawing.Font('Segoe UI', 11)
$grid.ColumnHeadersHeight = 36
$grid.RowTemplate.Height = 30

[void]$grid.Columns.Add('Num', '#')
[void]$grid.Columns.Add('Empresa', 'Empresa / Sucursal')
[void]$grid.Columns.Add('ArchivoNuevo', 'Certificado a Crear')
[void]$grid.Columns.Add('FolioAnt', 'Folio Anterior')
[void]$grid.Columns.Add('FolioNuevo', 'Folio Nuevo')
[void]$grid.Columns.Add('Estado', 'Estado')
$grid.Columns['Num'].FillWeight = 8
$grid.Columns['Empresa'].FillWeight = 34
$grid.Columns['ArchivoNuevo'].FillWeight = 28
$grid.Columns['FolioAnt'].FillWeight = 12
$grid.Columns['FolioNuevo'].FillWeight = 12
$grid.Columns['Estado'].FillWeight = 18
$form.Controls.Add($grid)

function Update-HeaderLabels() {
    $mOrig = [string]$cmbMes.SelectedItem
    $aOrig = [int]$numAnio.Value
    if (-not $mOrig) { return }
    $next1 = Get-NextMonth $mOrig $aOrig
    $mDest = $next1.Mes; $aDest = $next1.Anio
    $next2 = Get-NextMonth $mDest $aDest
    $mHasta = $next2.Mes; $aHasta = $next2.Anio

    $lblTargetInfo.Text = "MES A CREAR:  $mDest $aDest   (Validez: Desde el 01 de $mDest $aDest  al  15 de $mHasta $aHasta)"

    $invPath = Join-Path $tools ("inventario-{0}-{1}.csv" -f $mOrig.ToLowerInvariant(), $aOrig)
    if (Test-Path -LiteralPath $invPath) {
        $rows = @(Import-Csv -LiteralPath $invPath)
        $nums = @($rows | Where-Object { $_.FolioNum -match '^\d+$' } | ForEach-Object { [int]$_.FolioNum })
        if ($nums.Count -gt 0) {
            $minF = ($nums | Measure-Object -Minimum).Minimum
            $maxF = ($nums | Measure-Object -Maximum).Maximum
            $newStart = $maxF + 1
            $newEnd = $maxF + $rows.Count
            $lblSummaryInfo.Text = "Inventario de $mOrig $aOrig listo: $($rows.Count) certificados (Folios $(Format-Folio $minF) al $(Format-Folio $maxF))  ->  Nuevos folios: $(Format-Folio $newStart) al $(Format-Folio $newEnd)"
            return
        }
    }
    $lblSummaryInfo.Text = "Pulse 'PASO 1: Revisar y Simular' para leer los certificados de $mOrig $aOrig y preparar la lista."
}

function Load-SimulationGrid([bool]$forceExtract = $false) {
    $mOrig = [string]$cmbMes.SelectedItem
    $aOrig = [int]$numAnio.Value
    $next1 = Get-NextMonth $mOrig $aOrig
    $mDest = $next1.Mes; $aDest = $next1.Anio

    $invPath = Join-Path $tools ("inventario-{0}-{1}.csv" -f $mOrig.ToLowerInvariant(), $aOrig)
    if ($forceExtract -or -not (Test-Path -LiteralPath $invPath)) {
        $lblStatus.Text = "Leyendo certificados de $mOrig $aOrig en Word... Por favor espere..."
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()
        try {
            & (Join-Path $tools 'extraer-mes.ps1') -Mes $mOrig -Anio $aOrig -Raiz $raiz | Out-Null
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    if (-not (Test-Path -LiteralPath $invPath)) {
        [System.Windows.Forms.MessageBox]::Show("No se encontraron certificados para $mOrig $aOrig.", 'Aviso', 'OK', 'Warning') | Out-Null
        return
    }

    $rows = @(Import-Csv -LiteralPath $invPath)
    if ($rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No se encontraron archivos de $mOrig en las carpetas del año $aOrig.", 'Sin archivos', 'OK', 'Warning') | Out-Null
        return
    }

    # Si falta algun folio, ofrecer ingresarlo manualmente
    $faltantes = @($rows | Where-Object { $_.FolioNum -notmatch '^\d+$' })
    if ($faltantes.Count -gt 0) {
        $manualPath = Join-Path $tools 'folios-manuales.csv'
        $manualList = if (Test-Path -LiteralPath $manualPath) { @(Import-Csv -LiteralPath $manualPath) } else { @() }
        foreach ($f in $faltantes) {
            $val = Request-ManualFolio $f.Rel $f.FileName
            if ($val -match '^[0-9\.]+$') {
                $f.FolioRaw = $val
                $f.FolioNum = ($val -replace '\.', '')
                $manualList = @($manualList | Where-Object { $_.Rel -ne $f.Rel }) + @([pscustomobject]@{ Rel = $f.Rel; Folio = $val })
            }
        }
        if ($manualList.Count -gt 0) {
            $manualList | Export-Csv -LiteralPath $manualPath -NoTypeInformation -Encoding UTF8
        }
        $rows | Export-Csv -LiteralPath $invPath -NoTypeInformation -Encoding UTF8
    }

    $nums = @($rows | Where-Object { $_.FolioNum -match '^\d+$' } | ForEach-Object { [int]$_.FolioNum })
    $maxFolio = if ($nums.Count -gt 0) { ($nums | Measure-Object -Maximum).Maximum } else { 0 }
    $nextFolio = $maxFolio + 1

    $grid.Rows.Clear()
    $i = 0; $listos = 0; $existentes = 0
    foreach ($r in $rows) {
        $i++
        $newName = [regex]::Replace($r.FileName, $mOrig, $mDest, 'IgnoreCase').Replace("$aOrig", "$aDest")
        $srcDir = Split-Path $r.Path -Parent
        $dstDir = if ((Split-Path $srcDir -Leaf) -match '^\d{4}$' -and (Split-Path $srcDir -Leaf) -ne "$aDest") {
            Join-Path (Split-Path $srcDir -Parent) "$aDest"
        } else { $srcDir }
        $dstPath = Join-Path $dstDir $newName

        $newFolioStr = Format-Folio $nextFolio
        $nextFolio++

        $yaExiste = Test-Path -LiteralPath $dstPath
        $estado = if ($yaExiste) { $existentes++; 'YA EXISTE' } else { $listos++; 'LISTO PARA CREAR' }
        $emp = ($r.Rel.TrimStart('\') -split '\\')[0..1] -join ' \ '

        $rowIdx = $grid.Rows.Add($i, $emp, $newName, $r.FolioRaw, $newFolioStr, $estado)
        if ($yaExiste) {
            $grid.Rows[$rowIdx].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)
        }
    }

    Update-HeaderLabels
    $lblStatus.Text = "Revision lista: $($rows.Count) certificados encontrados ($listos por crear, $existentes ya creados)."
}

# Eventos
$cmbMes.Add_SelectedIndexChanged({ Update-HeaderLabels })
$numAnio.Add_ValueChanged({ Update-HeaderLabels })

$btnAuto.Add_Click({
    $det = Find-LatestBaseMonth
    $cmbMes.SelectedItem = $det.Mes
    $numAnio.Value = $det.Anio
    Update-HeaderLabels
    Load-SimulationGrid $false
})

$btnSimular.Add_Click({
    $r = [System.Windows.Forms.MessageBox]::Show("¿Desea revisar los documentos del mes seleccionado y mostrar la simulacion?`r`n(No se creara ni modificara ningun archivo).", 'Revisar y Simular', 'YesNo', 'Question')
    if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
        Load-SimulationGrid $true
    }
})

$btnGenerar.Add_Click({
    $mOrig = [string]$cmbMes.SelectedItem
    $aOrig = [int]$numAnio.Value
    $next1 = Get-NextMonth $mOrig $aOrig
    $mDest = $next1.Mes; $aDest = $next1.Anio
    $next2 = Get-NextMonth $mDest $aDest
    $mHasta = $next2.Mes; $aHasta = $next2.Anio

    Load-SimulationGrid $false
    if ($grid.Rows.Count -eq 0) { return }

    $msg = "¿Confirma que desea CREAR los certificados de $mDest $aDest?`r`n`r`n" +
           "- Mes origen: $mOrig $aOrig`r`n" +
           "- Mes nuevo: $mDest $aDest (Carpeta $aDest)`r`n" +
           "- Validez: Desde el 01 de $mDest $aDest hasta el 15 de $mHasta $aHasta`r`n" +
           "- Total de certificados: $($grid.Rows.Count)"
    $ans = [System.Windows.Forms.MessageBox]::Show($msg, "Confirmar Creacion de $mDest $aDest", 'YesNo', 'Question')
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnGenerar.Enabled = $false; $btnSimular.Enabled = $false; $btnDeshacer.Enabled = $false
    $prog.Minimum = 0; $prog.Maximum = $grid.Rows.Count; $prog.Value = 0

    try {
        $lblStatus.Text = "Creando y verificando los certificados de $mDest $aDest en Word... Por favor espere..."
        [System.Windows.Forms.Application]::DoEvents()

        & (Join-Path $tools 'generar-mes.ps1') -Origen $mOrig -OrigenAnio $aOrig -Destino $mDest -DestinoAnio $aDest -HastaMes $mHasta -HastaAnio $aHasta | Out-Null

        # Extraer automaticamente el inventario del nuevo mes para que quede listo para el proximo mes
        $prog.Value = [int]($grid.Rows.Count * 0.8)
        $lblStatus.Text = "Guardando inventario de $mDest $aDest para el proximo mes..."
        [System.Windows.Forms.Application]::DoEvents()
        & (Join-Path $tools 'extraer-mes.ps1') -Mes $mDest -Anio $aDest -Raiz $raiz | Out-Null

        $prog.Value = $grid.Rows.Count
        Load-SimulationGrid $false

        $logPath = Join-Path $tools ("log-{0}-{1}.csv" -f $mDest.ToLowerInvariant(), $aDest)
        $creados = 0; $errores = 0
        if (Test-Path -LiteralPath $logPath) {
            $lrows = @(Import-Csv -LiteralPath $logPath)
            $creados = @($lrows | Where-Object { $_.Status -eq 'CREATED-OK' -or $_.Status -eq 'SKIP-ya-existe' }).Count
            $errores = @($lrows | Where-Object { $_.Status -eq 'ERROR' }).Count
        }
        $lblStatus.Text = "¡Proceso terminado! Certificados de $mDest $aDest listos: $creados (Errores: $errores)."
        [System.Windows.Forms.MessageBox]::Show("¡Listo!`r`n`r`nCertificados de $mDest $aDest completados y verificados: $creados de $($grid.Rows.Count).`r`nErrores: $errores.", 'Proceso Completado', 'OK', 'Information') | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ocurrio un error: $($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnGenerar.Enabled = $true; $btnSimular.Enabled = $true; $btnDeshacer.Enabled = $true
    }
})

$btnCarpeta.Add_Click({
    if (Test-Path -LiteralPath $raiz) { Start-Process explorer.exe -ArgumentList "`"$raiz`"" }
})

$btnDeshacer.Add_Click({
    $mOrig = [string]$cmbMes.SelectedItem
    $aOrig = [int]$numAnio.Value
    $next1 = Get-NextMonth $mOrig $aOrig
    $mDest = $next1.Mes; $aDest = $next1.Anio

    $logPath = Join-Path $tools ("log-{0}-{1}.csv" -f $mDest.ToLowerInvariant(), $aDest)
    if (-not (Test-Path -LiteralPath $logPath)) {
        [System.Windows.Forms.MessageBox]::Show("No hay registro de certificados creados para $mDest $aDest.", 'Nada que deshacer', 'OK', 'Information') | Out-Null
        return
    }
    $logRows = @(Import-Csv -LiteralPath $logPath)
    $toDelete = @($logRows | Where-Object { $_.Status -match '^CREATED' -and $_.DestPath -and (Test-Path -LiteralPath $_.DestPath) })
    if ($toDelete.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No se encontraron archivos creados de $mDest $aDest para borrar.", 'Nada que deshacer', 'OK', 'Information') | Out-Null
        return
    }

    if (Confirm-DoubleVerification $mDest $aDest $toDelete.Count) {
        $borrados = 0
        foreach ($item in $toDelete) {
            if (Test-Path -LiteralPath $item.DestPath) {
                Remove-Item -LiteralPath $item.DestPath -Force
                $borrados++
            }
        }
        Remove-Item -LiteralPath $logPath -Force -ErrorAction SilentlyContinue
        $invDest = Join-Path $tools ("inventario-{0}-{1}.csv" -f $mDest.ToLowerInvariant(), $aDest)
        Remove-Item -LiteralPath $invDest -Force -ErrorAction SilentlyContinue

        Load-SimulationGrid $false
        $lblStatus.Text = "Se deshizo el mes $mDest $aDest ($borrados archivos eliminados)."
        [System.Windows.Forms.MessageBox]::Show("Se eliminaron correctamente los $borrados certificados de $mDest $aDest.", 'Deshacer completado', 'OK', 'Information') | Out-Null
    }
})

# Inicializar detectando el ultimo mes automaticamente
$init = Find-LatestBaseMonth
$cmbMes.SelectedItem = $init.Mes
$numAnio.Value = $init.Anio
Update-HeaderLabels
Load-SimulationGrid $false

[void]$form.ShowDialog()
