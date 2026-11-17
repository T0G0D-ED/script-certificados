Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

try {
    Add-Type -Name Win32 -Namespace Console -MemberDefinition '
        [DllImport("Kernel32.dll")] public static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    ' -ErrorAction SilentlyContinue
} catch {}

$ErrorActionPreference = 'Stop'
$tools = $PSScriptRoot
$proj = Split-Path $tools -Parent
$configFile = Join-Path $tools 'ruta-certificados.txt'

function Select-CertificatesFolder([string]$initialDir = '') {
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = 'Seleccione la carpeta principal donde estan las empresas y sus certificados (por ejemplo: CERTIFICADOS DICIEMBRE 2026)'
    $fbd.ShowNewFolderButton = $false
    if ($initialDir -and (Test-Path -LiteralPath $initialDir)) {
        $fbd.SelectedPath = $initialDir
    }
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $sel = $fbd.SelectedPath.TrimEnd('\')
        [IO.File]::WriteAllText($configFile, $sel, (New-Object Text.UTF8Encoding($false)))
        return $sel
    }
    return ''
}

$script:raiz = ''
if (Test-Path -LiteralPath $configFile) {
    $saved = (Get-Content -LiteralPath $configFile -TotalCount 1).Trim()
    if ($saved -and (Test-Path -LiteralPath $saved)) {
        $script:raiz = $saved.TrimEnd('\')
    }
}
if (-not $script:raiz) {
    $cand = Get-ChildItem -LiteralPath $proj -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^CERTIFICADOS' } | Select-Object -First 1
    if ($cand) {
        $script:raiz = $cand.FullName.TrimEnd('\')
    }
}
$raiz = $script:raiz

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
    if (-not $script:raiz -or -not (Test-Path -LiteralPath $script:raiz)) {
        return @{ Mes = 'ENERO'; Anio = 2027; Count = 0 }
    }
    # Primero revisar si ya existen inventarios CSV listos (instantaneo)
    $invFiles = @(Get-ChildItem -LiteralPath $tools -Filter 'inventario-*-20*.csv' -ErrorAction SilentlyContinue)
    $bestAnio = 0; $bestMesIdx = -1; $bestMes = ''
    foreach ($f in $invFiles) {
        if ($f.Name -match '^inventario-([a-z]+)-(\d{4})\.csv$') {
            $mCand = $Matches[1].ToUpperInvariant()
            $yCand = [int]$Matches[2]
            $mIdx = [array]::IndexOf($meses, $mCand)
            if ($mIdx -ge 0 -and ($yCand -gt $bestAnio -or ($yCand -eq $bestAnio -and $mIdx -gt $bestMesIdx))) {
                $bestAnio = $yCand; $bestMesIdx = $mIdx; $bestMes = $mCand
            }
        }
    }
    if ($bestAnio -gt 0) {
        return @{ Mes = $bestMes; Anio = $bestAnio; Count = 75 }
    }

    $yearDirs = @(Get-ChildItem -LiteralPath $script:raiz -Recurse -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{4}$' })
    $years = @($yearDirs | ForEach-Object { [int]$_.Name } | Sort-Object -Unique -Descending)
    foreach ($y in $years) {
        $files = @(Get-ChildItem -LiteralPath $script:raiz -Recurse -File -ErrorAction SilentlyContinue | Where-Object {
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

function Get-NextCompanyNumber() {
    $dirs = @(Get-ChildItem -LiteralPath $raiz -Directory)
    $nums = @($dirs | ForEach-Object { if ($_.Name -match '^(\d+)') { [int]$Matches[1] } })
    if ($nums.Count -eq 0) { return 1 }
    return (($nums | Measure-Object -Maximum).Maximum + 1)
}

function Show-NewCompanyOrBranchDialog([string]$mesOrig, [int]$anioOrig) {
    $invPath = Join-Path $tools ("inventario-{0}-{1}.csv" -f $mesOrig.ToLowerInvariant(), $anioOrig)
    $maxFolio = 37800
    if (Test-Path -LiteralPath $invPath) {
        $rws = @(Import-Csv -LiteralPath $invPath)
        $ns = @($rws | Where-Object { $_.FolioNum -match '^\d+$' } | ForEach-Object { [int]$_.FolioNum })
        if ($ns.Count -gt 0) { $maxFolio = ($ns | Measure-Object -Maximum).Maximum }
    }
    $suggestedFolio = Format-Folio ($maxFolio + 1)
    $nextNum = Get-NextCompanyNumber

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Agregar Nueva Empresa o Sucursal'
    $dlg.Size = New-Object System.Drawing.Size(750, 680)
    $dlg.MinimumSize = New-Object System.Drawing.Size(720, 500)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'Sizable'
    $dlg.MaximizeBox = $true
    $dlg.MinimizeBox = $false
    $dlg.AutoScroll = $true
    $dlg.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $dlg.Font = New-Object System.Drawing.Font('Segoe UI', 12)

    $lblHeader = New-Object System.Windows.Forms.Label
    $lblHeader.Text = "Crear carpeta y primer certificado base en $mesOrig $anioOrig"
    $lblHeader.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
    $lblHeader.ForeColor = [System.Drawing.Color]::FromArgb(20, 60, 140)
    $lblHeader.Location = New-Object System.Drawing.Point(22, 16)
    $lblHeader.Size = New-Object System.Drawing.Size(660, 32)
    $dlg.Controls.Add($lblHeader)

    $rbEmpresa = New-Object System.Windows.Forms.RadioButton
    $rbEmpresa.Text = 'Crear una EMPRESA NUEVA'
    $rbEmpresa.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
    $rbEmpresa.Location = New-Object System.Drawing.Point(25, 58)
    $rbEmpresa.Size = New-Object System.Drawing.Size(300, 30)
    $rbEmpresa.Checked = $true
    $dlg.Controls.Add($rbEmpresa)

    $rbSucursal = New-Object System.Windows.Forms.RadioButton
    $rbSucursal.Text = 'Agregar SUCURSAL a empresa existente'
    $rbSucursal.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
    $rbSucursal.Location = New-Object System.Drawing.Point(335, 58)
    $rbSucursal.Size = New-Object System.Drawing.Size(355, 30)
    $dlg.Controls.Add($rbSucursal)

    # Panel Empresa
    $lblEmp = New-Object System.Windows.Forms.Label
    $lblEmp.Text = "Nombre de carpeta de la nueva empresa (ej: $nextNum MI EMPRESA):"
    $lblEmp.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $lblEmp.Location = New-Object System.Drawing.Point(22, 102)
    $lblEmp.Size = New-Object System.Drawing.Size(650, 26)
    $dlg.Controls.Add($lblEmp)

    $txtNuevaEmp = New-Object System.Windows.Forms.TextBox
    $txtNuevaEmp.Text = "$nextNum "
    $txtNuevaEmp.Font = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
    $txtNuevaEmp.Location = New-Object System.Drawing.Point(22, 130)
    $txtNuevaEmp.Size = New-Object System.Drawing.Size(650, 34)
    $dlg.Controls.Add($txtNuevaEmp)

    $cmbEmpExistente = New-Object System.Windows.Forms.ComboBox
    $cmbEmpExistente.DropDownStyle = 'DropDownList'
    $cmbEmpExistente.Font = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
    $cmbEmpExistente.Location = New-Object System.Drawing.Point(22, 130)
    $cmbEmpExistente.Size = New-Object System.Drawing.Size(650, 34)
    $cmbEmpExistente.Visible = $false
    Get-ChildItem -LiteralPath $raiz -Directory | Sort-Object Name | ForEach-Object { [void]$cmbEmpExistente.Items.Add($_.Name) }
    if ($cmbEmpExistente.Items.Count -gt 0) { $cmbEmpExistente.SelectedIndex = 0 }
    $dlg.Controls.Add($cmbEmpExistente)

    # Sucursal
    $lblSuc = New-Object System.Windows.Forms.Label
    $lblSuc.Text = 'Nombre de Sucursal (opcional si la empresa no tiene sucursales):'
    $lblSuc.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $lblSuc.Location = New-Object System.Drawing.Point(22, 178)
    $lblSuc.Size = New-Object System.Drawing.Size(650, 26)
    $dlg.Controls.Add($lblSuc)

    $txtSucursal = New-Object System.Windows.Forms.TextBox
    $txtSucursal.Font = New-Object System.Drawing.Font('Segoe UI', 13)
    $txtSucursal.Location = New-Object System.Drawing.Point(22, 206)
    $txtSucursal.Size = New-Object System.Drawing.Size(650, 34)
    $dlg.Controls.Add($txtSucursal)

    # Datos dentro del certificado Word
    $lblInm = New-Object System.Windows.Forms.Label
    $lblInm.Text = 'Inmueble Tratado (aparece dentro del certificado):'
    $lblInm.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $lblInm.Location = New-Object System.Drawing.Point(22, 254)
    $lblInm.Size = New-Object System.Drawing.Size(650, 26)
    $dlg.Controls.Add($lblInm)

    $txtInmueble = New-Object System.Windows.Forms.TextBox
    $txtInmueble.Font = New-Object System.Drawing.Font('Segoe UI', 13)
    $txtInmueble.Location = New-Object System.Drawing.Point(22, 282)
    $txtInmueble.Size = New-Object System.Drawing.Size(650, 34)
    $dlg.Controls.Add($txtInmueble)

    $lblDir = New-Object System.Windows.Forms.Label
    $lblDir.Text = 'Direccion:'
    $lblDir.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $lblDir.Location = New-Object System.Drawing.Point(22, 330)
    $lblDir.Size = New-Object System.Drawing.Size(410, 26)
    $dlg.Controls.Add($lblDir)

    $txtDireccion = New-Object System.Windows.Forms.TextBox
    $txtDireccion.Font = New-Object System.Drawing.Font('Segoe UI', 13)
    $txtDireccion.Location = New-Object System.Drawing.Point(22, 358)
    $txtDireccion.Size = New-Object System.Drawing.Size(410, 34)
    $dlg.Controls.Add($txtDireccion)

    $lblCom = New-Object System.Windows.Forms.Label
    $lblCom.Text = 'Comuna:'
    $lblCom.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $lblCom.Location = New-Object System.Drawing.Point(450, 330)
    $lblCom.Size = New-Object System.Drawing.Size(222, 26)
    $dlg.Controls.Add($lblCom)

    $txtComuna = New-Object System.Windows.Forms.TextBox
    $txtComuna.Font = New-Object System.Drawing.Font('Segoe UI', 13)
    $txtComuna.Location = New-Object System.Drawing.Point(450, 358)
    $txtComuna.Size = New-Object System.Drawing.Size(222, 34)
    $dlg.Controls.Add($txtComuna)

    # Folio Inicial
    $lblFolio = New-Object System.Windows.Forms.Label
    $lblFolio.Text = "Folio para su certificado de $mesOrig $anioOrig (correlativo sugerido):"
    $lblFolio.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $lblFolio.Location = New-Object System.Drawing.Point(22, 406)
    $lblFolio.Size = New-Object System.Drawing.Size(650, 28)
    $dlg.Controls.Add($lblFolio)

    $txtFolio = New-Object System.Windows.Forms.TextBox
    $txtFolio.Text = $suggestedFolio
    $txtFolio.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
    $txtFolio.Location = New-Object System.Drawing.Point(22, 436)
    $txtFolio.Size = New-Object System.Drawing.Size(240, 36)
    $dlg.Controls.Add($txtFolio)

    $chkAbrir = New-Object System.Windows.Forms.CheckBox
    $chkAbrir.Text = 'Abrir el certificado creado en Word al terminar para revisarlo'
    $chkAbrir.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
    $chkAbrir.Checked = $true
    $chkAbrir.Location = New-Object System.Drawing.Point(22, 486)
    $chkAbrir.Size = New-Object System.Drawing.Size(650, 30)
    $dlg.Controls.Add($chkAbrir)

    $rbEmpresa.Add_CheckedChanged({
        if ($rbEmpresa.Checked) {
            $lblEmp.Text = "Nombre de carpeta de la nueva empresa (ej: $nextNum MI EMPRESA):"
            $txtNuevaEmp.Visible = $true
            $cmbEmpExistente.Visible = $false
            $lblSuc.Text = 'Nombre de Sucursal (opcional si la empresa no tiene sucursales):'
        } else {
            $lblEmp.Text = 'Seleccione la Empresa existente:'
            $txtNuevaEmp.Visible = $false
            $cmbEmpExistente.Visible = $true
            $lblSuc.Text = 'Nombre de la nueva Sucursal (obligatorio):'
        }
    })

    $btnCrear = New-Object System.Windows.Forms.Button
    $btnCrear.Text = 'CREAR CARPETA Y CERTIFICADO'
    $btnCrear.BackColor = [System.Drawing.Color]::FromArgb(25, 135, 60)
    $btnCrear.ForeColor = [System.Drawing.Color]::White
    $btnCrear.FlatStyle = 'Flat'
    $btnCrear.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
    $btnCrear.Location = New-Object System.Drawing.Point(22, 532)
    $btnCrear.Size = New-Object System.Drawing.Size(360, 52)
    $dlg.Controls.Add($btnCrear)

    $btnCerrar = New-Object System.Windows.Forms.Button
    $btnCerrar.Text = 'Cancelar'
    $btnCerrar.Font = New-Object System.Drawing.Font('Segoe UI', 12)
    $btnCerrar.Location = New-Object System.Drawing.Point(405, 532)
    $btnCerrar.Size = New-Object System.Drawing.Size(267, 52)
    $btnCerrar.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCerrar)

    $script:createdResult = $null
    $btnCrear.Add_Click({
        $modo = if ($rbEmpresa.Checked) { 'Empresa' } else { 'Sucursal' }
        $empName = if ($rbEmpresa.Checked) { $txtNuevaEmp.Text.Trim().ToUpperInvariant() } else { [string]$cmbEmpExistente.SelectedItem }
        $sucName = $txtSucursal.Text.Trim().ToUpperInvariant()
        if ($modo -eq 'Empresa' -and ($empName -match '^\d+$' -or $empName.Length -lt 3)) {
            [System.Windows.Forms.MessageBox]::Show('Escriba el nombre de la nueva empresa (por ejemplo: 42 PANADERIA CENTRAL).', 'Falta nombre', 'OK', 'Warning') | Out-Null
            return
        }
        if ($modo -eq 'Sucursal' -and -not $sucName) {
            [System.Windows.Forms.MessageBox]::Show('Escriba el nombre de la nueva sucursal.', 'Falta sucursal', 'OK', 'Warning') | Out-Null
            return
        }
        if ($txtFolio.Text.Trim() -notmatch '^[0-9\.]+$') {
            [System.Windows.Forms.MessageBox]::Show('El folio debe ser un numero valido (por ejemplo 37.874).', 'Folio invalido', 'OK', 'Warning') | Out-Null
            return
        }
        $inmText = if ($txtInmueble.Text.Trim()) { $txtInmueble.Text.Trim() } elseif ($sucName) { $sucName } else { ($empName -replace '^\d+[\-\s]*', '').Trim() }

        $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $btnCrear.Enabled = $false
        try {
            $script:createdResult = & (Join-Path $tools 'nueva-empresa-sucursal.ps1') `
                -Modo $modo `
                -CarpetaEmpresa $empName `
                -NombreSucursal $sucName `
                -Inmueble $inmText `
                -Direccion $txtDireccion.Text.Trim() `
                -Comuna $txtComuna.Text.Trim() `
                -MesBase $mesOrig `
                -AnioBase $anioOrig `
                -FolioInicial $txtFolio.Text.Trim() `
                -Raiz $raiz `
                -AbrirEnWord:($chkAbrir.Checked)
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $dlg.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("No se pudo crear: $($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
        } finally {
            $dlg.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnCrear.Enabled = $true
        }
    })

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $script:createdResult
    }
    return $null
}

# ==================== VENTANA PRINCIPAL ====================
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Generador Mensual de Certificados - Parra & Mauro Ltda.'
$form.Size = New-Object System.Drawing.Size(1120, 780)
$icoFile = Join-Path $tools 'icono-certificados.ico'
if (Test-Path -LiteralPath $icoFile) {
    try { $form.Icon = New-Object System.Drawing.Icon($icoFile) } catch {}
}
# Asegurar que el acceso directo .lnk tenga el icono personalizado de esta carpeta
try {
    $lnkLocal = Join-Path $proj 'Abrir Generador de Certificados.lnk'
    if ((Test-Path -LiteralPath $lnkLocal) -and (Test-Path -LiteralPath $icoFile)) {
        $wsh = New-Object -ComObject WScript.Shell
        $sc = $wsh.CreateShortcut($lnkLocal)
        if ($sc.IconLocation -ne "$icoFile,0") {
            $sc.IconLocation = "$icoFile,0"
            $sc.Save()
        }
    }
} catch {}
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
$btnAuto.Text = 'Detectar ultimo mes'
$btnAuto.Font = New-Object System.Drawing.Font('Segoe UI', 11)
$btnAuto.Location = New-Object System.Drawing.Point(705, 11)
$btnAuto.Size = New-Object System.Drawing.Size(175, 36)
$pnlTop.Controls.Add($btnAuto)

$btnCambiarCarpeta = New-Object System.Windows.Forms.Button
$btnCambiarCarpeta.Text = 'Cargar otra carpeta...'
$btnCambiarCarpeta.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
$btnCambiarCarpeta.Location = New-Object System.Drawing.Point(888, 11)
$btnCambiarCarpeta.Size = New-Object System.Drawing.Size(168, 36)
$pnlTop.Controls.Add($btnCambiarCarpeta)

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
$btnSimular.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$btnSimular.BackColor = [System.Drawing.Color]::FromArgb(30, 100, 200)
$btnSimular.ForeColor = [System.Drawing.Color]::White
$btnSimular.FlatStyle = 'Flat'
$btnSimular.Location = New-Object System.Drawing.Point(16, 178)
$btnSimular.Size = New-Object System.Drawing.Size(230, 58)
$form.Controls.Add($btnSimular)

$btnGenerar = New-Object System.Windows.Forms.Button
$btnGenerar.Text = 'PASO 2: Crear Certificados'
$btnGenerar.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$btnGenerar.BackColor = [System.Drawing.Color]::FromArgb(25, 135, 60)
$btnGenerar.ForeColor = [System.Drawing.Color]::White
$btnGenerar.FlatStyle = 'Flat'
$btnGenerar.Location = New-Object System.Drawing.Point(256, 178)
$btnGenerar.Size = New-Object System.Drawing.Size(240, 58)
$form.Controls.Add($btnGenerar)

$btnNuevaEmp = New-Object System.Windows.Forms.Button
$btnNuevaEmp.Text = '+ Nueva Empresa / Sucursal'
$btnNuevaEmp.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
$btnNuevaEmp.BackColor = [System.Drawing.Color]::FromArgb(105, 65, 198)
$btnNuevaEmp.ForeColor = [System.Drawing.Color]::White
$btnNuevaEmp.FlatStyle = 'Flat'
$btnNuevaEmp.Location = New-Object System.Drawing.Point(506, 178)
$btnNuevaEmp.Size = New-Object System.Drawing.Size(220, 58)
$form.Controls.Add($btnNuevaEmp)

$btnCarpeta = New-Object System.Windows.Forms.Button
$btnCarpeta.Text = 'Abrir Carpeta'
$btnCarpeta.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
$btnCarpeta.BackColor = [System.Drawing.Color]::FromArgb(230, 235, 242)
$btnCarpeta.Location = New-Object System.Drawing.Point(736, 178)
$btnCarpeta.Size = New-Object System.Drawing.Size(160, 58)
$form.Controls.Add($btnCarpeta)

$btnDeshacer = New-Object System.Windows.Forms.Button
$btnDeshacer.Text = 'Deshacer Mes Creado'
$btnDeshacer.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
$btnDeshacer.BackColor = [System.Drawing.Color]::FromArgb(250, 225, 225)
$btnDeshacer.ForeColor = [System.Drawing.Color]::FromArgb(160, 20, 20)
$btnDeshacer.FlatStyle = 'Flat'
$btnDeshacer.Location = New-Object System.Drawing.Point(906, 178)
$btnDeshacer.Size = New-Object System.Drawing.Size(180, 58)
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

$colAbrir = New-Object System.Windows.Forms.DataGridViewButtonColumn
$colAbrir.Name = 'Abrir'
$colAbrir.HeaderText = 'Word'
$colAbrir.Text = '▶ Abrir'
$colAbrir.UseColumnTextForButtonValue = $true
$colAbrir.FlatStyle = 'Flat'
$colAbrir.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(34, 150, 70)
$colAbrir.DefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$colAbrir.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(25, 120, 55)
$colAbrir.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
$colAbrir.DefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
[void]$grid.Columns.Add($colAbrir)

[void]$grid.Columns.Add('Num', '#')
[void]$grid.Columns.Add('Empresa', 'Empresa / Sucursal')
[void]$grid.Columns.Add('ArchivoNuevo', 'Certificado a Crear')
[void]$grid.Columns.Add('FolioAnt', 'Folio Anterior')
[void]$grid.Columns.Add('FolioNuevo', 'Folio Nuevo')
[void]$grid.Columns.Add('Estado', 'Estado')
$grid.Columns['Abrir'].FillWeight = 9
$grid.Columns['Num'].FillWeight = 6
$grid.Columns['Empresa'].FillWeight = 32
$grid.Columns['ArchivoNuevo'].FillWeight = 27
$grid.Columns['FolioAnt'].FillWeight = 11
$grid.Columns['FolioNuevo'].FillWeight = 11
$grid.Columns['Estado'].FillWeight = 16
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

function Set-BusyBanner([string]$msg, [int]$val = 0, [int]$max = 100) {
    $lblStatus.BackColor = [System.Drawing.Color]::FromArgb(255, 243, 190)
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(140, 70, 0)
    $lblStatus.Text = "  [TRABAJANDO - NO CIERRE LA VENTANA]  $msg"
    if ($max -gt 0) {
        $prog.Maximum = $max
        $prog.Value = [Math]::Min($val, $max)
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-IdleBanner([string]$msg) {
    $lblStatus.BackColor = [System.Drawing.Color]::Transparent
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
    $lblStatus.Text = $msg
    [System.Windows.Forms.Application]::DoEvents()
}

function Load-SimulationGrid([bool]$forceExtract = $false) {
    if (-not $script:raiz -or -not (Test-Path -LiteralPath $script:raiz)) { return }
    $mOrig = [string]$cmbMes.SelectedItem
    $aOrig = [int]$numAnio.Value
    if (-not $mOrig) { return }
    $next1 = Get-NextMonth $mOrig $aOrig
    $mDest = $next1.Mes; $aDest = $next1.Anio

    $invPath = Join-Path $tools ("inventario-{0}-{1}.csv" -f $mOrig.ToLowerInvariant(), $aOrig)
    # Si el inventario existe pero sus archivos no existen en esta carpeta raiz (ej. se abrio en otro PC), re-extraer
    $needReextract = $false
    if (-not $forceExtract -and (Test-Path -LiteralPath $invPath)) {
        $testRows = @(Import-Csv -LiteralPath $invPath)
        if ($testRows.Count -gt 0) {
            $firstPath = Join-Path $script:raiz $testRows[0].Rel.TrimStart('\')
            if (-not (Test-Path -LiteralPath $firstPath)) { $needReextract = $true }
        }
    }
    if ($forceExtract -or $needReextract -or -not (Test-Path -LiteralPath $invPath)) {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $btnGenerar.Enabled = $false; $btnSimular.Enabled = $false; $btnDeshacer.Enabled = $false; $btnCambiarCarpeta.Enabled = $false
        Set-BusyBanner "Iniciando Microsoft Word para leer $mOrig $aOrig... Por favor espere..." 0 75
        try {
            $cb = {
                param($idx, $tot, $rel)
                $pct = [int](($idx / [double]$tot) * 100)
                $shortName = Split-Path $rel -Leaf
                Set-BusyBanner "Leyendo documento $idx de $tot ($pct%): $shortName" $idx $tot
            }
            & (Join-Path $tools 'extraer-mes.ps1') -Mes $mOrig -Anio $aOrig -Raiz $script:raiz -OnProgress $cb | Out-Null
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $btnGenerar.Enabled = $true; $btnSimular.Enabled = $true; $btnDeshacer.Enabled = $true; $btnCambiarCarpeta.Enabled = $true
        }
    }

    if (-not (Test-Path -LiteralPath $invPath)) {
        [System.Windows.Forms.MessageBox]::Show("No se encontraron certificados para $mOrig $aOrig.", 'Aviso', 'OK', 'Warning') | Out-Null
        return
    }

    $rows = @(Import-Csv -LiteralPath $invPath)
    foreach ($r in $rows) {
        if ($r.Rel) { $r.Path = Join-Path $script:raiz $r.Rel.TrimStart('\') }
    }
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

        $rowIdx = $grid.Rows.Add('▶ Abrir', $i, $emp, $newName, $r.FolioRaw, $newFolioStr, $estado)
        $grid.Rows[$rowIdx].Tag = @{ SrcPath = $r.Path; DstPath = $dstPath }
        if ($yaExiste) {
            $grid.Rows[$rowIdx].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 245, 230)
        }
        $grid.Rows[$rowIdx].Cells['Abrir'].Style.BackColor = [System.Drawing.Color]::FromArgb(34, 150, 70)
        $grid.Rows[$rowIdx].Cells['Abrir'].Style.ForeColor = [System.Drawing.Color]::White
    }

    Update-HeaderLabels
    $prog.Maximum = $rows.Count; $prog.Value = $rows.Count
    Set-IdleBanner "Carpeta: $(Split-Path $script:raiz -Leaf)  |  $($rows.Count) certificados ($listos por crear, $existentes ya creados). Clic en '▶ Abrir' para ver en Word."
}

function Open-RowDocument([int]$rowIndex) {
    if ($rowIndex -lt 0 -or $rowIndex -ge $grid.Rows.Count) { return }
    $tag = $grid.Rows[$rowIndex].Tag
    if (-not $tag) { return }
    if ($tag.DstPath -and (Test-Path -LiteralPath $tag.DstPath)) {
        Start-Process -FilePath $tag.DstPath
    } elseif ($tag.SrcPath -and (Test-Path -LiteralPath $tag.SrcPath)) {
        Start-Process -FilePath $tag.SrcPath
    }
}

# Clic en el boton verde '▶ Abrir' (columna 0) abre el certificado con un solo clic
$grid.Add_CellContentClick({
    param($sender, $e)
    if ($e.RowIndex -ge 0 -and $grid.Columns[$e.ColumnIndex].Name -eq 'Abrir') {
        Open-RowDocument $e.RowIndex
    }
})

# Doble clic en cualquier parte de la fila tambien abre el certificado en Word
$grid.Add_CellDoubleClick({
    param($sender, $e)
    Open-RowDocument $e.RowIndex
})

# Eventos
$cmbMes.Add_SelectedIndexChanged({ Update-HeaderLabels; Load-SimulationGrid $false })
$numAnio.Add_ValueChanged({ Update-HeaderLabels; Load-SimulationGrid $false })

$btnCambiarCarpeta.Add_Click({
    $nueva = Select-CertificatesFolder $script:raiz
    if ($nueva) {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        Set-BusyBanner "Analizando carpeta seleccionada ($(Split-Path $nueva -Leaf))... Por favor espere..." 10 100
        try {
            $script:raiz = $nueva
            $raiz = $nueva
            $det = Find-LatestBaseMonth
            $cmbMes.SelectedItem = $det.Mes
            $numAnio.Value = $det.Anio
            Update-HeaderLabels
            # Primero intentar carga rapida usando el inventario existente si coincide con la carpeta
            Load-SimulationGrid $false
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }
})

$btnAuto.Add_Click({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Set-BusyBanner "Buscando el ultimo mes en las carpetas... Por favor espere..." 20 100
    try {
        $det = Find-LatestBaseMonth
        $cmbMes.SelectedItem = $det.Mes
        $numAnio.Value = $det.Anio
        Update-HeaderLabels
        Load-SimulationGrid $false
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnSimular.Add_Click({
    $mOrig = [string]$cmbMes.SelectedItem
    $aOrig = [int]$numAnio.Value
    $invPath = Join-Path $tools ("inventario-{0}-{1}.csv" -f $mOrig.ToLowerInvariant(), $aOrig)
    if (Test-Path -LiteralPath $invPath) {
        $r = [System.Windows.Forms.MessageBox]::Show(
            "El inventario de $mOrig $aOrig ya esta cargado en pantalla.`r`n`r`n" +
            "- Pulse 'SI' si modifico algun certificado manualmente en Word y desea volver a leer los 75 documentos desde cero (toma ~1 a 2 minutos).`r`n" +
            "- Pulse 'NO' para hacer una revision rapida instantanea.",
            'Revisar y Simular', 'YesNoCancel', 'Question', 'Button2'
        )
        if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
            Load-SimulationGrid $true
        } elseif ($r -eq [System.Windows.Forms.DialogResult]::No) {
            Load-SimulationGrid $false
        }
    } else {
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
    $btnGenerar.Enabled = $false; $btnSimular.Enabled = $false; $btnDeshacer.Enabled = $false; $btnCambiarCarpeta.Enabled = $false
    $totalDocs = $grid.Rows.Count
    $prog.Minimum = 0; $prog.Maximum = $totalDocs; $prog.Value = 0

    try {
        Set-BusyBanner "Iniciando Microsoft Word para crear $mDest $aDest (0 de $totalDocs)... Por favor espere..." 0 $totalDocs

        $cbGen = {
            param($idx, $tot, $rel)
            $pct = [int](($idx / [double]$tot) * 100)
            $shortName = Split-Path $rel -Leaf
            Set-BusyBanner "Creando y verificando $idx de $tot ($pct%): $shortName" $idx $tot
            if ($idx -le $grid.Rows.Count) {
                $grid.Rows[$idx - 1].Cells['Estado'].Value = 'CREANDO...'
                $grid.FirstDisplayedScrollingRowIndex = [Math]::Max(0, $idx - 5)
            }
        }
        & (Join-Path $tools 'generar-mes.ps1') -Origen $mOrig -OrigenAnio $aOrig -Destino $mDest -DestinoAnio $aDest -HastaMes $mHasta -HastaAnio $aHasta -Raiz $script:raiz -OnProgress $cbGen | Out-Null

        # Construir el inventario del nuevo mes instantaneamente a partir del log (sin volver a abrir 75 veces Word!)
        $logPath = Join-Path $tools ("log-{0}-{1}.csv" -f $mDest.ToLowerInvariant(), $aDest)
        $creados = 0; $errores = 0
        if (Test-Path -LiteralPath $logPath) {
            $lrows = @(Import-Csv -LiteralPath $logPath)
            $creados = @($lrows | Where-Object { $_.Status -eq 'CREATED-OK' -or $_.Status -eq 'SKIP-ya-existe' }).Count
            $errores = @($lrows | Where-Object { $_.Status -eq 'ERROR' }).Count

            if ($errores -eq 0 -and $lrows.Count -gt 0) {
                $newInvPath = Join-Path $tools ("inventario-{0}-{1}.csv" -f $mDest.ToLowerInvariant(), $aDest)
                $newInvRows = @()
                foreach ($lr in $lrows) {
                    $newRel = $lr.DestPath.Substring($script:raiz.Length)
                    $newInvRows += [pscustomobject]@{
                        Path = $lr.DestPath
                        Rel = $newRel
                        FileName = $lr.Dest
                        FolioRaw = $lr.NewFolio
                        FolioNum = ($lr.NewFolio -replace '\.', '')
                        Desde = "Desde el 01 de $mDest $aDest"
                        Hasta = "Hasta el 15 de $mHasta $aHasta"
                        MesLabel = "$mDest / $aDest"
                        Error = ''
                    }
                }
                $newInvRows | Export-Csv -LiteralPath $newInvPath -NoTypeInformation -Encoding UTF8
            }
        }

        $prog.Value = $totalDocs
        Load-SimulationGrid $false

        Set-IdleBanner "¡Proceso terminado! Certificados de $mDest $aDest listos: $creados (Errores: $errores)."
        [System.Windows.Forms.MessageBox]::Show("¡Listo!`r`n`r`nCertificados de $mDest $aDest completados y verificados: $creados de $totalDocs.`r`nErrores: $errores.", 'Proceso Completado', 'OK', 'Information') | Out-Null
    } catch {
        Set-IdleBanner "Error durante el proceso: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Ocurrio un error: $($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnGenerar.Enabled = $true; $btnSimular.Enabled = $true; $btnDeshacer.Enabled = $true; $btnCambiarCarpeta.Enabled = $true
    }
})

$btnCarpeta.Add_Click({
    if (Test-Path -LiteralPath $script:raiz) { Start-Process explorer.exe -ArgumentList "`"$script:raiz`"" }
})

$btnNuevaEmp.Add_Click({
    $mOrig = [string]$cmbMes.SelectedItem
    $aOrig = [int]$numAnio.Value
    $res = Show-NewCompanyOrBranchDialog $mOrig $aOrig
    if ($res) {
        Load-SimulationGrid $false
        $lblStatus.Text = "Agregado: $($res.Rel) con folio $($res.FolioAsignado). Total ahora: $($res.TotalEnInventario) certificados."
        [System.Windows.Forms.MessageBox]::Show(
            "¡Empresa / Sucursal creada con exito!`r`n`r`n" +
            "- Carpeta: $($res.CarpetaCreada)`r`n" +
            "- Certificado inicial: $($res.ArchivoCreado)`r`n" +
            "- Folio asignado en $mOrig ${aOrig}: $($res.FolioAsignado)`r`n" +
            "- Total de certificados en la secuencia: $($res.TotalEnInventario)`r`n`r`n" +
            "Ya quedo incluida automaticamente para las proximas generaciones mensuales.",
            'Creado con exito', 'OK', 'Information'
        ) | Out-Null
    }
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

# Inicializar al mostrarse la ventana (abre en menos de 1 segundo)
$cmbMes.SelectedItem = 'ENERO'
$numAnio.Value = 2027

$form.Add_Shown({
    # Ocultar la consola de fondo DESPUES de que la ventana grafica ya esta visible
    try {
        $consolePtr = [Console.Win32]::GetConsoleWindow()
        if ($consolePtr -ne [IntPtr]::Zero) { [void][Console.Win32]::ShowWindow($consolePtr, 0) }
    } catch {}
    $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
    $form.TopMost = $true
    $form.Activate()
    $form.TopMost = $false
    [System.Windows.Forms.Application]::DoEvents()
    if (-not $script:raiz -or -not (Test-Path -LiteralPath $script:raiz)) {
        [System.Windows.Forms.MessageBox]::Show(
            $form,
            "Bienvenido al Generador de Certificados.`r`n`r`nPor favor seleccione en la siguiente ventana la carpeta donde tiene guardados sus certificados.",
            'Seleccionar Carpeta de Certificados', 'OK', 'Information'
        ) | Out-Null
        $sel = Select-CertificatesFolder $proj
        if ($sel) {
            $script:raiz = $sel
            $raiz = $sel
        } else {
            $lblStatus.Text = "Sin carpeta seleccionada. Pulse 'Cargar otra carpeta...' arriba a la derecha para elegir sus certificados."
            return
        }
    }
    $init = Find-LatestBaseMonth
    $cmbMes.SelectedItem = $init.Mes
    $numAnio.Value = $init.Anio
    Update-HeaderLabels
    $invInit = Join-Path $tools ("inventario-{0}-{1}.csv" -f $init.Mes.ToLowerInvariant(), $init.Anio)
    if (Test-Path -LiteralPath $invInit) {
        Load-SimulationGrid $false
    } else {
        $lblStatus.Text = "Carpeta cargada: $(Split-Path $script:raiz -Leaf). Pulse 'PASO 1: Revisar y Simular' para leer los certificados de $($init.Mes) $($init.Anio)."
    }
})

[void]$form.ShowDialog()
