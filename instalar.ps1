# Instalador automatico - Generador Mensual de Certificados
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Windows.Forms

Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host " Instalando Generador Mensual de Certificados..." -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan

$desktop = [Environment]::GetFolderPath('Desktop')
$installDir = Join-Path $desktop 'Generador-Certificados'
$zipUrl = 'https://github.com/T0G0D-ED/script-certificados/releases/download/v1.0.0/Generador-Certificados-App.zip'
$tempZip = Join-Path $env:TEMP ("Generador-Certificados-" + [guid]::NewGuid().ToString('N') + ".zip")

Write-Host "1. Descargando aplicacion desde GitHub..."
Invoke-WebRequest -Uri $zipUrl -OutFile $tempZip -UseBasicParsing
Unblock-File -LiteralPath $tempZip -ErrorAction SilentlyContinue

Write-Host "2. Instalando en: $installDir"
if (-not (Test-Path -LiteralPath $installDir)) {
    New-Item -ItemType Directory -Path $installDir -Force | Out-Null
}
$tempExtract = Join-Path $env:TEMP ("cert_extract_" + [guid]::NewGuid().ToString('N'))
[System.IO.Compression.ZipFile]::ExtractToDirectory($tempZip, $tempExtract)
Copy-Item -Path "$tempExtract\*" -Destination $installDir -Recurse -Force
Remove-Item -LiteralPath $tempExtract -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $tempZip -Force -ErrorAction SilentlyContinue

# 3. Reescribir los scripts localmente para que Windows los registre como archivos locales
Get-ChildItem -LiteralPath (Join-Path $installDir '_herramientas') -Filter '*.ps1' -File | ForEach-Object {
    $txt = [System.IO.File]::ReadAllText($_.FullName)
    Remove-Item -LiteralPath $_.FullName -Force
    [System.IO.File]::WriteAllText($_.FullName, $txt, (New-Object System.Text.UTF8Encoding($true)))
}
Get-ChildItem -LiteralPath $installDir -Recurse -File | ForEach-Object {
    Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue
}

# 4. Crear Acceso Directo en el Escritorio y en la carpeta usando powershell.exe (firmado por Microsoft)
# con el icono Opcion 1 (icono-certificados.ico), para que OneDrive / Smart App Control nunca lo bloqueen.
$icoPath = Join-Path $installDir '_herramientas\icono-certificados.ico'
$appScript = Join-Path $installDir '_herramientas\app-certificados.ps1'
$psSigned = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"

$wsh = New-Object -ComObject WScript.Shell
foreach ($lnkDest in @(
    (Join-Path $desktop 'Generador de Certificados.lnk'),
    (Join-Path $installDir 'Abrir Generador de Certificados.lnk')
)) {
    $sc = $wsh.CreateShortcut($lnkDest)
    $sc.TargetPath = $psSigned
    $sc.Arguments = "-Sta -NoProfile -ExecutionPolicy Bypass -File `"$appScript`""
    $sc.WorkingDirectory = $installDir
    if (Test-Path -LiteralPath $icoPath) {
        $sc.IconLocation = "$icoPath,0"
    }
    $sc.Description = 'Abrir Generador Mensual de Certificados'
    $sc.Save()
}

Write-Host ""
Write-Host "INSTALACION COMPLETADA CON EXITO!" -ForegroundColor Green
Write-Host "- Se creo el icono 'Generador de Certificados' en su Escritorio." -ForegroundColor Green
Write-Host "- Abriendo la aplicacion..." -ForegroundColor Yellow

& $psSigned -Sta -NoProfile -ExecutionPolicy Bypass -File $appScript
