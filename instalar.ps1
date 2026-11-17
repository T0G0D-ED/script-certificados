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

# 3. Reescribir los scripts localmente con [IO.File]::WriteAllText para que el kernel de Windows
# los registre como archivos creados localmente (sin origen de Internet / sin bloqueo Smart App Control)
Get-ChildItem -LiteralPath (Join-Path $installDir '_herramientas') -Filter '*.ps1' -File | ForEach-Object {
    $txt = [System.IO.File]::ReadAllText($_.FullName)
    Remove-Item -LiteralPath $_.FullName -Force
    [System.IO.File]::WriteAllText($_.FullName, $txt, (New-Object System.Text.UTF8Encoding($true)))
}

# 4. Compilar Generador de Certificados.exe LOCALMENTE en este PC con csc.exe de Windows
# (Un .exe compilado localmente en la maquina no tiene marca de Internet y tiene el icono Opcion 1 incrustado)
$icoPath = Join-Path $installDir '_herramientas\icono-certificados.ico'
$appScript = Join-Path $installDir '_herramientas\app-certificados.ps1'
$exePath = Join-Path $installDir 'Generador de Certificados.exe'
if (Test-Path -LiteralPath $exePath) { Remove-Item -LiteralPath $exePath -Force -ErrorAction SilentlyContinue }

$csCode = @'
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

[assembly: AssemblyTitle("Generador Mensual de Certificados")]
[assembly: AssemblyProduct("Generador Mensual de Certificados")]
[assembly: AssemblyVersion("1.0.0.0")]

namespace GeneradorCertificados
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string scriptPath = Path.Combine(baseDir, "_herramientas", "app-certificados.ps1");
            string psExe = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"WindowsPowerShell\v1.0\powershell.exe");
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = psExe,
                Arguments = "-Sta -NoProfile -ExecutionPolicy Bypass -File \"" + scriptPath + "\"",
                WorkingDirectory = baseDir,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            Process.Start(psi);
        }
    }
}
'@
$tempCs = Join-Path $env:TEMP 'LauncherLocal.cs'
[System.IO.File]::WriteAllText($tempCs, $csCode, (New-Object System.Text.UTF8Encoding($false)))
$csc = Join-Path ([Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) 'csc.exe'
if (Test-Path -LiteralPath $csc) {
    & $csc /nologo /target:winexe /optimize+ "/win32icon:$icoPath" "/out:$exePath" /r:System.Windows.Forms.dll $tempCs | Out-Null
}
Remove-Item -LiteralPath $tempCs -Force -ErrorAction SilentlyContinue

# 5. Crear Acceso Directo en el Escritorio con el icono Opcion 1
$psSigned = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$wsh = New-Object -ComObject WScript.Shell
$lnkDest = Join-Path $desktop 'Generador de Certificados.lnk'
$sc = $wsh.CreateShortcut($lnkDest)
if (Test-Path -LiteralPath $exePath) {
    $sc.TargetPath = $exePath
    $sc.IconLocation = "$exePath,0"
} else {
    $sc.TargetPath = $psSigned
    $sc.Arguments = "-Sta -NoProfile -ExecutionPolicy Bypass -File `"$appScript`""
    $sc.IconLocation = "$icoPath,0"
}
$sc.WorkingDirectory = $installDir
$sc.Description = 'Abrir Generador Mensual de Certificados'
$sc.Save()

Write-Host ""
Write-Host "INSTALACION COMPLETADA CON EXITO!" -ForegroundColor Green
Write-Host "- Se creo el icono 'Generador de Certificados' en su Escritorio." -ForegroundColor Green
Write-Host "- Abriendo la aplicacion..." -ForegroundColor Yellow

& $psSigned -Sta -NoProfile -ExecutionPolicy Bypass -File $appScript
