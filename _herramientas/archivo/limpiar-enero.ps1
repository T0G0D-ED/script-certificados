
$ErrorActionPreference = 'Stop'

# Ruta de la carpeta de Diciembre
$root = 'C:\Users\Maxi-\OneDrive\Escritorio\proyecto-certificados\CERTIFICADOS DICIEMBRE 2026'

# Encontrar todos los archivos que contengan "ENERO" en el nombre y "2027"
$eneroFiles = Get-ChildItem -LiteralPath $root -Recurse -File |
    Where-Object {
        $_.Name -match 'ENERO' -and $_.Name -match '2027'
    }

Write-Host ("Archivos de Enero 2027 encontrados para borrar: {0}" -f $eneroFiles.Count)

if ($eneroFiles.Count -gt 0) {
    foreach ($f in $eneroFiles) {
        Write-Host ("Borrando: {0}" -f $f.FullName)
        Remove-Item -LiteralPath $f.FullName -Force
    }
    Write-Host "Todos los archivos generados de Enero han sido eliminados exitosamente."
} else {
    Write-Host "No se encontraron archivos de Enero para borrar."
}

# (Opcional) Borrar el archivo log de enero
$tools = 'C:\Users\Maxi-\OneDrive\Escritorio\proyecto-certificados\_herramientas'
$logPath = Join-Path $tools 'log-enero-2027.csv'

if (Test-Path -LiteralPath $logPath) {
    Remove-Item -LiteralPath $logPath -Force
    Write-Host "Archivo log-enero-2027.csv eliminado."
}
