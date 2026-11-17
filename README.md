# Generador Mensual de Certificados — Parra & Mauro Ltda.

Sistema para crear y verificar automáticamente los **75 certificados mensuales en Word (`.doc`)** manteniendo el formato original de cada empresa y sucursal, la numeración correlativa (`FOLIO Nº`) y las fechas de validez.

---

## 1. Uso rápido mes a mes (Interfaz Visual)

1. Haga doble clic en el archivo **`Abrir Generador de Certificados.bat`** (en la carpeta principal del proyecto).
2. Al abrirse la ventana, el programa **detecta automáticamente el último mes listo** y muestra arriba qué mes nuevo se va a crear, sus fechas de validez y el rango de folios nuevos (+75).
3. **PASO 1: Revisar y Simular** (botón azul):
   - Verifica que estén los 75 documentos del mes anterior y muestra en la tabla qué se va a crear, sin modificar ningún archivo.
   - Si algún documento no tuviera folio detectable, se abrirá una ventana para escribir el número manualmente (con un botón para abrir ese documento en Word).
4. **PASO 2: Crear Certificados** (botón verde):
   - Copia cada certificado del mes anterior a la carpeta del año correspondiente (`2026`, `2027`, etc.), actualiza el nombre del archivo, el número de folio (`+1` correlativo), la etiqueta `MES / AÑO` y el período (`Desde el 01 de <MES> <AÑO>` / `Hasta el 15 de <MES SIGUIENTE> <AÑO>`).
   - Reabre cada documento guardado para verificar que el folio y las fechas hayan quedado exactos y que el documento original no haya cambiado.
5. **Abrir Carpeta Certificados**:
   - Abre directamente la carpeta donde están todas las empresas y sucursales.
6. **Deshacer Mes Creado** (botón rojo con doble verificación):
   - Si necesita borrar el último mes generado para rehacerlo, pide **dos confirmaciones** (la segunda escribiendo la palabra `BORRAR`) y elimina únicamente los certificados creados en ese mes, sin tocar meses anteriores.

---

## 2. Estructura del proyecto

- **`Abrir Generador de Certificados.bat`**: acceso directo para abrir la aplicación visual con doble clic.
- **`CERTIFICADOS DICIEMBRE 2026/`**: carpeta raíz con las 30 empresas y sus sucursales (75 ubicaciones en total). Cada sucursal tiene sus carpetas por año (`2025`, `2026`, `2027`).
- **`INFORME-ENERO-2027.md`**: informe detallado de generación y auditoría de los 75 certificados.
- **`_herramientas/`**: scripts de automatización e inventarios:
  - `app-certificados.ps1`: aplicación visual en Windows Forms.
  - `extraer-mes.ps1`: lee los certificados de un mes y año (`-Mes ENERO -Anio 2027`) y genera `inventario-<mes>-<anio>.csv`. Detecta folios en el cuerpo, encabezados, pies de página y cuadros de texto.
  - `generar-mes.ps1`: genera los certificados del mes siguiente a partir del inventario, con soporte de simulación (`-DryRun`) y verificación automática post-guardado. Crea automáticamente la carpeta del nuevo año cuando se pasa de Diciembre a Enero.
  - `normalizar-carpetas.ps1`: inspecciona las carpetas de empresas y sucursales, detecta colisiones y normaliza espacios sobrantes (`-Aplicar`).
  - `folios-manuales.csv` *(opcional)*: archivo con columnas `Rel,Folio` para forzar manualmente el folio de un documento si fuera necesario.
  - `archivo/`: respaldo de scripts y registros antiguos.

---

## 3. Uso por consola (PowerShell, opcional)

Si en algún momento se desea ejecutar desde PowerShell:

```powershell
# 1. Extraer el inventario de un mes
powershell -ExecutionPolicy Bypass -File .\_herramientas\extraer-mes.ps1 -Mes ENERO -Anio 2027

# 2. Simular el mes siguiente (no crea ni modifica archivos)
powershell -ExecutionPolicy Bypass -File .\_herramientas\generar-mes.ps1 -Origen ENERO -OrigenAnio 2027 -Destino FEBRERO -DestinoAnio 2027 -HastaMes MARZO -HastaAnio 2027 -DryRun

# 3. Generar el mes siguiente
powershell -ExecutionPolicy Bypass -File .\_herramientas\generar-mes.ps1 -Origen ENERO -OrigenAnio 2027 -Destino FEBRERO -DestinoAnio 2027 -HastaMes MARZO -HastaAnio 2027
```

---

## 4. Requisitos

- Windows con **Microsoft Word** instalado (necesario para abrir y guardar los archivos `.doc` conservando logotipos, firmas, tablas y tabulaciones).
- Cerrar los certificados de Word antes de ejecutar la generación para evitar archivos bloqueados.
