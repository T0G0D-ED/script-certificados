# Generador Mensual de Certificados — Parra & Mauro Ltda.

Sistema con interfaz visual para crear y verificar automáticamente los **certificados mensuales en Word (`.doc`)** manteniendo el formato original de cada empresa y sucursal, la numeración correlativa (`FOLIO Nº`) y las fechas de validez.

---

## 1. Instalación rápida (1 comando en PowerShell)

Para instalar la aplicación en cualquier computador con Windows (crea automáticamente el ícono en el Escritorio y evita bloqueos de *Smart App Control*):

1. Abra el menú Inicio, escriba **`powershell`** y ábralo.
2. Copie, pegue esta línea y presione **Enter**:

```powershell
irm https://raw.githubusercontent.com/T0G0D-ED/script-certificados/main/instalar.ps1 | iex
```

*(También puede descargar `Generador-Certificados-App.zip` desde la sección **Releases**; si lo descarga con el navegador, haga clic derecho sobre el `.zip` → **Propiedades** → marque **☑ Desbloquear** antes de descomprimirlo).*

---

## 2. Uso rápido mes a mes (Interfaz Visual)

1. Haga doble clic en **`Generador de Certificados.exe`** (o en el ícono **`Generador de Certificados`** del Escritorio).
2. Si es la primera vez que lo abre en ese computador, seleccione la carpeta donde están sus certificados (o pulse **`Cargar otra carpeta...`** arriba a la derecha).
3. Al abrirse la ventana, el programa **detecta automáticamente el último mes listo** y muestra arriba qué mes nuevo se va a crear, sus fechas de validez y el rango de folios nuevos.
4. **PASO 1: Revisar y Simular** (botón azul):
   - Verifica los documentos del mes anterior y muestra en la tabla qué se va a crear, sin modificar ningún archivo.
   - Si algún documento no tuviera folio detectable, se abrirá una ventana para escribir el número manualmente (con un botón para abrir ese documento en Word).
5. **PASO 2: Crear Certificados** (botón verde):
   - Copia cada certificado del mes anterior a la carpeta del año correspondiente (`2026`, `2027`, etc.), actualiza el nombre del archivo, el número de folio (`+1` correlativo), la etiqueta `MES / AÑO` y el período (`Desde el 01 de <MES> <AÑO>` / `Hasta el 15 de <MES SIGUIENTE> <AÑO>`).
   - Muestra una barra de progreso en vivo documento por documento y reabre cada archivo guardado para verificar que el folio y las fechas hayan quedado exactos.
6. **Botón verde `▶ Abrir` en cada fila (o doble clic)**:
   - Abre inmediatamente ese certificado en Microsoft Word.
7. **`+ Nueva Empresa / Sucursal`** (botón morado):
   - Permite crear una empresa nueva o agregar una sucursal a una empresa existente, generando sus carpetas de año y su primer certificado base con el siguiente folio correlativo disponible.
8. **`Deshacer Mes Creado`** (botón rojo con doble verificación):
   - Si necesita borrar el último mes generado para rehacerlo, pide **dos confirmaciones** (la segunda escribiendo la palabra `BORRAR`) y elimina únicamente los certificados creados en ese mes, sin tocar meses anteriores.

---

## 3. Estructura del repositorio

- **`Generador de Certificados.exe`**: ejecutable principal en C# (.NET) con el ícono oficial incrustado.
- **`instalar.ps1`**: instalador automático de 1 línea para PowerShell.
- **`_herramientas/`**:
  - `app-certificados.ps1`: aplicación visual en Windows Forms.
  - `extraer-mes.ps1`: lee los certificados de un mes y año (`-Mes ENERO -Anio 2027`) y genera `inventario-<mes>-<anio>.csv`.
  - `generar-mes.ps1`: genera los certificados del mes siguiente a partir del inventario, con soporte de simulación (`-DryRun`) y verificación automática post-guardado.
  - `nueva-empresa-sucursal.ps1`: crea carpetas y certificado inicial para nuevas empresas o sucursales.
  - `normalizar-carpetas.ps1`: inspecciona y normaliza nombres de carpetas de empresas y sucursales.
  - `icono-certificados.ico`: ícono oficial de la aplicación.

---

## 4. Requisitos

- Windows con **Microsoft Word** instalado (necesario para abrir y guardar los archivos `.doc` conservando logotipos, firmas, tablas y tabulaciones).
- Cerrar los certificados de Word antes de ejecutar la generación para evitar archivos bloqueados.
