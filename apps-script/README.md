# Programación MP 2026 — Google Apps Script

Aplicación ligada a una hoja de **Google Sheets** que gestiona las mantenciones
preventivas (MP) de equipos médicos críticos. Lee las hojas `PMP_2026` y
`Registro_MP-2026` del archivo *Programación MP 2026* y genera **una fila por cada
mantención** en la hoja de salida `Programación MP`, con búsqueda de equipos y
registro guiado de la ejecución.

Especificación funcional completa: [`../docs/especificacion-programacion-mp.md`](../docs/especificacion-programacion-mp.md).

## Archivos

| Archivo | Rol |
|---------|-----|
| `appsscript.json` | Manifiesto (zona horaria, permisos). |
| `Codigo.gs` | Lógica: menú, importación, búsqueda y registro. |
| `Importar.html` | Diálogo para importar PMP / Registro. |
| `Buscar.html` | Barra lateral para buscar un equipo y registrar su MP. |

## Instalación

1. Crea (o abre) un Google Sheets que será la base **`Programación MP`**.
2. Menú **Extensiones → Apps Script**.
3. Crea los archivos con el mismo nombre y pega el contenido de cada uno
   (`Codigo.gs`, y los HTML `Importar` y `Buscar`). En el manifiesto
   (`appsscript.json`, visible con el ícono ⚙️ *Configuración → Mostrar
   "appsscript.json"*) pega el contenido de este repositorio.
4. Guarda y **recarga** la hoja de cálculo: aparecerá el menú **Programación MP**.

> Alternativa con [`clasp`](https://github.com/google/clasp):
> `clasp create --type sheets` y `clasp push` desde esta carpeta.

## Uso

### 1. Preparar la hoja de salida
Menú **Programación MP → Preparar hoja de salida**. Crea la hoja `Programación MP`
con encabezados y listas desplegables (Programa, Resultado, Ejecutor, Estado).

### 2. Importar PMP / Registro
1. Sube el archivo `Programación MP 2026.xlsm` a Google Drive y ábrelo con
   **Hojas de cálculo de Google** (crea una copia convertida con `PMP_2026` y
   `Registro_MP-2026`).
2. Menú **Programación MP → Importar PMP / Registro…**, pega la **URL** del libro
   convertido y pulsa **Importar**.
3. Se genera una fila por cada `(equipo × mes con marca)`. Si el `Registro_MP-2026`
   ya traía resultado, queda precargado. La reimportación **no duplica** filas
   (clave: Serie/Inventario + Mes) ni pisa lo registrado manualmente.

### 3. Buscar y registrar una MP
1. Menú **Programación MP → Buscar equipo y registrar MP…** (barra lateral).
2. Busca por **Serie** o **N° de Inventario** (coincidencia exacta como texto;
   `0024` ≠ `24`).
3. Elige una mantención **pendiente**, ingresa la **fecha**:
   - Si el **mes** de la fecha no coincide con el mes programado → se **bloquea**.
   - Si coincide → selecciona **Ejecutor**, **Resultado**, **Estado del equipo**,
     una **Observación** opcional y **Guarda**.
4. Al guardar se fija el `ID Programación` y la `Fecha de Registro` (distinta de la
   fecha de mantención).

## Notas técnicas

- **Series con ceros a la izquierda:** se leen como texto (`getDisplayValues`) y la
  columna Serie de la hoja de salida tiene formato texto (`@`). Las series que en el
  Excel original quedaron guardadas como número (p. ej. `298`) ya perdieron el cero
  en el origen y no se pueden recuperar.
- **Normalización:** los resultados se canonizan (`SI` → `Si`) y se validan contra
  la lista; valores fuera de catálogo se descartan.
- **Zona horaria:** `America/Santiago` (ver `appsscript.json`).
