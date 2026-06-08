# Especificación funcional — Sistema "Programación MP 2026"

Gestión del Programa de Mantenciones Preventivas (MP) de equipos médicos críticos
del Hospital Hernán Henríquez Aravena.

> Este documento es una redacción ordenada de los requerimientos. Está basado en
> la inspección real del archivo `Programación MP 2026.xlsm` (hojas `PMP_2026` y
> `Registro_MP-2026`).

---

## 1. Objetivo

Construir una aplicación que, a partir del archivo Excel `Programación MP 2026`,
permita:

1. **Importar** la programación anual y los registros de mantenciones preventivas.
2. **Generar un registro (una fila) por cada mantención** programada de cada equipo.
3. **Almacenar** el resultado en una hoja de **Google Sheets** llamada
   **`Programación MP`**.
4. **Buscar** un equipo por su **N° de Serie** o **N° de Inventario**.
5. **Registrar la ejecución** de una mantención con validación de fecha y captura
   guiada (listas desplegables) de ejecutor, resultado, estado del equipo y
   observación.

---

## 2. Modelo de datos del equipo

Cada equipo se describe con los siguientes atributos (provenientes del catálogo).

| # | Campo | Descripción |
|---|-------|-------------|
| 1 | **ID** | Orden del equipo dentro de la planilla. **No es identificador único de negocio.** |
| 2 | N° Carpeta | Número de carpeta del equipo |
| 3 | N° Inventario | Identificador de inventario |
| 4 | Equipo | Tipo / nombre del equipo |
| 5 | Servicio | Servicio clínico |
| 6 | Unidad | Unidad |
| 7 | Ubicación | Ubicación física |
| 8 | Procedencia | Propiedad / Arriendo / etc. |
| 9 | Marca | Fabricante |
| 10 | Modelo | Modelo |
| 11 | **Serie** | N° de serie. **Identificador único de negocio.** |
| 12 | Año Instalación | Año de instalación |
| 13 | Vida Útil Residual | Años de vida útil restantes |
| 14 | Clasificación | Fija / Móvil / etc. |
| 15 | ENU / Baja | Estado de baja / no ubicable |
| 16 | Frecuencia MP | Frecuencia de la mantención (p. ej. Trimestral) |

> **Identidad del equipo:** el **ID** es solo el orden en la planilla. Cada equipo
> se identifica unívocamente por su **N° de Serie** o, en su defecto, por su
> **N° de Inventario**.

### 2.1 Regla crítica: ceros a la izquierda en la Serie

Las series **deben tratarse siempre como texto**, preservando los ceros a la
izquierda: la serie `0024` **no** puede registrarse ni buscarse como `24`.

Esto es un problema real en el archivo fuente: algunas series están guardadas
como texto con su cero (`'00888'`, `'0100645'`) y otras quedaron guardadas como
número, perdiendo el cero (`298`, `13042`, `40361`). El importador debe:

- Leer la serie **forzando tipo texto** (sin conversión numérica).
- No recortar ceros a la izquierda ni espacios significativos.
- Normalizar caracteres espurios detectados en el archivo (p. ej. punto final
  `'02516.'`, `'0100638.'`) según criterio que defina el usuario.
- Comparar en la búsqueda de forma **exacta como cadena**.

---

## 3. Glosario de siglas y códigos

### 3.1 Programación (columna **P** del registro / celdas de mes en el plan)

| Sigla | Significado |
|-------|-------------|
| **X** | Mantención Preventiva Programada |
| **R** | Mantención Preventiva Reprogramada |
| **RA** | Mantención Preventiva Reprogramada del Año Anterior |
| **PM** | Puesta en Marcha |

### 3.2 Resultado (columna **R** del registro)

| Código | Significado |
|--------|-------------|
| **Si** | MP realizada |
| **C1–C8** | Causal de reprogramación (ver 3.3) |
| **Si-RA** | MP realizada correspondiente a reprogramación del año anterior |
| **FS** | Fuera de servicio |
| **No** | MP no realizada |
| **NU** | No ubicable |
| **Baja** | Equipo dado de baja |

> **Normalización requerida:** el archivo contiene inconsistencias de mayúsculas
> (`'Si'` y `'SI'`). El importador debe normalizar a un valor canónico
> (p. ej. `Si`). Valores fuera de la lista válida deben descartarse o marcarse
> para revisión.

### 3.3 Causales de reprogramación (C1–C8)

| Código | Descripción |
|--------|-------------|
| C1 | Imposibilidad de desocupar el equipo del paciente por indicación clínica |
| C2 | Equipo en servicio técnico |
| C3 | Equipo no operativo, a la espera de repuestos o accesorios |
| C4 | Equipo en préstamo a otro hospital o institución |
| C5 | No disponibilidad de horas-hombre del funcionario SEC por alta carga laboral |
| C6 | No disponibilidad de horas-hombre del servicio técnico externo |
| C7 | Ausencia justificada del funcionario SEC superior a 15 días |
| C8 | Contingencia hospitalaria |

**Reglas de reprogramación**

- **C2, C3, C4** → *no fijan fecha*; se espera el reintegro del equipo. La
  mantención se registra en el **mes real de ejecución**.
- **C1, C5, C6, C7, C8** → *reprogramar dentro de 30 días*. El mes original
  conserva `X` en **P** y la causal en **R**; el mes destino lleva `R` en **P**.

### 3.4 Estado del equipo

Valor capturado al registrar la ejecución. Lista cerrada:

`Operativo` · `No operativo` · `Servicio técnico` · `Desconocido`

---

## 4. Ejecutores (responsables de la MP)

Lista cerrada de 11 responsables (lista desplegable en el registro):

1. Carlos Bahamondes Seguel
2. Cristián Beltrán Oviedo
3. Cristina Rozas Urrutia
4. Daniel Díaz Neira
5. Ignacio Berner Bergara
6. Macarena Toledo
7. Marco Ulloa
8. Matías Soazo Garrido
9. Ricardo Matus Aroca
10. Tito Millapán Riquelme
11. Personal externo

---

## 5. Lectura del archivo `Programación MP 2026.xlsm`

El archivo tiene encabezados en la **fila 7** y datos desde la **fila 8**
(aprox. 966 equipos, filas 8 a 973). En ambas hojas se **ignoran las columnas
`Q` (Observación) y `S` (Responsable MP)**.

### 5.1 Hoja `PMP_2026` — Plan (Carta Gantt)

Bloque de identidad del equipo: columnas **B a P** + **R** (Frecuencia MP).
Meses (un valor por celda: `X` / `R` / `RA` / `PM`):

| Mes | Col | Mes | Col |
|-----|-----|-----|-----|
| Ene | T | Jul | Z |
| Feb | U | Ago | AA |
| Mar | V | Sep | AB |
| Abr | W | Oct | AC |
| May | X | Nov | AD |
| Jun | Y | Dic | AE |

> Lectura solicitada: rango de encabezados **`B7:AE7`** (datos `B8:AE973`),
> **sin** columnas `Q` ni `S`.

### 5.2 Hoja `Registro_MP-2026` — Ejecución

Mismo bloque de identidad (**B–P** + **R**, ignorando `Q` y `S`). Cada mes ocupa
**dos subcolumnas**: **P** (programa) y **R** (resultado).

| Mes | P | R | Mes | P | R |
|-----|---|---|-----|---|---|
| Ene | T | U | Jul | AF | AG |
| Feb | V | W | Ago | AH | AI |
| Mar | X | Y | Sep | AJ | AK |
| Abr | Z | AA | Oct | AL | AM |
| May | AB | AC | Nov | AN | AO |
| Jun | AD | AE | Dic | AP | AQ |

- Subcolumna **P**: código de programación (`X` / `R` / `RA` / `PM`).
- Subcolumna **R**: resultado (`Si` / `C1–C8` / `Si-RA` / `FS` / `No` / `NU` / `Baja`).
- Columna **AR**: Observación general (texto libre).

---

## 6. Generación de registros (una fila por mantención)

Al importar, el sistema genera **una fila por cada mantención** de cada equipo,
es decir, por cada combinación **(equipo × mes con marca)**:

- Desde `PMP_2026`: cada celda de mes con valor (`X`/`R`/`RA`/`PM`) origina una
  fila de **mantención programada**.
- Desde `Registro_MP-2026`: la subcolumna **P** define la programación y la
  subcolumna **R** completa el **resultado** de esa misma mantención.

Cada fila combina los **16 campos del equipo** (sección 2) con los campos de
ejecución de la sección 7.

---

## 7. Salida: Google Sheets `Programación MP`

La hoja de destino tendrá **todas las columnas del equipo** (sección 2) más las
siguientes columnas de gestión:

| Campo | Origen | Descripción |
|-------|--------|-------------|
| **ID Programación** | Sistema | Identificador único de la fila de programación |
| **Fecha de Registro** | Sistema | Fecha/hora en que se creó/registró la fila en el sistema. **Distinta** de la fecha de mantención |
| **Programa** | P / plan | Código de programación (`X` / `R` / `RA` / `PM`) y mes asociado |
| **Fecha de Mantención** | Usuario | Fecha real (o programada) de la mantención |
| **Resultado** | Usuario / R | `Si` / `C1–C8` / `Si-RA` / `FS` / `No` / `NU` / `Baja` |
| **Ejecutor** | Usuario | Uno de los 11 responsables (sección 4) |
| **Estado del Equipo** | Usuario | `Operativo` / `No operativo` / `Servicio técnico` / `Desconocido` |
| **Observaciones** | Usuario | Texto libre (opcional) |

> **Trazabilidad — dos fechas distintas:**
> - *Fecha de Mantención*: cuándo se realizó (o se programó) la mantención.
> - *Fecha de Registro*: cuándo se ingresó esa información al sistema.

---

## 8. Búsqueda de equipos

- Buscar por **N° de Serie** o **N° de Inventario**.
- Coincidencia **exacta como texto**, respetando ceros a la izquierda
  (`0024` ≠ `24`) — ver regla 2.1.
- Resultado de la búsqueda: ficha del equipo y sus mantenciones programadas del año.

---

## 9. Flujo de registro de una mantención

1. El usuario **busca** un equipo (serie o inventario).
2. El sistema **solicita la fecha** de la mantención.
3. **Validación de mes:**
   - Si el **mes de la fecha ingresada NO coincide** con el mes programado para
     ese equipo → el sistema **impide** el registro (mensaje de error).
   - Si el **mes coincide** → continúa.
4. Con el mes válido, el sistema solicita mediante **listas desplegables**:
   - **Ejecutor** (lista de 11, sección 4).
   - **Resultado** de la mantención (sección 3.2).
   - **Estado del equipo** (sección 3.4).
   - **Observación** (texto libre, **opcional**).
5. El usuario **guarda**. El sistema crea/actualiza la fila en `Programación MP`
   con un **ID Programación** y la **Fecha de Registro** (timestamp del sistema).

---

## 10. Reglas de validación (resumen)

| # | Regla |
|---|-------|
| V1 | La serie se maneja como texto; nunca se recortan ceros a la izquierda. |
| V2 | La búsqueda por serie/inventario es exacta como cadena. |
| V3 | No se permite registrar si el mes de la fecha ≠ mes programado. |
| V4 | Ejecutor, Resultado y Estado del equipo provienen de listas cerradas. |
| V5 | La Observación es opcional. |
| V6 | Cada fila guardada recibe ID Programación único y Fecha de Registro. |
| V7 | Códigos de programación válidos: `X`, `R`, `RA`, `PM` (otros se descartan/revisan). |
| V8 | Resultados se normalizan (`SI`→`Si`) y se validan contra la lista. |

---

## 11. Consideraciones técnicas y casos detectados en el archivo

- **Tipos mixtos en Serie:** texto con ceros (`'00888'`) y números sin cero
  (`298`). Forzar lectura como texto.
- **Caracteres espurios:** algunas series traen punto final (`'02516.'`) — definir
  limpieza.
- **Mayúsculas inconsistentes:** `'Si'` vs `'SI'` en resultados — normalizar.
- **Valores fuera de catálogo:** en las subcolumnas aparecen valores numéricos
  sueltos (provenientes de tablas auxiliares/estadísticas a la derecha del rango).
  Limitar la lectura estrictamente a los rangos de columnas indicados.
- **Otras hojas del libro** (`Servicio tecnico 2025`, `Bajas 2026/2025/2024/2023`)
  no forman parte del alcance de esta carga; confirmar si deben integrarse.

---

## 12. Decisiones pendientes (a confirmar con el usuario)

1. **Plataforma de la aplicación:** ¿Google Apps Script (sobre la propia hoja),
   aplicación web (p. ej. Python + API de Google Sheets) u otra?
2. **Identidad ante conflicto:** si Serie e Inventario apuntan a equipos distintos,
   ¿cuál prevalece? ¿Qué ocurre si la Serie está vacía?
3. **Formato del ID Programación** (p. ej. `MP-2026-000123`).
4. **Limpieza de series** con punto final u otros caracteres.
5. **Carga inicial:** ¿se generan filas desde `PMP_2026` (plan) y se completan con
   `Registro_MP-2026`, o solo desde una de las hojas?
6. **Tolerancia de fecha:** ¿la validación es por mes calendario exacto, o se
   admite la ventana de 30 días de las reprogramaciones C1/C5/C6/C7/C8?
