/**
 * Sistema "Programación MP 2026" — Google Apps Script
 * Gestión de mantenciones preventivas (MP) de equipos médicos críticos.
 *
 * Ligado a la hoja de Google Sheets "Programación MP" (la hoja de salida).
 * Lee las hojas PMP_2026 y Registro_MP-2026 de un libro fuente (el archivo
 * "Programación MP 2026" importado a Google Sheets) y genera una fila por cada
 * mantención. Permite buscar un equipo y registrar la ejecución de su MP.
 *
 * Ver docs/especificacion-programacion-mp.md para el detalle funcional.
 */

// ============================================================================
// Configuración
// ============================================================================

const CFG = {
  SRC_PMP: 'PMP_2026',            // Hoja plan (Carta Gantt)
  SRC_REG: 'Registro_MP-2026',   // Hoja de registro/ejecución
  OUT_SHEET: 'Programación MP',   // Hoja de salida (en este mismo libro)
  HEADER_ROW: 7,                  // Encabezados en fila 7
  DATA_START_ROW: 8,              // Datos desde fila 8
  ANIO: 2026,
  ID_PREFIX: 'MP-2026-',

  MESES: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'],
  PROGRAMAS: ['X', 'R', 'RA', 'PM'],
  RESULTADOS: ['Si', 'Si-RA', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'FS', 'No', 'NU', 'Baja'],
  ESTADOS: ['Operativo', 'No operativo', 'Servicio técnico', 'Desconocido'],
  EJECUTORES: [
    'Carlos Bahamondes Seguel',
    'Cristián Beltrán Oviedo',
    'Cristina Rozas Urrutia',
    'Daniel Díaz Neira',
    'Ignacio Berner Bergara',
    'Macarena Toledo',
    'Marco Ulloa',
    'Matías Soazo Garrido',
    'Ricardo Matus Aroca',
    'Tito Millapán Riquelme',
    'Personal externo'
  ]
};

// Columnas (1-based) del bloque de identidad del equipo. Comunes a ambas hojas.
// Se ignoran Q (17, Observación) y S (19, Responsable MP).
const EQ = {
  id: 2, carpeta: 3, inventario: 4, equipo: 5, servicio: 6, unidad: 7,
  ubicacion: 8, procedencia: 9, marca: 10, modelo: 11, serie: 12,
  anioInstalacion: 13, vidaUtil: 14, clasificacion: 15, enuBaja: 16, frecuencia: 18
};

// Orden de los campos de identidad tal como se escriben en la hoja de salida.
const EQ_FIELDS = [
  'id', 'carpeta', 'inventario', 'equipo', 'servicio', 'unidad', 'ubicacion',
  'procedencia', 'marca', 'modelo', 'serie', 'anioInstalacion', 'vidaUtil',
  'clasificacion', 'enuBaja', 'frecuencia'
];

// PMP_2026: una columna por mes (T..AE = 20..31).
const PMP_MES_COL = { Ene: 20, Feb: 21, Mar: 22, Abr: 23, May: 24, Jun: 25, Jul: 26, Ago: 27, Sep: 28, Oct: 29, Nov: 30, Dic: 31 };

// Registro_MP-2026: dos subcolumnas por mes [P (programa), R (resultado)].
const REG_MES_COL = {
  Ene: [20, 21], Feb: [22, 23], Mar: [24, 25], Abr: [26, 27], May: [28, 29], Jun: [30, 31],
  Jul: [32, 33], Ago: [34, 35], Sep: [36, 37], Oct: [38, 39], Nov: [40, 41], Dic: [42, 43]
};

// Encabezados de la hoja de salida "Programación MP".
const OUT_HEADERS = [
  'ID Programación', 'Fecha de Registro',
  'ID', 'N° Carpeta', 'N° Inventario', 'Equipo', 'Servicio', 'Unidad', 'Ubicación',
  'Procedencia', 'Marca', 'Modelo', 'Serie', 'Año Instalación', 'Vida Útil Residual',
  'Clasificación', 'ENU / Baja', 'Frecuencia MP',
  'Mes', 'Programa', 'Fecha de Mantención', 'Resultado', 'Ejecutor', 'Estado del Equipo', 'Observaciones'
];

// Índices (1-based) de columnas de gestión dentro de la hoja de salida.
const OUT = {
  idProg: 1, fechaRegistro: 2,
  serie: 13, inventario: 5,
  mes: 19, programa: 20, fechaMant: 21, resultado: 22, ejecutor: 23, estado: 24, observaciones: 25
};

// ============================================================================
// Menú
// ============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Programación MP')
    .addItem('Buscar equipo y registrar MP…', 'abrirBuscador')
    .addSeparator()
    .addItem('Importar PMP / Registro…', 'abrirImportador')
    .addItem('Preparar hoja de salida', 'prepararHojaSalida')
    .addToUI();
}

function abrirBuscador() {
  const html = HtmlService.createHtmlOutputFromFile('Buscar')
    .setTitle('Buscar equipo y registrar MP')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

function abrirImportador() {
  const html = HtmlService.createHtmlOutputFromFile('Importar')
    .setWidth(520).setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Importar PMP / Registro');
}

/** Datos de configuración para las listas desplegables del cliente. */
function getConfig() {
  return {
    ejecutores: CFG.EJECUTORES,
    resultados: CFG.RESULTADOS,
    estados: CFG.ESTADOS,
    meses: CFG.MESES
  };
}

// ============================================================================
// Hoja de salida
// ============================================================================

/** Crea (si no existe) la hoja de salida con encabezados, formato y validaciones. */
function prepararHojaSalida() {
  const sheet = getOutputSheet_();
  SpreadsheetApp.getActive().toast('Hoja "' + CFG.OUT_SHEET + '" lista.', 'Programación MP', 5);
  return sheet.getName();
}

function getOutputSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(CFG.OUT_SHEET);
  if (!sheet) sheet = ss.insertSheet(CFG.OUT_SHEET);

  // Encabezados
  const headerRange = sheet.getRange(1, 1, 1, OUT_HEADERS.length);
  headerRange.setValues([OUT_HEADERS]).setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  // La columna Serie se fuerza a texto para preservar ceros a la izquierda.
  sheet.getRange(2, OUT.serie, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
  sheet.getRange(2, OUT.inventario, sheet.getMaxRows() - 1, 1).setNumberFormat('@');

  // Validaciones (listas desplegables) en las columnas de gestión.
  setListValidation_(sheet, OUT.programa, CFG.PROGRAMAS);
  setListValidation_(sheet, OUT.resultado, CFG.RESULTADOS);
  setListValidation_(sheet, OUT.ejecutor, CFG.EJECUTORES);
  setListValidation_(sheet, OUT.estado, CFG.ESTADOS);

  return sheet;
}

function setListValidation_(sheet, col, values) {
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build();
  sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
}

// ============================================================================
// Importación: generación de una fila por mantención
// ============================================================================

/**
 * Importa desde un libro fuente (URL o ID del Google Sheet convertido del Excel).
 * Genera una fila por cada (equipo × mes con marca) y la agrega a la hoja de
 * salida, evitando duplicados por (clave de equipo + mes).
 */
function importarDesdeUrl(urlOrId) {
  const srcId = extraerSpreadsheetId_(urlOrId);
  if (!srcId) throw new Error('No se reconoce la URL o ID del libro fuente.');

  let src;
  try {
    src = SpreadsheetApp.openById(srcId);
  } catch (e) {
    throw new Error('No se pudo abrir el libro fuente. Verifica el enlace y los permisos.');
  }

  const pmp = leerHoja_(src.getSheetByName(CFG.SRC_PMP), CFG.SRC_PMP);
  const reg = leerHoja_(src.getSheetByName(CFG.SRC_REG), CFG.SRC_REG);

  // Mapa de equipos por clave (serie || inventario) con sus datos de identidad.
  // Mapa de mantenciones por clave -> { mes -> {programa, resultado} }.
  const equipos = {};   // key -> identidad
  const planes = {};    // key -> { mes -> {programa, resultado, observacion} }

  // --- PMP_2026 (plan): programa por mes ---
  pmp.forEach(function (row) {
    const eq = leerIdentidad_(row);
    const key = equipoKey_(eq);
    if (!key) return;
    equipos[key] = equipos[key] || eq;
    planes[key] = planes[key] || {};
    CFG.MESES.forEach(function (mes) {
      const code = textoCelda_(row[PMP_MES_COL[mes] - 1]);
      if (CFG.PROGRAMAS.indexOf(code) !== -1) {
        planes[key][mes] = planes[key][mes] || {};
        planes[key][mes].programa = code;
      }
    });
  });

  // --- Registro_MP-2026: programa (P) y resultado (R) por mes ---
  reg.forEach(function (row) {
    const eq = leerIdentidad_(row);
    const key = equipoKey_(eq);
    if (!key) return;
    equipos[key] = equipos[key] || eq;
    planes[key] = planes[key] || {};
    CFG.MESES.forEach(function (mes) {
      const cols = REG_MES_COL[mes];
      const prog = textoCelda_(row[cols[0] - 1]);
      const resRaw = textoCelda_(row[cols[1] - 1]);
      const res = normalizarResultado_(resRaw);
      if (CFG.PROGRAMAS.indexOf(prog) === -1 && !res) return;
      planes[key][mes] = planes[key][mes] || {};
      if (CFG.PROGRAMAS.indexOf(prog) !== -1) planes[key][mes].programa = prog;
      if (res) planes[key][mes].resultado = res;
    });
  });

  // --- Construir filas de salida, evitando duplicados ---
  const sheet = getOutputSheet_();
  const existentes = clavesExistentes_(sheet); // Set de "key||mes"
  let correlativo = siguienteCorrelativo_(sheet);
  const ahora = new Date();
  const nuevas = [];

  Object.keys(planes).forEach(function (key) {
    const eq = equipos[key];
    CFG.MESES.forEach(function (mes) {
      const p = planes[key][mes];
      if (!p || !p.programa) return; // solo mantenciones efectivamente programadas
      const dedupe = key + '||' + mes;
      if (existentes[dedupe]) return;
      existentes[dedupe] = true;
      nuevas.push(filaSalida_(correlativo++, ahora, eq, mes, p));
    });
  });

  if (nuevas.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, nuevas.length, OUT_HEADERS.length).setValues(nuevas);
  }
  return { agregadas: nuevas.length, equipos: Object.keys(equipos).length };
}

/** Construye una fila de la hoja de salida. */
function filaSalida_(correlativo, ahora, eq, mes, plan) {
  const fila = [];
  fila.push(CFG.ID_PREFIX + ('000000' + correlativo).slice(-6)); // ID Programación
  fila.push(ahora);                                              // Fecha de Registro
  EQ_FIELDS.forEach(function (f) { fila.push(eq[f]); });         // 16 campos de identidad
  fila.push(mes);                                                // Mes
  fila.push(plan.programa || '');                               // Programa
  fila.push('');                                                 // Fecha de Mantención (se llena al registrar)
  fila.push(plan.resultado || '');                             // Resultado (precargado si venía en el Registro)
  fila.push('');                                                 // Ejecutor
  fila.push('');                                                 // Estado del Equipo
  fila.push('');                                                 // Observaciones
  return fila;
}

// ============================================================================
// Búsqueda y registro
// ============================================================================

/**
 * Busca un equipo por Serie o N° de Inventario (coincidencia exacta como texto,
 * respetando ceros a la izquierda). Devuelve la ficha y sus mantenciones.
 */
function buscarEquipo(query) {
  const q = String(query == null ? '' : query).trim();
  if (!q) throw new Error('Ingresa una serie o número de inventario.');

  const sheet = getOutputSheet_();
  const last = sheet.getLastRow();
  if (last < 2) throw new Error('La hoja de salida está vacía. Primero importa el PMP / Registro.');

  const display = sheet.getRange(2, 1, last - 1, OUT_HEADERS.length).getDisplayValues();
  const mantenciones = [];
  let ficha = null;

  display.forEach(function (row, i) {
    const serie = String(row[OUT.serie - 1]).trim();
    const inv = String(row[OUT.inventario - 1]).trim();
    if (serie === q || inv === q) {
      if (!ficha) {
        ficha = {};
        EQ_FIELDS.forEach(function (f, idx) { ficha[f] = row[2 + idx]; }); // identidad arranca en col 3 (índice 2)
      }
      mantenciones.push({
        rowIndex: i + 2,
        mes: row[OUT.mes - 1],
        programa: row[OUT.programa - 1],
        fechaMant: row[OUT.fechaMant - 1],
        resultado: row[OUT.resultado - 1],
        ejecutor: row[OUT.ejecutor - 1],
        estado: row[OUT.estado - 1],
        observaciones: row[OUT.observaciones - 1],
        registrada: !!String(row[OUT.fechaMant - 1]).trim()
      });
    }
  });

  if (!ficha) throw new Error('No se encontró ningún equipo con serie/inventario "' + q + '".');
  return { ficha: ficha, mantenciones: mantenciones };
}

/**
 * Registra la ejecución de una mantención. Valida que el mes de la fecha
 * coincida con el mes programado de esa fila; si no coincide, rechaza.
 */
function registrarMantencion(payload) {
  const sheet = getOutputSheet_();
  const rowIndex = Number(payload.rowIndex);
  if (!rowIndex || rowIndex < 2) throw new Error('Fila de programación inválida.');

  const mesProgramado = String(sheet.getRange(rowIndex, OUT.mes).getValue()).trim();
  const fecha = parseFechaLocal_(payload.fecha);
  if (!fecha) throw new Error('Fecha inválida.');

  const mesFecha = CFG.MESES[fecha.getMonth()];
  if (mesFecha !== mesProgramado) {
    throw new Error('El mes de la fecha (' + mesFecha + ') no coincide con el mes programado (' +
      mesProgramado + '). No se puede registrar.');
  }

  // Validar valores contra listas cerradas.
  if (CFG.EJECUTORES.indexOf(payload.ejecutor) === -1) throw new Error('Ejecutor no válido.');
  if (CFG.RESULTADOS.indexOf(payload.resultado) === -1) throw new Error('Resultado no válido.');
  if (CFG.ESTADOS.indexOf(payload.estado) === -1) throw new Error('Estado del equipo no válido.');

  // Asegurar ID Programación y Fecha de Registro.
  if (!String(sheet.getRange(rowIndex, OUT.idProg).getValue()).trim()) {
    sheet.getRange(rowIndex, OUT.idProg).setValue(CFG.ID_PREFIX + ('000000' + (siguienteCorrelativo_(sheet))).slice(-6));
  }
  if (!String(sheet.getRange(rowIndex, OUT.fechaRegistro).getValue()).trim()) {
    sheet.getRange(rowIndex, OUT.fechaRegistro).setValue(new Date());
  }

  sheet.getRange(rowIndex, OUT.fechaMant).setValue(fecha);
  sheet.getRange(rowIndex, OUT.resultado).setValue(payload.resultado);
  sheet.getRange(rowIndex, OUT.ejecutor).setValue(payload.ejecutor);
  sheet.getRange(rowIndex, OUT.estado).setValue(payload.estado);
  sheet.getRange(rowIndex, OUT.observaciones).setValue(payload.observacion || '');

  return { ok: true, mes: mesProgramado };
}

// ============================================================================
// Utilidades
// ============================================================================

/** Lee una hoja fuente y devuelve sus filas de datos (display values, fila 8+). */
function leerHoja_(sheet, nombre) {
  if (!sheet) throw new Error('No se encontró la hoja "' + nombre + '" en el libro fuente.');
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 44);
  if (lastRow < CFG.DATA_START_ROW) return [];
  // getDisplayValues preserva el texto tal como se ve (ceros a la izquierda en series guardadas como texto).
  return sheet.getRange(CFG.DATA_START_ROW, 1, lastRow - CFG.DATA_START_ROW + 1, lastCol).getDisplayValues();
}

/** Extrae el bloque de identidad de una fila (array 0-based por columna). */
function leerIdentidad_(row) {
  const eq = {};
  EQ_FIELDS.forEach(function (f) { eq[f] = textoCelda_(row[EQ[f] - 1]); });
  return eq;
}

/** Clave de equipo: Serie y, si está vacía, N° de Inventario (texto exacto). */
function equipoKey_(eq) {
  const serie = String(eq.serie || '').trim();
  if (serie) return 'S:' + serie;
  const inv = String(eq.inventario || '').trim();
  if (inv) return 'I:' + inv;
  return '';
}

function textoCelda_(v) {
  if (v == null) return '';
  return String(v).trim();
}

/** Normaliza el resultado: 'SI' -> 'Si'; valida contra la lista; '' si no aplica. */
function normalizarResultado_(v) {
  const s = textoCelda_(v);
  if (!s) return '';
  if (s.toUpperCase() === 'SI') return 'Si';
  const up = s.toUpperCase();
  for (let i = 0; i < CFG.RESULTADOS.length; i++) {
    if (CFG.RESULTADOS[i].toUpperCase() === up) return CFG.RESULTADOS[i];
  }
  return ''; // valores fuera de catálogo (números sueltos, etc.) se descartan
}

/** Set de claves "key||mes" ya presentes en la hoja de salida. */
function clavesExistentes_(sheet) {
  const set = {};
  const last = sheet.getLastRow();
  if (last < 2) return set;
  const display = sheet.getRange(2, 1, last - 1, OUT_HEADERS.length).getDisplayValues();
  display.forEach(function (row) {
    const serie = String(row[OUT.serie - 1]).trim();
    const inv = String(row[OUT.inventario - 1]).trim();
    const key = serie ? 'S:' + serie : (inv ? 'I:' + inv : '');
    const mes = String(row[OUT.mes - 1]).trim();
    if (key && mes) set[key + '||' + mes] = true;
  });
  return set;
}

/** Siguiente correlativo a partir del mayor ID Programación existente. */
function siguienteCorrelativo_(sheet) {
  const last = sheet.getLastRow();
  let max = 0;
  if (last >= 2) {
    const ids = sheet.getRange(2, OUT.idProg, last - 1, 1).getDisplayValues();
    ids.forEach(function (r) {
      const m = String(r[0]).match(/(\d+)\s*$/);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  return max + 1;
}

function extraerSpreadsheetId_(urlOrId) {
  const s = String(urlOrId || '').trim();
  if (!s) return '';
  const m = s.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9-_]{20,}$/.test(s)) return s; // parece un ID directo
  return '';
}

/** Convierte 'YYYY-MM-DD' (input date) a Date local sin desfase de zona horaria. */
function parseFechaLocal_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) {
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}
