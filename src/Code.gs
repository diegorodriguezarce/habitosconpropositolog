/**
 * Hábitos con propósito · Bitácora de consultas
 *
 * Web app de Apps Script ligada a una Google Sheet. La usa solo la administradora:
 * registra cada atención en una única tabla (hoja "Registros"), lee tarifas y
 * catálogos desde hojas editables y genera reportes mensuales.
 */

const TZ = 'America/Costa_Rica';

const HOJAS = {
  registros: 'Registros',
  tarifas: 'Tarifas',
  ubicaciones: 'Ubicaciones',
  aseguradoras: 'Aseguradoras',
  descuentos: 'Descuentos',
  tipos: 'Tipos de cita',
  formas: 'Formas de pago',
  config: 'Config',
};

const ESTADOS_ASEGURADORA = ['Pendiente', 'Facturado', 'Pagado'];
const BASES_COMISION = ['Total facturado', 'Pago paciente'];
const NO_APLICA = 'No aplica';

const COLOR = {
  crema: '#F7F2E8',
  crema2: '#FBF8F1',
  linea: '#E6DECD',
  verdeClaro: '#D2DCA1',
  salvia: '#A8B77A',
  oscuro: '#304323',
  amarillo: '#F4D978',
  amarilloSuave: '#FBF0C6',
  gris: '#939A85',
  grisSuave: '#ECEAE4',
};

/** Claves de la hoja Config con su valor inicial. */
const CONFIG_INICIAL = [
  ['Nutricionista', 'Moni', 'Nombre de la nutricionista (saludo de la app y su parte en los reportes).'],
  ['Calendarios de Google', '', 'IDs o nombres de calendarios a revisar, separados por coma (incluye calendarios suscritos por iCal). Vacío = calendario principal. "ninguno" = no leer Google.'],
  ['iCal (enlaces .ics)', '', 'Opcional: enlaces .ics / webcal:// de la agenda, separados por coma. Se leen directo (más actualizado que la suscripción de Google).'],
  ['Solo eventos que contengan', '', 'Opcional: muestra solo eventos cuyo título contenga este texto.'],
  ['Carpeta de reportes', 'Hábitos con propósito · Reportes', 'Carpeta de Drive donde se guardan los reportes exportados.'],
];

/** Tarifas iniciales (por persona en pareja). Solo se cargan si la hoja Tarifas está vacía. */
const TARIFAS_INICIALES = {
  'Primera vez': { 'Escazú': 50000, 'Curridabat': 40000, 'Domicilio': 35000 },
  'Control': { 'Escazú': 35000, 'Curridabat': 30000, 'Domicilio': 30000 },
  'Pareja primera vez': { 'Escazú': 47000, 'Curridabat': 35000, 'Domicilio': 30000 },
  'Pareja control': { 'Escazú': 32000, 'Curridabat': 27000, 'Domicilio': 25000 },
  'Mediciones': { 'Escazú': 14000, 'Curridabat': NO_APLICA, 'Domicilio': 12000 },
};

const MONEDA = '"₡"#,##0';
const FECHA = 'dd/mm/yyyy';
const FECHA_HORA = 'dd/mm/yyyy hh:mm';

const MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto',
  'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];

/** Columnas de la hoja Registros. Se localizan por encabezado, no por posición. */
const COLUMNAS = [
  { key: 'id', header: 'ID', ancho: 80, formato: '@' },
  { key: 'fecha', header: 'Fecha', ancho: 90, formato: FECHA },
  { key: 'hora', header: 'Hora', ancho: 60, formato: '@' },
  { key: 'paciente', header: 'Paciente', ancho: 180, formato: '@' },
  { key: 'paciente2', header: 'Paciente 2', ancho: 160, formato: '@' },
  { key: 'tipo', header: 'Tipo de cita', ancho: 140, formato: '@' },
  { key: 'ubicacion', header: 'Ubicación', ancho: 100, formato: '@' },
  { key: 'conSeguro', header: 'Con seguro', ancho: 80, formato: '@' },
  { key: 'aseguradora', header: 'Aseguradora', ancho: 110, formato: '@' },
  { key: 'tarifa', header: 'Tarifa de lista', ancho: 100, formato: MONEDA },
  { key: 'descuento', header: 'Descuento', ancho: 90, formato: '@' },
  { key: 'montoDescuento', header: 'Monto descuento', ancho: 110, formato: MONEDA },
  { key: 'pago1', header: 'Pago 1', ancho: 90, formato: MONEDA },
  { key: 'forma1', header: 'Forma de pago 1', ancho: 110, formato: '@' },
  { key: 'pago2', header: 'Pago 2', ancho: 90, formato: MONEDA },
  { key: 'forma2', header: 'Forma de pago 2', ancho: 110, formato: '@' },
  { key: 'totalPaciente', header: 'Total paciente', ancho: 100, formato: MONEDA },
  { key: 'montoAseguradora', header: 'Monto aseguradora', ancho: 120, formato: MONEDA },
  { key: 'estadoAseguradora', header: 'Estado aseguradora', ancho: 120, formato: '@' },
  { key: 'totalFacturado', header: 'Total facturado', ancho: 110, formato: MONEDA },
  { key: 'notas', header: 'Notas', ancho: 220, formato: '@' },
  { key: 'eventoId', header: 'Evento Calendar', ancho: 120, formato: '@' },
  { key: 'creado', header: 'Registrado', ancho: 130, formato: FECHA_HORA },
  { key: 'modificado', header: 'Modificado', ancho: 130, formato: FECHA_HORA },
];

/* ───────────────────────── Setup de la hoja ───────────────────────── */

/**
 * Crea las hojas que falten con encabezados, valores iniciales y formato.
 * Es seguro correrlo varias veces: no borra datos; solo agrega columnas que
 * falten y completa celdas vacías de los valores iniciales.
 */
function setup() {
  const ss = ss_();
  ss.setSpreadsheetTimeZone(TZ);

  crearCatalogo_(ss, HOJAS.tipos, ['Tipo de cita', 'Es pareja', 'Activo'], [
    ['Primera vez', 'No', 'Sí'],
    ['Control', 'No', 'Sí'],
    ['Pareja primera vez', 'Sí', 'Sí'],
    ['Pareja control', 'Sí', 'Sí'],
    ['Mediciones', 'No', 'Sí'],
  ], [180, 90, 70], function (sh) {
    sh.getRange(1, 2).setNote('Sí = cita de dos personas: pide dos nombres y la tarifa se cobra por persona.');
  });

  crearCatalogo_(ss, HOJAS.ubicaciones, ['Ubicación', 'Encargado / a quién se paga', '% comisión', 'Base de comisión', 'Activa'], [
    ['Escazú', 'Andrea (DS · Diseñando tu Salud)', 0.4, 'Total facturado', 'Sí'],
    ['Curridabat', 'Raque', 0.4, 'Total facturado', 'Sí'],
    ['Domicilio', 'Andrea (DS · Diseñando tu Salud)', 0.2, 'Total facturado', 'Sí'],
  ], [130, 240, 100, 140, 70], function (sh) {
    sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setNumberFormat('0.##%');
    sh.getRange(2, 4, sh.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(BASES_COMISION, true).build());
    sh.getRange(1, 2).setNote('Si el mismo nombre aparece en varias ubicaciones, los reportes suman lo que se le paga en total.');
    sh.getRange(1, 3).setNote('Porcentaje que se le paga al encargado. El resto es de la nutricionista. Ej.: 40%.');
    sh.getRange(1, 4).setNote('"Total facturado" = paciente + aseguradora. "Pago paciente" = solo lo que paga el paciente.');
  });

  crearCatalogo_(ss, HOJAS.aseguradoras, ['Aseguradora', 'Monto aseguradora', 'Monto paciente', 'Ubicaciones', 'Activa'], [
    ['Sana Sana', 8500, 16000, 'Escazú', 'Sí'],
  ], [160, 140, 130, 200, 70], function (sh) {
    sh.getRange(2, 2, sh.getMaxRows() - 1, 2).setNumberFormat(MONEDA);
    sh.getRange(1, 2).setNote('Monto fijo que paga la aseguradora por consulta. En la app se puede editar.');
    sh.getRange(1, 3).setNote('Monto que se le cobra al paciente cuando usa este seguro. En la app se puede editar.');
    sh.getRange(1, 4).setNote('Ubicaciones donde se acepta, separadas por coma. Vacío = todas.');
  }, { 'Monto por consulta': 'Monto aseguradora' });

  crearCatalogo_(ss, HOJAS.descuentos, ['Descuento', '% descuento', 'Ubicaciones', 'Excepto tipos', 'Activo'], [
    ['VOLK', 0.1, 'Domicilio', 'Mediciones', 'Sí'],
  ], [140, 110, 200, 200, 70], function (sh) {
    sh.getRange(2, 2, sh.getMaxRows() - 1, 1).setNumberFormat('0.##%');
    sh.getRange(1, 3).setNote('Ubicaciones donde aplica, separadas por coma. Vacío = todas.');
    sh.getRange(1, 4).setNote('Tipos de cita a los que NO aplica, separados por coma.');
  });

  crearCatalogo_(ss, HOJAS.formas, ['Forma de pago', 'Activa'], [
    ['SINPE', 'Sí'],
    ['Tarjeta', 'Sí'],
  ], [160, 70]);

  const config = crearCatalogo_(ss, HOJAS.config, ['Clave', 'Valor', 'Descripción'], CONFIG_INICIAL, [200, 260, 420]);
  asegurarClavesConfig_(config);

  crearRegistros_(ss);
  sincronizarTarifas();

  // Orden de pestañas
  [HOJAS.registros, HOJAS.tarifas, HOJAS.ubicaciones, HOJAS.aseguradoras, HOJAS.descuentos, HOJAS.tipos, HOJAS.formas, HOJAS.config]
    .forEach(function (n, i) {
      const sh = ss.getSheetByName(n);
      if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
    });
  ss.setActiveSheet(ss.getSheetByName(HOJAS.registros));

  // Borrar la hoja vacía que trae un archivo nuevo
  ['Hoja 1', 'Sheet1', 'Hoja1'].forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
}

function crearRegistros_(ss) {
  let sh = ss.getSheetByName(HOJAS.registros);
  if (!sh) sh = ss.insertSheet(HOJAS.registros);
  const nCols = Math.max(sh.getLastColumn(), 1);
  const actuales = sh.getRange(1, 1, 1, nCols).getValues()[0].map(function (h) { return String(h).trim(); });
  const idx = indiceColumnas_(actuales);
  const headers = actuales.filter(function (h) { return h; });
  COLUMNAS.forEach(function (c) { if (idx[c.key] < 0) headers.push(c.header); });
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  estiloEncabezado_(sh, headers.length);

  const fin = indiceColumnas_(headers);
  const filas = sh.getMaxRows() - 1;
  COLUMNAS.forEach(function (c) {
    const col = fin[c.key] + 1;
    sh.setColumnWidth(col, c.ancho);
    sh.getRange(2, col, filas, 1).setNumberFormat(c.formato);
  });
  sh.setFrozenColumns(Math.max(fin.paciente + 1, 0));

  // Listas desplegables (aceptan valores fuera de la lista, solo avisan)
  const lista = function (hoja) {
    const h = ss.getSheetByName(hoja);
    return SpreadsheetApp.newDataValidation().requireValueInRange(h.getRange(2, 1, h.getMaxRows() - 1, 1), true)
      .setAllowInvalid(true).build();
  };
  const opciones = function (arr) {
    return SpreadsheetApp.newDataValidation().requireValueInList(arr, true).setAllowInvalid(true).build();
  };
  const validar = function (key, regla) { if (fin[key] >= 0) sh.getRange(2, fin[key] + 1, filas, 1).setDataValidation(regla); };
  validar('tipo', lista(HOJAS.tipos));
  validar('ubicacion', lista(HOJAS.ubicaciones));
  validar('aseguradora', lista(HOJAS.aseguradoras));
  validar('descuento', lista(HOJAS.descuentos));
  validar('forma1', lista(HOJAS.formas));
  validar('forma2', lista(HOJAS.formas));
  validar('conSeguro', opciones(['Sí', 'No']));
  validar('estadoAseguradora', opciones(ESTADOS_ASEGURADORA));

  try {
    if (!sh.getBandings().length) {
      sh.getRange(1, 1, sh.getMaxRows(), headers.length)
        .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREEN, true, false)
        .setHeaderRowColor(COLOR.oscuro).setFirstRowColor('#FFFFFF').setSecondRowColor(COLOR.crema2);
    }
  } catch (e) { /* el banding es solo estético */ }
  return sh;
}

/**
 * Crea o actualiza una hoja de catálogo. Si ya existe: renombra encabezados viejos,
 * agrega columnas que falten y completa solo las celdas vacías de las filas iniciales.
 */
function crearCatalogo_(ss, nombre, headers, iniciales, anchos, extra, renombrar) {
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (iniciales.length) sh.getRange(2, 1, iniciales.length, headers.length).setValues(iniciales);
  } else {
    completarCatalogo_(sh, headers, iniciales, renombrar || {});
  }
  estiloEncabezado_(sh, Math.max(headers.length, sh.getLastColumn()));
  anchos.forEach(function (a, i) { sh.setColumnWidth(i + 1, a); });
  if (extra) extra(sh);
  return sh;
}

function completarCatalogo_(sh, headers, iniciales, renombrar) {
  const nCols = Math.max(sh.getLastColumn(), 1);
  const v = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), nCols).getValues().map(function (r) { return r.slice(); });
  let cab = v[0].map(function (h) { const t = String(h).trim(); return renombrar[t] || t; });
  headers.forEach(function (h) {
    if (cab.indexOf(h) < 0) {
      cab.push(h);
      v.forEach(function (r) { r.push(''); });
    }
  });
  v[0] = cab;
  iniciales.forEach(function (ini) {
    const clave = normTxt_(ini[0]);
    const fila = v.slice(1).filter(function (r) { return normTxt_(r[0]) === clave; })[0];
    if (!fila) return;
    headers.forEach(function (h, i) {
      const j = cab.indexOf(h);
      if ((fila[j] === '' || fila[j] == null) && ini[i] !== '') fila[j] = ini[i];
    });
  });
  sh.getRange(1, 1, v.length, cab.length).setValues(v);
}

/** Agrega a Config las claves nuevas que no existan (sin tocar valores). */
function asegurarClavesConfig_(sh) {
  const n = sh.getLastRow();
  const claves = n >= 2 ? sh.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  const faltan = CONFIG_INICIAL.filter(function (f) { return claves.indexOf(f[0]) < 0; });
  if (faltan.length) sh.getRange(n + 1, 1, faltan.length, 3).setValues(faltan);
}

/** Ajusta la hoja Tarifas a los tipos y ubicaciones activos, sin perder montos. */
function sincronizarTarifas() {
  const ss = ss_();
  const cats = catalogos_();
  let sh = ss.getSheetByName(HOJAS.tarifas);
  if (!sh) sh = ss.insertSheet(HOJAS.tarifas);

  let v = sh.getLastRow() ? sh.getRange(1, 1, sh.getLastRow(), Math.max(sh.getLastColumn(), 1)).getValues() : [['Tipo de cita']];
  v = v.map(function (r) { return r.slice(); });
  v[0][0] = 'Tipo de cita';
  const cols = v[0].map(function (h) { return String(h).trim(); });
  cats.ubicaciones.forEach(function (u) {
    if (cols.indexOf(u.nombre) < 0) {
      cols.push(u.nombre);
      v.forEach(function (r, i) { r.push(i === 0 ? u.nombre : ''); });
    }
  });
  const tipos = v.slice(1).map(function (r) { return String(r[0]).trim(); });
  cats.tipos.forEach(function (t) {
    if (tipos.indexOf(t.nombre) < 0) v.push([t.nombre].concat(cols.slice(1).map(function () { return ''; })));
  });

  // Hoja recién creada (sin ningún monto): cargar las tarifas iniciales.
  const vacia = v.slice(1).every(function (r) { return r.slice(1).every(function (x) { return x === '' || x == null; }); });
  if (vacia) {
    v.slice(1).forEach(function (r) {
      const ini = TARIFAS_INICIALES[String(r[0]).trim()] || {};
      for (let j = 1; j < cols.length; j++) if (ini[cols[j]] !== undefined) r[j] = ini[cols[j]];
    });
  }

  sh.getRange(1, 1, v.length, cols.length).setValues(v);
  estiloEncabezado_(sh, cols.length);
  sh.setFrozenColumns(1);
  sh.setColumnWidth(1, 180);
  for (let c = 2; c <= cols.length; c++) sh.setColumnWidth(c, 120);
  if (v.length > 1 && cols.length > 1) {
    const datos = sh.getRange(2, 2, v.length - 1, cols.length - 1);
    datos.setNumberFormat(MONEDA).setHorizontalAlignment('right');
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenCellEmpty().setBackground(COLOR.amarilloSuave).setRanges([datos]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(NO_APLICA).setBackground(COLOR.grisSuave)
        .setFontColor(COLOR.gris).setRanges([datos]).build(),
    ]);
  }
  sh.getRange(1, 1).setNote('Tarifa por tipo de cita y ubicación. En citas de pareja es el monto POR PERSONA. ' +
    'Celda vacía (amarilla) = sin tarifa: en la app se digita el monto. "' + NO_APLICA + '" = ese tipo no se ofrece en esa ubicación.');
}

function estiloEncabezado_(sh, n) {
  sh.getRange(1, 1, 1, n)
    .setBackground(COLOR.oscuro).setFontColor(COLOR.crema).setFontWeight('bold')
    .setFontFamily('Montserrat').setVerticalAlignment('middle');
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 30);
}

/* ───────────────────────────── Web app ───────────────────────────── */

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Hábitos con propósito · Bitácora')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(nombre) {
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Hábitos')
    .addItem('Abrir la app', 'abrirApp')
    .addSeparator()
    .addItem('Configurar / reparar hojas', 'setup')
    .addItem('Sincronizar tarifas con catálogos', 'sincronizarTarifas')
    .addSeparator()
    .addItem('Borrar todos los registros (datos de prueba)', 'limpiarRegistros')
    .addToUi();
}

/* ─────────────────────── API para el cliente ─────────────────────── */

/** Todo lo que el formulario necesita al cargar. */
function getBootstrap() {
  if (!ss_().getSheetByName(HOJAS.registros)) setup();
  const regs = leerRegistros_().map(function (x) { return x.reg; });
  const cfg = config_();
  return {
    cats: catalogos_(),
    tarifas: tarifas_(),
    pacientes: pacientes_(regs),
    recientes: recientes_(regs, 10),
    hoy: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'),
    nutricionista: cfg['Nutricionista'] || '',
  };
}

/**
 * Citas de un día (yyyy-MM-dd) desde los calendarios de Google configurados y
 * los enlaces iCal (.ics), con sugerencias para llenar el formulario.
 */
function getEventos(fechaStr) {
  const ini = aFecha_(fechaStr);
  if (!ini) throw new Error('Fecha inválida.');
  const fin = new Date(ini.getFullYear(), ini.getMonth(), ini.getDate() + 1);
  const cfg = config_();
  const avisos = [];
  let crudos = [];

  calendarios_(cfg, avisos).forEach(function (cal) {
    try {
      cal.getEvents(ini, fin).forEach(function (e) {
        if (e.isAllDayEvent()) return;
        crudos.push({ id: e.getId(), titulo: e.getTitle(), lugar: e.getLocation() || '',
          inicio: e.getStartTime(), fin: e.getEndTime(), fuente: 'Google' });
      });
    } catch (err) {
      avisos.push('No se pudo leer un calendario de Google: ' + err.message);
    }
  });

  listaConfig_(cfg['iCal (enlaces .ics)']).forEach(function (url) {
    try {
      eventosIcs_(url, ini, fin).forEach(function (e) { crudos.push(e); });
    } catch (err) {
      avisos.push('No se pudo leer el iCal: ' + err.message);
    }
  });

  // La misma cita puede venir de Google (calendario suscrito) y del .ics: se deja una.
  const vistos = {};
  crudos = crudos
    .sort(function (a, b) { return a.inicio - b.inicio; })
    .filter(function (e) {
      const k = normTxt_(e.titulo) + '|' + Utilities.formatDate(e.inicio, TZ, 'HH:mm');
      if (vistos[k]) return false;
      vistos[k] = true;
      return true;
    });

  const filtro = normTxt_(cfg['Solo eventos que contengan'] || '');
  const registrados = {};
  leerRegistros_().forEach(function (x) { if (x.reg.eventoId) registrados[x.reg.eventoId] = true; });
  const cats = catalogos_();

  const eventos = crudos
    .filter(function (e) { return !filtro || normTxt_(e.titulo).indexOf(filtro) >= 0; })
    .map(function (e) {
      const id = e.id + '|' + Utilities.formatDate(e.inicio, TZ, "yyyy-MM-dd'T'HH:mm");
      const sug = sugerirDesdeEvento_(e.titulo, e.lugar, cats);
      return {
        id: id,
        titulo: e.titulo,
        lugar: e.lugar,
        fuente: e.fuente,
        hora: Utilities.formatDate(e.inicio, TZ, 'HH:mm'),
        horaFin: e.fin ? Utilities.formatDate(e.fin, TZ, 'HH:mm') : '',
        registrado: !!registrados[id],
        paciente: sug.paciente,
        paciente2: sug.paciente2,
        tipo: sug.tipo,
        ubicacion: sug.ubicacion,
      };
    });
  return { eventos: eventos, error: avisos.join(' ') };
}

/** Crea (sin id) o actualiza (con id) un registro. */
function guardarRegistro(data) {
  const reg = normalizarRegistro_(data || {}, catalogos_(), tarifas_());

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let guardado;
  try {
    const sh = hoja_(HOJAS.registros);
    const nCols = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, nCols).getValues()[0];
    const idx = indiceColumnas_(headers);
    const ahora = new Date();
    let fila;
    let actual;
    if (reg.id) {
      fila = buscarFila_(sh, idx, reg.id);
      if (!fila) throw new Error('No se encontró el registro a editar (¿lo borraron de la hoja?).');
      actual = sh.getRange(fila, 1, 1, nCols).getValues()[0];
      reg.creado = idx.creado >= 0 ? actual[idx.creado] : '';
    } else {
      reg.id = Utilities.getUuid().slice(0, 8);
      fila = sh.getLastRow() + 1;
      actual = headers.map(function () { return ''; });
      reg.creado = ahora;
    }
    reg.modificado = ahora;

    const valores = actual.slice();
    COLUMNAS.forEach(function (c) {
      if (idx[c.key] >= 0) valores[idx[c.key]] = valorCelda_(c.key, reg[c.key]);
    });
    const rango = sh.getRange(fila, 1, 1, nCols);
    rango.setNumberFormats([formatosFila_(sh, fila, nCols, idx)]);
    rango.setValues([valores]);
    guardado = filaAObjeto_(valores, idx);
  } finally {
    lock.releaseLock();
  }
  const regs = leerRegistros_().map(function (x) { return x.reg; });
  return { registro: guardado, recientes: recientes_(regs, 10), pacientes: pacientes_(regs) };
}

function getRecientes(n) {
  return recientes_(leerRegistros_().map(function (x) { return x.reg; }), n || 10);
}

function eliminarRegistro(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = hoja_(HOJAS.registros);
    const idx = indiceColumnas_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
    const fila = buscarFila_(sh, idx, id);
    if (!fila) throw new Error('No se encontró el registro.');
    sh.deleteRow(fila);
  } finally {
    lock.releaseLock();
  }
  const regs = leerRegistros_().map(function (x) { return x.reg; });
  return { recientes: recientes_(regs, 10), pacientes: pacientes_(regs) };
}

/** Resumen del mes (yyyy-MM) para la pantalla de reportes. */
function getReporteMes(mes) {
  validarMes_(mes);
  const regs = registrosDelMes_(mes);
  const r = resumen_(regs, catalogos_());
  r.mes = mes;
  r.etiqueta = etiquetaMes_(mes);
  return r;
}

/**
 * Cambia el estado de cobro de las consultas con seguro de un mes
 * (p. ej. Pendiente → Facturado al enviar la factura a la aseguradora).
 */
function marcarEstadoAseguradora(mes, aseguradora, desde, hacia) {
  validarMes_(mes);
  if (ESTADOS_ASEGURADORA.indexOf(hacia) < 0) throw new Error('Estado inválido.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let cambiados = 0;
  try {
    const sh = hoja_(HOJAS.registros);
    const last = sh.getLastRow();
    if (last >= 2) {
      const nCols = sh.getLastColumn();
      const valores = sh.getRange(1, 1, last, nCols).getValues();
      const idx = indiceColumnas_(valores[0]);
      if (idx.estadoAseguradora < 0) throw new Error('Falta la columna "Estado aseguradora".');
      const colEstado = [];
      const colMod = [];
      const ahora = new Date();
      for (let i = 1; i < valores.length; i++) {
        const r = filaAObjeto_(valores[i], idx);
        let estado = valores[i][idx.estadoAseguradora];
        let mod = idx.modificado >= 0 ? valores[i][idx.modificado] : '';
        if (r.conSeguro && r.fecha.slice(0, 7) === mes && r.aseguradora === aseguradora &&
            (r.estadoAseguradora || 'Pendiente') === desde) {
          estado = hacia;
          mod = ahora;
          cambiados++;
        }
        colEstado.push([estado]);
        colMod.push([mod]);
      }
      sh.getRange(2, idx.estadoAseguradora + 1, colEstado.length, 1).setValues(colEstado);
      if (idx.modificado >= 0) sh.getRange(2, idx.modificado + 1, colMod.length, 1).setValues(colMod);
    }
  } finally {
    lock.releaseLock();
  }
  return { cambiados: cambiados, reporte: getReporteMes(mes) };
}

/**
 * Genera un reporte como Google Sheet nueva (carpeta de reportes en Drive)
 * y devuelve enlaces para abrirla o descargarla en PDF / Excel.
 * opc: { tipo: 'general' | 'sede' | 'encargado' | 'aseguradora', mes: 'yyyy-MM', filtro, nombres: bool }
 */
function exportarReporte(opc) {
  opc = opc || {};
  validarMes_(opc.mes);
  const cats = catalogos_();
  const regs = registrosDelMes_(opc.mes);
  const etiqueta = etiquetaMes_(opc.mes);
  let titulo;
  if (opc.tipo === 'sede' || opc.tipo === 'encargado') titulo = 'Liquidación ' + opc.filtro + ' · ' + etiqueta;
  else if (opc.tipo === 'aseguradora') titulo = 'Cobro ' + opc.filtro + ' · ' + etiqueta;
  else titulo = 'Reporte general · ' + etiqueta;

  const ss = SpreadsheetApp.create(titulo);
  ss.setSpreadsheetTimeZone(TZ);
  const sh = ss.getSheets()[0];
  sh.setName('Reporte');
  const w = escritor_(sh);
  w.encabezado(titulo, 'Generado el ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy'));

  if (opc.tipo === 'sede') reporteSede_(w, regs, cats, opc);
  else if (opc.tipo === 'encargado') reporteEncargado_(w, regs, cats, opc);
  else if (opc.tipo === 'aseguradora') reporteAseguradora_(w, regs, opc);
  else reporteGeneral_(w, regs, cats);

  w.cerrar();
  moverACarpeta_(ss.getId());
  SpreadsheetApp.flush();
  return {
    nombre: titulo,
    url: ss.getUrl(),
    pdf: urlExport_(ss.getId(), sh.getSheetId(), 'pdf'),
    xlsx: urlExport_(ss.getId(), sh.getSheetId(), 'xlsx'),
  };
}

/* ───────────────────────── Registros ───────────────────────── */

/** Valida y completa un registro; calcula tarifa de lista y descuento. */
function normalizarRegistro_(d, cats, tarifas) {
  const s = function (v) { return v == null ? '' : String(v).trim(); };
  const faltan = [];

  const fecha = s(d.fecha);
  if (!aFecha_(fecha)) faltan.push('fecha');
  const hora = s(d.hora);
  if (hora && !/^\d{1,2}:\d{2}$/.test(hora)) faltan.push('hora (HH:MM)');

  const tipo = s(d.tipo);
  if (!tipo) faltan.push('tipo de cita');
  const tipoCat = cats.tipos.filter(function (t) { return t.nombre === tipo; })[0];
  const pareja = tipoCat ? tipoCat.pareja : /pareja/i.test(tipo);

  const ubicacion = s(d.ubicacion);
  if (!ubicacion) faltan.push('ubicación');
  const paciente = s(d.paciente);
  if (!paciente) faltan.push('nombre del paciente');
  const paciente2 = pareja ? s(d.paciente2) : '';
  if (pareja && !paciente2) faltan.push('nombre de la segunda persona');

  const conSeguro = d.conSeguro === true || esSi_(d.conSeguro);
  const aseguradora = conSeguro ? s(d.aseguradora) : '';
  if (conSeguro && !aseguradora) faltan.push('aseguradora');

  const dosPagos = d.dosPagos === true || esSi_(d.dosPagos);
  const pago1 = monto_(d.pago1);
  const pago2 = dosPagos ? monto_(d.pago2) : 0;
  const montoAseguradora = conSeguro ? monto_(d.montoAseguradora) : 0;
  if (isNaN(pago1) || isNaN(pago2) || isNaN(montoAseguradora)) faltan.push('montos válidos (números enteros ≥ 0)');

  const forma1 = pago1 > 0 ? s(d.forma1) : '';
  if (pago1 > 0 && !forma1) faltan.push('forma de pago');
  const forma2 = dosPagos && pago2 > 0 ? s(d.forma2) : '';
  if (dosPagos && pago2 > 0 && !forma2) faltan.push('forma del segundo pago');

  if (faltan.length) throw new Error('Falta completar: ' + faltan.join(', ') + '.');

  // Reglas de catálogo: tipo ofrecido en la ubicación, seguro y descuento permitidos.
  const tarifaBase = tarifas[tipo] ? tarifas[tipo][ubicacion] : null;
  if (tarifaBase === false) throw new Error(tipo + ' no se ofrece en ' + ubicacion + '.');
  const aseg = cats.aseguradoras.filter(function (a) { return a.nombre === aseguradora; })[0];
  if (conSeguro && aseg && !aplicaEn_(aseg.ubicaciones, ubicacion)) {
    throw new Error(aseguradora + ' solo aplica en ' + aseg.ubicaciones.join(', ') + '.');
  }
  const personas = pareja ? 2 : 1;
  const tarifa = typeof tarifaBase === 'number' ? tarifaBase * personas : '';
  const descuento = conSeguro ? '' : s(d.descuento);
  let montoDescuento = '';
  if (descuento) {
    const desc = cats.descuentos.filter(function (x) { return x.nombre === descuento; })[0];
    if (desc && !descuentoAplica_(desc, ubicacion, tipo)) {
      throw new Error('El descuento ' + descuento + ' no aplica a ' + tipo + ' en ' + ubicacion + '.');
    }
    if (desc && tarifa !== '') montoDescuento = Math.round(tarifa * desc.porcentaje);
  }

  const estado = conSeguro
    ? (ESTADOS_ASEGURADORA.indexOf(s(d.estadoAseguradora)) >= 0 ? s(d.estadoAseguradora) : 'Pendiente')
    : '';
  const totalPaciente = pago1 + pago2;
  return {
    id: s(d.id),
    fecha: fecha,
    hora: hora ? pad2_(hora.split(':')[0]) + ':' + hora.split(':')[1] : '',
    paciente: paciente,
    paciente2: paciente2,
    tipo: tipo,
    ubicacion: ubicacion,
    conSeguro: conSeguro,
    aseguradora: aseguradora,
    tarifa: tarifa,
    descuento: descuento,
    montoDescuento: montoDescuento,
    pago1: pago1,
    forma1: forma1,
    pago2: dosPagos ? pago2 : '',
    forma2: forma2,
    totalPaciente: totalPaciente,
    montoAseguradora: conSeguro ? montoAseguradora : '',
    estadoAseguradora: estado,
    totalFacturado: totalPaciente + montoAseguradora,
    notas: s(d.notas),
    eventoId: s(d.eventoId),
  };
}

/** Convierte un valor del registro al valor que se escribe en la celda. */
function valorCelda_(key, v) {
  if (key === 'fecha') return aFecha_(v);
  if (key === 'conSeguro') return v ? 'Sí' : 'No';
  if (v === undefined || v === null) return '';
  return v;
}

function leerRegistros_() {
  const sh = hoja_(HOJAS.registros);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const valores = sh.getRange(1, 1, last, sh.getLastColumn()).getValues();
  const idx = indiceColumnas_(valores[0]);
  const out = [];
  for (let i = 1; i < valores.length; i++) {
    const reg = filaAObjeto_(valores[i], idx);
    if (reg.id || reg.paciente) out.push({ fila: i + 1, reg: reg });
  }
  return out;
}

/** Fila de la hoja → objeto plano (sin Date, apto para google.script.run). */
function filaAObjeto_(fila, idx) {
  const g = function (k) { return idx[k] >= 0 ? fila[idx[k]] : ''; };
  const s = function (k) { const v = g(k); return v == null ? '' : String(v).trim(); };
  const conSeguro = esSi_(g('conSeguro'));
  const pago1 = num_(g('pago1'));
  const pago2 = num_(g('pago2'));
  const montoAseguradora = conSeguro ? num_(g('montoAseguradora')) : 0;
  // Los totales se recalculan siempre, por si alguien edita montos directo en la hoja.
  const totalPaciente = pago1 + pago2;
  return {
    id: s('id'),
    fecha: fechaStr_(g('fecha')),
    hora: horaStr_(g('hora')),
    paciente: s('paciente'),
    paciente2: s('paciente2'),
    tipo: s('tipo'),
    ubicacion: s('ubicacion'),
    conSeguro: conSeguro,
    aseguradora: conSeguro ? s('aseguradora') : '',
    tarifa: g('tarifa') === '' ? null : num_(g('tarifa')),
    descuento: s('descuento'),
    montoDescuento: num_(g('montoDescuento')),
    pago1: pago1,
    forma1: s('forma1'),
    pago2: pago2,
    forma2: s('forma2'),
    totalPaciente: totalPaciente,
    montoAseguradora: montoAseguradora,
    estadoAseguradora: conSeguro ? (s('estadoAseguradora') || 'Pendiente') : '',
    totalFacturado: totalPaciente + montoAseguradora,
    notas: s('notas'),
    eventoId: s('eventoId'),
    creado: fechaHoraStr_(g('creado')),
    modificado: fechaHoraStr_(g('modificado')),
  };
}

function indiceColumnas_(headers) {
  const norm = headers.map(function (h) { return normTxt_(h); });
  const idx = {};
  COLUMNAS.forEach(function (c) { idx[c.key] = norm.indexOf(normTxt_(c.header)); });
  return idx;
}

function buscarFila_(sh, idx, id) {
  const last = sh.getLastRow();
  if (last < 2 || idx.id < 0) return 0;
  const ids = sh.getRange(2, idx.id + 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === String(id).trim()) return i + 2;
  return 0;
}

function formatosFila_(sh, fila, nCols, idx) {
  const formatos = sh.getRange(fila, 1, 1, nCols).getNumberFormats()[0].slice();
  COLUMNAS.forEach(function (c) { if (idx[c.key] >= 0) formatos[idx[c.key]] = c.formato; });
  return formatos;
}

function recientes_(regs, n) {
  return regs.slice()
    .map(function (r, i) { return { r: r, i: i }; })
    .sort(function (a, b) {
      if (a.r.creado !== b.r.creado) return a.r.creado < b.r.creado ? 1 : -1;
      return b.i - a.i;
    })
    .slice(0, n)
    .map(function (x) { return x.r; });
}

/** Pacientes conocidos con su última consulta, para autocompletar. */
function pacientes_(regs) {
  const mapa = {};
  regs.forEach(function (r) {
    [r.paciente, r.paciente2].forEach(function (n) {
      if (!n) return;
      const k = normTxt_(n);
      if (!mapa[k] || mapa[k].u < r.fecha) mapa[k] = { n: n, u: r.fecha };
    });
  });
  return Object.keys(mapa).map(function (k) { return mapa[k]; })
    .sort(function (a, b) { return a.n.localeCompare(b.n, 'es'); });
}

function registrosDelMes_(mes) {
  return leerRegistros_()
    .map(function (x) { return x.reg; })
    .filter(function (r) { return r.fecha.slice(0, 7) === mes; })
    .sort(function (a, b) { return (a.fecha + a.hora) < (b.fecha + b.hora) ? -1 : 1; });
}

/* ───────────────────────── Catálogos ───────────────────────── */

function catalogos_() {
  const activo = function (r, col) { return !esNo_(r[col]); };
  const tipos = leerTabla_(HOJAS.tipos)
    .filter(function (r) { return r['Tipo de cita'] && activo(r, 'Activo'); })
    .map(function (r) { return { nombre: String(r['Tipo de cita']).trim(), pareja: esSi_(r['Es pareja']) }; });
  const ubicaciones = leerTabla_(HOJAS.ubicaciones)
    .filter(function (r) { return r['Ubicación'] && activo(r, 'Activa'); })
    .map(function (r) {
      return {
        nombre: String(r['Ubicación']).trim(),
        encargado: String(r['Encargado / a quién se paga'] || '').trim(),
        porcentaje: pct_(r['% comisión']),
        base: BASES_COMISION.indexOf(String(r['Base de comisión']).trim()) >= 0
          ? String(r['Base de comisión']).trim() : BASES_COMISION[0],
      };
    });
  const montoONull = function (m) { return m === '' || m == null ? null : num_(m); };
  const aseguradoras = leerTabla_(HOJAS.aseguradoras)
    .filter(function (r) { return r['Aseguradora'] && activo(r, 'Activa'); })
    .map(function (r) {
      return {
        nombre: String(r['Aseguradora']).trim(),
        monto: montoONull(r['Monto aseguradora'] !== undefined ? r['Monto aseguradora'] : r['Monto por consulta']),
        montoPaciente: montoONull(r['Monto paciente']),
        ubicaciones: listaConfig_(r['Ubicaciones']),
      };
    });
  const descuentos = leerTabla_(HOJAS.descuentos)
    .filter(function (r) { return r['Descuento'] && activo(r, 'Activo'); })
    .map(function (r) {
      return {
        nombre: String(r['Descuento']).trim(),
        porcentaje: pct_(r['% descuento']),
        ubicaciones: listaConfig_(r['Ubicaciones']),
        exceptoTipos: listaConfig_(r['Excepto tipos']),
      };
    });
  const formas = leerTabla_(HOJAS.formas)
    .filter(function (r) { return r['Forma de pago'] && activo(r, 'Activa'); })
    .map(function (r) { return String(r['Forma de pago']).trim(); });
  return {
    tipos: tipos, ubicaciones: ubicaciones, aseguradoras: aseguradoras, descuentos: descuentos,
    formas: formas, estados: ESTADOS_ASEGURADORA,
  };
}

/** Lista de ubicaciones vacía = aplica en todas. */
function aplicaEn_(lista, ubicacion) {
  if (!lista || !lista.length) return true;
  const u = normTxt_(ubicacion);
  return lista.some(function (x) { return normTxt_(x) === u; });
}

function descuentoAplica_(desc, ubicacion, tipo) {
  const t = normTxt_(tipo);
  return aplicaEn_(desc.ubicaciones, ubicacion) &&
    !desc.exceptoTipos.some(function (x) { return normTxt_(x) === t; });
}

/**
 * Tarifas en matriz: filas = tipo de cita, columnas = ubicación (por persona en pareja).
 * Valor: número, null (vacío = sin tarifa) o false ("No aplica" = no se ofrece ahí).
 */
function tarifas_() {
  const sh = ss_().getSheetByName(HOJAS.tarifas);
  const out = {};
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < 2) return out;
  const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const ubic = v[0].map(function (h) { return String(h).trim(); });
  for (let i = 1; i < v.length; i++) {
    const tipo = String(v[i][0]).trim();
    if (!tipo) continue;
    out[tipo] = {};
    for (let j = 1; j < ubic.length; j++) {
      if (!ubic[j]) continue;
      const x = v[i][j];
      if (x === '' || x == null) out[tipo][ubic[j]] = null;
      else if (typeof x === 'string' && !/\d/.test(x)) out[tipo][ubic[j]] = false;
      else out[tipo][ubic[j]] = num_(x);
    }
  }
  return out;
}

function config_() {
  const out = {};
  leerTabla_(HOJAS.config).forEach(function (r) {
    if (r['Clave']) out[String(r['Clave']).trim()] = r['Valor'] == null ? '' : String(r['Valor']).trim();
  });
  return out;
}

function leerTabla_(nombre) {
  const sh = ss_().getSheetByName(nombre);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const headers = v[0].map(function (h) { return String(h).trim(); });
  return v.slice(1)
    .filter(function (r) { return r.some(function (x) { return x !== '' && x != null; }); })
    .map(function (r) {
      const o = {};
      headers.forEach(function (h, i) { if (h) o[h] = r[i]; });
      return o;
    });
}

/* ───────────────────────── Calendar ───────────────────────── */

/**
 * Calendarios de Google a revisar (Config → "Calendarios de Google", separados por coma).
 * Vacío = calendario principal. "ninguno" = no leer Google (solo iCal).
 */
function calendarios_(cfg, avisos) {
  const lista = listaConfig_(cfg['Calendarios de Google']);
  if (!lista.length) lista.push('principal');
  const out = [];
  lista.forEach(function (v) {
    const n = normTxt_(v);
    if (n === 'ninguno') return;
    try {
      const cal = n === 'principal'
        ? CalendarApp.getDefaultCalendar()
        : CalendarApp.getCalendarById(v) || CalendarApp.getCalendarsByName(v)[0];
      if (cal) out.push(cal);
      else avisos.push('No se encontró el calendario "' + v + '".');
    } catch (err) {
      avisos.push('No se pudo abrir el calendario "' + v + '": ' + err.message);
    }
  });
  return out;
}

function listaConfig_(v) {
  return String(v || '').split(/[,\n]/).map(function (s) { return s.trim(); }).filter(Boolean);
}

/* ─────────────── iCal (.ics) ─────────────── */

/** Eventos de un enlace .ics que empiezan dentro de [ini, fin). */
function eventosIcs_(url, ini, fin) {
  const texto = descargarIcs_(url);
  const out = [];
  parsearIcs_(texto).forEach(function (ev) {
    if (!ev.inicio || ev.todoElDia || ev.cancelado) return;
    const dur = ev.fin ? ev.fin - ev.inicio : 60 * 60000;
    ocurrencias_(ev, ini, fin).forEach(function (inicio) {
      out.push({ id: 'ics:' + ev.uid, titulo: ev.titulo, lugar: ev.lugar, inicio: inicio,
        fin: new Date(inicio.getTime() + dur), fuente: 'iCal' });
    });
  });
  return out;
}

/** Descarga el .ics (webcal:// → https://) con caché de 5 minutos. */
function descargarIcs_(url) {
  url = url.replace(/^webcal:\/\//i, 'https://');
  const cache = CacheService.getScriptCache();
  const clave = 'ics:' + Utilities.base64Encode(url).slice(0, 200);
  const enCache = cache.get(clave);
  if (enCache) return enCache;
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  if (resp.getResponseCode() !== 200) throw new Error('el enlace respondió ' + resp.getResponseCode());
  const texto = resp.getContentText();
  try { cache.put(clave, texto, 300); } catch (e) { /* demasiado grande para la caché */ }
  return texto;
}

/** Lector simple de iCalendar (RFC 5545): VEVENT con UID, SUMMARY, LOCATION, DTSTART, DTEND, RRULE. */
function parsearIcs_(texto) {
  const lineas = String(texto).replace(/\r\n[ \t]/g, '').replace(/\n[ \t]/g, '').split(/\r?\n/);
  const eventos = [];
  let ev = null;
  let anidado = 0;
  lineas.forEach(function (linea) {
    if (linea === 'BEGIN:VEVENT') { ev = { uid: '', titulo: '', lugar: '', exdates: {} }; anidado = 0; return; }
    if (!ev) return;
    if (linea === 'END:VEVENT') { eventos.push(ev); ev = null; return; }
    if (/^BEGIN:/.test(linea)) { anidado++; return; } // VALARM, etc.
    if (/^END:/.test(linea)) { anidado--; return; }
    if (anidado > 0) return;
    const m = /^([A-Z-]+)((?:;[^:]*)?):(.*)$/.exec(linea);
    if (!m) return;
    const prop = m[1];
    const params = m[2];
    const valor = m[3];
    const tzid = (/;TZID=("?)([^;:"]+)\1/.exec(params) || [])[2];
    if (prop === 'UID') ev.uid = valor;
    else if (prop === 'SUMMARY') ev.titulo = textoIcs_(valor);
    else if (prop === 'LOCATION') ev.lugar = textoIcs_(valor);
    else if (prop === 'STATUS') ev.cancelado = valor.toUpperCase() === 'CANCELLED';
    else if (prop === 'RRULE') ev.rrule = valor;
    else if (prop === 'RECURRENCE-ID') ev.recurrenceId = fechaIcs_(valor, tzid);
    else if (prop === 'DTSTART') {
      ev.todoElDia = /VALUE=DATE(?!-)/.test(params) || /^\d{8}$/.test(valor);
      ev.inicio = fechaIcs_(valor, tzid);
    } else if (prop === 'DTEND') ev.fin = fechaIcs_(valor, tzid);
    else if (prop === 'EXDATE') {
      valor.split(',').forEach(function (v) {
        const d = fechaIcs_(v, tzid);
        if (d) ev.exdates[d.getTime()] = true;
      });
    }
  });
  // Una instancia modificada de una serie (RECURRENCE-ID) reemplaza a la ocurrencia original.
  const series = {};
  eventos.forEach(function (e) {
    e.uid = e.uid || ('sin-uid-' + e.titulo + (e.inicio ? e.inicio.getTime() : ''));
    if (e.rrule && !e.recurrenceId) series[e.uid] = e;
  });
  eventos.forEach(function (e) {
    if (!e.recurrenceId) return;
    if (series[e.uid]) series[e.uid].exdates[e.recurrenceId.getTime()] = true;
    e.uid = e.uid + '#' + e.recurrenceId.getTime();
    e.rrule = '';
  });
  return eventos;
}

function textoIcs_(v) {
  return String(v).replace(/\\n/gi, ' ').replace(/\\([,;\\])/g, '$1').trim();
}

/** 20261013T080000Z (UTC) · 20261013T080000 con TZID o flotante · 20261013 (día). */
function fechaIcs_(v, tzid) {
  const m = /^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?(Z)?)?$/.exec(String(v).trim());
  if (!m) return null;
  if (!m[4]) return new Date(+m[1], +m[2] - 1, +m[3]);
  if (m[7]) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0)));
  const texto = m[1] + '-' + m[2] + '-' + m[3] + ' ' + m[4] + ':' + m[5] + ':' + (m[6] || '00');
  try {
    return Utilities.parseDate(texto, tzid || TZ, 'yyyy-MM-dd HH:mm:ss');
  } catch (e) {
    return Utilities.parseDate(texto, TZ, 'yyyy-MM-dd HH:mm:ss');
  }
}

/**
 * Inicios del evento dentro de [ini, fin). Soporta repeticiones DAILY / WEEKLY / MONTHLY
 * con INTERVAL, COUNT, UNTIL, BYDAY (semanal) y EXDATE, que es lo que usan las agendas normales.
 */
function ocurrencias_(ev, ini, fin) {
  const dentro = function (d) { return d >= ini && d < fin && !ev.exdates[d.getTime()]; };
  if (!ev.rrule) return dentro(ev.inicio) ? [ev.inicio] : [];

  const r = {};
  ev.rrule.split(';').forEach(function (p) { const kv = p.split('='); r[kv[0].toUpperCase()] = kv[1]; });
  const freq = r.FREQ;
  const intervalo = Math.max(1, +(r.INTERVAL || 1));
  const limite = r.COUNT ? +r.COUNT : Infinity;
  const hasta = r.UNTIL ? fechaIcs_(r.UNTIL) : null;
  const dias = { SU: 0, MO: 1, TU: 2, WE: 3, TH: 4, FR: 5, SA: 6 };
  const porDia = freq === 'WEEKLY' && r.BYDAY
    ? r.BYDAY.split(',').map(function (d) { return dias[d.replace(/^[+-]?\d+/, '')]; })
    : null;

  const s = ev.inicio;
  const conHora = function (y, mo, d) {
    return new Date(y, mo, d, s.getHours(), s.getMinutes(), s.getSeconds());
  };
  const out = [];
  let n = 0;
  for (let paso = 0; paso < 3000; paso++) {
    let candidatos;
    if (freq === 'DAILY') candidatos = [conHora(s.getFullYear(), s.getMonth(), s.getDate() + paso * intervalo)];
    else if (freq === 'WEEKLY') {
      const lunes = s.getDate() - ((s.getDay() + 6) % 7) + paso * 7 * intervalo;
      candidatos = (porDia || [s.getDay()]).map(function (dw) {
        return conHora(s.getFullYear(), s.getMonth(), lunes + ((dw + 6) % 7));
      }).sort(function (a, b) { return a - b; });
    } else if (freq === 'MONTHLY') {
      const c = conHora(s.getFullYear(), s.getMonth() + paso * intervalo, s.getDate());
      candidatos = c.getDate() === s.getDate() ? [c] : [];
    } else return dentro(s) ? [s] : [];

    for (let i = 0; i < candidatos.length; i++) {
      const c = candidatos[i];
      if (c < s) continue;
      if ((hasta && c > hasta) || n >= limite || c >= fin) return out;
      n++;
      if (dentro(c)) out.push(c);
    }
  }
  return out;
}

/** Intenta deducir paciente(s), tipo y ubicación del título y lugar de un evento. */
function sugerirDesdeEvento_(titulo, lugar, cats) {
  titulo = String(titulo || '');
  const t = normTxt_(titulo);
  const l = normTxt_(lugar || '');
  const tiene = function (re) { return re.test(t); };
  const buscar = function (fn) {
    const x = cats.tipos.filter(function (tp) { return fn(normTxt_(tp.nombre), tp); })[0];
    return x ? x.nombre : '';
  };

  const pareja = tiene(/\bpareja\b/);
  const control = tiene(/control|seguimiento/);
  const primera = tiene(/primera|inicial|\bnuev[oa]\b/);
  let tipo = '';
  if (tiene(/medici/)) tipo = buscar(function (n) { return n.indexOf('medic') >= 0; });
  else if (pareja && control) tipo = buscar(function (n, tp) { return tp.pareja && n.indexOf('control') >= 0; });
  else if (pareja && primera) tipo = buscar(function (n, tp) { return tp.pareja && n.indexOf('primera') >= 0; });
  else if (!pareja && control) tipo = buscar(function (n, tp) { return !tp.pareja && n.indexOf('control') >= 0; });
  else if (!pareja && primera) tipo = buscar(function (n, tp) { return !tp.pareja && n.indexOf('primera') >= 0; });

  let ubicacion = '';
  cats.ubicaciones.forEach(function (u) {
    const n = normTxt_(u.nombre);
    if (!ubicacion && n && (l.indexOf(n) >= 0 || t.indexOf(n) >= 0)) ubicacion = u.nombre;
  });

  // Nombre: quitar palabras clave en cualquier lugar y conectores sueltos en los extremos.
  const clave = {};
  ['cita', 'consulta', 'nutricion', 'nutri', 'control', 'seguimiento', 'primera', 'vez', 'pareja',
    'mediciones', 'medicion', 'inicial', 'nuevo', 'nueva', 'paciente', 'sesion', 'virtual', 'presencial']
    .forEach(function (k) { clave[k] = true; });
  cats.ubicaciones.forEach(function (u) { normTxt_(u.nombre).split(' ').forEach(function (k) { if (k) clave[k] = true; }); });
  const conectores = { de: 1, del: 1, en: 1, con: 1, la: 1, el: 1, los: 1, las: 1, y: 1, a: 1 };
  let tokens = titulo.replace(/[-–—|:/()[\]·•]+/g, ' ').split(/\s+/)
    .filter(function (tk) { return tk && !clave[normTxt_(tk).replace(/[.,]/g, '')]; });
  while (tokens.length && conectores[normTxt_(tokens[0])]) tokens.shift();
  while (tokens.length && conectores[normTxt_(tokens[tokens.length - 1])]) tokens.pop();
  const nombre = tokens.join(' ').replace(/\s*,\s*$/, '').trim();

  let paciente = nombre;
  let paciente2 = '';
  if (pareja) {
    const partes = nombre.split(/\s+(?:y|e)\s+|\s*[&+,]\s*/i).filter(Boolean);
    if (partes.length === 2) { paciente = partes[0]; paciente2 = partes[1]; }
  }
  return { paciente: paciente, paciente2: paciente2, tipo: tipo, ubicacion: ubicacion };
}

/* ───────────────────────── Reportes ───────────────────────── */

function resumen_(regs, cats) {
  const tot = { consultas: 0, personas: 0, pagoPacientes: 0, aseguradoras: 0, facturado: 0, descuentos: 0, comisiones: 0, porCobrar: 0 };
  const porUbic = {};
  const porAseg = {};
  const porTipo = {};
  const porForma = {};
  const porEnc = {};
  const cfgUbic = configUbicaciones_(cats);
  const obtener = function (mapa, k, init) { if (!mapa[k]) mapa[k] = init(); return mapa[k]; };

  regs.forEach(function (r) {
    const personas = r.paciente2 ? 2 : 1;
    const comision = comisionDe_(r, cfgUbic);
    tot.consultas++;
    tot.personas += personas;
    tot.pagoPacientes += r.totalPaciente;
    tot.aseguradoras += r.montoAseguradora;
    tot.facturado += r.totalFacturado;
    tot.descuentos += r.montoDescuento || 0;
    tot.comisiones += comision;

    const nombreU = r.ubicacion || '(sin ubicación)';
    const u = obtener(porUbic, nombreU, function () {
      return { nombre: nombreU, consultas: 0, personas: 0, pagoPacientes: 0, aseguradoras: 0, facturado: 0, montoBase: 0, comision: 0 };
    });
    u.consultas++;
    u.personas += personas;
    u.pagoPacientes += r.totalPaciente;
    u.aseguradoras += r.montoAseguradora;
    u.facturado += r.totalFacturado;
    u.montoBase += baseComision_(r, cfgUbic);
    u.comision += comision;

    const enc = (cfgUbic[r.ubicacion] || {}).encargado;
    if (enc) {
      const e = obtener(porEnc, enc, function () { return { nombre: enc, ubicaciones: [], consultas: 0, facturado: 0, comision: 0 }; });
      if (e.ubicaciones.indexOf(r.ubicacion) < 0) e.ubicaciones.push(r.ubicacion);
      e.consultas++;
      e.facturado += r.totalFacturado;
      e.comision += comision;
    }

    if (r.conSeguro) {
      const a = obtener(porAseg, r.aseguradora || '(sin aseguradora)', function () {
        return { nombre: r.aseguradora || '(sin aseguradora)', consultas: 0, monto: 0, Pendiente: 0, Facturado: 0, Pagado: 0 };
      });
      const estado = ESTADOS_ASEGURADORA.indexOf(r.estadoAseguradora) >= 0 ? r.estadoAseguradora : 'Pendiente';
      a.consultas++;
      a.monto += r.montoAseguradora;
      a[estado] += r.montoAseguradora;
      if (estado !== 'Pagado') tot.porCobrar += r.montoAseguradora;
    }

    const tp = obtener(porTipo, r.tipo || '(sin tipo)', function () {
      return { nombre: r.tipo || '(sin tipo)', consultas: 0, facturado: 0 };
    });
    tp.consultas++;
    tp.facturado += r.totalFacturado;

    [[r.pago1, r.forma1], [r.pago2, r.forma2]].forEach(function (p) {
      if (!p[0]) return;
      const f = obtener(porForma, p[1] || '(sin forma)', function () { return { nombre: p[1] || '(sin forma)', pagos: 0, monto: 0 }; });
      f.pagos++;
      f.monto += p[0];
    });
  });

  const ubicaciones = ordenar_(porUbic, cats.ubicaciones.map(function (u) { return u.nombre; }))
    .map(function (u) {
      const cfg = cfgUbic[u.nombre] || {};
      u.encargado = cfg.encargado || '';
      u.porcentaje = cfg.porcentaje || 0;
      u.base = cfg.base || BASES_COMISION[0];
      u.nutricionista = u.facturado - u.comision;
      return u;
    });
  tot.neto = tot.facturado - tot.comisiones;

  return {
    totales: tot,
    ubicaciones: ubicaciones,
    encargados: ordenar_(porEnc, []),
    aseguradoras: ordenar_(porAseg, cats.aseguradoras.map(function (a) { return a.nombre; })),
    tipos: ordenar_(porTipo, cats.tipos.map(function (t) { return t.nombre; })),
    formas: ordenar_(porForma, cats.formas),
  };
}

function configUbicaciones_(cats) {
  const m = {};
  cats.ubicaciones.forEach(function (u) { m[u.nombre] = u; });
  return m;
}

function baseComision_(r, cfgUbic) {
  const cfg = cfgUbic[r.ubicacion];
  if (!cfg) return 0;
  return cfg.base === 'Pago paciente' ? r.totalPaciente : r.totalFacturado;
}

/** Lo que se le paga al encargado de la ubicación por esta consulta. */
function comisionDe_(r, cfgUbic) {
  const cfg = cfgUbic[r.ubicacion];
  return cfg ? Math.round(baseComision_(r, cfgUbic) * cfg.porcentaje) : 0;
}

/** Valores del mapa en el orden del catálogo; lo que no esté en el catálogo va al final. */
function ordenar_(mapa, orden) {
  const out = [];
  orden.forEach(function (k) { if (mapa[k]) { out.push(mapa[k]); delete mapa[k]; } });
  Object.keys(mapa).sort().forEach(function (k) { out.push(mapa[k]); });
  return out;
}

function nombreNutricionista_() {
  return config_()['Nutricionista'] || 'la nutricionista';
}

function reporteSede_(w, regs, cats, opc) {
  const u = cats.ubicaciones.filter(function (x) { return x.nombre === opc.filtro; })[0] ||
    { nombre: opc.filtro, encargado: '', porcentaje: 0, base: BASES_COMISION[0] };
  const filas = regs.filter(function (r) { return r.ubicacion === u.nombre; });
  const res = resumen_(filas, cats);
  const ub = res.ubicaciones[0] || { facturado: 0, montoBase: 0, comision: 0, consultas: 0, nutricionista: 0 };

  w.subtitulo('Detalle de consultas');
  const conNombres = opc.nombres !== false;
  const headers = ['Fecha', 'Hora'].concat(conNombres ? ['Paciente(s)'] : [])
    .concat(['Tipo de cita', 'Pago paciente', 'Aseguradora', 'Total facturado']);
  const formatos = [FECHA, '@'].concat(conNombres ? ['@'] : []).concat(['@', MONEDA, MONEDA, MONEDA]);
  const rows = filas.map(function (r) {
    return [aFecha_(r.fecha), r.hora].concat(conNombres ? [nombres_(r)] : [])
      .concat([r.tipo + (r.descuento ? ' (' + r.descuento + ')' : ''), r.totalPaciente, r.montoAseguradora, r.totalFacturado]);
  });
  const total = ['Total', ''].concat(conNombres ? [''] : [])
    .concat([filas.length + ' consultas', res.totales.pagoPacientes, res.totales.aseguradoras, res.totales.facturado]);
  w.tabla(headers, rows, formatos, total);

  w.subtitulo('Liquidación');
  w.pares([
    ['Consultas atendidas', ub.consultas, '0'],
    ['Total facturado', ub.facturado || 0, MONEDA],
    ['Base de cálculo', u.base, '@'],
    ['Porcentaje acordado', u.porcentaje, '0.##%'],
    ['Monto para ' + nombreNutricionista_(), ub.nutricionista || 0, MONEDA],
  ]);
  w.destacado('Monto a pagar' + (u.encargado ? ' a ' + u.encargado : ''), ub.comision || 0, MONEDA);
}

/** Liquidación de una persona que cobra en una o varias ubicaciones (p. ej. Andrea: Escazú + Domicilio). */
function reporteEncargado_(w, regs, cats, opc) {
  const cfgUbic = configUbicaciones_(cats);
  const ubics = cats.ubicaciones.filter(function (u) { return u.encargado === opc.filtro; });
  const nombresU = ubics.map(function (u) { return u.nombre; });
  const filas = regs.filter(function (r) { return nombresU.indexOf(r.ubicacion) >= 0; });
  const conNombres = opc.nombres !== false;

  w.subtitulo('Detalle de consultas');
  const headers = ['Fecha', 'Ubicación'].concat(conNombres ? ['Paciente(s)'] : [])
    .concat(['Tipo de cita', 'Total facturado', '%', 'A pagar']);
  const formatos = [FECHA, '@'].concat(conNombres ? ['@'] : []).concat(['@', MONEDA, '0.##%', MONEDA]);
  let facturado = 0;
  let pagar = 0;
  const rows = filas.map(function (r) {
    const c = comisionDe_(r, cfgUbic);
    facturado += r.totalFacturado;
    pagar += c;
    return [aFecha_(r.fecha), r.ubicacion].concat(conNombres ? [nombres_(r)] : [])
      .concat([r.tipo + (r.descuento ? ' (' + r.descuento + ')' : ''), r.totalFacturado, cfgUbic[r.ubicacion].porcentaje, c]);
  });
  const total = ['Total', ''].concat(conNombres ? [''] : []).concat([filas.length + ' consultas', facturado, '', pagar]);
  w.tabla(headers, rows, formatos, total);

  w.subtitulo('Resumen por ubicación');
  const res = resumen_(filas, cats);
  w.pares(res.ubicaciones.map(function (u) {
    return [u.nombre + ' · ' + u.consultas + ' consultas · ' + Math.round(u.porcentaje * 100) + '% de ' + formatoColones_(u.facturado),
      u.comision, MONEDA];
  }));
  w.destacado('Total a pagar a ' + opc.filtro, pagar, MONEDA);
}

function reporteAseguradora_(w, regs, opc) {
  const filas = regs.filter(function (r) { return r.conSeguro && r.aseguradora === opc.filtro; });
  const conNombres = opc.nombres !== false;
  w.subtitulo('Consultas con seguro');
  const headers = ['Fecha'].concat(conNombres ? ['Paciente(s)'] : [])
    .concat(['Tipo de cita', 'Ubicación', 'Monto aseguradora', 'Estado']);
  const formatos = [FECHA].concat(conNombres ? ['@'] : []).concat(['@', '@', MONEDA, '@']);
  const rows = filas.map(function (r) {
    return [aFecha_(r.fecha)].concat(conNombres ? [nombres_(r)] : [])
      .concat([r.tipo, r.ubicacion, r.montoAseguradora, r.estadoAseguradora]);
  });
  const suma = function (estado) {
    return filas.reduce(function (s, r) { return s + (!estado || r.estadoAseguradora === estado ? r.montoAseguradora : 0); }, 0);
  };
  const total = ['Total'].concat(conNombres ? [''] : []).concat([filas.length + ' consultas', '', suma(), '']);
  w.tabla(headers, rows, formatos, total);

  w.subtitulo('Estado de cobro');
  w.pares(ESTADOS_ASEGURADORA.map(function (e) { return [e, suma(e), MONEDA]; }));
  w.destacado('Total a cobrar a ' + opc.filtro, suma(), MONEDA);
}

function reporteGeneral_(w, regs, cats) {
  const res = resumen_(regs, cats);
  const t = res.totales;
  const nutri = nombreNutricionista_();
  w.subtitulo('Resumen del mes');
  w.pares([
    ['Consultas', t.consultas, '0'],
    ['Personas atendidas', t.personas, '0'],
    ['Pagado por pacientes', t.pagoPacientes, MONEDA],
    ['Monto aseguradoras', t.aseguradoras, MONEDA],
    ['Total facturado', t.facturado, MONEDA],
    ['Descuentos aplicados', t.descuentos, MONEDA],
    ['Pagos a colaboradoras', t.comisiones, MONEDA],
    ['Por cobrar a aseguradoras', t.porCobrar, MONEDA],
  ]);
  w.destacado('Monto para ' + nutri, t.neto, MONEDA);

  w.subtitulo('Por ubicación');
  w.tabla(['Ubicación', 'Consultas', 'Pago pacientes', 'Aseguradoras', 'Total facturado', '% sede', 'Sede', nutri],
    res.ubicaciones.map(function (u) {
      return [u.nombre, u.consultas, u.pagoPacientes, u.aseguradoras, u.facturado, u.porcentaje, u.comision, u.nutricionista];
    }),
    ['@', '0', MONEDA, MONEDA, MONEDA, '0.##%', MONEDA, MONEDA],
    ['Total', t.consultas, t.pagoPacientes, t.aseguradoras, t.facturado, '', t.comisiones, t.neto]);

  w.subtitulo('Pagos a colaboradoras');
  w.tabla(['Persona', 'Ubicaciones', 'Consultas', 'Total facturado', 'A pagar'],
    res.encargados.map(function (e) { return [e.nombre, e.ubicaciones.join(', '), e.consultas, e.facturado, e.comision]; }),
    ['@', '@', '0', MONEDA, MONEDA],
    ['Total', '', '', '', t.comisiones]);

  w.subtitulo('Por aseguradora');
  w.tabla(['Aseguradora', 'Consultas', 'Monto', 'Pendiente', 'Facturado', 'Pagado'],
    res.aseguradoras.map(function (a) { return [a.nombre, a.consultas, a.monto, a.Pendiente, a.Facturado, a.Pagado]; }),
    ['@', '0', MONEDA, MONEDA, MONEDA, MONEDA], null);

  w.subtitulo('Por tipo de cita');
  w.tabla(['Tipo de cita', 'Consultas', 'Total facturado'],
    res.tipos.map(function (x) { return [x.nombre, x.consultas, x.facturado]; }),
    ['@', '0', MONEDA], null);

  w.subtitulo('Pagos de pacientes por forma de pago');
  w.tabla(['Forma de pago', 'Pagos', 'Monto'],
    res.formas.map(function (x) { return [x.nombre, x.pagos, x.monto]; }),
    ['@', '0', MONEDA], null);
}

function formatoColones_(n) {
  return '₡' + String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
}

/** Pequeño ayudante para escribir reportes con el estilo de la marca. */
function escritor_(sh) {
  let fila = 1;
  let maxCols = 2;
  const fuente = 'Montserrat';
  return {
    encabezado: function (titulo, sub) {
      sh.getRange(fila, 1).setValue('Hábitos con propósito')
        .setFontFamily('DM Serif Display').setFontSize(20).setFontColor(COLOR.gris);
      fila++;
      sh.getRange(fila, 1).setValue(titulo)
        .setFontFamily('Poppins').setFontWeight('bold').setFontSize(14).setFontColor(COLOR.oscuro);
      fila++;
      sh.getRange(fila, 1).setValue(sub).setFontFamily(fuente).setFontSize(9).setFontColor(COLOR.gris);
      fila += 2;
    },
    subtitulo: function (texto) {
      sh.getRange(fila, 1).setValue(texto)
        .setFontFamily('Poppins').setFontWeight('bold').setFontSize(11).setFontColor(COLOR.oscuro);
      fila++;
    },
    tabla: function (headers, rows, formatos, total) {
      const nc = headers.length;
      maxCols = Math.max(maxCols, nc);
      const inicio = fila;
      sh.getRange(fila, 1, 1, nc).setValues([headers])
        .setBackground(COLOR.salvia).setFontColor(COLOR.oscuro).setFontWeight('bold').setFontFamily(fuente);
      fila++;
      if (rows.length) {
        sh.getRange(fila, 1, rows.length, nc)
          .setNumberFormats(rows.map(function () { return formatos; }))
          .setValues(rows).setFontFamily(fuente);
        fila += rows.length;
      } else {
        sh.getRange(fila, 1).setValue('Sin registros en este periodo')
          .setFontStyle('italic').setFontColor(COLOR.gris).setFontFamily(fuente);
        fila++;
      }
      if (total) {
        sh.getRange(fila, 1, 1, nc).setNumberFormats([formatos.map(function (f) { return f === FECHA ? '@' : f; })])
          .setValues([total]).setFontWeight('bold').setBackground(COLOR.verdeClaro).setFontFamily(fuente);
        fila++;
      }
      sh.getRange(inicio, 1, fila - inicio, nc)
        .setBorder(true, true, true, true, false, true, COLOR.linea, SpreadsheetApp.BorderStyle.SOLID);
      fila++;
    },
    pares: function (pares) {
      pares.forEach(function (p) {
        sh.getRange(fila, 1).setValue(p[0]).setFontFamily(fuente).setFontColor(COLOR.oscuro);
        sh.getRange(fila, 2).setNumberFormat(p[2] || '@').setValue(p[1]).setFontFamily(fuente).setHorizontalAlignment('right');
        fila++;
      });
    },
    destacado: function (etiqueta, valor, formato) {
      sh.getRange(fila, 1, 1, 2).setBackground(COLOR.amarilloSuave).setFontWeight('bold')
        .setFontFamily('Poppins').setFontColor(COLOR.oscuro);
      sh.getRange(fila, 1).setValue(etiqueta);
      sh.getRange(fila, 2).setNumberFormat(formato).setValue(valor).setHorizontalAlignment('right');
      fila += 2;
    },
    cerrar: function () {
      sh.setHiddenGridlines(true);
      sh.setColumnWidth(1, 230);
      for (let c = 2; c <= maxCols; c++) sh.setColumnWidth(c, 125);
    },
  };
}

function nombres_(r) {
  return r.paciente2 ? r.paciente + ' y ' + r.paciente2 : r.paciente;
}

function moverACarpeta_(id) {
  try {
    const nombre = config_()['Carpeta de reportes'] || 'Hábitos con propósito · Reportes';
    const it = DriveApp.getFoldersByName(nombre);
    const carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(nombre);
    DriveApp.getFileById(id).moveTo(carpeta);
  } catch (e) {
    console.warn('No se pudo mover el reporte a la carpeta: ' + e.message);
  }
}

function urlExport_(id, gid, formato) {
  const base = 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=' + formato + '&gid=' + gid;
  if (formato !== 'pdf') return base;
  return base + '&size=letter&portrait=true&fitw=true&gridlines=false&printtitle=false' +
    '&sheetnames=false&pagenum=UNDEFINED&top_margin=0.6&bottom_margin=0.6&left_margin=0.6&right_margin=0.6';
}

/* ───────────────────────── Menú ───────────────────────── */

function abrirApp() {
  const ui = SpreadsheetApp.getUi();
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    ui.alert('La app todavía no está publicada.\n\nEn el editor de Apps Script: Implementar → Nueva implementación → Aplicación web.');
    return;
  }
  const html = HtmlService.createHtmlOutput(
    '<p style="font-family:Montserrat,Arial,sans-serif">' +
    '<a href="' + url + '" target="_blank" style="color:#304323;font-weight:bold">Abrir la bitácora →</a></p>')
    .setWidth(320).setHeight(90);
  ui.showModalDialog(html, 'Hábitos con propósito');
}

function limpiarRegistros() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('Borrar registros',
    'Se van a borrar TODAS las filas de la hoja Registros (los catálogos y tarifas se mantienen). ¿Continuar?',
    ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  const sh = hoja_(HOJAS.registros);
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
}

/* ───────────────────────── Utilidades ───────────────────────── */

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function hoja_(nombre) {
  const sh = ss_().getSheetByName(nombre);
  if (!sh) throw new Error('No existe la hoja "' + nombre + '". Corra Hábitos → Configurar / reparar hojas.');
  return sh;
}

function esFecha_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}

function aFecha_(v) {
  if (esFecha_(v)) return v;
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(v || '').trim());
  if (!m) return null;
  const d = new Date(+m[1], +m[2] - 1, +m[3]);
  return d.getMonth() === +m[2] - 1 ? d : null;
}

function fechaStr_(v) {
  if (esFecha_(v)) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  let m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
  if (m) return m[3] + '-' + pad2_(m[2]) + '-' + pad2_(m[1]);
  return '';
}

function horaStr_(v) {
  if (esFecha_(v)) return Utilities.formatDate(v, TZ, 'HH:mm');
  const m = /^(\d{1,2}):(\d{2})/.exec(String(v || '').trim());
  return m ? pad2_(m[1]) + ':' + m[2] : '';
}

function fechaHoraStr_(v) {
  return esFecha_(v) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss') : String(v || '');
}

function validarMes_(mes) {
  if (!/^\d{4}-\d{2}$/.test(String(mes || ''))) throw new Error('Mes inválido.');
}

function etiquetaMes_(mes) {
  return MESES[+mes.slice(5, 7) - 1] + ' ' + mes.slice(0, 4);
}

function pad2_(n) {
  return ('0' + String(n).trim()).slice(-2);
}

/** Monto de entrada: entero ≥ 0, vacío = 0, inválido = NaN. */
function monto_(v) {
  if (v === '' || v == null) return 0;
  const n = typeof v === 'number' ? v : Number(String(v).replace(/[₡\s,]/g, ''));
  if (!isFinite(n) || n < 0) return NaN;
  return Math.round(n);
}

/** Número leído de una celda (tolerante). */
function num_(v) {
  if (typeof v === 'number') return v;
  const n = Number(String(v == null ? '' : v).replace(/[₡\s,]/g, ''));
  return isFinite(n) ? n : 0;
}

/** Porcentaje: acepta 0.2, 20 o "20%". */
function pct_(v) {
  if (v === '' || v == null) return 0;
  const n = typeof v === 'number' ? v : Number(String(v).replace(/[%\s]/g, '').replace(',', '.'));
  if (!isFinite(n)) return 0;
  return typeof v === 'string' && v.indexOf('%') >= 0 ? n / 100 : (n > 1 ? n / 100 : n);
}

function esSi_(v) {
  if (v === true) return true;
  return ['si', 's', 'true', 'x', 'yes'].indexOf(normTxt_(v)) >= 0;
}

function esNo_(v) {
  if (v === false) return true;
  return ['no', 'n', 'false'].indexOf(normTxt_(v)) >= 0;
}

function normTxt_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}
