/*
 * Pruebas de Code.gs con servicios de Google simulados.
 * Uso: npm test   (o: TZ=America/Costa_Rica node test/test.js)
 */
'use strict';
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { crearEntornoGAS } = require('../dev/gas-mock.js');

// Los objetos creados dentro del contexto vm tienen otro prototipo: se comparan como JSON.
const igual = (a, b, msg) => assert.deepStrictEqual(JSON.parse(JSON.stringify(a)), b, msg);

const CODE = fs.readFileSync(path.join(__dirname, '..', 'src', 'Code.gs'), 'utf8');

function nuevoEntorno() {
  const env = crearEntornoGAS();
  const ctx = vm.createContext(Object.assign({ console }, env));
  vm.runInContext(CODE, ctx, { filename: 'Code.gs' });
  ctx.setup();
  return ctx;
}

function hoja(ctx, nombre) {
  return ctx._libro.getSheetByName(nombre);
}

/** Columna (1-based) de un encabezado en la hoja Registros. */
function col(ctx, header) {
  const sh = hoja(ctx, 'Registros');
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf(header) + 1;
}

function sinFechas(o, ruta = 'resultado') {
  if (Object.prototype.toString.call(o) === '[object Date]') throw new Error('Date en ' + ruta + ' (google.script.run no lo permite)');
  if (o && typeof o === 'object') Object.keys(o).forEach((k) => sinFechas(o[k], ruta + '.' + k));
}

const SIN_SEGURO = ['Pareja primera vez', 'Pareja control', 'Mediciones'];

const pruebas = [];
const prueba = (nombre, fn) => pruebas.push({ nombre, fn });

prueba('setup crea las hojas con tarifas, comisiones, seguro y descuento iniciales', () => {
  const ctx = nuevoEntorno();
  ctx.setup();
  const nombres = ctx._libro.getSheets().map((h) => h.getName());
  igual(nombres, ['Registros', 'Tarifas', 'Ubicaciones', 'Aseguradoras', 'Descuentos', 'Tipos de cita', 'Formas de pago', 'Config']);
  const t = hoja(ctx, 'Tarifas');
  igual(t.getRange(1, 1, 6, 4).getValues(), [
    ['Tipo de cita', 'Escazú', 'Curridabat', 'Domicilio'],
    ['Primera vez', 50000, 40000, 35000],
    ['Control', 35000, 30000, 30000],
    ['Pareja primera vez', 47000, 35000, 30000],
    ['Pareja control', 32000, 27000, 25000],
    ['Mediciones', 14000, 'No aplica', 12000],
  ]);
  const reg = hoja(ctx, 'Registros');
  assert.strictEqual(reg.getRange(1, 1, 1, reg.getLastColumn()).getValues()[0].length, 24);
  assert.strictEqual(hoja(ctx, 'Config').getLastRow(), 6, 'Config no debe duplicar claves');
  igual(hoja(ctx, 'Formas de pago').getRange(2, 1, 2, 1).getValues(), [['SINPE'], ['Tarjeta']]);
  igual(hoja(ctx, 'Ubicaciones').getRange(2, 1, 3, 3).getValues(), [
    ['Escazú', 'Andrea R.', 0.4],
    ['Curridabat', 'Raquel D.', 0.4],
    ['Domicilio', 'Andrea R.', 0.2],
  ]);
});

prueba('setup sobre una hoja vieja completa vacíos sin pisar datos', () => {
  const env = crearEntornoGAS();
  const ctx = vm.createContext(Object.assign({ console }, env));
  vm.runInContext(CODE, ctx, { filename: 'Code.gs' });
  const libro = ctx._libro;
  const a = libro.insertSheet('Aseguradoras');
  a.getRange(1, 1, 2, 3).setValues([['Aseguradora', 'Monto por consulta', 'Activa'], ['Sana Sana', '', 'Sí']]);
  const u = libro.insertSheet('Ubicaciones');
  u.getRange(1, 1, 3, 5).setValues([
    ['Ubicación', 'Encargado / a quién se paga', '% comisión', 'Base de comisión', 'Activa'],
    ['Escazú', 'Andrea', '', 'Total facturado', 'Sí'],
    ['Curridabat', 'Encargado Curridabat', 0.35, 'Total facturado', 'Sí'],
  ]);
  ctx.setup();
  igual(a.getRange(1, 1, 2, 6).getValues(), [
    ['Aseguradora', 'Monto aseguradora', 'Activa', 'Monto paciente', 'Ubicaciones', 'Excepto tipos'],
    ['Sana Sana', 8500, 'Sí', 16000, 'Escazú', 'Pareja primera vez, Pareja control, Mediciones'],
  ]);
  igual(u.getRange(2, 1, 2, 3).getValues(), [['Escazú', 'Andrea', 0.4], ['Curridabat', 'Encargado Curridabat', 0.35]]);
  const b = ctx.getBootstrap();
  igual(b.cats.aseguradoras, [{ nombre: 'Sana Sana', monto: 8500, montoPaciente: 16000, ubicaciones: ['Escazú'], exceptoTipos: SIN_SEGURO }]);
});

prueba('sincronizarTarifas agrega ubicaciones nuevas sin perder montos ni recargar', () => {
  const ctx = nuevoEntorno();
  const t = hoja(ctx, 'Tarifas');
  t.getRange(3, 2).setValue(36000);
  hoja(ctx, 'Ubicaciones').getRange(5, 1, 1, 5).setValues([['Virtual', '', 0, 'Total facturado', 'Sí']]);
  ctx.sincronizarTarifas();
  const tar = ctx.tarifas_();
  assert.strictEqual(tar['Control']['Escazú'], 36000);
  assert.strictEqual(tar['Control']['Virtual'], null);
  assert.strictEqual(tar['Mediciones']['Curridabat'], false);
});

prueba('bootstrap entrega catálogos y nada con Date', () => {
  const ctx = nuevoEntorno();
  const b = ctx.getBootstrap();
  sinFechas(b);
  igual(b.cats.formas, ['SINPE', 'Tarjeta']);
  igual(b.cats.ubicaciones.map((u) => [u.nombre, u.porcentaje]), [['Escazú', 0.4], ['Curridabat', 0.4], ['Domicilio', 0.2]]);
  igual(b.cats.aseguradoras, [{ nombre: 'Sana Sana', monto: 8500, montoPaciente: 16000, ubicaciones: ['Escazú'], exceptoTipos: SIN_SEGURO }]);
  igual(b.cats.descuentos, [{ nombre: 'VOLK', porcentaje: 0.1, ubicaciones: ['Domicilio'], exceptoTipos: ['Mediciones'] }]);
  assert.strictEqual(b.tarifas['Primera vez']['Curridabat'], 40000);
  assert.strictEqual(b.tarifas['Mediciones']['Curridabat'], false);
  assert.strictEqual(b.cats.tipos.filter((t) => t.pareja).length, 2);
});

prueba('consulta simple: fecha real, hora texto, tarifa de lista', () => {
  const ctx = nuevoEntorno();
  const r = ctx.guardarRegistro({
    fecha: '2026-10-13', hora: '8:05', paciente: ' María Fernández ', tipo: 'Control', ubicacion: 'Escazú',
    conSeguro: false, pago1: 35000, forma1: 'SINPE',
  });
  sinFechas(r);
  assert.strictEqual(r.registro.paciente, 'María Fernández');
  assert.strictEqual(r.registro.hora, '08:05');
  assert.strictEqual(r.registro.totalFacturado, 35000);
  assert.strictEqual(r.registro.tarifa, 35000);
  const fila = hoja(ctx, 'Registros').getRange(2, 1, 1, 24).getValues()[0];
  assert.ok(Object.prototype.toString.call(fila[1]) === '[object Date]', 'Fecha debe guardarse como fecha');
  assert.strictEqual(fila[1].getDate(), 13);
  assert.strictEqual(fila[col(ctx, 'Con seguro') - 1], 'No');
  assert.strictEqual(r.recientes.length, 1);
  assert.strictEqual(r.pacientes[0].n, 'María Fernández');
});

prueba('pareja: tarifa por persona y dos pagos', () => {
  const ctx = nuevoEntorno();
  const r = ctx.guardarRegistro({
    fecha: '2026-10-14', paciente: 'Ana Solís', paciente2: 'Luis Mora', tipo: 'Pareja primera vez', ubicacion: 'Curridabat',
    dosPagos: true, pago1: 35000, forma1: 'SINPE', pago2: 35000, forma2: 'Tarjeta',
  }).registro;
  assert.strictEqual(r.tarifa, 70000);
  assert.strictEqual(r.totalPaciente, 70000);
  assert.strictEqual(r.paciente2, 'Luis Mora');
});

prueba('Sana Sana: 16 000 paciente + 8 500 aseguradora, solo en Escazú y no en pareja ni Mediciones', () => {
  const ctx = nuevoEntorno();
  const base = { fecha: '2026-10-14', paciente: 'Pedro', tipo: 'Control', conSeguro: true, aseguradora: 'Sana Sana',
    montoAseguradora: 8500, pago1: 16000, forma1: 'SINPE' };
  const r = ctx.guardarRegistro(Object.assign({ ubicacion: 'Escazú' }, base)).registro;
  assert.strictEqual(r.totalFacturado, 24500);
  assert.strictEqual(r.estadoAseguradora, 'Pendiente');
  assert.throws(() => ctx.guardarRegistro(Object.assign({ ubicacion: 'Curridabat' }, base)), /solo aplica en Escazú/);
  assert.throws(() => ctx.guardarRegistro(Object.assign({}, base, { ubicacion: 'Escazú', tipo: 'Mediciones' })),
    /Sana Sana no aplica a Mediciones/);
  assert.throws(() => ctx.guardarRegistro(Object.assign({}, base, { ubicacion: 'Escazú', tipo: 'Pareja control', paciente2: 'Y' })),
    /Sana Sana no aplica a Pareja control/);
  ctx.guardarRegistro(Object.assign({}, base, { ubicacion: 'Escazú', tipo: 'Primera vez' }));
});

prueba('Mediciones no se ofrece en Curridabat', () => {
  const ctx = nuevoEntorno();
  assert.throws(() => ctx.guardarRegistro({ fecha: '2026-10-14', paciente: 'X', tipo: 'Mediciones', ubicacion: 'Curridabat', pago1: 0 }),
    /Mediciones no se ofrece en Curridabat/);
});

prueba('descuento VOLK: 10% en Domicilio, no en Mediciones ni otras sedes', () => {
  const ctx = nuevoEntorno();
  const g = (d) => ctx.guardarRegistro(Object.assign({ fecha: '2026-10-14', paciente: 'X', forma1: 'SINPE', descuento: 'VOLK' }, d)).registro;
  const r = g({ tipo: 'Control', ubicacion: 'Domicilio', pago1: 27000 });
  assert.strictEqual(r.tarifa, 30000);
  assert.strictEqual(r.montoDescuento, 3000);
  assert.strictEqual(r.descuento, 'VOLK');
  const p = g({ tipo: 'Pareja control', ubicacion: 'Domicilio', paciente2: 'Y', pago1: 45000 });
  assert.strictEqual(p.tarifa, 50000);
  assert.strictEqual(p.montoDescuento, 5000);
  assert.throws(() => g({ tipo: 'Mediciones', ubicacion: 'Domicilio', pago1: 12000 }), /VOLK no aplica a Mediciones/);
  assert.throws(() => g({ tipo: 'Control', ubicacion: 'Escazú', pago1: 31500 }), /no aplica a Control en Escazú/);
});

prueba('descuento puntual "Otro" con monto digitado', () => {
  const ctx = nuevoEntorno();
  const base = { fecha: '2026-10-02', paciente: 'Mónica', tipo: 'Primera vez', ubicacion: 'Escazú', forma1: 'SINPE', descuento: 'Otro' };
  const r = ctx.guardarRegistro(Object.assign({ montoDescuento: 10000, pago1: 40000 }, base)).registro;
  igual([r.tarifa, r.descuento, r.montoDescuento, r.totalFacturado], [50000, 'Otro', 10000, 40000]);
  assert.throws(() => ctx.guardarRegistro(Object.assign({ pago1: 40000 }, base)), /monto del descuento/);
});

prueba('validaciones', () => {
  const ctx = nuevoEntorno();
  const base = { fecha: '2026-10-13', paciente: 'X', tipo: 'Control', ubicacion: 'Escazú', pago1: 1000, forma1: 'SINPE' };
  const falla = (cambio, texto) => assert.throws(() => ctx.guardarRegistro(Object.assign({}, base, cambio)), (e) => e.message.includes(texto));
  falla({ fecha: '13/10/2026' }, 'fecha');
  falla({ paciente: ' ' }, 'nombre del paciente');
  falla({ tipo: 'Pareja control' }, 'segunda persona');
  falla({ conSeguro: true }, 'aseguradora');
  falla({ forma1: '' }, 'forma de pago');
  falla({ pago1: -5 }, 'montos válidos');
  falla({ hora: '8am' }, 'hora');
  // Paciente que no paga (cubre todo la aseguradora) es válido sin forma de pago
  ctx.guardarRegistro(Object.assign({}, base, { pago1: 0, forma1: '', conSeguro: true, aseguradora: 'Sana Sana', montoAseguradora: 20000 }));
});

prueba('editar y eliminar conservan la fila correcta', () => {
  const ctx = nuevoEntorno();
  const a = ctx.guardarRegistro({ fecha: '2026-10-01', paciente: 'A', tipo: 'Control', ubicacion: 'Escazú', pago1: 35000, forma1: 'SINPE' }).registro;
  const b = ctx.guardarRegistro({ fecha: '2026-10-02', paciente: 'B', tipo: 'Control', ubicacion: 'Escazú', pago1: 35000, forma1: 'SINPE' }).registro;
  const e = ctx.guardarRegistro(Object.assign({}, a, { id: a.id, paciente: 'A editada', pago1: 30000 })).registro;
  assert.strictEqual(e.id, a.id);
  assert.strictEqual(e.creado, a.creado);
  assert.strictEqual(hoja(ctx, 'Registros').getLastRow(), 3);
  assert.strictEqual(hoja(ctx, 'Registros').getRange(2, col(ctx, 'Paciente')).getValue(), 'A editada');
  const rest = ctx.eliminarRegistro(a.id);
  igual(rest.recientes.map((r) => r.id), [b.id]);
  assert.throws(() => ctx.guardarRegistro(Object.assign({}, a, { id: 'noexiste' })), /No se encontró/);
});

prueba('totales se recalculan si se editan montos directo en la hoja', () => {
  const ctx = nuevoEntorno();
  ctx.guardarRegistro({ fecha: '2026-10-01', paciente: 'A', tipo: 'Control', ubicacion: 'Escazú', pago1: 35000, forma1: 'SINPE' });
  hoja(ctx, 'Registros').getRange(2, col(ctx, 'Pago 1')).setValue(40000);
  assert.strictEqual(ctx.getReporteMes('2026-10').totales.facturado, 40000);
});

/** Mes de ejemplo con las reglas reales. */
function mesDeEjemplo(ctx) {
  const g = (d) => ctx.guardarRegistro(Object.assign({ hora: '10:00', forma1: 'SINPE' }, d));
  g({ fecha: '2026-10-01', paciente: 'A', tipo: 'Primera vez', ubicacion: 'Escazú', pago1: 50000 });
  g({ fecha: '2026-10-02', paciente: 'B', tipo: 'Control', ubicacion: 'Escazú', pago1: 16000, conSeguro: true, aseguradora: 'Sana Sana', montoAseguradora: 8500 });
  g({ fecha: '2026-10-03', paciente: 'C', tipo: 'Control', ubicacion: 'Curridabat', pago1: 30000, forma1: 'Tarjeta' });
  g({ fecha: '2026-10-04', paciente: 'D', paciente2: 'E', tipo: 'Pareja control', ubicacion: 'Domicilio', descuento: 'VOLK',
    dosPagos: true, pago1: 22500, pago2: 22500, forma2: 'Tarjeta' });
  g({ fecha: '2026-11-01', paciente: 'F', tipo: 'Control', ubicacion: 'Escazú', pago1: 99999 });
}

prueba('reporte del mes: parte de cada sede, de Moni y por colaboradora', () => {
  const ctx = nuevoEntorno();
  mesDeEjemplo(ctx);
  const r = ctx.getReporteMes('2026-10');
  sinFechas(r);
  assert.strictEqual(r.etiqueta, 'Octubre 2026');
  const t = r.totales;
  igual([t.consultas, t.personas, t.facturado, t.descuentos, t.porCobrar], [4, 5, 149500, 5000, 8500]);
  igual(r.ubicaciones.map((u) => [u.nombre, u.facturado, u.comision, u.nutricionista]), [
    ['Escazú', 74500, 29800, 44700],
    ['Curridabat', 30000, 12000, 18000],
    ['Domicilio', 45000, 9000, 36000],
  ]);
  assert.strictEqual(t.comisiones, 50800);
  assert.strictEqual(t.neto, 98700);
  igual(r.encargados.map((e) => [e.nombre, e.ubicaciones, e.consultas, e.comision]), [
    ['Andrea R.', ['Escazú', 'Domicilio'], 3, 38800],
    ['Raquel D.', ['Curridabat'], 1, 12000],
  ]);
  igual(r.formas.map((f) => [f.nombre, f.monto]), [['SINPE', 88500], ['Tarjeta', 52500]]);

  // Base "Pago paciente" en Escazú: la comisión no incluye lo de la aseguradora
  hoja(ctx, 'Ubicaciones').getRange(2, 4).setValue('Pago paciente');
  assert.strictEqual(ctx.getReporteMes('2026-10').ubicaciones[0].comision, 26400);
  hoja(ctx, 'Ubicaciones').getRange(2, 4).setValue('Total facturado');

  const m = ctx.marcarEstadoAseguradora('2026-10', 'Sana Sana', 'Pendiente', 'Facturado');
  assert.strictEqual(m.cambiados, 1);
  assert.strictEqual(m.reporte.aseguradoras[0].Facturado, 8500);
  assert.strictEqual(m.reporte.totales.porCobrar, 8500);
  ctx.marcarEstadoAseguradora('2026-10', 'Sana Sana', 'Facturado', 'Pagado');
  assert.strictEqual(ctx.getReporteMes('2026-10').totales.porCobrar, 0);
});

prueba('exportar reportes: general, sede, colaboradora y aseguradora', () => {
  const ctx = nuevoEntorno();
  mesDeEjemplo(ctx);
  const texto = (i) => {
    const sh = ctx._creados[i].getSheets()[0];
    return JSON.stringify(sh.getRange(1, 1, sh.getLastRow(), 8).getValues());
  };
  const casos = [
    [{ tipo: 'general' }, ['Monto para Moni', '98700', '"Comisiones",50800']],
    [{ tipo: 'sede', filtro: 'Curridabat' }, ['Monto a pagar a Raquel D.', '12000', 'Monto para Moni', '18000']],
    [{ tipo: 'encargado', filtro: 'Andrea R.' }, ['Total a pagar a Andrea R.', '38800', '"Domicilio"', '(VOLK)']],
    [{ tipo: 'aseguradora', filtro: 'Sana Sana' }, ['"B"', '8500', 'Total a cobrar a Sana Sana']],
  ];
  casos.forEach(([opc, esperado], i) => {
    const r = ctx.exportarReporte(Object.assign({ mes: '2026-10', nombres: true }, opc));
    assert.ok(r.pdf.includes('format=pdf'));
    const tx = texto(i);
    assert.ok(tx.includes('Hábitos con propósito'));
    esperado.forEach((e) => assert.ok(tx.includes(e), opc.tipo + ': falta ' + e + '\n' + tx));
  });
  assert.ok(ctx.exportarReporte({ tipo: 'encargado', filtro: 'Raquel D.', mes: '2026-10', nombres: false }).nombre.includes('Liquidación Raquel D.'));
  assert.ok(!texto(4).includes('"C"'), 'sin nombres de pacientes');
});

prueba('importar registros desde la pestaña Importar: por nombre de columna y sin duplicar', () => {
  const ctx = nuevoEntorno();
  // Registros con columnas en otro orden (como una hoja creada con una versión anterior)
  const reg = hoja(ctx, 'Registros');
  const cab = reg.getRange(1, 1, 1, reg.getLastColumn()).getValues()[0];
  const movidas = cab.filter((h) => h !== 'Descuento' && h !== 'Monto descuento').concat(['Descuento', 'Monto descuento']);
  reg.getRange(1, 1, 1, movidas.length).setValues([movidas]);

  const imp = ctx._libro.insertSheet('Importar');
  imp.getRange(1, 1, 3, 9).setValues([
    ['ID', 'Fecha', 'Paciente', 'Tipo de cita', 'Ubicación', 'Descuento', 'Pago 1', 'Forma de pago 1', 'Con seguro'],
    ['H0001', new Date(2026, 8, 7), 'Carlos', 'Primera vez', 'Domicilio', 'VOLK', 31500, 'Tarjeta', 'No'],
    ['H0002', '2026-09-26', 'Paula', 'Control', 'Escazú', '', 35000, 'SINPE', 'No'],
  ]);
  const r1 = ctx.importarDesdeHoja_('Importar');
  assert.strictEqual(r1.importados, 2);
  const regs = ctx.getRecientes(10);
  const carlos = regs.find((r) => r.id === 'H0001');
  igual([carlos.fecha, carlos.descuento, carlos.pago1, carlos.forma1, carlos.ubicacion], ['2026-09-07', 'VOLK', 31500, 'Tarjeta', 'Domicilio']);
  assert.strictEqual(regs.find((r) => r.id === 'H0002').fecha, '2026-09-26');
  assert.strictEqual(reg.getRange(2, movidas.indexOf('Descuento') + 1).getValue(), 'VOLK');

  const r2 = ctx.importarDesdeHoja_('Importar');
  igual([r2.importados, r2.repetidos], [0, 2]);
  assert.strictEqual(ctx.getReporteMes('2026-09').totales.facturado, 66500);
});

prueba('eventos de Google: sugerencias de paciente, tipo y ubicación', () => {
  const ctx = nuevoEntorno();
  const { eventos, error } = ctx.getEventos('2026-10-13');
  assert.strictEqual(error, '');
  sinFechas(eventos);
  const por = (h) => eventos.find((e) => e.hora === h);
  igual(pick(por('08:00')), { paciente: 'María Fernández', paciente2: '', tipo: 'Control', ubicacion: 'Escazú' });
  igual(pick(por('09:30')), { paciente: 'Laura Jiménez', paciente2: '', tipo: 'Primera vez', ubicacion: 'Escazú' });
  igual(pick(por('11:00')), { paciente: 'Ana Solís', paciente2: 'Luis Mora', tipo: 'Pareja primera vez', ubicacion: 'Curridabat' });
  igual(pick(por('14:00')), { paciente: 'Sofía Rojas', paciente2: '', tipo: 'Control', ubicacion: 'Curridabat' });
  igual(pick(por('16:00')), { paciente: 'Carlos Vargas', paciente2: '', tipo: 'Mediciones', ubicacion: 'Domicilio' });
  assert.strictEqual(por('11:00').horaFin, '13:00');

  // Marcar como registrado
  const ev = por('08:00');
  ctx.guardarRegistro({ fecha: '2026-10-13', paciente: 'María', tipo: 'Control', ubicacion: 'Escazú', pago1: 0, eventoId: ev.id });
  assert.strictEqual(ctx.getEventos('2026-10-13').eventos.find((e) => e.hora === '08:00').registrado, true);

  // Filtro por palabra clave
  const cfg = hoja(ctx, 'Config');
  const fila = cfg.getRange(1, 1, cfg.getLastRow(), 1).getValues().findIndex((r) => r[0] === 'Solo eventos que contengan') + 1;
  cfg.getRange(fila, 2).setValue('control');
  assert.strictEqual(ctx.getEventos('2026-10-13').eventos.length, 2);
});

function pick(e) {
  return { paciente: e.paciente, paciente2: e.paciente2, tipo: e.tipo, ubicacion: e.ubicacion };
}

const ICS = [
  'BEGIN:VCALENDAR',
  'VERSION:2.0',
  'BEGIN:VEVENT',
  'UID:uno@agenda',
  'DTSTART;TZID=America/Costa_Rica:20261013T070000',
  'DTEND;TZID=America/Costa_Rica:20261013T080000',
  'SUMMARY:Control - Pedro Castro\\, Escazú',
  'BEGIN:VALARM',
  'SUMMARY:alarma que no es evento',
  'END:VALARM',
  'END:VEVENT',
  'BEGIN:VEVENT',
  'UID:dos@agenda',
  'DTSTART:20261013T140000Z', // 08:00 en Costa Rica (UTC-6), duplicado del evento de Google
  'DTEND:20261013T150000Z',
  'SUMMARY:María Fernández - Control',
  'END:VEVENT',
  'BEGIN:VEVENT',
  'UID:semanal@agenda',
  'DTSTART:20260901T190000',
  'DTEND:20260901T200000',
  'RRULE:FREQ=WEEKLY;BYDAY=TU;COUNT=20',
  'EXDATE:20261020T190000',
  'SUMMARY:Pareja control Rosa y Juan',
  'LOCATION:Currida',
  ' bat',
  'END:VEVENT',
  'BEGIN:VEVENT',
  'UID:semanal@agenda',
  'RECURRENCE-ID:20261027T190000',
  'DTSTART:20261027T170000',
  'DTEND:20261027T180000',
  'SUMMARY:Pareja control Rosa y Juan',
  'END:VEVENT',
  'BEGIN:VEVENT',
  'UID:cancelada@agenda',
  'DTSTART:20261013T100000',
  'STATUS:CANCELLED',
  'SUMMARY:Cancelada',
  'END:VEVENT',
  'BEGIN:VEVENT',
  'UID:dia@agenda',
  'DTSTART;VALUE=DATE:20261013',
  'SUMMARY:Feriado',
  'END:VEVENT',
  'END:VCALENDAR',
].join('\r\n');

prueba('iCal: zonas horarias, repeticiones, excepciones y duplicados', () => {
  const ctx = nuevoEntorno();
  ctx._icsPorUrl['https://agenda.example/cal.ics'] = ICS;
  const cfg = hoja(ctx, 'Config');
  const fila = cfg.getRange(1, 1, cfg.getLastRow(), 1).getValues().findIndex((r) => r[0] === 'iCal (enlaces .ics)') + 1;
  cfg.getRange(fila, 2).setValue('webcal://agenda.example/cal.ics');

  const dia13 = ctx.getEventos('2026-10-13').eventos;
  const ics13 = dia13.filter((e) => e.fuente === 'iCal');
  igual(ics13.map((e) => e.hora + ' ' + e.titulo), [
    '07:00 Control - Pedro Castro, Escazú',
    '19:00 Pareja control Rosa y Juan',
  ]);
  assert.strictEqual(dia13.filter((e) => e.hora === '08:00').length, 1, 'el duplicado Google/iCal se muestra una vez');
  const rosa = ics13[1];
  igual(pick(rosa), { paciente: 'Rosa', paciente2: 'Juan', tipo: 'Pareja control', ubicacion: 'Curridabat' });
  assert.strictEqual(pick(ics13[0]).paciente, 'Pedro Castro');

  const solo = (f) => ctx.getEventos(f).eventos.filter((e) => e.fuente === 'iCal').map((e) => e.hora);
  igual(solo('2026-10-20'), [], 'EXDATE');
  igual(solo('2026-10-27'), ['17:00'], 'RECURRENCE-ID reemplaza la ocurrencia');
  igual(solo('2026-10-14'), [], 'solo martes');
  igual(solo('2027-01-19'), [], 'COUNT=20 termina el 12 de enero');
  igual(solo('2027-01-12'), ['19:00']);

  // Enlace caído: avisa pero no rompe las citas de Google
  cfg.getRange(fila, 2).setValue('https://agenda.example/otra.ics');
  const r = ctx.getEventos('2026-10-13');
  assert.ok(r.error.includes('404'));
  assert.ok(r.eventos.length > 0);
});

let fallas = 0;
pruebas.forEach(({ nombre, fn }) => {
  try {
    fn();
    console.log('  ✓ ' + nombre);
  } catch (e) {
    fallas++;
    console.log('  ✗ ' + nombre + '\n    ' + (e.stack || e.message).split('\n').slice(0, 30).join('\n    '));
  }
});
console.log(`\n${pruebas.length - fallas}/${pruebas.length} pruebas pasaron`);
process.exit(fallas ? 1 : 0);
