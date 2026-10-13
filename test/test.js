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

/** Pone montos de ejemplo en Tarifas, Ubicaciones y Aseguradoras. */
function configurarMontos(ctx) {
  const t = hoja(ctx, 'Tarifas');
  const v = t.getRange(1, 1, t.getLastRow(), t.getLastColumn()).getValues();
  const precio = { 'Primera vez': 30000, Control: 22000, 'Pareja primera vez': 50000, 'Pareja control': 40000, Mediciones: 10000 };
  for (let i = 1; i < v.length; i++) {
    for (let j = 1; j < v[0].length; j++) v[i][j] = precio[v[i][0]] + (v[0][j] === 'Curridabat' ? 3000 : 0);
  }
  t.getRange(1, 1, v.length, v[0].length).setValues(v);
  const u = hoja(ctx, 'Ubicaciones');
  u.getRange(2, 3, 3, 1).setValues([[0.2], ['15%'], [0]]);
  hoja(ctx, 'Aseguradoras').getRange(2, 2).setValue(18000);
}

const pruebas = [];
const prueba = (nombre, fn) => pruebas.push({ nombre, fn });

prueba('setup crea las hojas, es idempotente y deja tarifas en matriz', () => {
  const ctx = nuevoEntorno();
  ctx.setup();
  const nombres = ctx._libro.getSheets().map((h) => h.getName());
  igual(nombres, ['Registros', 'Tarifas', 'Ubicaciones', 'Aseguradoras', 'Tipos de cita', 'Formas de pago', 'Config']);
  const t = hoja(ctx, 'Tarifas');
  igual(t.getRange(1, 1, 1, 4).getValues()[0], ['Tipo de cita', 'Escazú', 'Curridabat', 'Domicilio']);
  assert.strictEqual(t.getLastRow(), 6);
  const reg = hoja(ctx, 'Registros');
  assert.strictEqual(reg.getRange(1, 1, 1, reg.getLastColumn()).getValues()[0].length, 22);
  const cfg = hoja(ctx, 'Config');
  assert.strictEqual(cfg.getLastRow(), 6, 'Config no debe duplicar claves');
});

prueba('sincronizarTarifas agrega ubicaciones nuevas sin perder montos', () => {
  const ctx = nuevoEntorno();
  configurarMontos(ctx);
  hoja(ctx, 'Ubicaciones').getRange(5, 1, 1, 5).setValues([['Virtual', '', 0, 'Total facturado', 'Sí']]);
  ctx.sincronizarTarifas();
  const tar = ctx.tarifas_();
  assert.strictEqual(tar['Control']['Escazú'], 22000);
  assert.strictEqual(tar['Control']['Virtual'], null);
});

prueba('bootstrap entrega catálogos y nada con Date', () => {
  const ctx = nuevoEntorno();
  configurarMontos(ctx);
  const b = ctx.getBootstrap();
  assert.strictEqual(JSON.stringify(b).includes('"__date'), false);
  igual(b.cats.formas, ['Efectivo', 'SINPE', 'Tarjeta']);
  assert.strictEqual(b.cats.ubicaciones[0].porcentaje, 0.2);
  assert.strictEqual(b.cats.ubicaciones[1].porcentaje, 0.15);
  assert.strictEqual(b.cats.aseguradoras[0].monto, 18000);
  assert.strictEqual(b.tarifas['Primera vez']['Curridabat'], 33000);
  assert.strictEqual(b.cats.tipos.filter((t) => t.pareja).length, 2);
  sinFechas(b);
});

function sinFechas(o, ruta = 'resultado') {
  if (Object.prototype.toString.call(o) === '[object Date]') throw new Error('Date en ' + ruta + ' (google.script.run no lo permite)');
  if (o && typeof o === 'object') Object.keys(o).forEach((k) => sinFechas(o[k], ruta + '.' + k));
}

prueba('guardar consulta simple: fecha real, hora texto, totales', () => {
  const ctx = nuevoEntorno();
  configurarMontos(ctx);
  const r = ctx.guardarRegistro({
    fecha: '2026-10-13', hora: '8:05', paciente: ' María Fernández ', tipo: 'Control', ubicacion: 'Escazú',
    conSeguro: false, pago1: 22000, forma1: 'SINPE',
  });
  sinFechas(r);
  assert.strictEqual(r.registro.paciente, 'María Fernández');
  assert.strictEqual(r.registro.hora, '08:05');
  assert.strictEqual(r.registro.totalFacturado, 22000);
  assert.strictEqual(r.registro.tarifa, 22000);
  const fila = hoja(ctx, 'Registros').getRange(2, 1, 1, 22).getValues()[0];
  assert.ok(Object.prototype.toString.call(fila[1]) === '[object Date]', 'Fecha debe guardarse como fecha');
  assert.strictEqual(fila[1].getDate(), 13);
  assert.strictEqual(fila[7], 'No');
  assert.strictEqual(r.recientes.length, 1);
  assert.strictEqual(r.pacientes[0].n, 'María Fernández');
});

prueba('pareja con dos pagos y seguro', () => {
  const ctx = nuevoEntorno();
  configurarMontos(ctx);
  const r = ctx.guardarRegistro({
    fecha: '2026-10-14', hora: '11:00', paciente: 'Ana Solís', paciente2: 'Luis Mora', tipo: 'Pareja primera vez',
    ubicacion: 'Curridabat', conSeguro: true, aseguradora: 'Sana Sana', montoAseguradora: 18000,
    dosPagos: true, pago1: 20000, forma1: 'Efectivo', pago2: 15000, forma2: 'Tarjeta',
  }).registro;
  assert.strictEqual(r.totalPaciente, 35000);
  assert.strictEqual(r.montoAseguradora, 18000);
  assert.strictEqual(r.totalFacturado, 53000);
  assert.strictEqual(r.estadoAseguradora, 'Pendiente');
  assert.strictEqual(r.paciente2, 'Luis Mora');
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
  configurarMontos(ctx);
  const a = ctx.guardarRegistro({ fecha: '2026-10-01', paciente: 'A', tipo: 'Control', ubicacion: 'Escazú', pago1: 22000, forma1: 'SINPE' }).registro;
  const b = ctx.guardarRegistro({ fecha: '2026-10-02', paciente: 'B', tipo: 'Control', ubicacion: 'Escazú', pago1: 22000, forma1: 'SINPE' }).registro;
  const e = ctx.guardarRegistro(Object.assign({}, a, { id: a.id, paciente: 'A editada', pago1: 20000 })).registro;
  assert.strictEqual(e.id, a.id);
  assert.strictEqual(e.creado, a.creado);
  assert.strictEqual(hoja(ctx, 'Registros').getLastRow(), 3);
  assert.strictEqual(hoja(ctx, 'Registros').getRange(2, 4).getValue(), 'A editada');
  const rest = ctx.eliminarRegistro(a.id);
  igual(rest.recientes.map((r) => r.id), [b.id]);
  assert.throws(() => ctx.guardarRegistro(Object.assign({}, a, { id: 'noexiste' })), /No se encontró/);
});

prueba('totales se recalculan si se editan montos directo en la hoja', () => {
  const ctx = nuevoEntorno();
  ctx.guardarRegistro({ fecha: '2026-10-01', paciente: 'A', tipo: 'Control', ubicacion: 'Escazú', pago1: 22000, forma1: 'SINPE' });
  hoja(ctx, 'Registros').getRange(2, 11).setValue(25000);
  assert.strictEqual(ctx.getReporteMes('2026-10').totales.facturado, 25000);
});

prueba('reporte del mes: comisiones por sede, aseguradoras y formas de pago', () => {
  const ctx = nuevoEntorno();
  configurarMontos(ctx);
  const g = (d) => ctx.guardarRegistro(Object.assign({ hora: '10:00', forma1: 'SINPE' }, d));
  g({ fecha: '2026-10-01', paciente: 'A', tipo: 'Primera vez', ubicacion: 'Escazú', pago1: 30000 });
  g({ fecha: '2026-10-02', paciente: 'B', tipo: 'Control', ubicacion: 'Escazú', pago1: 4000, conSeguro: true, aseguradora: 'Sana Sana', montoAseguradora: 18000 });
  g({ fecha: '2026-10-03', paciente: 'C', tipo: 'Control', ubicacion: 'Curridabat', pago1: 25000, forma1: 'Efectivo' });
  g({ fecha: '2026-10-04', paciente: 'D', paciente2: 'E', tipo: 'Pareja control', ubicacion: 'Domicilio', dosPagos: true, pago1: 20000, pago2: 20000, forma2: 'Tarjeta' });
  g({ fecha: '2026-11-01', paciente: 'F', tipo: 'Control', ubicacion: 'Escazú', pago1: 99999 });

  const r = ctx.getReporteMes('2026-10');
  sinFechas(r);
  assert.strictEqual(r.etiqueta, 'Octubre 2026');
  assert.strictEqual(r.totales.consultas, 4);
  assert.strictEqual(r.totales.personas, 5);
  assert.strictEqual(r.totales.facturado, 30000 + 22000 + 25000 + 40000);
  const esc = r.ubicaciones.find((u) => u.nombre === 'Escazú');
  assert.strictEqual(esc.facturado, 52000);
  assert.strictEqual(esc.comision, 10400); // 20% del total facturado
  const cur = r.ubicaciones.find((u) => u.nombre === 'Curridabat');
  assert.strictEqual(cur.comision, 3750); // 15%
  assert.strictEqual(r.totales.comisiones, 14150);
  assert.strictEqual(r.totales.porCobrar, 18000);
  igual(r.formas.map((f) => [f.nombre, f.monto]), [['Efectivo', 25000], ['SINPE', 54000], ['Tarjeta', 20000]]);

  // Base "Pago paciente" en Escazú: la comisión no incluye lo de la aseguradora
  hoja(ctx, 'Ubicaciones').getRange(2, 4).setValue('Pago paciente');
  assert.strictEqual(ctx.getReporteMes('2026-10').ubicaciones[0].comision, 6800);

  const m = ctx.marcarEstadoAseguradora('2026-10', 'Sana Sana', 'Pendiente', 'Facturado');
  assert.strictEqual(m.cambiados, 1);
  assert.strictEqual(m.reporte.aseguradoras[0].Facturado, 18000);
  assert.strictEqual(m.reporte.totales.porCobrar, 18000);
  ctx.marcarEstadoAseguradora('2026-10', 'Sana Sana', 'Facturado', 'Pagado');
  assert.strictEqual(ctx.getReporteMes('2026-10').totales.porCobrar, 0);
});

prueba('exportar reportes crea hojas con datos', () => {
  const ctx = nuevoEntorno();
  configurarMontos(ctx);
  ctx.guardarRegistro({ fecha: '2026-10-02', paciente: 'B', tipo: 'Control', ubicacion: 'Curridabat', pago1: 4000, forma1: 'SINPE', conSeguro: true, aseguradora: 'Sana Sana', montoAseguradora: 18000 });
  ['general', 'sede', 'aseguradora'].forEach((tipo, i) => {
    const filtro = { sede: 'Curridabat', aseguradora: 'Sana Sana' }[tipo];
    const r = ctx.exportarReporte({ tipo, mes: '2026-10', filtro, nombres: true });
    assert.ok(r.pdf.includes('format=pdf'));
    const sh = ctx._creados[i].getSheets()[0];
    const texto = JSON.stringify(sh.getRange(1, 1, sh.getLastRow(), 7).getValues());
    assert.ok(texto.includes('Hábitos con propósito'));
    if (tipo === 'sede') assert.ok(texto.includes('Monto a pagar a Encargado Curridabat') && texto.includes('3300'), texto);
    if (tipo === 'aseguradora') assert.ok(texto.includes('"B"') && texto.includes('18000'));
  });
  const sinNombres = ctx.exportarReporte({ tipo: 'sede', mes: '2026-10', filtro: 'Curridabat', nombres: false });
  const sh = ctx._creados[3].getSheets()[0];
  assert.ok(!JSON.stringify(sh.getRange(1, 1, sh.getLastRow(), 7).getValues()).includes('"B"'));
  assert.ok(sinNombres.url);
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
