/*
 * Arma dev/preview.html: la app completa corriendo en el navegador con servicios
 * de Google simulados y datos de ejemplo. Sirve para revisar diseño y flujo sin
 * publicar nada.  Uso: npm run preview  → abrir dev/preview.html
 */
'use strict';
const fs = require('fs');
const path = require('path');

const raiz = path.join(__dirname, '..');
const leer = (p) => fs.readFileSync(path.join(raiz, p), 'utf8');

const shim = `
<script>${leer('dev/gas-mock.js')}</script>
<script>
  Object.assign(window, crearEntornoGAS());
</script>
<script>${leer('src/Code.gs')}</script>
<script>
(function () {
  // Datos de ejemplo
  setup();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t = ss.getSheetByName('Tarifas');
  const v = t.getRange(1, 1, t.getLastRow(), t.getLastColumn()).getValues();
  const precio = { 'Primera vez': 30000, 'Control': 22000, 'Pareja primera vez': 50000, 'Pareja control': 40000, 'Mediciones': 10000 };
  for (let i = 1; i < v.length; i++) for (let j = 1; j < v[0].length; j++) v[i][j] = precio[v[i][0]] + (v[0][j] === 'Curridabat' ? 3000 : 0);
  v[5][3] = '';
  t.getRange(1, 1, v.length, v[0].length).setValues(v);
  ss.getSheetByName('Ubicaciones').getRange(2, 3, 2, 1).setValues([[0.2], [0.15]]);
  ss.getSheetByName('Aseguradoras').getRange(2, 2).setValue(18000);

  const hoy = new Date();
  const f = (d) => Utilities.formatDate(new Date(hoy.getFullYear(), hoy.getMonth(), d), '', 'yyyy-MM-dd');
  const dia = Math.max(1, hoy.getDate() - 1);
  const nombres = ['Valeria Quesada', 'Daniela Arias', 'José Pablo Solano', 'Mariana Ulate', 'Andrés Brenes', 'Fernanda Mora', 'Gabriela Chaves', 'Sebastián Rojas'];
  const datos = [
    ['Primera vez', 'Escazú', 30000, 'SINPE'], ['Control', 'Curridabat', 25000, 'Efectivo'],
    ['Control', 'Escazú', 22000, 'Tarjeta'], ['Mediciones', 'Escazú', 10000, 'SINPE'],
    ['Primera vez', 'Curridabat', 33000, 'SINPE'], ['Control', 'Domicilio', 22000, 'SINPE'],
  ];
  datos.forEach((d, i) => guardarRegistro({ fecha: f(1 + (i * 3) % dia), hora: (8 + i) + ':00', paciente: nombres[i],
    tipo: d[0], ubicacion: d[1], pago1: d[2], forma1: d[3] }));
  guardarRegistro({ fecha: f(Math.min(dia, 5)), hora: '15:00', paciente: nombres[6], tipo: 'Control', ubicacion: 'Escazú',
    conSeguro: true, aseguradora: 'Sana Sana', montoAseguradora: 18000, pago1: 7000, forma1: 'SINPE' });
  guardarRegistro({ fecha: f(dia), hora: '17:00', paciente: nombres[7], paciente2: 'Paula Jiménez', tipo: 'Pareja control',
    ubicacion: 'Curridabat', dosPagos: true, pago1: 21500, forma1: 'SINPE', pago2: 21500, forma2: 'Efectivo' });

  // google.script.run simulado (asíncrono y con la misma restricción de no devolver Date)
  function revisar(o, ruta) {
    if (Object.prototype.toString.call(o) === '[object Date]') throw new Error('Date en ' + ruta + ': google.script.run no lo permite');
    if (o && typeof o === 'object') Object.keys(o).forEach((k) => revisar(o[k], ruta + '.' + k));
  }
  function runner(ok, fail) {
    return new Proxy({}, { get(_, fn) {
      if (fn === 'withSuccessHandler') return (f) => runner(f, fail);
      if (fn === 'withFailureHandler') return (f) => runner(ok, f);
      return function () {
        const args = JSON.parse(JSON.stringify(Array.from(arguments)));
        setTimeout(() => {
          try {
            const r = window[fn].apply(null, args);
            revisar(r, fn);
            ok && ok(r === undefined ? null : JSON.parse(JSON.stringify(r)));
          } catch (e) { console.error(e); fail && fail(e); }
        }, 200);
      };
    } });
  }
  window.google = { script: { run: runner() } };
})();
</script>`;

let html = leer('src/index.html')
  .replace("<?!= include('styles'); ?>", () => leer('src/styles.html'))
  .replace("<?!= include('app'); ?>", () => shim + '\n' + leer('src/app.html'));

if (/<\?/.test(html)) throw new Error('Quedó una etiqueta de plantilla sin reemplazar');
const salida = path.join(__dirname, 'preview.html');
fs.writeFileSync(salida, html);
console.log('Vista previa: ' + salida);
