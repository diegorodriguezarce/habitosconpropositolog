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
  // Datos de ejemplo (tarifas, comisiones, Sana Sana y VOLK vienen de setup)
  setup();
  const hoy = new Date();
  const f = (d) => Utilities.formatDate(new Date(hoy.getFullYear(), hoy.getMonth(), d), '', 'yyyy-MM-dd');
  const dia = Math.max(1, hoy.getDate() - 1);
  const g = (d) => guardarRegistro(Object.assign({ forma1: 'SINPE' }, d));
  g({ fecha: f(1), hora: '08:00', paciente: 'Valeria Quesada', tipo: 'Primera vez', ubicacion: 'Escazú', pago1: 50000 });
  g({ fecha: f(Math.min(dia, 3)), hora: '09:00', paciente: 'Daniela Arias', tipo: 'Control', ubicacion: 'Curridabat', pago1: 30000, forma1: 'Tarjeta' });
  g({ fecha: f(Math.min(dia, 5)), hora: '10:00', paciente: 'José Pablo Solano', tipo: 'Control', ubicacion: 'Escazú',
    conSeguro: true, aseguradora: 'Sana Sana', montoAseguradora: 8500, pago1: 16000 });
  g({ fecha: f(Math.min(dia, 6)), hora: '11:00', paciente: 'Mariana Ulate', tipo: 'Mediciones', ubicacion: 'Escazú', pago1: 14000 });
  g({ fecha: f(Math.min(dia, 7)), hora: '15:00', paciente: 'Andrés Brenes', tipo: 'Control', ubicacion: 'Domicilio', descuento: 'VOLK', pago1: 27000 });
  g({ fecha: f(Math.min(dia, 8)), hora: '16:00', paciente: 'Gabriela Chaves', tipo: 'Primera vez', ubicacion: 'Curridabat', pago1: 40000, forma1: 'Tarjeta' });
  g({ fecha: f(dia), hora: '17:00', paciente: 'Sebastián Rojas', paciente2: 'Paula Jiménez', tipo: 'Pareja control',
    ubicacion: 'Escazú', dosPagos: true, pago1: 32000, pago2: 32000, forma2: 'Tarjeta' });

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
