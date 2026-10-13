# Hábitos con propósito · Bitácora de consultas

App web (Google Apps Script + Google Sheets) para que Moni registre cada consulta desde el celular o la PC. Todo queda en **una sola tabla** y desde la PC se generan los reportes del mes por sede y por aseguradora.

- **Registrar**: formulario con autollenado. La tarifa sale de *tipo de cita × ubicación*, las citas del día salen de Google Calendar o iCal y el app sugiere *Primera vez / Control* según si el paciente ya vino.
- **Seguro**: al marcar "Sí" aparecen la aseguradora, el total de la consulta con seguro, lo que paga la aseguradora (con monto fijo editable) y lo que paga el paciente, que se calcula solo. El estado del cobro es *Pendiente → Facturado → Pagado*.
- **Pareja**: dos nombres y uno o dos pagos, cada uno con su forma de pago.
- **Recientes**: los últimos 10 registros, editables o eliminables.
- **Reportes**: totales del mes por ubicación (con la comisión de cada sede), por aseguradora, por tipo de cita y por forma de pago. Se exportan a una hoja nueva en Drive y se descargan en PDF o Excel.

Solo la usa la administradora: la app corre con su cuenta y nadie más tiene acceso.

---

## Estructura

```
src/
  Code.gs          lógica del servidor (hojas, calendario, iCal, reportes)
  index.html       estructura de la app
  styles.html      estilo de la marca (colores y fuentes en variables CSS al inicio)
  app.html         lógica del formulario en el navegador
  appsscript.json  zona horaria, permisos y configuración de la web app
dev/
  gas-mock.js      simulación de los servicios de Google para probar sin publicar
  preview.js       arma dev/preview.html: la app completa con datos de ejemplo
test/
  test.js          pruebas automáticas del servidor
```

## Hojas que crea la app

| Hoja | Para qué |
|---|---|
| **Registros** | La tabla única: una fila por atención. Se puede ordenar, filtrar o usar en tablas dinámicas y Looker Studio. Las columnas se buscan por nombre, así que se pueden agregar columnas propias. |
| **Tarifas** | Matriz *tipo de cita × ubicación*. Celdas amarillas = sin tarifa (en la app se digita el monto). En pareja es el monto por ambas personas. |
| **Ubicaciones** | Encargado, **% de comisión** y base de cálculo (*Total facturado* = paciente + aseguradora, o *Pago paciente*). |
| **Aseguradoras** | Nombre y monto fijo por consulta (editable en cada registro). |
| **Tipos de cita** | Lista de tipos; *Es pareja = Sí* activa el segundo nombre. |
| **Formas de pago** | Efectivo, SINPE, Tarjeta… |
| **Config** | Nombre de la nutricionista, calendarios de Google, enlaces iCal, filtro de eventos y carpeta de reportes. |

Para agregar una ubicación, aseguradora o tipo nuevo basta con escribirlo en su hoja y correr **Hábitos → Sincronizar tarifas con catálogos**, que agrega la fila o columna en Tarifas. Para desactivar algo sin borrarlo, ponga *No* en la columna *Activo/Activa*.

---

## Instalación (la primera vez, en cualquier cuenta de Google)

1. Cree una **Google Sheet nueva** (por ejemplo "Bitácora Hábitos con propósito").
2. Abra **Extensiones → Apps Script**.
3. En el editor:
   - Reemplace el contenido de `Código.gs` con el de `src/Code.gs`.
   - Cree tres archivos HTML (**+ → HTML**) llamados exactamente `index`, `styles` y `app`, y pegue en cada uno el contenido de `src/index.html`, `src/styles.html` y `src/app.html`.
   - En **Configuración del proyecto** (ícono ⚙️) active *"Mostrar el archivo de manifiesto appsscript.json"* y reemplace su contenido con `src/appsscript.json`.
4. Arriba, elija la función **`setup`** y presione **Ejecutar**. Google pide permisos: acéptelos. Si sale *"Google no verificó esta app"*, entre a *Configuración avanzada → Ir a … (no seguro)*; es normal porque la app es suya. `setup` crea todas las hojas.
5. Vuelva a la hoja y complete:
   - **Tarifas**: los 15 montos.
   - **Ubicaciones**: % de comisión de Escazú (Andrea · DS) y de Curridabat.
   - **Aseguradoras**: monto fijo de Sana Sana.
6. Publique la app: **Implementar → Nueva implementación → ⚙️ Aplicación web**.
   - *Ejecutar como*: **Yo**
   - *Quién tiene acceso*: **Solo yo**
   - Copie la URL que termina en `/exec`.
7. En el celular, abra esa URL en Chrome o Safari y use **Agregar a pantalla de inicio**.

> Cuando cambie el código, publique una nueva versión en **Implementar → Administrar implementaciones → ✏️ → Versión: Nueva**. Así la URL no cambia.

### Alternativa con `clasp` (desde la terminal)

```bash
npm i -g @google/clasp
clasp login
cp .clasp.json.example .clasp.json   # pegar el ID del script (Configuración del proyecto → ID de la secuencia de comandos)
clasp push
```

## Calendario e iCal

En la hoja **Config**:

- **Calendarios de Google**: vacío usa el calendario principal. Se pueden poner varios IDs o nombres separados por coma, incluidos los calendarios suscritos a un iCal (en Google Calendar: *Configuración → el calendario → ID del calendario*). Con `ninguno` no se lee Google.
- **iCal (enlaces .ics)**: enlaces `https://…ics` o `webcal://…` de la agenda, separados por coma. El app los lee directo, así que las citas nuevas aparecen al instante (una suscripción en Google Calendar puede tardar varias horas en actualizarse). Si una cita viene por Google y por iCal a la vez, se muestra una sola vez.
- **Solo eventos que contengan**: opcional, para ocultar eventos personales.

El app intenta sacar del título del evento el nombre, el tipo y la ubicación, por ejemplo *"María Pérez – Control Escazú"* o *"Ana y Luis · Pareja primera vez"*. Siempre se puede corregir antes de guardar. Las citas ya registradas aparecen con ✓.

## Pasar de la cuenta de prueba a la cuenta de Hábitos

**Opción A (recomendada):** copiar la hoja completa.

1. Comparta la hoja de prueba con la cuenta de Hábitos (o descárguela).
2. Desde la cuenta de Hábitos: **Archivo → Hacer una copia**. La copia incluye el script.
3. En la copia: **Hábitos → Borrar todos los registros (datos de prueba)**.
4. Abra **Extensiones → Apps Script**, ejecute `setup` una vez para autorizar con la nueva cuenta y repita el paso 6 (publicar). La URL nueva es la que va en el celular de Moni.
5. Revise **Config**: los calendarios ahora son los de la cuenta de Hábitos.

**Opción B:** instalación limpia en la cuenta de Hábitos siguiendo los pasos de arriba, y copiar a mano las hojas Tarifas, Ubicaciones y Aseguradoras.

---

## Desarrollo y pruebas

```bash
npm test          # pruebas del servidor con Google simulado
npm run preview   # genera dev/preview.html: la app con datos de ejemplo, se abre en el navegador
```

Las pruebas cubren: creación de hojas, tarifas, registro simple, pareja con dos pagos y seguro, validaciones, edición y eliminación, reportes y comisiones, cambio de estado de la aseguradora, exportación, sugerencias desde Calendar y lectura de iCal (zonas horarias, repeticiones semanales, excepciones, cancelaciones y duplicados).

### Notas técnicas

- Las fechas se guardan como fecha real y la hora como texto `HH:mm`, en la zona horaria `America/Costa_Rica`.
- Los totales (*Total paciente*, *Total facturado*) se recalculan al leer, así que si se corrige un monto directo en la hoja, los reportes salen bien.
- Las comisiones se calculan con el % vigente en la hoja Ubicaciones al momento de generar el reporte.
- Permisos que pide la app: Hojas de cálculo (bitácora y reportes), Calendar en solo lectura, Drive (carpeta de reportes), interfaz de la hoja (menú) y conexiones externas (leer el iCal).
