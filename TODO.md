# Pendientes · Bitácora Hábitos con propósito

## 1. Agenda con autoservicio del paciente (prioridad)

**Idea:** que Moni casi no tenga que digitar. El paciente pide su cita y llena sus datos; Moni solo **revisa y valida**.

### Flujo propuesto
1. **Paciente** (página pública, separada de la app de Moni):
   - Elige ubicación y tipo de cita.
   - Ve solo los horarios libres, que salen del Calendar y respetan la duración y el margen de cada tipo.
   - Llena nombre, teléfono, correo y cédula. Si es pareja, los datos de la segunda persona. Si usa Sana Sana (solo Escazú, no pareja ni Mediciones), lo marca.
   - Acepta el uso de sus datos.
   - La solicitud queda **Pendiente de validar** en una hoja nueva, *Solicitudes*.
2. **Moni** (nueva pestaña "Por validar" en la app):
   - Revisa la solicitud. Puede aprobarla, cambiarle la hora o rechazarla.
   - Al aprobar, se crea el evento en Google Calendar y se le confirma al paciente.
3. **Después de la consulta:**
   - El registro ya viene lleno con los datos de la solicitud: paciente, tipo, ubicación y tarifa.
   - Moni solo confirma el monto y la forma de pago, y guarda.

### Por definir con Moni
- Horarios de atención por sede y por día.
- Duración de cada tipo de cita y margen entre citas. Ya se sabe: pareja 2 h, primera vez con 30 min de margen, controles seguidos sin margen. Falta Mediciones.
- Con cuánta anticipación se puede pedir o cancelar una cita.
- Cómo se confirma al paciente: correo automático, enlace de WhatsApp o ambos.
- Si se cobra algo por adelantado (por ejemplo, SINPE para reservar).
- Qué datos se piden y el texto de consentimiento.

### Notas técnicas
- La página del paciente se publica como una implementación aparte con acceso "Cualquier persona". Esa implementación solo puede crear solicitudes y consultar horarios libres; nunca leer la bitácora ni los reportes.
- Permisos nuevos: Calendar con escritura, para crear eventos, y posiblemente envío de correo.
- Protección contra spam: límite por teléfono o correo, y validación de campos.
- Reutilizar tarifas, tipos, ubicaciones y reglas de seguro y descuento de las hojas que ya existen.

## 2. Otros pendientes

- [ ] **Pasar a la cuenta de Hábitos:** copiar la hoja, borrar los registros de prueba, correr `setup`, publicar e importar la bitácora (ver README).
- [ ] **Importar octubre:** volver a correr `tools/importar_bitacora.py` con el Excel actualizado.
- [ ] **Base de datos de pacientes en un Google Sheet aparte, sincronizada con todo.** Detalle en la sección 3.
- [ ] **Reportes:** comparar entre meses, o un tablero en Looker Studio sobre la hoja Registros.
- [ ] **Recordatorios:** avisar a los pacientes antes de la cita (correo o WhatsApp).

## 3. Base de datos de pacientes (Sheet aparte, sincronizada)

**Idea:** tener los pacientes en un Google Sheet **separado** de la bitácora. Así los datos personales quedan en un solo lugar, se pueden manejar aparte y la bitácora y los reportes no los exponen. La agenda, la bitácora y los reportes leen y escriben ahí.

### Qué guarda
- **ID de paciente**, nombre completo, teléfono, correo, cédula y fecha de nacimiento (opcional).
- **Origen o referencia:** la marca y la visitadora que hoy están en la *Base de datos* del Excel, VOLK, etc.
- Si es pareja, con quién viene (ID de la otra persona).
- Aseguradora (Sana Sana) y número de afiliado, si aplica.
- Consentimiento de uso de datos, con fecha.
- Primera consulta, última consulta y total de consultas (calculados desde la bitácora).
- Notas.

### Cómo se sincroniza
- **Agenda (autoservicio):** cuando un paciente pide cita, se busca por cédula, correo o teléfono.
  - Si ya existe, se actualizan sus datos.
  - Si no existe, se crea como paciente nuevo.
  - Moni valida los cambios junto con la cita.
- **Bitácora:** cada registro guarda el **ID de paciente**, además del nombre.
  - El formulario autocompleta desde la base de pacientes, no desde los registros.
  - *Primera vez / Control* se decide con el historial real del paciente.
- **Calendar:** el evento lleva el ID del paciente en la descripción, para enlazar cita → solicitud → registro.
- **Reportes:** siguen leyendo solo la bitácora. Los nombres en los PDF salen del ID, y se puede seguir eligiendo no incluirlos.

### Migración
- Importar la hoja *Base de datos* del Excel actual (nombre, número, correo, cédula, lugar, marca, visitadora).
- Unir duplicados y variaciones de nombre (con y sin tilde, nombres incompletos como "Marie Paz Lemus").
- Asignar el ID de paciente a los registros ya importados.

### Por definir
- **Dónde vive el Sheet de pacientes:** misma cuenta de Hábitos, con el ID guardado en Config.
- **Quién más puede verlo:** por defecto solo Moni.
- **Qué dato identifica a un paciente sin duda:** cédula o teléfono.
- **Cuánto tiempo se guardan los datos:** y cómo borrar los de un paciente que lo pida.
