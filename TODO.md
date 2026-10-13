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
- [ ] **Hoja Pacientes:** teléfono, correo y cédula, tomados de la *Base de datos* del Excel; autocompletar en el formulario.
- [ ] **Reportes:** comparar entre meses, o un tablero en Looker Studio sobre la hoja Registros.
- [ ] **Recordatorios:** avisar a los pacientes antes de la cita (correo o WhatsApp).
