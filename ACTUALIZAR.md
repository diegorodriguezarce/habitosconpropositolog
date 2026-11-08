# Cómo actualizar la app de Moni

**Flujo:** Diego hace los cambios en este repo de GitHub. Las actualizaciones se cargan desde la compu de Moni, con la cuenta de Hábitos.

> ⚠️ **La versión nueva siempre la publica la cuenta de Hábitos.** El app corre como *"Ejecutar como: Yo"* y *"Solo yo"*. Si publica otra cuenta, el app puede quedar a su nombre y dejar de abrirle a Moni.

## Pasos (en la compu de Moni, con la cuenta de Hábitos)

1. Abrir la hoja **Bitácora Hábitos con propósito** → **Extensiones → Apps Script**.
2. Por cada archivo que diga la versión (tabla de abajo):
   1. Abrirlo en GitHub, en la carpeta [`src/`](https://github.com/diegorodriguezarce/habitosconpropositolog/tree/claude/relaxed-lovelace-x3hb4k/src).
   2. Tocar el botón de **copiar** (*Copy raw file*).
   3. En Apps Script, abrir el archivo con el mismo nombre (`Code.gs` → `Código.gs`), seleccionar todo (Ctrl + A), pegar (Ctrl + V) y guardar (Ctrl + S).
3. Si la versión lo indica, ir a la hoja y correr **Hábitos → Configurar / reparar hojas**.
4. **Implementar → Administrar implementaciones → ✏️ (editar) → Versión: Nueva versión → Implementar.** La URL no cambia.
5. Abrir el app y revisar que **al pie diga la versión nueva**. Si sigue la anterior, recargar la página.

## Versiones

| Versión | Qué cambió | Archivos a pegar | ¿Reparar hojas? |
|---|---|---|---|
| **1.1** | Número de versión al pie del app. Desde aquí cada actualización se anota en esta tabla. | `Code.gs`, `index`, `styles`, `app` | No |
| 1.0 | Base desplegada: registro con tarifas, Sana Sana, VOLK y "Otro"; calendarios por sede; pestaña *No registradas* (solo citas); reportes y liquidaciones (resumen por persona); logo oficial; "Doctora Moni"; filtro en Registros; pestañas ocultas respetadas. | Todos | Sí |

> **Para Diego:** en cada cambio, subir `VERSION_APP` en `src/Code.gs` y agregar una fila arriba en esta tabla.
