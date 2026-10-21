# Transparencia Tributaria — Fundación Apertura

Mapa interactivo del Derecho de Registro e Inspección (DReI) en Rosario y su área metropolitana.

- **Página:** `index.html` (archivo único, sin dependencias locales).
- **Datos:** se leen en vivo de la planilla de carga en Google Sheets (pestañas `localidades`, `actividades`, `notas`, `config`). Si la planilla no responde, la página usa la copia de datos incluida.
- **Formularios de aportes:** envían a la aplicación web de Google Apps Script cuyo código está en `apps-script/Codigo.gs`. Los aportes se guardan en una planilla privada; las localidades nuevas se agregan a `localidades` con estado `aportada` («en revisión»).
- **Mapa:** Leaflet; límites municipales del IGN (WFS) y ubicación con la API Georef de datos.gob.ar; fondo Esri gris claro y Argenmap (IGN).

## Configuración

Al comienzo del script de `index.html`, en `TT_CONFIG`:

| Clave | Uso |
|---|---|
| `SHEET_ID` | ID de la planilla de carga |
| `APPS_SCRIPT_URL` | Dirección de la aplicación web que recibe los formularios |
| `BASEMAP_STYLE` | `claro`, `gris`, `calles` o `ninguno` |

Los textos (título, bajada, fecha de corte, localidad de referencia, escala de colores) se editan en la pestaña `config` de la planilla.

## Actualizar

Los cambios de datos se hacen en la planilla y se ven al recargar la página. Solo hay que volver a subir `index.html` si cambia el diseño o el código.
