# Búsqueda laboral · Dai — Tracker (Google Sheets + Apps Script)

Tracker de búsqueda de trabajo en **Google Sheets + Google Apps Script**.
Dos pestañas: una para cargar postulaciones y otra con un dashboard que
muestra el resumen a simple vista.

> Proyecto en **versión simplificada**: se eliminaron las pestañas
> Calendario, Ayuda y Entrevistas. Todo vive en 2 pestañas.

## 😴 Qué hace

- **Postulaciones**: una fila por postulación con empresa, puesto, fecha,
  canal, estado, salarios, próxima entrevista y notas. Los estados se
  colorean solos y hay filtros automáticos por columna.
- **Dashboard**: 4 tarjetas (total, este mes, en proceso, próximas
  entrevistas), barras por estado y por canal, y la lista de entrevistas de
  los próximos 7 días.
- **Aviso automático**: al abrir la planilla, si tenés una entrevista dentro
  de las próximas 48 h, te salta una alerta.
- **Formulario lateral**: un menú `🚀 Búsqueda ▸ ➕ Nueva postulación` abre un
  formulario desde donde carga cada postulación sin tocar la hoja.

## 📦 Archivos

| Archivo | Qué hace |
|---|---|
| `Code.gs` | Lógica completa: menús, setup, dashboard, alertas, API para el formulario |
| `postulacion.html` | Formulario lateral para cargar postulaciones |
| `VISTA-PREVIA.html` | Vista previa local (navegador) de cómo queda la planilla |

## 🚀 Instalación (3 minutos)

1. Andá a [sheets.new](https://sheets.new) y llamala "Búsqueda laboral · Dai".
2. **Extensiones ▸ Apps Script**.
3. Reemplazá el contenido de `Code.gs` con el de este proyecto.
4. Click en **+** ▸ **HTML** → nombrá el archivo `postulacion` (sin extensión)
   → pegá el contenido de `postulacion.html`.
5. **Ctrl+S** y ejecutá la función `setupBuscador` (▶). Autorizá cuando lo pida.
6. Volvé a la planilla: ya tenés las 2 pestañas listas.

> ⚠️ **Importante:** ejecutar `setupBuscador` borra las hojas existentes.
> Correlo una sola vez sobre una planilla nueva.

El formulario y el dashboard se actualizan solos. Para abrir el formulario
en cualquier momento: menú **🚀 Búsqueda ▸ ➕ Nueva postulación**.

## 🤔 Dónde ver el resultado

Google Sheets no permite generar planillas desde texto; el repositorio trae
`VISTA-PREVIA.html` que podés abrir en el navegador para ver el diseño
(datos de ejemplo) antes de instalar el script.