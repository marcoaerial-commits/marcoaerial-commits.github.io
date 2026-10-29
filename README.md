# Marco FPV Aerial Films · Sitio web

Sitio estático para GitHub Pages: servicios, portafolio, cotizador con cálculo de distancia y contacto por WhatsApp.

## Qué editar

Todo está al inicio del bloque `<script>` de `index.html`, en `CONFIG` y `PORTFOLIO`:

| Dato | Dónde |
|---|---|
| Número de WhatsApp | `CONFIG.whatsapp` (ej. `56912345678`) |
| Correo que recibe cotizaciones | `AJUSTES.miCorreo` en `apps-script/Codigo.gs` |
| Dirección de salida | `CONFIG.origen` |
| Precios, extras y costos internos | planilla privada **Precios MarcoFPV** en Google Drive (no en el código) |
| Precio de la bencina 95 | `combustible.json` (se actualiza solo cada viernes) |
| Videos del portafolio | `PORTFOLIO` (ID de YouTube y miniatura en `assets/`) |

## Archivos multimedia

Súbelos a la carpeta `assets/`:

- `hero.mp4`: reel de portada, 10–15 s, sin audio, menos de 8 MB
- `hero-poster.jpg`: imagen que se ve mientras carga el video
- `p1.jpg` … `p6.jpg`: miniaturas del portafolio

## Servicios externos (gratuitos)

- **FormSubmit**: envía el formulario a tu correo. El primer envío te llega como correo de activación.
- **OpenStreetMap (Nominatim)** y **OSRM**: buscan la dirección del evento y calculan la distancia en auto.

## Correos desde Gmail (Google Apps Script)

El archivo `apps-script/Codigo.gs` recibe cada cotización, la guarda en la planilla **Cotizaciones MarcoFPV** de Google Drive y envía desde Gmail un aviso a Marco y un correo de respaldo al cliente.

1. Abre la planilla **Cotizaciones MarcoFPV** → **Extensiones → Apps Script**.
2. Borra el contenido de `Código.gs`, pega el de `apps-script/Codigo.gs` y guarda.
3. Elige la función `probar` y toca **Ejecutar**. Autoriza los permisos con tu cuenta (en "Google no verificó esta app": **Configuración avanzada → Ir a…**).
4. **Implementar → Nueva implementación → Aplicación web**. Ejecutar como: **Yo**. Quién tiene acceso: **Cualquier usuario**. Toca **Implementar** y copia la URL que termina en `/exec`.
5. Pega esa URL en `CONFIG.endpoint` de `index.html`.

Mientras `CONFIG.endpoint` esté vacío, el formulario usa FormSubmit (sin correo de respaldo al cliente).
