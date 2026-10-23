# Marco FPV Aerial Films · Sitio web

Sitio estático para GitHub Pages: servicios, portafolio, cotizador con cálculo de distancia y contacto por WhatsApp.

## Qué editar

Todo está al inicio del bloque `<script>` de `index.html`, en `CONFIG` y `PORTFOLIO`:

| Dato | Dónde |
|---|---|
| Número de WhatsApp | `CONFIG.whatsapp` (ej. `56912345678`) |
| Correo que recibe cotizaciones | atributo `action` del formulario (`formsubmit.co/TU_CORREO`) |
| Dirección de salida | `CONFIG.origen` |
| Precios por jornada y extras | `CONFIG.services` y `CONFIG.extras` |
| Rendimiento del auto, bencina y comida | `CONFIG.kmPorLitro`, `CONFIG.precioLitro`, `CONFIG.comidaPorPersona` |
| Videos del portafolio | `PORTFOLIO` (ID de YouTube y miniatura en `assets/`) |

## Archivos multimedia

Súbelos a la carpeta `assets/`:

- `hero.mp4`: reel de portada, 10–15 s, sin audio, menos de 8 MB
- `hero-poster.jpg`: imagen que se ve mientras carga el video
- `p1.jpg` … `p6.jpg`: miniaturas del portafolio

## Servicios externos (gratuitos)

- **FormSubmit**: envía el formulario a tu correo. El primer envío te llega como correo de activación.
- **OpenStreetMap (Nominatim)** y **OSRM**: buscan la dirección del evento y calculan la distancia en auto.
