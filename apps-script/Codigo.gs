/**
 * Marco FPV Aerial Films · Cotizaciones
 *
 * Recibe las solicitudes de https://marcofpv.cl/cotizar.html,
 * calcula el valor con los precios de la planilla privada "Precios MarcoFPV",
 * guarda la solicitud en "Cotizaciones MarcoFPV" y envía desde tu Gmail:
 *   - al cliente: su cotización con el total (sin desglose de costos internos)
 *   - a ti: el aviso con el desglose completo
 *
 * Los precios NO están en este código (que es público en GitHub):
 * se leen de la planilla "Precios MarcoFPV" en tu Google Drive.
 * El precio de la bencina se lee de combustible.json, que se actualiza cada semana.
 *
 * Instalación: ver README.md del repositorio (sección "Correos desde Gmail").
 */

const AJUSTES = {
  planillaCotizaciones: "19zwupJG8FBwBzVkBFGLaVrUqJOxZJXb9XgH63xm5M1g",
  planillaPrecios: "1-01StYibo8LUbQBospbiI3v7o6mlnZ5whhkm7FbRJEU",
  combustible: "https://marcofpv.cl/combustible.json",
  miCorreo: "marcoaerial@gmail.com",
  nombreRemitente: "Marco FPV Aerial Films",
  whatsapp: "56939070008",
  web: "https://marcofpv.cl",
  logo: "https://marcofpv.cl/assets/logo.png",
  validezDias: 15,
  carpetaPdf: "1TadG7Z_poljNRY58X2qvkFmG-RZbue52"   // carpeta "Cotizaciones MarcoFPV - PDF" en tu Drive
  // Titular, RUT y correo del PDF se leen de la planilla privada (filas con Tipo "dato"),
  // para que el RUT no quede publicado en GitHub.
};

const COLUMNAS = ["Fecha de envío", "Nombre", "WhatsApp", "Correo", "Servicio", "Jornadas", "Extras",
  "Dirección del evento", "Km ida desde Maipú", "Fecha del evento", "Total estimado", "Mensaje", "Estado",
  "Folio", "Desglose interno", "PDF"];

/* ================= recepción ================= */

function doPost(e) {
  try {
    const p = e.parameter || {};
    if (p._gotcha) return responder({ success: true }); // trampa anti-spam

    const precios = leerPrecios();
    const ids = String(p.Servicios_ids || "").split(",").map(s => s.trim()).filter(Boolean);
    const exIds = (e.parameters && e.parameters["Extras[]"]) || [];

    const d = {
      nombre: limpiar(p.Nombre, 80),
      telefono: limpiar(p.Telefono, 30),
      correo: limpiar(p.email, 120),
      jornadas: Math.min(30, Math.max(1, parseInt(p.Jornadas, 10) || 1)),
      direccion: limpiar(p.Direccion_evento, 200),
      km: Math.min(3000, Math.max(0, parseFloat(p.Km_ida_desde_Maipu) || 0)),
      fecha: limpiar(p.Fecha, 10),
      mensaje: limpiar(p.Mensaje, 2000),
      servicios: ids.map(id => precios.servicios[id]).filter(Boolean),
      extras: exIds.map(id => precios.extras[limpiar(id, 20)]).filter(Boolean)
    };

    if (!d.nombre || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(d.correo)) {
      return responder({ success: false, message: "Faltan el nombre o un correo válido." });
    }
    if (!d.servicios.length) return responder({ success: false, message: "Elige al menos un servicio." });

    d.fechaTexto = d.fecha ? d.fecha.split("-").reverse().join("-") : "Por definir";
    d.c = calcular(d, precios.ajustes);
    d.titular = {
      nombre: precios.datos.titular || "Marco Salazar",
      rut: precios.datos.rut || "",
      correo: precios.datos.correo || AJUSTES.miCorreo
    };

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      d.folio = siguienteFolio();
      d.pdf = crearPdf(d);
      guardarEnPlanilla(d);
    } finally { lock.releaseLock(); }

    cotizacionAlCliente(d);
    avisarAMarco(d);
    return responder({ success: true, folio: d.folio });
  } catch (err) {
    console.error(err);
    return responder({ success: false, message: "Error interno" });
  }
}

function doGet() {
  return responder({ ok: true, servicio: "Cotizaciones MarcoFPV" });
}

/* ================= precios ================= */

function leerPrecios() {
  const filas = SpreadsheetApp.openById(AJUSTES.planillaPrecios).getSheets()[0].getDataRange().getValues().slice(1);
  const out = { servicios: {}, extras: {}, ajustes: {}, datos: {} };
  filas.forEach(([tipo, clave, nombre, valor]) => {
    tipo = String(tipo).trim().toLowerCase(); clave = String(clave).trim();
    const v = Number(String(valor).replace(/[^\d.,-]/g, "").replace(/\./g, "").replace(",", ".")) || 0;
    if (!clave) return;
    if (tipo === "servicio") out.servicios[clave] = { id: clave, nombre: String(nombre), precio: v };
    else if (tipo === "extra") out.extras[clave] = { id: clave, nombre: String(nombre), precio: v };
    else if (tipo === "ajuste") out.ajustes[clave] = v;
    else if (tipo === "dato") out.datos[clave] = String(nombre).trim();
  });
  return out;
}

function precioBencina(respaldo) {
  const cache = CacheService.getScriptCache();
  const enCache = cache.get("precioLitro");
  if (enCache) return Number(enCache);
  try {
    const j = JSON.parse(UrlFetchApp.fetch(AJUSTES.combustible, { muteHttpExceptions: true }).getContentText());
    const v = Number(j.precioLitro);
    if (v > 500 && v < 5000) { cache.put("precioLitro", String(v), 6 * 3600); return v; }
  } catch (e) { console.warn("No se pudo leer combustible.json", e); }
  return respaldo || 1578;
}

function calcular(d, a) {
  const dias = d.jornadas, km = d.km;
  const sumaServ = d.servicios.reduce((s, x) => s + x.precio, 0) * dias;
  const desc = d.servicios.length > 1 ? (a.descuentoMultiServicio || 0) / 100 : 0;
  const servicios = Math.round(sumaServ * (1 - desc));
  const extras = d.extras.reduce((s, x) => s + x.precio, 0);

  const litro = precioBencina(a.precioLitroRespaldo);
  const ciudad = Math.min(km, a.kmCiudadPorTramo || 25);
  const carretera = Math.max(0, km - ciudad);
  const litrosTramo = ciudad / (a.rendimientoCiudad || 12) + carretera / (a.rendimientoCarretera || 14);
  const trasladoDia = a.trasladoFijo > 0 ? a.trasladoFijo : Math.max(a.trasladoMinimo || 0, litrosTramo * 2 * litro);
  const traslado = Math.round(trasladoDia * dias);
  const peajes = km > (a.peajesDesdeKm || 30)
    ? Math.round(Math.max(0, km - (a.kmCiudadPorTramo || 25)) * (a.peajesPorKm || 0) * 2 / 100) * 100 * dias : 0;
  const comida = (a.comidaPorPersona || 0) * (a.operadores || 0) * dias;

  return { servicios, descuentoPct: desc * 100, extras, traslado, peajes, comida, litro,
    total: servicios + extras + traslado + peajes + comida };
}

/* ================= planilla ================= */

function hojaCotizaciones() {
  const hoja = SpreadsheetApp.openById(AJUSTES.planillaCotizaciones).getSheets()[0];
  if (hoja.getRange(1, COLUMNAS.length).getValue() === "") hoja.getRange(1, 1, 1, COLUMNAS.length).setValues([COLUMNAS]);
  return hoja;
}

function siguienteFolio() {
  return "COT-" + new Date().getFullYear() + "-" + String(hojaCotizaciones().getLastRow()).padStart(4, "0");
}

function guardarEnPlanilla(d) {
  const hoja = hojaCotizaciones();
  const folio = d.folio;
  hoja.appendRow([
    new Date(), d.nombre, d.telefono, d.correo, d.servicios.map(x => x.nombre).join(" + "), d.jornadas,
    d.extras.map(x => x.nombre).join(", "), d.direccion, Math.round(d.km), d.fechaTexto, d.c.total, d.mensaje,
    "Nueva", folio, desgloseTexto(d), d.pdf ? d.pdf.url : ""
  ]);
  return folio;
}

/* ================= PDF ================= */

function crearPdf(d) {
  try {
    const html = HtmlService.createHtmlOutput(htmlPdf(d)).getContent();
    const blob = Utilities.newBlob(html, "text/html", "cotizacion.html").getAs("application/pdf")
      .setName(`${d.folio} - ${d.nombre}.pdf`);
    const raiz = DriveApp.getFolderById(AJUSTES.carpetaPdf);
    const nombreCarpeta = d.nombre.replace(/[\/:*?"|]/g, "").trim() || "Sin nombre";
    const it = raiz.getFoldersByName(nombreCarpeta);
    const carpeta = it.hasNext() ? it.next() : raiz.createFolder(nombreCarpeta);
    const archivo = carpeta.createFile(blob);
    return { blob: archivo.getBlob(), url: archivo.getUrl() };
  } catch (e) {
    console.error("No se pudo crear el PDF", e);
    return null;
  }
}

function logoDataUri() {
  try {
    const b = UrlFetchApp.fetch(AJUSTES.logo).getBlob();
    return "data:" + b.getContentType() + ";base64," + Utilities.base64Encode(b.getBytes());
  } catch (e) { return ""; }
}

function htmlPdf(d) {
  const c = d.c, t = d.titular;
  const hoy = Utilities.formatDate(new Date(), "America/Santiago", "dd-MM-yyyy");
  const vence = Utilities.formatDate(new Date(Date.now() + AJUSTES.validezDias * 864e5), "America/Santiago", "dd-MM-yyyy");
  const logo = logoDataUri();
  const fila = (k, v, fuerte) => `<tr><td class="k">${k}</td><td class="v${fuerte ? " f" : ""}">${v}</td></tr>`;
  const items = [
    ...d.servicios.map(x => fila(x.nombre, `${d.jornadas} jornada${d.jornadas > 1 ? "s" : ""} de 8 h`)),
    ...(c.descuentoPct ? [fila(`Descuento por ${d.servicios.length} servicios`, `${c.descuentoPct}%`)] : []),
    ...d.extras.map(x => fila(x.nombre, "Incluido")),
    fila(c.peajes ? "Traslado, peajes y alimentación del equipo" : "Traslado y alimentación del equipo", "Incluido")
  ].join("");
  return `<html><head><meta charset="utf-8"><style>
    @page { size: A4; margin: 0; }
    body { margin:0; font-family: Helvetica, Arial, sans-serif; color:#1a1a1a; font-size:12px; }
    .top { background:#000; padding:26px 40px; }
    .top img { height:70px; }
    .wrap { padding:28px 40px; }
    h1 { font-size:20px; font-weight:300; letter-spacing:1px; margin:0 0 4px; }
    .muted { color:#777; }
    .cols { width:100%; border-collapse:collapse; margin:18px 0 8px; }
    .cols td { vertical-align:top; width:50%; padding:0 12px 0 0; line-height:1.6; }
    .lbl { font-size:9px; letter-spacing:2px; text-transform:uppercase; color:#888; margin-bottom:4px; }
    table.items { width:100%; border-collapse:collapse; margin-top:18px; }
    table.items th { text-align:left; font-size:9px; letter-spacing:2px; text-transform:uppercase; color:#888; border-bottom:1px solid #1a1a1a; padding:8px 0; }
    td.k { padding:10px 0; border-bottom:1px solid #e3e3e3; }
    td.v { padding:10px 0; border-bottom:1px solid #e3e3e3; text-align:right; white-space:nowrap; }
    .total { margin-top:18px; text-align:right; }
    .total .n { font-size:26px; font-weight:300; }
    .cond { margin-top:26px; padding-top:14px; border-top:1px solid #e3e3e3; line-height:1.7; color:#555; font-size:11px; }
    .foot { margin-top:26px; color:#999; font-size:10px; text-align:center; letter-spacing:1px; }
  </style></head><body>
  <div class="top">${logo ? `<img src="${logo}">` : `<span style="color:#fff;font-size:22px;letter-spacing:4px">MARCO · FPV AERIAL FILMS</span>`}</div>
  <div class="wrap">
    <h1>Cotización ${d.folio}</h1>
    <div class="muted">Emitida el ${hoy} · Válida hasta el ${vence}</div>
    <table class="cols"><tr>
      <td><div class="lbl">De</div>${t.nombre}${t.rut ? `<br>RUT ${t.rut}` : ""}<br>${t.correo}<br>WhatsApp +56 9 3907 0008<br>marcofpv.cl</td>
      <td><div class="lbl">Para</div>${d.nombre}<br>${d.correo}${d.telefono ? `<br>${d.telefono}` : ""}<br><br>
        <div class="lbl">Evento</div>${d.direccion || "Dirección por confirmar"}<br>Fecha: ${d.fechaTexto}</td>
    </tr></table>
    <table class="items"><tr><th>Detalle</th><th style="text-align:right">Cantidad</th></tr>${items}</table>
    <div class="total"><div class="lbl">Total</div><div class="n">${clp(c.total)}</div><div class="muted">Pesos chilenos · No incluye IVA</div></div>
    <div class="cond"><div class="lbl">Condiciones</div>
      • Cotización referencial, válida por ${AJUSTES.validezDias} días. El valor final se confirma al revisar el lugar y los detalles del evento.<br>
      • La fecha queda reservada al confirmar la cotización. Forma de pago a convenir.<br>
      • Si el clima no permite volar con seguridad, el servicio se reprograma sin costo adicional.<br>
      • Entrega del material en 24 a 48 horas por link de descarga.
    </div>
    <div class="foot">MARCO FPV AERIAL FILMS · marcofpv.cl</div>
  </div></body></html>`;
}

function desgloseTexto(d) {
  const c = d.c;
  return `Servicios ${clp(c.servicios)}${c.descuentoPct ? ` (con ${c.descuentoPct}% dcto.)` : ""} · Extras ${clp(c.extras)} · ` +
    `Traslado ${clp(c.traslado)} (${Math.round(d.km)} km ida, bencina ${clp(c.litro)}/l) · Peajes ${clp(c.peajes)} · Alimentación ${clp(c.comida)}`;
}

/* ================= correo al cliente ================= */

function cotizacionAlCliente(d) {
  const c = d.c;
  const filas = [
    ...d.servicios.map(x => [x.nombre, "✓"]),
    ["Jornadas", `${d.jornadas} de 8 h`],
    ...(c.descuentoPct ? [[`Descuento por ${d.servicios.length} servicios`, `${c.descuentoPct}%`]] : []),
    ...d.extras.map(x => [x.nombre, "✓"]),
    [c.peajes ? "Traslado, peajes y alimentación del equipo" : "Traslado y alimentación del equipo", "Incluido"],
    ["Dirección del evento", d.direccion || "Por confirmar"],
    ["Fecha", d.fechaTexto],
    ["Total", `<b style="font-size:18px">${clp(c.total)}</b>`]
  ];
  const wa = "https://wa.me/" + AJUSTES.whatsapp + "?text=" + encodeURIComponent(`Hola Marco, te escribo por la cotización ${d.folio}.`);
  MailApp.sendEmail({
    to: d.correo,
    replyTo: AJUSTES.miCorreo,
    name: AJUSTES.nombreRemitente,
    subject: `Tu cotización ${d.folio} · Marco FPV Aerial Films`,
    attachments: d.pdf ? [d.pdf.blob] : [],
    htmlBody: plantilla(
      `Hola ${d.nombre}, aquí está tu cotización`,
      `Folio <b>${d.folio}</b>. Gracias por considerarnos para tu proyecto.${d.pdf ? " Adjuntamos la cotización en PDF." : ""}`,
      filas,
      `<p style="margin:18px 0 0;color:#8f8f8f;font-size:13px;line-height:1.6">Valor en pesos chilenos, no incluye IVA. ` +
      `Cotización válida por ${AJUSTES.validezDias} días. Te confirmamos el valor final al revisar el lugar y los detalles del evento.</p>
       <p style="margin:24px 0 0"><a href="${wa}" style="display:inline-block;background:#ffffff;color:#000000;text-decoration:none;padding:12px 22px;font-size:12px;letter-spacing:3px;text-transform:uppercase">Reservar por WhatsApp</a></p>
       <p style="margin:14px 0 0;color:#8f8f8f;font-size:13px">También puedes responder este correo.</p>`
    )
  });
}

/* ================= aviso para Marco ================= */

function avisarAMarco(d) {
  const c = d.c;
  const tel = d.telefono.replace(/\D/g, "");
  const wa = tel ? "https://wa.me/" + (tel.startsWith("56") ? tel : "56" + tel) : "";
  const filas = [
    ["Folio", d.folio],
    ["Nombre", d.nombre],
    ["WhatsApp", d.telefono + (wa ? ` · <a href="${wa}" style="color:#ffffff">Abrir chat</a>` : "")],
    ["Correo", d.correo],
    ["Servicios", `${d.servicios.map(x => x.nombre).join(" + ")} · ${d.jornadas} jornada(s)`],
    ["Extras", d.extras.map(x => x.nombre).join(", ") || "—"],
    ["Dirección", d.direccion || "—"],
    ["Fecha del evento", d.fechaTexto],
    ["Total enviado al cliente", `<b>${clp(c.total)}</b>`],
    ["Servicios", clp(c.servicios) + (c.descuentoPct ? ` (con ${c.descuentoPct}% dcto.)` : "")],
    ["Extras", clp(c.extras)],
    ["Traslado", `${clp(c.traslado)} · ${Math.round(d.km)} km ida · bencina ${clp(c.litro)}/l`],
    ["Peajes", clp(c.peajes)],
    ["Alimentación", clp(c.comida)],
    ["Mensaje", d.mensaje || "—"],
    ["PDF", d.pdf ? `<a href="${d.pdf.url}" style="color:#ffffff">Ver cotización en Drive</a>` : "No se pudo generar"]
  ];
  MailApp.sendEmail({
    to: AJUSTES.miCorreo,
    replyTo: d.correo,
    name: "Web MarcoFPV",
    subject: `Nueva cotización ${d.folio}: ${d.nombre} · ${d.fechaTexto} · ${clp(c.total)}`,
    htmlBody: plantilla("Nueva cotización desde la web",
      "El cliente ya recibió su cotización con el total. Responde este correo para escribirle directamente.", filas)
  });
}

/* ================= utilidades ================= */

function plantilla(titulo, intro, filas, pie) {
  const tabla = filas.map(([k, v]) =>
    `<tr><td style="padding:10px 0;border-bottom:1px solid #222;color:#8f8f8f;font-size:13px;width:42%;vertical-align:top">${k}</td>` +
    `<td style="padding:10px 0;border-bottom:1px solid #222;color:#f2f2f2;font-size:14px">${v}</td></tr>`
  ).join("");
  return `
  <div style="background:#000000;padding:32px 16px;font-family:Helvetica,Arial,sans-serif">
    <div style="max-width:560px;margin:0 auto">
      <p style="text-align:center;margin:0 0 28px"><img src="${AJUSTES.logo}" alt="Marco FPV Aerial Films" width="220" style="max-width:220px"></p>
      <div style="background:#0f0f0f;border:1px solid #222;padding:28px">
        <h1 style="margin:0 0 10px;color:#ffffff;font-size:20px;font-weight:300;letter-spacing:.5px">${titulo}</h1>
        <p style="margin:0 0 18px;color:#bdbdbd;font-size:14px;line-height:1.6">${intro}</p>
        <table style="width:100%;border-collapse:collapse">${tabla}</table>
        ${pie || ""}
      </div>
      <p style="text-align:center;margin:22px 0 0;color:#5f5f5f;font-size:11px;letter-spacing:3px;text-transform:uppercase">
        <a href="${AJUSTES.web}" style="color:#8f8f8f;text-decoration:none">marcofpv.cl</a>
      </p>
    </div>
  </div>`;
}

function clp(n) {
  return "$" + Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

function limpiar(v, max) {
  return String(v || "").replace(/[<>]/g, "").trim().slice(0, max);
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Ejecuta esta función una vez desde el editor para autorizar los permisos y probar el envío. */
function probar() {
  doPost({
    parameter: {
      Nombre: "Prueba", Telefono: "939070008", email: AJUSTES.miCorreo,
      Servicios_ids: "video_full", Jornadas: "1", Direccion_evento: "Viña del Mar",
      Km_ida_desde_Maipu: "120", Fecha: "2026-11-15", Mensaje: "Correo de prueba"
    },
    parameters: { "Extras[]": ["vert"] }
  });
}
