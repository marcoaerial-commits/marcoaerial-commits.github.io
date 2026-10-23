/**
 * Marco FPV Aerial Films · Receptor de cotizaciones
 *
 * Recibe las cotizaciones de https://marcoaerial-commits.github.io,
 * las guarda en la planilla "Cotizaciones MarcoFPV" y envía dos correos
 * desde tu Gmail: uno de aviso para ti y uno de respaldo para el cliente.
 *
 * Instalación: ver README.md del repositorio (sección "Correos desde Gmail").
 */

const AJUSTES = {
  planillaId: "19zwupJG8FBwBzVkBFGLaVrUqJOxZJXb9XgH63xm5M1g",
  hoja: "Cotizaciones MarcoFPV",
  miCorreo: "marcoaerial@gmail.com",
  nombreRemitente: "Marco FPV Aerial Films",
  whatsapp: "56939070008",
  web: "https://marcoaerial-commits.github.io",
  logo: "https://marcoaerial-commits.github.io/assets/logo.png"
};

function doPost(e) {
  try {
    const p = e.parameter || {};
    const extras = (e.parameters && e.parameters["Extras[]"]) || [];

    // Trampa anti-spam: los bots rellenan este campo oculto
    if (p._gotcha) return responder({ success: true });

    const d = {
      nombre: limpiar(p.Nombre, 80),
      telefono: limpiar(p.Telefono, 30),
      correo: limpiar(p.email, 120),
      servicio: limpiar(p.Servicio, 80),
      jornadas: limpiar(p.Jornadas, 5),
      extras: extras.map(x => limpiar(x, 80)).join(", "),
      direccion: limpiar(p.Direccion_evento, 200),
      km: limpiar(p.Km_ida_desde_Maipu, 8),
      fecha: limpiar(p.Fecha, 10),
      total: limpiar(p.Total_estimado, 20),
      mensaje: limpiar(p.Mensaje, 2000)
    };

    if (!d.nombre || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(d.correo)) {
      return responder({ success: false, message: "Faltan el nombre o un correo válido." });
    }
    d.fechaTexto = d.fecha ? d.fecha.split("-").reverse().join("-") : "Por definir";

    guardarEnPlanilla(d);
    avisarAMarco(d);
    respaldoAlCliente(d);

    return responder({ success: true });
  } catch (err) {
    console.error(err);
    return responder({ success: false, message: String(err) });
  }
}

function doGet() {
  return responder({ ok: true, servicio: "Cotizaciones MarcoFPV" });
}

/* ---------- planilla ---------- */

function guardarEnPlanilla(d) {
  const libro = SpreadsheetApp.openById(AJUSTES.planillaId);
  const hoja = libro.getSheetByName(AJUSTES.hoja) || libro.getSheets()[0];
  hoja.appendRow([
    new Date(), d.nombre, d.telefono, d.correo, d.servicio, d.jornadas,
    d.extras, d.direccion, d.km, d.fechaTexto, d.total, d.mensaje, "Nueva"
  ]);
}

/* ---------- correo para Marco ---------- */

function avisarAMarco(d) {
  const tel = d.telefono.replace(/\D/g, "");
  const wa = tel ? "https://wa.me/" + (tel.startsWith("56") ? tel : "56" + tel) : "";
  const filas = [
    ["Nombre", d.nombre],
    ["WhatsApp", d.telefono + (wa ? ` · <a href="${wa}">Abrir chat</a>` : "")],
    ["Correo", d.correo],
    ["Servicio", `${d.servicio} · ${d.jornadas} jornada(s)`],
    ["Extras", d.extras || "—"],
    ["Dirección", d.direccion || "—"],
    ["Km ida desde Maipú", d.km || "—"],
    ["Fecha del evento", d.fechaTexto],
    ["Total estimado", d.total],
    ["Mensaje", d.mensaje || "—"]
  ];
  MailApp.sendEmail({
    to: AJUSTES.miCorreo,
    replyTo: d.correo,
    name: "Web MarcoFPV",
    subject: `Nueva cotización: ${d.nombre} · ${d.fechaTexto} · ${d.total}`,
    htmlBody: plantilla(
      "Nueva cotización desde la web",
      "Responde este correo para escribirle directamente al cliente.",
      filas
    )
  });
}

/* ---------- correo de respaldo para el cliente ---------- */

function respaldoAlCliente(d) {
  const filas = [
    ["Servicio", `${d.servicio} · ${d.jornadas} jornada(s)`],
    ["Extras", d.extras || "Sin extras"],
    ["Dirección del evento", d.direccion || "Por confirmar"],
    ["Fecha", d.fechaTexto],
    ["Total estimado", `<b>${d.total}</b>`]
  ];
  const wa = "https://wa.me/" + AJUSTES.whatsapp;
  MailApp.sendEmail({
    to: d.correo,
    replyTo: AJUSTES.miCorreo,
    name: AJUSTES.nombreRemitente,
    subject: "Recibimos tu solicitud de cotización · Marco FPV Aerial Films",
    htmlBody: plantilla(
      `Hola ${d.nombre}, gracias por cotizar con nosotros`,
      "Recibimos tu solicitud y te contactaremos pronto para confirmar los detalles. " +
      "Este es el resumen de lo que nos enviaste:",
      filas,
      `<p style="margin:18px 0 0;color:#8f8f8f;font-size:13px">El total es un valor referencial y no incluye IVA. ` +
      `Te confirmaremos el precio final al revisar el lugar y los detalles del evento.</p>
       <p style="margin:24px 0 0"><a href="${wa}" style="display:inline-block;background:#ffffff;color:#000000;text-decoration:none;padding:12px 22px;font-size:12px;letter-spacing:3px;text-transform:uppercase">Escríbenos por WhatsApp</a></p>
       <p style="margin:14px 0 0;color:#8f8f8f;font-size:13px">También puedes responder este correo.</p>`
    )
  });
}

/* ---------- utilidades ---------- */

function plantilla(titulo, intro, filas, pie) {
  const tabla = filas.map(([k, v]) =>
    `<tr><td style="padding:10px 0;border-bottom:1px solid #222;color:#8f8f8f;font-size:13px;width:38%;vertical-align:top">${k}</td>` +
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
        <a href="${AJUSTES.web}" style="color:#8f8f8f;text-decoration:none">marcoaerial-commits.github.io</a>
      </p>
    </div>
  </div>`;
}

function limpiar(v, max) {
  return String(v || "").replace(/[<>]/g, "").trim().slice(0, max);
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Ejecuta esta función una vez desde el editor para autorizar los permisos y probar el envío. */
function probar() {
  doPost({
    parameter: {
      Nombre: "Prueba", Telefono: "939070008", email: AJUSTES.miCorreo,
      Servicio: "Filmación aérea / FPV", Jornadas: "1", Direccion_evento: "Maipú",
      Km_ida_desde_Maipu: "0", Fecha: "2026-11-01", Total_estimado: "$220.000", Mensaje: "Correo de prueba"
    },
    parameters: { "Extras[]": [] }
  });
}
