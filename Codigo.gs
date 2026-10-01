/* =====================================================================
   BEAUTY BY ANTO — Mailing automático
   =====================================================================
   Qué hace este script:

   1. Carrito abandonado: si alguien llega a "Métodos de pago" con el
      carrito cargado y no confirma la compra, a las 48hs le llega un
      mail recordándoselo.

   2. Agradecimiento por compra: NO se manda apenas la clienta toca
      "Confirmar por WhatsApp" (eso solo significa que dijo que iba a
      pagar, no que pagó). Se manda cuando Anto, desde la página de
      estadísticas del sitio, marca el pedido como "pagado" después de
      chequear que el dinero realmente entró — ahí se avisa al instante.
      Además, una vez por día este script revisa si quedó algún pago
      marcado sin su mail de agradecimiento (por ejemplo, por un corte
      de conexión justo en ese momento) y lo manda igual, como red de
      seguridad.

   3. Promos y códigos de descuento: desde el menú "💌 Beauty By Anto"
      de esta planilla, podés mandar un mail a todas las personas
      registradas cuando quieras avisar de un descuento.

   No hace falta tocar código para usarlo día a día: todo se maneja
   desde el menú de la planilla. Las únicas dos cosas que SÍ hay que
   revisar antes de la primera vez están marcadas más abajo con "⚠️".
   ===================================================================== */

// ⚠️ ID del proyecto de Firebase. No debería hacer falta cambiarlo.
const PROJECT_ID = "beauty-by-anto";

const FIRESTORE_BASE = "https://firestore.googleapis.com/v1/projects/" + PROJECT_ID + "/databases/(default)/documents";
const REMINDER_HOURS = 48;
const FROM_NAME = "Beauty By Anto";
const SHOP_URL = "https://beauty-by-anto.web.app/productos.html";

// ⚠️ Tiene que ser EXACTAMENTE igual al MARK_PAID_TOKEN que está en
// metodos-pago.html — es lo que permite que el botón "Marcar pagado"
// de la notificación del celular funcione sin que tengas que loguearte
// en nada.
const MARK_PAID_TOKEN = "OuGoaGhlMLtusLzwOyuUtEK1VFQh3t1A";
const NTFY_TOPIC = "beautybyanto987";

/* ---------------------------------------------------------------------
   Menú de la planilla
   --------------------------------------------------------------------- */
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu("💌 Beauty By Anto")
    .addItem("✉️ Mandarme una prueba primero", "enviarPromoDePrueba")
    .addItem("📧 Enviar promo a todas las registradas", "enviarPromoATodas")
    .addItem("🔁 Probar ahora: recordatorios y agradecimientos", "procesarRecordatoriosYAgradecimientosManual")
    .addSeparator()
    .addItem("✅ Activar envío diario automático", "activarTriggerDiario")
    .addItem("⛔ Desactivar envío diario automático", "desactivarTriggerDiario")
    .addToUi();
}

/* ---------------------------------------------------------------------
   1 y 2. Recordatorios de carrito + red de seguridad de agradecimientos
   (se corre sola todos los días una vez activado el trigger, o se
   puede probar a mano desde el menú)
   --------------------------------------------------------------------- */
function procesarRecordatoriosYAgradecimientos(){
  const ahora = new Date();
  const limiteMs = REMINDER_HOURS * 60 * 60 * 1000;
  let recordatoriosEnviados = 0;
  let agradecimientosEnviados = 0;

  // --- Carritos abandonados (más de 48hs en estado "pendiente") ---
  const pendientesDocs = fsRunQuery_("pedidos", "estado", "EQUAL", { stringValue: "pendiente" });
  pendientesDocs.forEach(doc => {
    const pedido = docToPlano_(doc);
    if(pedido.recordatorioEnviado === true) return;
    if(!pedido.creadoEl) return;
    const creadoEl = pedido.creadoEl instanceof Date ? pedido.creadoEl : new Date(pedido.creadoEl);
    if((ahora - creadoEl) < limiteMs) return; // todavía no pasaron las 48hs

    if(pedido.clienteEmail){
      enviarMailRecordatorio_(pedido);
      recordatoriosEnviados++;
    }
    fsPatchFields_(pedido.path, { recordatorioEnviado: true });
  });

  // --- Red de seguridad: pagos verificados por Anto que por algún motivo
  // no llegaron a avisar al instante (por ejemplo, se cerró la pestaña
  // de estadísticas justo en ese momento) ---
  const pagadosDocs = fsRunQuery_("pedidos", "pagoConfirmado", "EQUAL", { booleanValue: true });
  pagadosDocs.forEach(doc => {
    const pedido = docToPlano_(doc);
    if(pedido.agradecimientoEnviado === true) return;

    if(pedido.clienteEmail){
      enviarMailAgradecimiento_(pedido);
      agradecimientosEnviados++;
    }
    fsPatchFields_(pedido.path, { agradecimientoEnviado: true });
  });

  Logger.log("Recordatorios enviados: " + recordatoriosEnviados + " — Agradecimientos (red de seguridad): " + agradecimientosEnviados);
  return { recordatoriosEnviados, agradecimientosEnviados };
}

function procesarRecordatoriosYAgradecimientosManual(){
  const r = procesarRecordatoriosYAgradecimientos();
  SpreadsheetApp.getUi().alert(
    "Listo ✅\n\nRecordatorios de carrito enviados: " + r.recordatoriosEnviados +
    "\nAgradecimientos enviados (red de seguridad): " + r.agradecimientosEnviados
  );
}

function activarTriggerDiario(){
  const yaExiste = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === "procesarRecordatoriosYAgradecimientos");
  if(yaExiste){
    SpreadsheetApp.getUi().alert("El envío diario ya estaba activado. 👍");
    return;
  }
  ScriptApp.newTrigger("procesarRecordatoriosYAgradecimientos")
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .inTimezone("America/Argentina/Buenos_Aires")
    .create();
  SpreadsheetApp.getUi().alert(
    "Listo ✅ Todos los días, alrededor de las 10hs (hora Argentina), se van a revisar los carritos abandonados y mandar los mails automáticamente."
  );
}

function desactivarTriggerDiario(){
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "procesarRecordatoriosYAgradecimientos");
  if(triggers.length === 0){
    SpreadsheetApp.getUi().alert("El envío diario no estaba activado.");
    return;
  }
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getUi().alert("Envío diario desactivado.");
}

/* ---------------------------------------------------------------------
   Web App (GET): el botón "✅ Marcar pagado" de la notificación del
   celular llega acá. Marca el pedido como pagado, manda el mail de
   agradecimiento, y te avisa por notificación que quedó hecho.
   --------------------------------------------------------------------- */
function doGet(e){
  const accion = e.parameter.accion;

  if(accion === "marcarPagado"){
    const pedidoId = e.parameter.pedidoId;
    const token = e.parameter.token;

    if(token !== MARK_PAID_TOKEN){
      return ContentService.createTextOutput("Token inválido.");
    }
    if(!pedidoId){
      return ContentService.createTextOutput("Falta el pedido.");
    }

    const pedido = fsGetDoc_("pedidos/" + pedidoId);
    if(!pedido){
      return ContentService.createTextOutput("No se encontró ese pedido.");
    }

    if(pedido.pagoConfirmado === true){
      enviarNtfy_("Ese pedido ya estaba marcado como pagado", "");
      return ContentService.createTextOutput("Ya estaba marcado como pagado.");
    }

    fsPatchFields_(pedido.path, { pagoConfirmado: true });

    if(pedido.clienteEmail){
      enviarMailAgradecimiento_(pedido);
    }

    const nombre = pedido.clienteNombre || pedido.clienteEmail || "";
    const total = Number(pedido.total || 0).toLocaleString("es-AR");
    enviarNtfy_("Pago verificado", (nombre ? nombre + " — " : "") + "$" + total);

    return ContentService.createTextOutput("Listo, marcado como pagado.");
  }

  return ContentService.createTextOutput("Acción no reconocida.");
}

function enviarNtfy_(titulo, texto){
  try{
    UrlFetchApp.fetch("https://ntfy.sh/" + NTFY_TOPIC, {
      method: "post",
      headers: { "Title": titulo, "Tags": "white_check_mark", "Priority": "default" },
      payload: texto || "",
      muteHttpExceptions: true
    });
  }catch(err){
    Logger.log("Error mandando notificación ntfy: " + err);
  }
}

/* ---------------------------------------------------------------------
   Web App (POST): recibe el aviso instantáneo cuando Anto marca un
   pedido como pagado desde estadisticas.html.
   --------------------------------------------------------------------- */
function doPost(e){
  try{
    const data = JSON.parse(e.postData.contents);

    if(data.tipo === "agradecimiento"){
      const pedido = {
        clienteEmail: data.clienteEmail,
        clienteNombre: data.clienteNombre,
        items: data.items,
        total: data.total,
        metodoPago: data.metodoPago
      };
      if(pedido.clienteEmail){
        enviarMailAgradecimiento_(pedido);
      }
      if(data.pedidoId){
        fsPatchFields_("pedidos/" + data.pedidoId, { agradecimientoEnviado: true });
      }
    }

    return ContentService.createTextOutput("ok");
  }catch(err){
    Logger.log("Error en doPost: " + err);
    return ContentService.createTextOutput("error");
  }
}

/* ---------------------------------------------------------------------
   3. Promo a todas las personas registradas (manual, desde el menú)
   --------------------------------------------------------------------- */

// Arma el asunto/cuerpo/HTML de un mail de promo a partir del mensaje
// que escribe Anto, para un destinatario puntual. Lo usan tanto el
// envío de prueba como el envío real, así los dos se ven exactamente
// igual.
function construirMailPromo_(nombreDestinatario, mensaje){
  const nombre = nombreDestinatario ? String(nombreDestinatario).split(" ")[0] : "";
  const saludo = nombre ? ("¡Hola, " + nombre + "!") : "¡Hola!";
  const cuerpoPersonalizado = mensaje.replace(/\{nombre\}/g, nombre || "");
  const body = saludo + "\n\n" + cuerpoPersonalizado + "\n\nBeauty By Anto\nLa belleza de encontrarte";
  const htmlBody = emailWrapper_(
    "<p style=\"margin:0 0 14px;font-family:Georgia,serif;font-size:20px;color:#9c0b5f;\">" + saludo + "</p>" +
    parrafosHtml_(cuerpoPersonalizado)
  );
  return { body, htmlBody };
}

// Pide asunto y mensaje por los cuadros de diálogo de siempre. Devuelve
// null si Anto cancela en cualquiera de los dos pasos.
function pedirAsuntoYMensaje_(ui){
  const asuntoResp = ui.prompt("Asunto del mail", "Ej: ¡15% OFF con el código BBAPRIMAVERA! 🌷", ui.ButtonSet.OK_CANCEL);
  if(asuntoResp.getSelectedButton() !== ui.Button.OK) return null;
  const asunto = asuntoResp.getResponseText().trim();
  if(!asunto){ ui.alert("Falta el asunto."); return null; }

  const mensajeResp = ui.prompt(
    "Mensaje",
    "Escribí el mensaje en texto simple (sin HTML). Podés usar {nombre} donde quieras que vaya el nombre de la persona.",
    ui.ButtonSet.OK_CANCEL
  );
  if(mensajeResp.getSelectedButton() !== ui.Button.OK) return null;
  const mensaje = mensajeResp.getResponseText();
  if(!mensaje.trim()){ ui.alert("Falta el mensaje."); return null; }

  return { asunto, mensaje };
}

// Manda el mail de promo SOLO a vos, para ver cómo queda antes de
// mandarlo a todas. No toca la lista de usuarias ni cuenta como envío
// real.
function enviarPromoDePrueba(){
  const ui = SpreadsheetApp.getUi();

  const emailResp = ui.prompt(
    "Mandarme una prueba",
    "¿A qué mail te la mando? (poné el tuyo)",
    ui.ButtonSet.OK_CANCEL
  );
  if(emailResp.getSelectedButton() !== ui.Button.OK) return;
  const emailDestino = emailResp.getResponseText().trim();
  if(!emailDestino){ ui.alert("Falta el mail."); return; }

  const datos = pedirAsuntoYMensaje_(ui);
  if(!datos) return;

  const { body, htmlBody } = construirMailPromo_("Anto", datos.mensaje);

  try{
    MailApp.sendEmail({ to: emailDestino, subject: "[PRUEBA] " + datos.asunto, body, htmlBody, name: FROM_NAME });
    ui.alert("Listo ✅ Te mandé la prueba a " + emailDestino + ".");
  }catch(err){
    Logger.log("Error enviando prueba: " + err);
    ui.alert("No se pudo mandar la prueba. Mirá el registro de ejecuciones para más detalle.");
  }
}

function enviarPromoATodas(){
  const ui = SpreadsheetApp.getUi();

  const datos = pedirAsuntoYMensaje_(ui);
  if(!datos) return;
  const { asunto, mensaje } = datos;

  const usuarios = fsGetAllDocs_("usuarios").map(docToPlano_).filter(u => u.email);

  if(usuarios.length === 0){
    ui.alert("No hay personas registradas todavía.");
    return;
  }

  const confirmacion = ui.alert(
    "Confirmar envío",
    "Se va a mandar este mail a " + usuarios.length + " personas registradas. ¿Confirmás?",
    ui.ButtonSet.YES_NO
  );
  if(confirmacion !== ui.Button.YES) return;

  const cuotaDisponible = MailApp.getRemainingDailyQuota();
  if(cuotaDisponible < usuarios.length){
    ui.alert(
      "Ojo: hoy te quedan " + cuotaDisponible + " mails disponibles (el límite diario de Gmail es 100) " +
      "y querés mandarle a " + usuarios.length + " personas. Se va a mandar solo a las primeras " +
      cuotaDisponible + " — mañana podés correrlo de nuevo para el resto."
    );
  }

  let enviados = 0;
  for(const u of usuarios){
    if(enviados >= cuotaDisponible) break;
    const { body, htmlBody } = construirMailPromo_(u.nombre, mensaje);
    try{
      MailApp.sendEmail({ to: u.email, subject: asunto, body, htmlBody, name: FROM_NAME });
      enviados++;
    }catch(err){
      Logger.log("Error enviando a " + u.email + ": " + err);
    }
  }

  ui.alert("Listo ✅ Se mandó el mail a " + enviados + " personas.");
}

/* ---------------------------------------------------------------------
   Plantillas de mail — diseño de marca (rosa, dorado, el imagotipo del
   sitio) con una versión en texto plano como respaldo para los
   clientes de mail que no muestran HTML.
   --------------------------------------------------------------------- */

// Logo incrustado directo en el mail (en vez de apuntar a una URL del
// sitio): así se ve siempre, sin depender de que esa imagen exista en
// el sitio en ese momento exacto ni de problemas de carga externa.
// Si en algún momento querés cambiar el logo, reemplazá este texto
// larguísimo completo por el de una imagen nueva (te paso el proceso
// si hace falta — no se edita a mano).
const LOGO_BASE64 = "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";
const LOGO_DATA_URI = "data:image/png;base64," + LOGO_BASE64;

function listaItemsTexto_(items){
  if(!items || !items.length) return "";
  return items.map(it =>
    "• " + it.title + " x" + it.qty + " — $" + Number(it.price * it.qty).toLocaleString("es-AR")
  ).join("\n");
}

function listaItemsHtml_(items){
  if(!items || !items.length) return "";
  const filas = items.map(it =>
    "<tr>" +
      "<td style=\"padding:8px 0;border-bottom:1px solid #ffd6e9;font-size:14px;color:#331420;\">" +
        it.title + " <span style=\"color:#8a6673;\">x" + it.qty + "</span>" +
      "</td>" +
      "<td style=\"padding:8px 0;border-bottom:1px solid #ffd6e9;font-size:14px;color:#331420;text-align:right;white-space:nowrap;\">" +
        "$" + Number(it.price * it.qty).toLocaleString("es-AR") +
      "</td>" +
    "</tr>"
  ).join("");
  return "<table role=\"presentation\" width=\"100%\" style=\"border-collapse:collapse;margin:14px 0;\">" + filas + "</table>";
}

// Convierte texto escrito por Anto (con saltos de línea simples) en
// párrafos HTML, para el mail de promo.
function parrafosHtml_(texto){
  return String(texto)
    .split(/\n{2,}/)
    .map(p => "<p style=\"margin:0 0 14px;font-size:15px;line-height:1.6;color:#331420;\">" + p.replace(/\n/g, "<br>") + "</p>")
    .join("");
}

function botonHtml_(texto, url){
  return "<div style=\"text-align:center;margin:22px 0 6px;\">" +
    "<a href=\"" + url + "\" style=\"display:inline-block;background:#ee0f82;color:#ffffff;text-decoration:none;" +
    "padding:13px 28px;border-radius:999px;font-family:Georgia,serif;font-weight:bold;font-size:14.5px;\">" +
    texto + "</a></div>";
}

// Envoltorio con el diseño de la marca: header rosa con el imagotipo,
// tarjeta blanca con el contenido, pie de página con el eslogan.
function emailWrapper_(contenidoHtml){
  return (
    "<!doctype html><html><head><meta charset=\"utf-8\"></head><body style=\"margin:0;padding:0;background:#fff8fb;\">" +
    "<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"background:#fff8fb;padding:28px 12px;\">" +
      "<tr><td align=\"center\">" +
        "<table role=\"presentation\" width=\"100%\" style=\"max-width:480px;background:#ffffff;border-radius:16px;overflow:hidden;border:1px solid #ffd6e9;font-family:Georgia,'Times New Roman',serif;\">" +
          "<tr><td style=\"background:#f6dfe9;padding:24px;text-align:center;border-bottom:2px solid #c9a24b;\">" +
            "<img src=\"" + LOGO_DATA_URI + "\" alt=\"Beauty By Anto\" width=\"260\" style=\"max-width:260px;width:100%;height:auto;display:inline-block;border:0;\">" +
          "</td></tr>" +
          "<tr><td style=\"padding:28px 26px 10px;\">" + contenidoHtml + "</td></tr>" +
          "<tr><td style=\"background:#fff0f6;padding:18px 26px;text-align:center;border-top:1px solid #ffd6e9;\">" +
            "<p style=\"margin:0;color:#9c0b5f;font-size:13px;font-weight:bold;\">Beauty By Anto</p>" +
            "<p style=\"margin:4px 0 0;color:#8a6673;font-size:12px;font-style:italic;\">La belleza de encontrarte</p>" +
          "</td></tr>" +
        "</table>" +
      "</td></tr>" +
    "</table>" +
    "</body></html>"
  );
}

function enviarMailRecordatorio_(pedido){
  if(!pedido.clienteEmail) return;
  const nombre = pedido.clienteNombre ? String(pedido.clienteNombre).split(" ")[0] : "";
  const saludo = nombre ? ("¡Hola, " + nombre + "!") : "¡Hola!";
  const itemsTxt = listaItemsTexto_(pedido.items);
  const total = Number(pedido.total || 0).toLocaleString("es-AR");

  const subject = "Dejaste algo en tu carrito 💕 — Beauty By Anto";

  const body =
    saludo + "\n\n" +
    "Vimos que armaste un pedido en Beauty By Anto y quedó esperándote:\n\n" +
    itemsTxt + "\n\n" +
    "Total: $" + total + "\n\n" +
    "Si todavía lo querés, podés terminarlo acá: " + SHOP_URL + "\n\n" +
    "Cualquier duda, escribinos por WhatsApp, ¡un gusto ayudarte! 🌷\n\n" +
    "Beauty By Anto\nLa belleza de encontrarte";

  const htmlBody = emailWrapper_(
    "<p style=\"margin:0 0 4px;font-family:Georgia,serif;font-size:21px;color:#9c0b5f;\">" + saludo + " 🌷</p>" +
    "<p style=\"margin:0 0 6px;font-size:15px;line-height:1.6;color:#331420;\">Vimos que armaste un pedido y quedó esperándote:</p>" +
    listaItemsHtml_(pedido.items) +
    "<p style=\"margin:10px 0 0;text-align:right;font-size:16px;font-weight:bold;color:#9c0b5f;\">Total: $" + total + "</p>" +
    botonHtml_("Completar mi compra", SHOP_URL) +
    "<p style=\"margin:18px 0 0;font-size:13px;color:#8a6673;\">Cualquier duda, escribinos por WhatsApp, ¡un gusto ayudarte!</p>"
  );

  MailApp.sendEmail({ to: pedido.clienteEmail, subject, body, htmlBody, name: FROM_NAME });
}

function enviarMailAgradecimiento_(pedido){
  if(!pedido.clienteEmail) return;
  const nombre = pedido.clienteNombre ? String(pedido.clienteNombre).split(" ")[0] : "";
  const saludo = nombre ? ("¡Gracias, " + nombre + "!") : "¡Gracias!";
  const itemsTxt = listaItemsTexto_(pedido.items);
  const total = Number(pedido.total || 0).toLocaleString("es-AR");

  const subject = "¡Gracias por tu compra! 💖 — Beauty By Anto";

  const body =
    saludo + "\n\n" +
    "Tu pedido en Beauty By Anto quedó confirmado:\n\n" +
    itemsTxt + "\n\n" +
    "Total: $" + total + "\n" +
    "Método de pago: " + (pedido.metodoPago || "a coordinar") + "\n\n" +
    "En breve nos ponemos en contacto por WhatsApp para coordinar la entrega. ¡Gracias por elegirnos! 🌷\n\n" +
    "Beauty By Anto\nLa belleza de encontrarte";

  const htmlBody = emailWrapper_(
    "<p style=\"margin:0 0 4px;font-family:Georgia,serif;font-size:21px;color:#9c0b5f;\">" + saludo + " 💖</p>" +
    "<p style=\"margin:0 0 6px;font-size:15px;line-height:1.6;color:#331420;\">Tu pedido en Beauty By Anto quedó confirmado:</p>" +
    listaItemsHtml_(pedido.items) +
    "<p style=\"margin:10px 0 0;text-align:right;font-size:16px;font-weight:bold;color:#9c0b5f;\">Total: $" + total + "</p>" +
    "<p style=\"margin:4px 0 0;text-align:right;font-size:13px;color:#8a6673;\">Método de pago: " + (pedido.metodoPago || "a coordinar") + "</p>" +
    "<p style=\"margin:20px 0 0;font-size:14px;line-height:1.6;color:#331420;\">En breve nos ponemos en contacto por WhatsApp para coordinar la entrega. ¡Gracias por elegirnos! 🌷</p>"
  );

  MailApp.sendEmail({ to: pedido.clienteEmail, subject, body, htmlBody, name: FROM_NAME });
}

/* ---------------------------------------------------------------------
   Funciones internas para hablar con Firestore por su API REST, sin
   necesitar ninguna clave: como las reglas de Firestore ya permiten
   lectura y escritura abierta en estas colecciones, alcanza con pedir
   los datos directamente.
   --------------------------------------------------------------------- */
function fsValue_(v){
  if(v == null) return null;
  if("stringValue" in v) return v.stringValue;
  if("integerValue" in v) return Number(v.integerValue);
  if("doubleValue" in v) return v.doubleValue;
  if("booleanValue" in v) return v.booleanValue;
  if("timestampValue" in v) return new Date(v.timestampValue);
  if("nullValue" in v) return null;
  if("arrayValue" in v) return (v.arrayValue.values || []).map(fsValue_);
  if("mapValue" in v) return fsFieldsToObj_(v.mapValue.fields || {});
  return null;
}

function fsFieldsToObj_(fields){
  const out = {};
  Object.keys(fields || {}).forEach(k => { out[k] = fsValue_(fields[k]); });
  return out;
}

function jsToFsValue_(val){
  if(typeof val === "boolean") return { booleanValue: val };
  if(typeof val === "number") return { doubleValue: val };
  if(val instanceof Date) return { timestampValue: val.toISOString() };
  return { stringValue: String(val) };
}

function docToPlano_(doc){
  const path = doc.name.split("/documents/")[1]; // ej: "pedidos/abc123"
  const id = path.split("/").pop();
  return Object.assign({ id, path }, fsFieldsToObj_(doc.fields || {}));
}

function fsRunQuery_(collectionId, fieldPath, op, fsValueWrapped){
  const url = FIRESTORE_BASE + ":runQuery";
  const body = {
    structuredQuery: {
      from: [{ collectionId }],
      where: { fieldFilter: { field: { fieldPath }, op, value: fsValueWrapped } },
      limit: 300
    }
  };
  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  if(resp.getResponseCode() >= 300){
    Logger.log("Error runQuery " + collectionId + ": " + resp.getContentText());
    return [];
  }
  const data = JSON.parse(resp.getContentText());
  return data.filter(r => r.document).map(r => r.document);
}

function fsGetDoc_(path){
  const url = FIRESTORE_BASE + "/" + path;
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if(resp.getResponseCode() >= 300){
    Logger.log("Error GET " + path + ": " + resp.getContentText());
    return null;
  }
  return docToPlano_(JSON.parse(resp.getContentText()));
}

function fsGetAllDocs_(collectionId){
  let docs = [];
  let pageToken = null;
  do{
    let url = FIRESTORE_BASE + "/" + collectionId + "?pageSize=300";
    if(pageToken) url += "&pageToken=" + encodeURIComponent(pageToken);
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if(resp.getResponseCode() >= 300){
      Logger.log("Error GET " + collectionId + ": " + resp.getContentText());
      break;
    }
    const data = JSON.parse(resp.getContentText());
    docs = docs.concat(data.documents || []);
    pageToken = data.nextPageToken || null;
  }while(pageToken);
  return docs;
}

function fsPatchFields_(path, fieldsObj){
  const fields = {};
  Object.keys(fieldsObj).forEach(k => { fields[k] = jsToFsValue_(fieldsObj[k]); });
  const maskParams = Object.keys(fieldsObj).map(k => "updateMask.fieldPaths=" + encodeURIComponent(k)).join("&");
  const url = FIRESTORE_BASE + "/" + path + "?" + maskParams;
  const resp = UrlFetchApp.fetch(url, {
    method: "patch",
    contentType: "application/json",
    payload: JSON.stringify({ fields }),
    muteHttpExceptions: true
  });
  if(resp.getResponseCode() >= 300){
    Logger.log("Error PATCH " + path + ": " + resp.getContentText());
  }
}
