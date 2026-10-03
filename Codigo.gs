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

  prepararHojaPromo_();
}

// Prepara una hoja llamada "Promo" con dos celdas fijas para escribir
// el asunto y el mensaje de la próxima promo — ahí SÍ podés escribir
// varios renglones (Alt+Enter dentro de la celda B2 para bajar de
// línea), cosa que el cuadro de diálogo de Apps Script no permite.
// Si la hoja ya existe, no la toca.
function prepararHojaPromo_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName("Promo");
  if(hoja) return;

  hoja = ss.insertSheet("Promo");
  hoja.getRange("A1").setValue("Asunto:");
  hoja.getRange("B1").setValue("¡15% OFF con el código BBAPRIMAVERA! 🌷");
  hoja.getRange("A2").setValue("Mensaje:");
  hoja.getRange("B2").setValue("Escribí acá el mensaje. Para bajar de línea dentro de esta celda usá Alt+Enter (no Enter solo). Podés usar {nombre} donde quieras que vaya el nombre de la persona.");
  hoja.setColumnWidth(1, 90);
  hoja.setColumnWidth(2, 520);
  hoja.getRange("B2").setWrap(true);
  hoja.setRowHeight(2, 140);
  hoja.getRange("A1:A2").setFontWeight("bold");
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

// Lee el asunto y el mensaje de la hoja "Promo" (celdas B1 y B2), en
// vez de un cuadro de diálogo — así sí se pueden escribir varios
// renglones (con Alt+Enter dentro de la celda), cosa que el cuadro de
// diálogo de Apps Script no deja hacer. Devuelve null si falta algo,
// después de avisar con un cartel.
function pedirAsuntoYMensaje_(ui){
  prepararHojaPromo_();
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Promo");

  const asunto = String(hoja.getRange("B1").getValue() || "").trim();
  const mensaje = String(hoja.getRange("B2").getValue() || "").trim();

  if(!asunto){
    ui.alert("Falta el asunto. Escribilo en la hoja \"Promo\", celda B1.");
    return null;
  }
  if(!mensaje){
    ui.alert("Falta el mensaje. Escribilo en la hoja \"Promo\", celda B2.");
    return null;
  }

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

// Logo alojado en el sitio. OJO: tiene que ser una imagen que ya
// esté subida y publicada en beauty-by-anto.web.app — Gmail (web y
// apps) NO muestra imágenes incrustadas directo en el código del mail,
// solo las que apunten a una URL real. Si en algún momento cambiás el
// logo, subí el archivo nuevo al sitio (mismo lugar que el resto de
// las imágenes) y actualizá esta línea con el nombre del archivo.
const LOGO_URL = "https://beauty-by-anto.web.app/logo-mail-v3.png";

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

// Convierte texto escrito por Anto en párrafos HTML, para el mail de
// promo. Funciona tanto si dejó una línea en blanco entre ideas (Enter
// dos veces) como si solo apretó Enter una vez por idea — cualquier
// salto de línea no vacío se trata como el final de un párrafo.
function parrafosHtml_(texto){
  return String(texto)
    .split(/\n+/)
    .map(p => p.trim())
    .filter(p => p.length > 0)
    .map(p => "<p style=\"margin:0 0 14px;font-size:15px;line-height:1.6;color:#331420;\">" + p + "</p>")
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
            "<img src=\"" + LOGO_URL + "\" alt=\"Beauty By Anto\" width=\"260\" style=\"max-width:260px;width:100%;height:auto;display:inline-block;border:0;\">" +
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
