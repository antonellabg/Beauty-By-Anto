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
    .addItem("🔧 Probar conexión (no manda ningún mail)", "probarConexionFirestore")
    .addItem("🔁 Correr ahora: recordatorios y agradecimientos (manda mails reales)", "procesarRecordatoriosYAgradecimientosManual")
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
    // Si ya está pagado, dejó de ser un carrito abandonado aunque el
    // estado nunca haya pasado a "confirmado_whatsapp" (por ejemplo, si
    // Anto lo marcó pagado directo sin que la clienta tocara el botón).
    if(pedido.pagoConfirmado === true) return;
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

// Verifica que la credencial de service account funcione, sin mandar
// ningún mail ni tocar ningún pedido real — solo pide el token y hace
// una lectura de prueba. Usá este botón para confirmar que todo está
// bien conectado antes de cerrar las reglas de Firestore, o después de
// cambiar la clave.
function probarConexionFirestore(){
  const ui = SpreadsheetApp.getUi();

  const token = obtenerTokenServiceAccount_();
  if(!token){
    ui.alert(
      "❌ No se pudo obtener el token.\n\n" +
      "Revisá que SERVICE_ACCOUNT_EMAIL y SERVICE_ACCOUNT_KEY estén bien pegados " +
      "(sin espacios de más, con el private_key completo). Mirá también Extensiones → " +
      "Registro de ejecuciones para ver el error exacto."
    );
    return;
  }

  const docs = fsGetAllDocs_("pedidos");
  ui.alert(
    "✅ Conexión exitosa.\n\n" +
    "Se pudo leer la colección de pedidos: " + docs.length + " pedidos encontrados.\n\n" +
    "No se mandó ningún mail ni se modificó nada — esto fue solo una lectura de prueba."
  );
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
    "<p style=\"margin:0 0 14px;font-family:" + FUENTE_TITULO + ";font-size:20px;color:" + COLOR_TITULO + ";\">" + saludo + "</p>" +
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
// ícono, subí el archivo nuevo al sitio (mismo lugar que el resto de
// las imágenes) y actualizá esta línea con el nombre del archivo.
//
// Esta vez es SOLO el tulipán (el isotipo), chico — el nombre "Beauty
// By Anto" y la frase van como texto real en el encabezado (no como
// parte de una imagen), igual que ya funciona en el pie de página.
const LOGO_URL = "https://beauty-by-anto.web.app/isotipo-mail.png";

// Paleta oscura (vino/bordó) para TODO el mail, no solo el logo. La
// razón: Gmail en la app del celular oscurece automáticamente los
// fondos CLAROS que escribimos en el código (sin que ninguna de las
// técnicas estándar lo evite), pero nunca toca los fondos que YA son
// oscuros. Si el mail entero arranca oscuro, no queda nada para que
// Gmail "corrija" — y se ve igual prenda en modo claro y oscuro.
// Diseño clarito (blanco + rosa pastel), tal como se ve en la compu.
// En el celular con modo oscuro, Gmail lo va a oscurecer solo (es su
// comportamiento automático, no hay forma de evitarlo) — queda
// aceptado así, sin pelearle más a eso.
const COLOR_FONDO = "#ffffff";
const COLOR_TEXTO = "#331420";
const COLOR_TITULO = "#9c0b5f";
const COLOR_MUTED = "#8a6673";
const COLOR_LINEA = "#ffd6e9";
const COLOR_DORADO = "#c9a24b";
const COLOR_HEADER = "#f6dfe9";
const COLOR_HEADER_TEXTO = COLOR_TITULO;

// Tipografías de la marca (las mismas del sitio): Great Vibes para el
// nombre cursiva, DM Serif Display para saludos/títulos, Libre
// Baskerville para el cuerpo del texto. Gmail no carga fuentes de
// Google en la mayoría de los casos, así que esto funciona de verdad
// en los clientes de mail que sí las soportan (Apple Mail, Outlook
// nuevo, etc.) y cae en una tipografía serif parecida en los demás —
// nunca se rompe, solo se ve "la opción B" en los que no las cargan.
const FUENTE_CURSIVA = "'Great Vibes', 'Brush Script MT', cursive";
const FUENTE_TITULO = "'DM Serif Display', Georgia, 'Times New Roman', serif";
const FUENTE_CUERPO = "'Libre Baskerville', Georgia, 'Times New Roman', serif";
const IMPORT_FUENTES =
  "<style>@import url('https://fonts.googleapis.com/css2?family=Great+Vibes&family=DM+Serif+Display&family=Libre+Baskerville:wght@400;700&display=swap');</style>";

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
      "<td style=\"padding:8px 0;border-bottom:1px solid " + COLOR_LINEA + ";font-size:14px;color:" + COLOR_TEXTO + ";\">" +
        it.title + " <span style=\"color:" + COLOR_MUTED + ";\">x" + it.qty + "</span>" +
      "</td>" +
      "<td style=\"padding:8px 0;border-bottom:1px solid " + COLOR_LINEA + ";font-size:14px;color:" + COLOR_TEXTO + ";text-align:right;white-space:nowrap;\">" +
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
    .map(p => "<p style=\"margin:0 0 14px;font-size:15px;line-height:1.6;color:" + COLOR_TEXTO + ";\">" + p + "</p>")
    .join("");
}

function botonHtml_(texto, url){
  return "<div style=\"text-align:center;margin:22px 0 6px;\">" +
    "<a href=\"" + url + "\" style=\"display:inline-block;background:#ee0f82;color:#ffffff;text-decoration:none;" +
    "padding:13px 28px;border-radius:999px;font-family:" + FUENTE_TITULO + ";font-weight:bold;font-size:14.5px;\">" +
    texto + "</a></div>";
}

// Envoltorio con el diseño de la marca: todo en tono vino, con el
// banner del logo arriba y el eslogan abajo — un solo bloque de color
// oscuro de punta a punta, para que no haya ningún fondo claro que
// Gmail pueda "corregir" por su cuenta.
function emailWrapper_(contenidoHtml){
  return (
    "<!doctype html><html><head><meta charset=\"utf-8\">" +
    "<meta name=\"color-scheme\" content=\"light\">" +
    "<meta name=\"supported-color-schemes\" content=\"light\">" +
    IMPORT_FUENTES +
    "</head><body style=\"margin:0;padding:0;background:" + COLOR_FONDO + ";\">" +
    "<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" bgcolor=\"" + COLOR_FONDO + "\" style=\"background:" + COLOR_FONDO + ";padding:28px 12px;\">" +
      "<tr><td align=\"center\">" +
        "<table role=\"presentation\" width=\"100%\" bgcolor=\"" + COLOR_FONDO + "\" style=\"max-width:480px;background:" + COLOR_FONDO + ";border-radius:16px;overflow:hidden;border:1px solid " + COLOR_LINEA + ";font-family:" + FUENTE_CUERPO + ";\">" +
          "<tr><td bgcolor=\"" + COLOR_HEADER + "\" align=\"center\" style=\"background:" + COLOR_HEADER + ";padding:26px 24px 20px;text-align:center;border-bottom:2px solid " + COLOR_DORADO + ";\">" +
            "<table role=\"presentation\" align=\"center\" style=\"margin:0 auto;\"><tr><td align=\"center\" style=\"text-align:center;\">" +
              "<img src=\"" + LOGO_URL + "\" alt=\"\" width=\"56\" style=\"display:block;width:56px;height:auto;border:0;margin:0 auto 10px;\">" +
              "<div style=\"font-family:" + FUENTE_CURSIVA + ";font-size:34px;color:" + COLOR_HEADER_TEXTO + ";line-height:1.2;text-align:center;\">Beauty By Anto</div>" +
              "<div style=\"font-family:" + FUENTE_CUERPO + ";font-size:11px;letter-spacing:1.5px;color:" + COLOR_HEADER_TEXTO + ";margin-top:6px;text-align:center;\">LA BELLEZA DE ENCONTRARTE</div>" +
            "</td></tr></table>" +
          "</td></tr>" +
          "<tr><td bgcolor=\"" + COLOR_FONDO + "\" style=\"background:" + COLOR_FONDO + ";padding:28px 26px 10px;\">" + contenidoHtml + "</td></tr>" +
          "<tr><td bgcolor=\"" + COLOR_FONDO + "\" style=\"background:" + COLOR_FONDO + ";padding:18px 26px;text-align:center;border-top:1px solid " + COLOR_LINEA + ";\">" +
            "<p style=\"margin:0;color:" + COLOR_TITULO + ";font-size:13px;font-weight:bold;\">Beauty By Anto</p>" +
            "<p style=\"margin:4px 0 0;color:" + COLOR_MUTED + ";font-size:12px;font-style:italic;\">La belleza de encontrarte</p>" +
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
    "<p style=\"margin:0 0 4px;font-family:" + FUENTE_TITULO + ";font-size:21px;color:" + COLOR_TITULO + ";\">" + saludo + " 🌷</p>" +
    "<p style=\"margin:0 0 6px;font-size:15px;line-height:1.6;color:" + COLOR_TEXTO + ";\">Vimos que armaste un pedido y quedó esperándote:</p>" +
    listaItemsHtml_(pedido.items) +
    "<p style=\"margin:10px 0 0;text-align:right;font-size:16px;font-weight:bold;color:" + COLOR_TITULO + ";\">Total: $" + total + "</p>" +
    botonHtml_("Completar mi compra", SHOP_URL) +
    "<p style=\"margin:18px 0 0;font-size:13px;color:" + COLOR_MUTED + ";\">Cualquier duda, escribinos por WhatsApp, ¡un gusto ayudarte!</p>"
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
    "<p style=\"margin:0 0 4px;font-family:" + FUENTE_TITULO + ";font-size:21px;color:" + COLOR_TITULO + ";\">" + saludo + " 💖</p>" +
    "<p style=\"margin:0 0 6px;font-size:15px;line-height:1.6;color:" + COLOR_TEXTO + ";\">Tu pedido en Beauty By Anto quedó confirmado:</p>" +
    listaItemsHtml_(pedido.items) +
    "<p style=\"margin:10px 0 0;text-align:right;font-size:16px;font-weight:bold;color:" + COLOR_TITULO + ";\">Total: $" + total + "</p>" +
    "<p style=\"margin:4px 0 0;text-align:right;font-size:13px;color:" + COLOR_MUTED + ";\">Método de pago: " + (pedido.metodoPago || "a coordinar") + "</p>" +
    "<p style=\"margin:20px 0 0;font-size:14px;line-height:1.6;color:" + COLOR_TEXTO + ";\">En breve nos ponemos en contacto por WhatsApp para coordinar la entrega. ¡Gracias por elegirnos! 🌷</p>"
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

/* ---------------------------------------------------------------------
   Autenticación como "cuenta de servicio" (service account).
   ---------------------------------------------------------------------
   Desde que Firestore dejó de ser público, este script ya no puede leer
   ni escribir pedidos sin credenciales — necesita autenticarse como un
   "robot" con permiso especial, aparte de las dos cuentas humanas
   admin. Completá estas dos constantes con los datos del archivo que
   bajás de Firebase Console (ver guía de Anto):

   ⚠️ SERVICE_ACCOUNT_EMAIL = el valor "client_email" del archivo JSON
   ⚠️ SERVICE_ACCOUNT_KEY   = el valor "private_key" del archivo JSON,
      tal cual (con los \n adentro, entre comillas) */
const SERVICE_ACCOUNT_EMAIL = "PEGA_ACA_EL_client_email_DEL_JSON";
const SERVICE_ACCOUNT_KEY = "PEGA_ACA_EL_private_key_DEL_JSON";

function obtenerTokenServiceAccount_(){
  const cache = CacheService.getScriptCache();
  const cacheado = cache.get("fs_token");
  if(cacheado) return cacheado;

  const ahora = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT" };
  const claimSet = {
    iss: SERVICE_ACCOUNT_EMAIL,
    scope: "https://www.googleapis.com/auth/datastore",
    aud: "https://oauth2.googleapis.com/token",
    exp: ahora + 3600,
    iat: ahora
  };

  const base64Header = Utilities.base64EncodeWebSafe(JSON.stringify(header)).replace(/=+$/, "");
  const base64Claim = Utilities.base64EncodeWebSafe(JSON.stringify(claimSet)).replace(/=+$/, "");
  const entrada = base64Header + "." + base64Claim;
  const firmaBytes = Utilities.computeRsaSha256Signature(entrada, SERVICE_ACCOUNT_KEY);
  const firma = Utilities.base64EncodeWebSafe(firmaBytes).replace(/=+$/, "");
  const jwt = entrada + "." + firma;

  const resp = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
    method: "post",
    payload: {
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt
    },
    muteHttpExceptions: true
  });

  if(resp.getResponseCode() >= 300){
    Logger.log("Error obteniendo token de service account: " + resp.getContentText());
    return null;
  }

  const token = JSON.parse(resp.getContentText()).access_token;
  cache.put("fs_token", token, 3300); // ~55 min, por debajo de la hora real de validez
  return token;
}

function authHeaders_(){
  const token = obtenerTokenServiceAccount_();
  return token ? { "Authorization": "Bearer " + token } : {};
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
    headers: authHeaders_(),
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
  const resp = UrlFetchApp.fetch(url, { headers: authHeaders_(), muteHttpExceptions: true });
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
    const resp = UrlFetchApp.fetch(url, { headers: authHeaders_(), muteHttpExceptions: true });
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
    headers: authHeaders_(),
    payload: JSON.stringify({ fields }),
    muteHttpExceptions: true
  });
  if(resp.getResponseCode() >= 300){
    Logger.log("Error PATCH " + path + ": " + resp.getContentText());
  }
}
