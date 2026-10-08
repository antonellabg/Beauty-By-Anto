// Importar Firebase
import { initializeApp } from "https://www.gstatic.com/firebasejs/11.10.0/firebase-app.js";
import {
  getAnalytics,
  isSupported as analyticsIsSupported,
  logEvent
} from "https://www.gstatic.com/firebasejs/11.10.0/firebase-analytics.js";
import {
  getAuth,
  GoogleAuthProvider,
  signInWithPopup,
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  signOut,
  onAuthStateChanged,
  updateProfile,
  sendPasswordResetEmail
} from "https://www.gstatic.com/firebasejs/11.10.0/firebase-auth.js";
import {
  getFirestore,
  collection,
  addDoc,
  updateDoc,
  deleteDoc,
  doc,
  getDoc,
  setDoc,
  getDocs,
  query,
  where,
  increment,
  serverTimestamp,
  onSnapshot
} from "https://www.gstatic.com/firebasejs/11.10.0/firebase-firestore.js";

// Configuración de tu proyecto
const firebaseConfig = {
  apiKey: "AIzaSyABNrotV9AvTLrkNrCiqkk18iVmxU4DOZA",
  authDomain: "beauty-by-anto.firebaseapp.com",
  projectId: "beauty-by-anto",
  storageBucket: "beauty-by-anto.firebasestorage.app",
  messagingSenderId: "199971429084",
  appId: "1:199971429084:web:24b4e2c91bf8b51592f219",
  measurementId: "G-NDSKDV8N1G"
};

// Inicializar Firebase
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
const provider = new GoogleAuthProvider();
// Siempre deja elegir cuenta de Google, en vez de reusar la última sesión.
provider.setCustomParameters({ prompt: "select_account" });

// Google Analytics (GA4): mide visitas a cada página de forma automática
// apenas se inicializa (no hace falta código extra por página), y lo
// usamos también para contar inicios de sesión y registros a mano con
// trackEvent(). analyticsIsSupported() chequea que el navegador lo
// soporte (falla, por ejemplo, si la clienta bloquea cookies/trackers),
// así que si no está disponible, trackEvent() simplemente no hace nada
// en vez de romper la página.
let analytics = null;
analyticsIsSupported().then(ok => {
  if(ok) analytics = getAnalytics(app);
}).catch(() => { /* navegador sin soporte: seguimos sin analytics */ });

window.trackEvent = function(nombreEvento, params){
  if(!analytics) return;
  try{
    logEvent(analytics, nombreEvento, params || {});
  }catch(err){
    console.error("No se pudo registrar el evento de Analytics:", err);
  }
};

// Hacer disponibles las funciones para el resto de la página
window.firebaseAuth = {
  auth,
  provider,
  signInWithPopup,
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  signOut,
  onAuthStateChanged,
  updateProfile,
  sendPasswordResetEmail
};

// Firestore: se usa para guardar un registro de cada pedido apenas la
// clienta llega a "Métodos de pago" con productos en el carrito, ANTES
// de que mande (o no) el mensaje de confirmación por WhatsApp. Así Anto
// nunca pierde el rastro de un pedido, aunque la clienta cierre la
// pestaña antes de confirmar.
window.firebaseDb = {
  db,
  collection,
  addDoc,
  updateDoc,
  deleteDoc,
  doc,
  getDoc,
  setDoc,
  getDocs,
  query,
  where,
  increment,
  serverTimestamp,
  onSnapshot
};

// Stock en vivo: vive en la colección "stock" de Firestore (un documento
// por producto, con el id del catálogo como id del documento y un campo
// "cantidad"). Esto permite que Anto lo suba o baje desde stock.html sin
// tener que subir un products-data.js nuevo a GitHub cada vez. Acá lo
// mezclamos con el catálogo estático apenas carga cualquier página que
// muestre productos — mutamos los objetos de window.PRODUCTS en el lugar
// (no reemplazamos el array) para que el resto del código, que ya tiene
// referencias a esos mismos objetos, vea el cambio sin tocar nada más.
(async function aplicarStockEnVivo(){
  if(!window.PRODUCTS) return;
  try{
    const snap = await getDocs(collection(db, "stock"));
    snap.forEach(docSnap => {
      const producto = window.PRODUCTS.find(p => String(p.id) === docSnap.id);
      const datos = docSnap.data();
      if(producto && typeof datos.cantidad === "number"){
        producto.stock = datos.cantidad;
      }
    });
  }catch(err){
    // Si falla (sin conexión, por ejemplo), seguimos con el stock del
    // catálogo estático — no es ideal, pero no rompe la página.
    console.error("No se pudo cargar el stock en vivo:", err);
  }
  window.dispatchEvent(new Event("stockReady"));
})();

// Si quedó guardado el id de un pedido (se guarda apenas se crea, en
// metodos-pago.html), chequeamos acá — en CUALQUIER página del sitio,
// apenas carga — si ya se marcó como pagado. Si es así, vaciamos el
// carrito recién en este momento: es la única forma de "avisarle" al
// navegador de la clienta que Anto confirmó el pago, ya que eso pasa
// en un dispositivo totalmente distinto (el de Anto), en otro momento.
// Cartel flotante de "pago confirmado" — se arma con estilos propios en
// JS (no depende del CSS de cada página) para poder aparecer en
// cualquiera de ellas sin tener que tocar los 7 archivos HTML.
function mostrarCartelPagoConfirmado(){
  const el = document.createElement("div");
  el.setAttribute("role", "status");
  el.style.cssText = [
    "position:fixed", "left:50%", "bottom:18px", "transform:translateX(-50%)",
    "z-index:99999", "max-width:min(92vw,420px)", "background:#ffffff",
    "border:1px solid #ffd6e9", "border-radius:16px",
    "box-shadow:0 14px 40px rgba(238,15,130,0.18)", "padding:16px 44px 16px 18px",
    "font-family:Georgia,serif", "font-size:14px", "line-height:1.5",
    "color:#331420"
  ].join(";");
  el.innerHTML =
    '<strong style="color:#9c0b5f;">¡Confirmamos tu pago! 💖</strong><br>' +
    'Gracias por tu compra — ya estamos coordinando tu entrega por WhatsApp.';
  const closeBtn = document.createElement("button");
  closeBtn.textContent = "✕";
  closeBtn.setAttribute("aria-label", "Cerrar");
  closeBtn.style.cssText = [
    "position:absolute", "top:10px", "right:12px", "border:0", "background:transparent",
    "color:#9c0b5f", "font-size:15px", "cursor:pointer", "line-height:1"
  ].join(";");
  closeBtn.onclick = () => el.remove();
  el.appendChild(closeBtn);
  document.body.appendChild(el);
  setTimeout(() => el.remove(), 15000);
}

// Si quedó guardado el id de un pedido (se guarda apenas se crea, en
// metodos-pago.html), chequeamos acá — en CUALQUIER página del sitio,
// apenas carga, sea cuando sea que la clienta vuelva — si ya se marcó
// como pagado. No hace falta que haya dejado ninguna pestaña abierta:
// esto se fija de nuevo la próxima vez que abra el sitio, pase lo que
// pase mientras tanto.
(async function revisarPedidoActivo(){
  const pedidoId = localStorage.getItem("pedidoActivo");
  if(!pedidoId) return;
  try{
    const snap = await getDoc(doc(db, "pedidos", pedidoId));
    if(snap.exists() && snap.data().pagoConfirmado === true){
      localStorage.removeItem("cart");
      localStorage.removeItem("pedidoActivo");
      // Esperamos a que el body exista antes de insertar el cartel,
      // por si este script corre antes de que termine de parsear el HTML.
      if(document.body) mostrarCartelPagoConfirmado();
      else document.addEventListener("DOMContentLoaded", mostrarCartelPagoConfirmado, { once:true });
    }
  }catch(err){
    // Si falla (por ejemplo, sin conexión en ese momento), no pasa
    // nada grave — simplemente se vuelve a intentar la próxima vez
    // que cargue una página del sitio.
    console.error("No se pudo revisar el estado del pedido activo:", err);
  }
})();

// ---------------------------------------------------------------------
// Estadísticas propias (Firestore), además de lo que ya mide Google
// Analytics. Se guardan agrupadas por día para poder filtrarlas en
// estadisticas.html igual que los pedidos ("Últimos 7/30 días").
// ---------------------------------------------------------------------

// Fecha de hoy en formato YYYY-MM-DD, en horario de Argentina (así los
// días coinciden con cuando Anto revisa el panel, sin importar en qué
// zona horaria esté el navegador de la clienta).
function fechaHoyAR_(){
  return new Date().toLocaleDateString("en-CA", { timeZone: "America/Argentina/Buenos_Aires" });
}

// Cuenta una visita a la web. Se llama una sola vez por pestaña/sesión
// (usamos sessionStorage para no inflar el número cada vez que la misma
// persona navega entre páginas o recarga) y queda agrupada por día.
function registrarVisita(){
  try{
    if(sessionStorage.getItem("bbaVisitaContada")) return;
    sessionStorage.setItem("bbaVisitaContada", "1");
  }catch(e){ /* sin sessionStorage disponible: seguimos igual, sin el límite */ }
  setDoc(doc(db, "visitasDiarias", fechaHoyAR_()), { cantidad: increment(1) }, { merge: true })
    .catch(err => console.error("No se pudo registrar la visita:", err));
}
window.registrarVisita = registrarVisita;

// Cuenta un "lead": alguien que inició sesión o se registró. Se llama
// desde los mismos puntos donde ya se llama a trackEvent("login"/"sign_up"),
// así queda también en nuestras propias estadísticas y no solo en Google
// Analytics.
function registrarLead(tipo){
  setDoc(doc(db, "leadsDiarios", fechaHoyAR_()), { cantidad: increment(1) }, { merge: true })
    .catch(err => console.error("No se pudo registrar el lead:", err));
}
window.registrarLead = registrarLead;

// Cuenta una vista de la ficha de un producto (producto.html), para
// poder ver cuáles son los que más interés generan. Se guarda un
// documento por producto con el total acumulado.
function registrarClickProducto(id, titulo){
  if(id === undefined || id === null) return;
  setDoc(doc(db, "clicksProductos", String(id)), {
    cantidad: increment(1),
    titulo: titulo || ""
  }, { merge: true }).catch(err => console.error("No se pudo registrar el click del producto:", err));
}
window.registrarClickProducto = registrarClickProducto;

registrarVisita();

// Como este archivo se carga como <script type="module">, se ejecuta en
// forma diferida (después de parsear el HTML) y en un momento distinto al
// del resto de los scripts de la página. Avisamos con un evento cuando
// window.firebaseAuth ya está listo para usarse, así el script principal
// puede engancharse tanto si llega antes como después de este momento.
window.dispatchEvent(new Event("firebaseReady"));
