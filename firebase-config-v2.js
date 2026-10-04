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
  serverTimestamp
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
  serverTimestamp
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
(async function revisarPedidoActivo(){
  const pedidoId = localStorage.getItem("pedidoActivo");
  if(!pedidoId) return;
  try{
    const snap = await getDoc(doc(db, "pedidos", pedidoId));
    if(snap.exists() && snap.data().pagoConfirmado === true){
      localStorage.removeItem("cart");
      localStorage.removeItem("pedidoActivo");
    }
  }catch(err){
    // Si falla (por ejemplo, sin conexión en ese momento), no pasa
    // nada grave — simplemente se vuelve a intentar la próxima vez
    // que cargue una página del sitio.
    console.error("No se pudo revisar el estado del pedido activo:", err);
  }
})();

// Como este archivo se carga como <script type="module">, se ejecuta en
// forma diferida (después de parsear el HTML) y en un momento distinto al
// del resto de los scripts de la página. Avisamos con un evento cuando
// window.firebaseAuth ya está listo para usarse, así el script principal
// puede engancharse tanto si llega antes como después de este momento.
window.dispatchEvent(new Event("firebaseReady"));
