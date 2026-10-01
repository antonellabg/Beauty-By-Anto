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
  doc,
  getDoc,
  setDoc,
  getDocs,
  query,
  where,
  increment,
  serverTimestamp
};

// Como este archivo se carga como <script type="module">, se ejecuta en
// forma diferida (después de parsear el HTML) y en un momento distinto al
// del resto de los scripts de la página. Avisamos con un evento cuando
// window.firebaseAuth ya está listo para usarse, así el script principal
// puede engancharse tanto si llega antes como después de este momento.
window.dispatchEvent(new Event("firebaseReady"));
