// Importar Firebase
import { initializeApp } from "https://www.gstatic.com/firebasejs/11.10.0/firebase-app.js";
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
const provider = new GoogleAuthProvider();
// Siempre deja elegir cuenta de Google, en vez de reusar la última sesión.
provider.setCustomParameters({ prompt: "select_account" });

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

// Como este archivo se carga como <script type="module">, se ejecuta en
// forma diferida (después de parsear el HTML) y en un momento distinto al
// del resto de los scripts de la página. Avisamos con un evento cuando
// window.firebaseAuth ya está listo para usarse, así el script principal
// puede engancharse tanto si llega antes como después de este momento.
window.dispatchEvent(new Event("firebaseReady"));
