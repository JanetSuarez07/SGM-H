import { initializeApp } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-app.js";
import { getAuth, signInWithEmailAndPassword } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-auth.js";

// Configuración secreta de tu proyecto
const firebaseConfig = {
    apiKey: "AIzaSyDUR0GP_8Z48uQQ0XZA86rBl5fqPVmPA68",
    authDomain: "sgm-h-1909c.firebaseapp.com",
    projectId: "sgm-h-1909c",
    storageBucket: "sgm-h-1909c.firebasestorage.app",
    messagingSenderId: "616664914241",
    appId: "1:616664914241:web:828b20032af3a5303ca9c0"
};

// Inicialización
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);

// Evento de clic
document.getElementById("loginBtn").addEventListener("click", () => {
    const email = document.getElementById("email").value;
    const password = document.getElementById("password").value;

    signInWithEmailAndPassword(auth, email, password)
        .then((userCredential) => {
            alert("Acceso concedido.");
            // Aquí luego haremos la redirección al dashboard
        })
        .catch((error) => {
            alert("Error: Verifica tus credenciales.");
        });
});