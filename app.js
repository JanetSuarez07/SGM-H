import { initializeApp } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-app.js";
import { getAuth, signInWithEmailAndPassword } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-auth.js";
import { getFirestore, collection, addDoc } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-firestore.js";

const firebaseConfig = { /* ... tu configuración ... */ };
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

// --- LÓGICA PARA EL LOGIN (Solo corre si existe el botón loginBtn) ---
const loginBtn = document.getElementById("loginBtn");
if (loginBtn) {
    loginBtn.addEventListener("click", () => {
        const email = document.getElementById("email").value;
        const password = document.getElementById("password").value;
        signInWithEmailAndPassword(auth, email, password)
            .then(() => window.location.href = "dashboard.html")
            .catch(() => alert("Error: Verifica tus credenciales."));
    });
}

// --- LÓGICA PARA EL REGISTRO (Solo corre si estamos en registro.html) ---
if (document.getElementById("guardarBtn")) {
    
    // 1. Autocompletar al cargar
    window.onload = () => {
        const ahora = new Date();
        document.getElementById("fecha").value = ahora.toLocaleDateString('es-MX');
        document.getElementById("hora").value = ahora.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });
        document.getElementById("folio").value = "IBM-" + Math.floor(Math.random() * 10000);
    };

    // 2. Sincronizar solicitante con "Recibe Trabajo"
    document.getElementById("solicitante_nombre").addEventListener("input", (e) => {
        document.getElementById("recibe_trabajo").value = e.target.value;
    });

    // 3. Guardar en Firestore
    document.getElementById("guardarBtn").addEventListener("click", async () => {
        const ordenData = {
            folio: document.getElementById("folio").value,
            fecha: document.getElementById("fecha").value,
            solicitante: document.getElementById("solicitante_nombre").value,
            area: document.getElementById("area").value,
            descripcion: document.getElementById("descripcion_problema").value,
            recibe_trabajo: document.getElementById("recibe_trabajo").value
            // Agrega aquí el resto de los campos...
        };

        try {
            await addDoc(collection(db, "ordenes"), ordenData);
            alert("Orden guardada exitosamente");
        } catch (e) {
            console.error("Error al guardar: ", e);
        }
    });
}