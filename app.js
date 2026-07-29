import { initializeApp } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-app.js";
import { getAuth, signInWithEmailAndPassword, signOut } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-auth.js";
import { getFirestore, collection, addDoc, getDocs, query, orderBy } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-firestore.js";

const firebaseConfig = {
  apiKey: "AIzaSyDUR0GP_8Z48uQQ0XZA86rBl5fqPVmPA68",
  authDomain: "sgm-h-1909c.firebaseapp.com",
  projectId: "sgm-h-1909c",
  storageBucket: "sgm-h-1909c.firebasestorage.app",
  messagingSenderId: "616664914241",
  appId: "1:616664914241:web:828b20032af3a5303ca9c0",
  measurementId: "G-MBCL36RGDS"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

// --- EL LOGIN ---
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

// ---  CERRAR SESIÓN ---
// La hacemos accesible al HTML mediante 'window'
window.cerrarSesion = () => {
    signOut(auth).then(() => {
        window.location.href = "index.html";
    }).catch((error) => console.error("Error al cerrar sesión:", error));
};

// --EL REGISTRO  ---
const guardarBtn = document.getElementById("guardarBtn");
if (guardarBtn) {
    // Autocompletar al cargar la página
    document.addEventListener("DOMContentLoaded", () => {
        const ahora = new Date();
        document.getElementById("fecha").value = ahora.toLocaleDateString('es-MX');
        document.getElementById("hora").value = ahora.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });
        document.getElementById("folio").value = "IBM-";
    });

    // Sincronizar solicitante
    const solInput = document.getElementById("solicitante_nombre");
    if (solInput) {
        solInput.addEventListener("input", (e) => {
            document.getElementById("recibe_trabajo").value = e.target.value;
        });
    }

    // Guardar en Firestore
    guardarBtn.addEventListener("click", async () => {
        const estadoAntesSeleccionado = document.querySelector('input[name="estado_antes"]:checked');
        const tipoMttSeleccionado = document.querySelector('input[name="tipo_mtt"]:checked');

        const ordenData = {
            // Información Automática
            folio: document.getElementById("folio")?.value ?? "",
            fecha: document.getElementById("fecha")?.value ?? "",
            hora: document.getElementById("hora")?.value ?? "",

            // Datos de la Solicitud
            solicitante: document.getElementById("solicitante_nombre")?.value ?? "",
            area: document.getElementById("area")?.value ?? "",
            descripcion: document.getElementById("descripcion_problema")?.value ?? "",
            trabajo_realizado: document.getElementById("trabajo_realizado")?.value ?? "",
            refacciones_usadas: document.getElementById("refacciones_usadas")?.value ?? "",

            // Estado y Mantenimiento (extraído de los radios)
            estado_antes: estadoAntesSeleccionado ? estadoAntesSeleccionado.value : "",
            tipo_mtt: tipoMttSeleccionado ? tipoMttSeleccionado.value : "",

            // Firmas y Responsables
            recibe_orden: document.getElementById("recibe_orden")?.value ?? "",
            jefe_taller: document.getElementById("jefe_taller")?.value ?? "",
            realiza_trabajo: document.getElementById("realiza_trabajo")?.value ?? "",
            recibe_trabajo: document.getElementById("recibe_trabajo")?.value ?? ""
        };

        try {
            await addDoc(collection(db, "ordenes"), ordenData);
            alert("Orden guardada exitosamente");
        } catch (e) {
            console.error("Error al guardar: ", e);
            alert("Hubo un error al guardar, intenta de nuevo.");
        }
    });
}

// --- PARA GESTIÓN--- 
const tablaOrdenesBody = document.getElementById("tablaOrdenes");
if (tablaOrdenesBody) {

    let ordenesCache = []; // guardamos todas las órdenes en memoria para filtrar sin volver a consultar Firestore

    const pintarFilas = (lista) => {
        tablaOrdenesBody.innerHTML = "";

        if (lista.length === 0) {
            tablaOrdenesBody.innerHTML = `<tr><td colspan="6" style="text-align:center;">No hay órdenes que coincidan.</td></tr>`;
            return;
        }

        lista.forEach((data) => {
            const fila = document.createElement("tr");
            fila.innerHTML = `
                <td>${data.folio ?? ""}</td>
                <td>${data.fecha ?? ""}</td>
                <td>${data.solicitante ?? ""}</td>
                <td>${data.area ?? ""}</td>
                <td>${data.descripcion ?? ""}</td>
                <td>${data.recibe_trabajo ?? ""}</td>
            `;
            tablaOrdenesBody.appendChild(fila);
        });
    };

    const cargarOrdenes = async () => {
        const estadoCarga = document.getElementById("estadoCarga");
        const tablaContainer = document.getElementById("tablaOrdenesContainer");

        try {
            const q = query(collection(db, "ordenes"), orderBy("fecha", "desc"));
            const snapshot = await getDocs(q);

            ordenesCache = snapshot.docs.map((doc) => doc.data());

            pintarFilas(ordenesCache);

            estadoCarga.style.display = "none";
            tablaContainer.style.display = "table";
        } catch (e) {
            console.error("Error al cargar órdenes: ", e);
            estadoCarga.textContent = "Hubo un error al cargar las órdenes.";
        }
    };

    // Filtros en vivo
    const filtroFolio = document.getElementById("filtroFolio");
    const filtroArea = document.getElementById("filtroArea");

    const aplicarFiltros = () => {
        const textoFolio = filtroFolio.value.trim().toLowerCase();
        const areaSeleccionada = filtroArea.value;

        const filtradas = ordenesCache.filter((data) => {
            const coincideFolio = (data.folio ?? "").toLowerCase().includes(textoFolio);
            const coincideArea = areaSeleccionada === "" || data.area === areaSeleccionada;
            return coincideFolio && coincideArea;
        });

        pintarFilas(filtradas);
    };

    filtroFolio.addEventListener("input", aplicarFiltros);
    filtroArea.addEventListener("change", aplicarFiltros);

    cargarOrdenes();
}

// --- BORRAR EL DASHBOARD ---
const limpiarBtn = document.getElementById("limpiarBtn");
if (limpiarBtn) {
    limpiarBtn.addEventListener("click", () => {
        // 1. Limpiar inputs de texto y áreas
        document.getElementById("solicitante_nombre").value = "";
        document.getElementById("area").value = "";
        document.getElementById("descripcion_problema").value = "";
        document.getElementById("trabajo_realizado").value = "";
        document.getElementById("refacciones_usadas").value = "";
        document.getElementById("realiza_trabajo").value = "";
        document.getElementById("recibe_trabajo").value = "";

        // 2. Desmarcar los radio buttons de estado y mantenimiento
        document.querySelectorAll('input[name="estado_antes"]').forEach(radio => radio.checked = false);
        document.querySelectorAll('input[name="tipo_mtt"]').forEach(radio => radio.checked = false);

        // 3. Generar un nuevo folio y actualizar fecha/hora actual de forma automática
        const ahora = new Date();
        document.getElementById("fecha").value = ahora.toLocaleDateString('es-MX');
        document.getElementById("hora").value = ahora.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });
        document.getElementById("folio").value = "IBM-";

        alert("Formulario limpiado correctamente.");
    });
}