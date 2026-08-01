import { initializeApp } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-app.js";
import { getAuth, signInWithEmailAndPassword, signOut } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-auth.js";
import { getFirestore, collection, addDoc, getDocs, query, orderBy } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-firestore.js";
import { PDFDocument, rgb } from "https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.esm.js";

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

// --- FUNCIÓN AUXILIAR PARA CAPTURAR LOS DATOS ---
function capturarDatosFormulario() {
    const estadoAntesSeleccionado = document.querySelector('input[name="estado_antes"]:checked');
    const tipoMttSeleccionado = document.querySelector('input[name="tipo_mtt"]:checked');

    return {
        folio: document.getElementById("folio")?.value ?? "",
        fecha: document.getElementById("fecha")?.value ?? "",
        hora: document.getElementById("hora")?.value ?? "",
        solicitante: document.getElementById("solicitante_nombre")?.value ?? "",
        area: document.getElementById("area")?.value ?? "",
        descripcion: document.getElementById("descripcion_problema")?.value ?? "",
        trabajo_realizado: document.getElementById("trabajo_realizado")?.value ?? "",
        refacciones_usadas: document.getElementById("refacciones_usadas")?.value ?? "",
        estado_antes: estadoAntesSeleccionado ? estadoAntesSeleccionado.value : "",
        tipo_mtt: tipoMttSeleccionado ? tipoMttSeleccionado.value : "",
        recibe_orden: document.getElementById("recibe_orden")?.value ?? "",
        jefe_taller: document.getElementById("jefe_taller")?.value ?? "",
        realiza_trabajo: document.getElementById("realiza_trabajo")?.value ?? "",
        recibe_trabajo: document.getElementById("recibe_trabajo")?.value ?? ""
    };
}

// ---  BOTÓN GUARDAR EN FIREBASE ---
const guardarBtn = document.getElementById("guardarBtn");
if (guardarBtn) {
    guardarBtn.addEventListener("click", async () => {
        const ordenData = capturarDatosFormulario();

        try {
            await addDoc(collection(db, "ordenes"), ordenData);
            alert("Orden guardada exitosamente en la base de datos");
        } catch (e) {
            console.error("Error al guardar: ", e);
            alert("Hubo un error al guardar, intenta de nuevo.");
        }
    });
}

// --- BOTÓN GENERAR Y DESCARGAR PDF ---
const pdfBtn = document.getElementById("pdfBtn");
if (pdfBtn) {
    pdfBtn.addEventListener("click", async () => {
        const ordenData = capturarDatosFormulario();
        
        if (!ordenData.folio || ordenData.folio === "IBM-") {
            alert("Por favor ingresa un número de folio válido antes de generar el PDF.");
            return;
        }

        try {
            await rellenarPlantillaPDF(ordenData);
        } catch (e) {
            console.error("Error al generar PDF: ", e);
            alert("Hubo un error al generar el PDF. Asegúrate de tener 'plantilla.pdf' en la carpeta.");
        }
    });
}

// --- 3. BOTÓN LIMPIAR FORMULARIO ---
const limpiarBtn = document.getElementById("limpiarBtn");
if (limpiarBtn) {
    limpiarBtn.addEventListener("click", () => {
        document.getElementById("folio").value = "IBM-";
        document.getElementById("solicitante_nombre").value = "";
        document.getElementById("area").value = "";
        document.getElementById("descripcion_problema").value = "";
        document.getElementById("trabajo_realizado").value = "";
        document.getElementById("refacciones_usadas").value = "";
        document.getElementById("realiza_trabajo").value = "";
        document.getElementById("recibe_trabajo").value = "";

        document.querySelectorAll('input[name="estado_antes"]').forEach(radio => radio.checked = false);
        document.querySelectorAll('input[name="tipo_mtt"]').forEach(radio => radio.checked = false);

        alert("Formulario limpiado correctamente.");
    });
}

// Función para rellenar la plantilla oficial con coordenadas fijas
async function rellenarPlantillaPDF(data) {
    // 1. Cargar tu archivo "plantilla.pdf" que debe estar en la misma carpeta raíz
    const urlPlantilla = "plantilla.pdf";
    const existingPdfBytes = await fetch(urlPlantilla).then(res => res.arrayBuffer());

    // 2. Cargar el documento con pdf-lib
    const pdfDoc = await PDFDocument.load(existingPdfBytes);
    const pages = pdfDoc.getPages();
    const firstPage = pages[0];

    // Color de texto negro estándar
    const textColor = rgb(0, 0, 0);

    // --- ESCRIBIR TEXTOS EN LAS COORDENADAS EXACTAS DE LA PLANTILLA ---
    
    // Folio (Esquina superior derecha, cuadro IBMD / Folio)
    firstPage.drawText(data.folio, { x: 440, y: 735, size: 9, color: textColor });
    
    // Fecha y hora (si tu plantilla tiene apartado de fecha)
    firstPage.drawText(data.fecha, { x: 430, y: 705, size: 8, color: textColor });

    // Nombre del Solicitante y Área
    firstPage.drawText(data.solicitante, { x: 80, y: 655, size: 9, color: textColor });
    firstPage.drawText(data.area, { x: 280, y: 655, size: 9, color: textColor });

    // Descripción del Trabajo Requerido (Bloque grande superior)
    firstPage.drawText(data.descripcion, { x: 70, y: 580, size: 9, color: textColor });

    // Descripción del Trabajo Realizado (Bloque central)
    firstPage.drawText(data.trabajo_realizado, { x: 70, y: 480, size: 9, color: textColor });

    // Refacciones, accesorios y/o químicos utilizados
    firstPage.drawText(data.refacciones_usadas, { x: 70, y: 390, size: 9, color: textColor });

    // --- MARCAR CON UNA "X" LOS CHECKBOXES SEGÚN SELECCIÓN ---
    // Estado del equipo antes del mantenimiento: "funcionamiento", "fuera_servicio", "NA"
    if (data.estado_antes === "funcionamiento") {
        firstPage.drawText("X", { x: 132, y: 275, size: 10, color: textColor });
    } else if (data.estado_antes === "fuera_servicio") {
        firstPage.drawText("X", { x: 232, y: 275, size: 10, color: textColor });
    } else if (data.estado_antes === "NA") {
        firstPage.drawText("X", { x: 335, y: 275, size: 10, color: textColor });
    }

    // Tipo de mantenimiento: "MP", "MC", "MC_MP", "NA"
    if (data.tipo_mtt === "MP") {
        firstPage.drawText("X", { x: 445, y: 275, size: 10, color: textColor });
    } else if (data.tipo_mtt === "MC") {
        firstPage.drawText("X", { x: 502, y: 275, size: 10, color: textColor });
    } else if (data.tipo_mtt === "MC_MP") {
        firstPage.drawText("X", { x: 562, y: 275, size: 10, color: textColor });
    } else if (data.tipo_mtt === "NA") {
        firstPage.drawText("X", { x: 622, y: 275, size: 10, color: textColor });
    }

    // Firmas y Responsables (Parte inferior de la tabla)
    firstPage.drawText(data.recibe_orden, { x: 70, y: 120, size: 7, color: textColor });
    firstPage.drawText(data.realiza_trabajo, { x: 190, y: 120, size: 8, color: textColor });
    firstPage.drawText(data.recibe_trabajo, { x: 310, y: 120, size: 8, color: textColor });
    firstPage.drawText(data.jefe_taller, { x: 430, y: 120, size: 8, color: textColor });

    // 3. Serializar y descargar el PDF resultante de forma automática
    const pdfBytes = await pdfDoc.save();
    const blob = new Blob([pdfBytes], { type: "application/pdf" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `Orden_${data.folio}.pdf`;
    link.click();
}