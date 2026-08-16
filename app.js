// ULTIMA CARGA

import { initializeApp } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-app.js";
import { getAuth, signInWithEmailAndPassword, signOut } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-auth.js";
import { getFirestore, collection, addDoc, updateDoc, getDoc, getDocs, query, where, orderBy, deleteDoc, doc } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-firestore.js";
import { PDFDocument, rgb } from "https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.esm.js";
import fontkit from "https://esm.sh/@pdf-lib/fontkit@1.1.1";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, ImageRun, VerticalMergeType, VerticalAlign, Header, TableLayoutType, HeightRule } from "https://esm.sh/docx@8.5.0";

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
const loginForm = document.getElementById("loginForm");
if (loginForm) {
    loginForm.addEventListener("submit", (e) => {
        e.preventDefault(); // evita que el formulario recargue la página
        const email = document.getElementById("email").value;
        const password = document.getElementById("password").value;
        signInWithEmailAndPassword(auth, email, password)
            .then(() => window.location.href = "dashboard.html")
            .catch(() => alert("Error: Verifica tus credenciales."));
    });
}

// --- MOSTRAR/OCULTAR CONTRASEÑA (el "ojito") EN EL LOGIN ---
const togglePasswordBtn = document.getElementById("togglePassword");
if (togglePasswordBtn) {
    togglePasswordBtn.addEventListener("click", () => {
        const passwordInput = document.getElementById("password");
        const iconoOjo = togglePasswordBtn.querySelector(".icon-eye");
        const iconoOjoTachado = togglePasswordBtn.querySelector(".icon-eye-off");

        const seEstaMostrando = passwordInput.type === "text";

        passwordInput.type = seEstaMostrando ? "password" : "text";
        togglePasswordBtn.setAttribute("aria-label", seEstaMostrando ? "Mostrar contraseña" : "Ocultar contraseña");

        if (iconoOjo && iconoOjoTachado) {
            iconoOjo.style.display = seEstaMostrando ? "block" : "none";
            iconoOjoTachado.style.display = seEstaMostrando ? "none" : "block";
        }
    });
}

// ---  CERRAR SESIÓN ---
// La hacemos accesible al HTML mediante 'window'
window.cerrarSesion = () => {
    signOut(auth).then(() => {
        window.location.href = "index.html";
    }).catch((error) => console.error("Error al cerrar sesión:", error));
};

// --- PARA GESTIÓN--- 
// --- CONVIERTE EL VALOR INTERNO DEL SELECT DE ÁREA A UN TEXTO LEGIBLE ---
// Ej: "medicina_interna" -> "Medicina Interna"
// Si agregas más áreas al <select id="area"> en el HTML, agrégalas también aquí.
function formatearArea(valor) {
    const mapaAreas = {
        urgencias_adultos: "Urgencias Adultos",
        urgencias_pediatricas: "Urgencias Pediátricas",
        medicina_interna: "Medicina Interna",
        cirugia: "Cirugía",
        quirofano: "Quirófano",
        ceye: "CEyE",
        terapia_adultos: "Terapia Adultos",
        ucip: "UCIP",
        pediatria: "Pediatría",
        traumatologia: "Traumatología",
        rehabilitacion: "Rehabilitación",
        imagenologia: "Imagenología",
        laboratorio: "Laboratorio",
        consulta_externa: "Consulta Externa",
        hospitalizacion: "Hospitalización",
        departamento_biomedico: "Departamento Biomédico"
    };
    return mapaAreas[valor] ?? valor ?? "";
}

const tablaOrdenesBody = document.getElementById("tablaOrdenes");
if (tablaOrdenesBody) {

    let ordenesCache = []; // guardamos todas las órdenes en memoria para filtrar sin volver a consultar Firestore

    const pintarFilas = (lista) => {
        tablaOrdenesBody.innerHTML = "";

        if (lista.length === 0) {
            tablaOrdenesBody.innerHTML = `<tr><td colspan="5" style="text-align:center;">No hay órdenes que coincidan.</td></tr>`;
            return;
        }

        lista.forEach((data) => {
            const fila = document.createElement("tr");
            fila.innerHTML = `
                <td>${data.folio ?? ""}</td>
                <td>${data.fecha ?? ""}</td>
                <td>${data.solicitante ?? ""}</td>
                <td>${formatearArea(data.area)}</td>
                <td>
                    <button type="button" class="btn-corregir-fila" style="background-color:#f0ad4e; color:white;">Corregir</button>
                    <button type="button" class="btn-pdf-fila">Descargar PDF</button>
                    <button type="button" class="btn-word-fila">Descargar Word</button>
                    <button type="button" class="btn-eliminar-fila" style="background-color:#e03538; color:white;">Eliminar</button>
                </td>
            `;

            // Botón: Corregir esta orden (te manda al formulario con los datos ya cargados)
            const btnCorregirFila = fila.querySelector(".btn-corregir-fila");
            btnCorregirFila.addEventListener("click", () => {
                // Guardamos el ID de la orden que se va a corregir para que el formulario
                // sepa que debe cargar estos datos y, al guardar, actualizar en vez de crear una nueva.
                localStorage.setItem("ordenEditandoId", data.id);
                // OJO: Ajusta "dashboard.html" si tu formulario de captura vive en otro archivo.
                window.location.href = "dashboard.html";
            });

            // Botón: Descargar PDF de esta orden
            const btnPdfFila = fila.querySelector(".btn-pdf-fila");
            btnPdfFila.addEventListener("click", async () => {
                try {
                    await rellenarPlantillaPDF(data);
                } catch (e) {
                    console.error("Error al generar PDF: ", e);
                    alert("Hubo un error al generar el PDF. Asegúrate de tener 'plantilla.pdf' en la carpeta.");
                }
            });

            // Botón: Descargar Word de esta orden
            const btnWordFila = fila.querySelector(".btn-word-fila");
            btnWordFila.addEventListener("click", async () => {
                try {
                    await generarWordOrden(data);
                } catch (e) {
                    console.error("Error al generar Word: ", e);
                    alert("Hubo un error al generar el documento de Word.");
                }
            });

            // Botón: Eliminar esta orden (de la tabla y de Firestore)
            const btnEliminarFila = fila.querySelector(".btn-eliminar-fila");
            btnEliminarFila.addEventListener("click", async () => {
                const confirmar = confirm(`¿Seguro que quieres eliminar la orden con folio "${data.folio}"? Esta acción no se puede deshacer.`);
                if (!confirmar) return;

                try {
                    await deleteDoc(doc(db, "ordenes", data.id));
                    // Quitamos la orden de la memoria local y volvemos a pintar la tabla, sin recargar todo desde Firestore
                    ordenesCache = ordenesCache.filter((o) => o.id !== data.id);
                    aplicarFiltros();
                    alert("Orden eliminada correctamente.");
                } catch (e) {
                    console.error("Error al eliminar la orden: ", e);
                    alert("Hubo un error al eliminar la orden. Intenta de nuevo.");
                }
            });

            tablaOrdenesBody.appendChild(fila);
        });
    };

    const cargarOrdenes = async () => {
        const estadoCarga = document.getElementById("estadoCarga");
        const tablaContainer = document.getElementById("tablaOrdenesContainer");

        try {
            const q = query(collection(db, "ordenes"), orderBy("fecha", "desc"));
            const snapshot = await getDocs(q);

            ordenesCache = snapshot.docs.map((docSnap) => ({ id: docSnap.id, ...docSnap.data() }));

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

// --- SECCIÓN: EXPORTAR ÓRDENES A CSV POR RANGO DE FECHAS ---
// Pensado para una página/apartado nuevo (ej. "reportes.html"), con estos elementos en el HTML:
//   <input type="date" id="filtroFechaInicio">
//   <input type="date" id="filtroFechaFin">
//   <button type="button" id="generarCSVBtn">Generar CSV</button>
//   <p id="resultadoCSV"></p>   (opcional, para mostrar cuántas órdenes se encontraron)
const generarCSVBtn = document.getElementById("generarCSVBtn");
if (generarCSVBtn) {

    // Convierte un valor a texto seguro para CSV: si contiene comas, comillas o saltos de
    // línea, lo envuelve entre comillas dobles y escapa las comillas internas duplicándolas.
    function escaparCampoCSV(valor) {
        const texto = String(valor ?? "");
        if (/[",\n]/.test(texto)) {
            return `"${texto.replace(/"/g, '""')}"`;
        }
        return texto;
    }

    // Reconstruye una fecha real (objeto Date) a partir de los campos dia/mes/anio que ya
    // guarda cada orden, para poder comparar fechas de forma confiable (el campo "fecha" es
    // solo texto tipo "13/8/2026" y no sirve para comparar rangos).
    function construirFechaDeOrden(data) {
        const dia = parseInt(data.dia, 10);
        const mes = parseInt(data.mes, 10);
        const anio = parseInt(data.anio, 10);
        if (!dia || !mes || !anio) return null;
        return new Date(anio, mes - 1, dia);
    }

    generarCSVBtn.addEventListener("click", async () => {
        const inicioInput = document.getElementById("filtroFechaInicio");
        const finInput = document.getElementById("filtroFechaFin");
        const resultadoEl = document.getElementById("resultadoCSV");

        const fechaInicioValor = inicioInput?.value; // formato "YYYY-MM-DD" (nativo de <input type="date">)
        const fechaFinValor = finInput?.value;

        if (!fechaInicioValor || !fechaFinValor) {
            alert("Selecciona una fecha de inicio y una fecha de término.");
            return;
        }

        const fechaInicio = new Date(fechaInicioValor + "T00:00:00");
        const fechaFin = new Date(fechaFinValor + "T23:59:59");

        if (fechaInicio > fechaFin) {
            alert("La fecha de inicio no puede ser posterior a la fecha de término.");
            return;
        }

        try {
            // Traemos todas las órdenes y filtramos aquí mismo en el navegador (ver nota arriba
            // sobre por qué no se puede filtrar directamente en la consulta a Firestore).
            const snapshot = await getDocs(collection(db, "ordenes"));
            const ordenes = snapshot.docs.map((docSnap) => ({ id: docSnap.id, ...docSnap.data() }));

            const ordenesFiltradas = ordenes.filter((data) => {
                const fechaOrden = construirFechaDeOrden(data);
                return fechaOrden && fechaOrden >= fechaInicio && fechaOrden <= fechaFin;
            });

            if (ordenesFiltradas.length === 0) {
                alert("No se encontraron órdenes en ese rango de fechas.");
                if (resultadoEl) resultadoEl.textContent = "0 órdenes encontradas.";
                return;
            }

            // De la más antigua a la más reciente, para que el CSV se lea en orden cronológico
            ordenesFiltradas.sort((a, b) => construirFechaDeOrden(a) - construirFechaDeOrden(b));

            const fechaInicioTexto = fechaInicio.toLocaleDateString('es-MX');
            const fechaFinTexto = fechaFin.toLocaleDateString('es-MX');

            const encabezados = [
                "Folio", "Fecha", "Hora", "Solicitante", "Área",
                "Descripción del Trabajo Requerido", "Trabajo Realizado",
                "Refacciones/Accesorios/Químicos", "Estado del Equipo (Antes)",
                "Tipo de Mantenimiento", "Recibe Orden", "Jefe de Taller",
                "Realiza Trabajo", "Recibe Trabajo", "Observaciones",
            ];

            const filasCSV = [];
            filasCSV.push(escaparCampoCSV(`REGISTRO DE ÓRDENES DE SERVICIO DEL ${fechaInicioTexto} AL ${fechaFinTexto}`));
            filasCSV.push(escaparCampoCSV(`Número de órdenes: ${ordenesFiltradas.length}`));
            filasCSV.push(""); // fila en blanco antes de la tabla de datos
            filasCSV.push(encabezados.map(escaparCampoCSV).join(","));

            ordenesFiltradas.forEach((data) => {
                const fila = [
                    data.folio, data.fecha, data.hora, data.solicitante, formatearArea(data.area),
                    data.descripcion, data.trabajo_realizado, data.refacciones_usadas,
                    data.estado_antes, data.tipo_mtt, data.recibe_orden, data.jefe_taller,
                    data.realiza_trabajo, data.recibe_trabajo, data.observaciones,
                ].map(escaparCampoCSV).join(",");
                filasCSV.push(fila);
            });

            // El BOM (\uFEFF) al inicio hace que Excel reconozca los acentos (á, é, í, ó, ú, ñ) bien
            const contenidoCSV = "\uFEFF" + filasCSV.join("\n");
            const blob = new Blob([contenidoCSV], { type: "text/csv;charset=utf-8;" });
            const link = document.createElement("a");
            link.href = URL.createObjectURL(blob);
            link.download = `Registro_Ordenes_${fechaInicioValor}_a_${fechaFinValor}.csv`;
            link.click();

            if (resultadoEl) resultadoEl.textContent = `${ordenesFiltradas.length} orden(es) encontradas y exportadas.`;

        } catch (e) {
            console.error("Error al generar el CSV: ", e);
            alert("Hubo un error al generar el archivo CSV.");
        }
    });
}

// Sincronizar solicitante
const solInput = document.getElementById("solicitante_nombre");
if (solInput) {
    solInput.addEventListener("input", (e) => {
        const recibeTrabajoInput = document.getElementById("recibe_trabajo");
        if (recibeTrabajoInput) {
            recibeTrabajoInput.value = e.target.value;
        }
    });
}
// --- AUTOCOMPLETAR FECHA, HORA Y FOLIO AL CARGAR LA PÁGINA ---
// (fuera de cualquier función, para que corra apenas se carga app.js)
const campoFecha = document.getElementById("fecha");
const campoHora = document.getElementById("hora");
const campoFolio = document.getElementById("folio");

if (campoFecha && campoHora) {
    const ahora = new Date();
    campoFecha.value = ahora.toLocaleDateString('es-MX');
    campoHora.value = ahora.toLocaleTimeString('es-MX', { hour: '2-digit', minute: '2-digit' });
}
if (campoFolio && !campoFolio.value) {
    campoFolio.value = "IBM-";
}

// --- SUGERIR AUTOMÁTICAMENTE EL SIGUIENTE FOLIO (último número registrado + 1) ---
// El usuario sigue pudiendo editarlo libremente; esto solo prellena el campo.
// Si estamos en "modo corrección" (llegamos aquí con el botón "Corregir" de Gestión),
// NO tocamos el folio: cargarOrdenParaEditar() más abajo se encarga de poner el folio
// real de esa orden, y no queremos que esta sugerencia se lo pise.
async function inicializarFolioSugerido() {
    if (!campoFolio) return;
    if (localStorage.getItem("ordenEditandoId")) return; // estamos corrigiendo una orden, no sugerir folio nuevo

    try {
        const snapshot = await getDocs(collection(db, "ordenes"));
        let mayorNumero = 0;

        snapshot.forEach((docSnap) => {
            const folioGuardado = (docSnap.data().folio ?? "").trim();
            if (!folioGuardado) return;

            // Extrae el número al final del folio (ej. "IBM-89" -> 89, "IBM-2222" -> 2222)
            const coincidencia = folioGuardado.match(/(\d+)\s*$/);
            if (coincidencia) {
                const numero = parseInt(coincidencia[1], 10);
                if (numero > mayorNumero) mayorNumero = numero;
            }
        });

        // Solo sugerimos un número si SÍ hay órdenes registradas con folio numérico.
        // Si no hay ninguna, dejamos el campo tal cual ("IBM-"), sin inventar un número.
        if (mayorNumero > 0) {
            campoFolio.value = `IBM-${mayorNumero + 1}`;
        }
    } catch (e) {
        console.error("Error al calcular el siguiente folio: ", e);
        // Si falla la consulta, dejamos el folio como estaba ("IBM-") y no bloqueamos el formulario.
    }
}

inicializarFolioSugerido();

// --- MODO CORRECCIÓN: SI VENIMOS DEL BOTÓN "CORREGIR" DE LA TABLA DE GESTIÓN ---
// Guarda en esta variable el ID de la orden que se está corrigiendo (si aplica).
// Mientras tenga un valor, el botón "Guardar" actualizará esa orden en Firestore
// en vez de crear una nueva.
let ordenEditandoId = null;

async function cargarOrdenParaEditar() {
    const idAEditar = localStorage.getItem("ordenEditandoId");
    if (!idAEditar || !campoFolio) return; // no hay nada que editar, o no estamos en el formulario

    try {
        const refOrden = doc(db, "ordenes", idAEditar);
        const snap = await getDoc(refOrden);

        if (!snap.exists()) {
            alert("No se encontró la orden que intentas corregir (puede que ya haya sido eliminada).");
            localStorage.removeItem("ordenEditandoId");
            return;
        }

        const data = snap.data();
        ordenEditandoId = idAEditar;

        // Rellenamos todos los campos del formulario con los datos guardados
        if (document.getElementById("folio")) document.getElementById("folio").value = data.folio ?? "";
        if (document.getElementById("fecha")) document.getElementById("fecha").value = data.fecha ?? "";
        if (document.getElementById("hora")) document.getElementById("hora").value = data.hora ?? "";
        if (document.getElementById("solicitante_nombre")) document.getElementById("solicitante_nombre").value = data.solicitante ?? "";
        if (document.getElementById("area")) document.getElementById("area").value = data.area ?? "";
        if (document.getElementById("descripcion_problema")) document.getElementById("descripcion_problema").value = data.descripcion ?? "";
        if (document.getElementById("trabajo_realizado")) document.getElementById("trabajo_realizado").value = data.trabajo_realizado ?? "";
        if (document.getElementById("refacciones_usadas")) document.getElementById("refacciones_usadas").value = data.refacciones_usadas ?? "";
        if (document.getElementById("recibe_orden")) document.getElementById("recibe_orden").value = data.recibe_orden ?? "";
        if (document.getElementById("jefe_taller")) document.getElementById("jefe_taller").value = data.jefe_taller ?? "";
        if (document.getElementById("realiza_trabajo")) document.getElementById("realiza_trabajo").value = data.realiza_trabajo ?? "";
        if (document.getElementById("recibe_trabajo")) document.getElementById("recibe_trabajo").value = data.recibe_trabajo ?? "";
        if (document.getElementById("observaciones_inf")) document.getElementById("observaciones_inf").value = data.observaciones ?? "";

        document.querySelectorAll('input[name="estado_antes"]').forEach((radio) => {
            radio.checked = (radio.value === data.estado_antes);
        });
        document.querySelectorAll('input[name="tipo_mtt"]').forEach((radio) => {
            radio.checked = (radio.value === data.tipo_mtt);
        });

        // Avisamos visualmente que estamos corrigiendo, no creando una orden nueva
        const guardarBtnRef = document.getElementById("guardarBtn");
        if (guardarBtnRef) guardarBtnRef.textContent = "Actualizar Orden";

    } catch (e) {
        console.error("Error al cargar la orden para corregir: ", e);
        alert("Hubo un error al cargar la orden para corregir.");
    }
}

cargarOrdenParaEditar();

// --- FUNCIÓN TEMPORAL: GENERA LA PLANTILLA CON CUADRÍCULA PARA CALIBRAR COORDENADAS ---
// Esta función es solo para ayudarte a encontrar las coordenadas (x, y) exactas.
// Una vez que termines de ajustar rellenarPlantillaPDF, puedes borrar esta función y su botón.
async function generarCuadriculaPDF() {
    const urlPlantilla = "plantilla1.pdf";
    const existingPdfBytes = await fetch(urlPlantilla).then(res => res.arrayBuffer());

    const pdfDoc = await PDFDocument.load(existingPdfBytes);
    const firstPage = pdfDoc.getPages()[0];
    const { width, height } = firstPage.getSize();

    const colorLinea = rgb(1, 0, 0);      // rojo, semi-visible
    const colorTexto = rgb(0, 0, 1);      // azul para los números

    const paso = 25; // separación de la cuadrícula en puntos (ajusta si quieres más o menos densidad)

    // Líneas y números verticales (eje X)
    for (let x = 0; x <= width; x += paso) {
        firstPage.drawLine({
            start: { x, y: 0 },
            end: { x, y: height },
            thickness: 0.3,
            color: colorLinea,
            opacity: 0.4,
        });
        firstPage.drawText(String(x), {
            x: x + 1,
            y: height - 10,
            size: 5,
            color: colorTexto,
        });
    }

    // Líneas y números horizontales (eje Y)
    for (let y = 0; y <= height; y += paso) {
        firstPage.drawLine({
            start: { x: 0, y },
            end: { x: width, y },
            thickness: 0.3,
            color: colorLinea,
            opacity: 0.4,
        });
        firstPage.drawText(String(y), {
            x: 2,
            y: y + 1,
            size: 5,
            color: colorTexto,
        });
    }

    const pdfBytes = await pdfDoc.save();
    const blob = new Blob([pdfBytes], { type: "application/pdf" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "plantilla_cuadricula.pdf";
    link.click();
}

// Botón temporal para generar la cuadrícula (bórralo cuando termines de calibrar)
const cuadriculaBtn = document.getElementById("cuadriculaBtn");
if (cuadriculaBtn) {
    cuadriculaBtn.addEventListener("click", async () => {
        try {
            await generarCuadriculaPDF();
        } catch (e) {
            console.error("Error al generar cuadrícula: ", e);
            alert("Hubo un error al generar la cuadrícula.");
        }
    });
}

// --- FUNCIÓN AUXILIAR PARA CAPTURAR LOS DATOS ---
function capturarDatosFormulario() {
    const estadoAntesSeleccionado = document.querySelector('input[name="estado_antes"]:checked');
    const tipoMttSeleccionado = document.querySelector('input[name="tipo_mtt"]:checked');

    // La fecha se guarda como texto tipo "5/8/2026" (día/mes/año, formato es-MX).
    // Aquí la separamos en tres partes para los cuadritos de Día / Mes / Año del PDF.
    const fechaTexto = document.getElementById("fecha")?.value ?? "";
    const [dia = "", mes = "", anio = ""] = fechaTexto.split("/");

    return {
        folio: document.getElementById("folio")?.value ?? "",
        fecha: fechaTexto,
        hora: document.getElementById("hora")?.value ?? "",
        dia: dia,
        mes: mes,
        anio: anio,
        solicitante: document.getElementById("solicitante_nombre")?.value ?? "",
        area: formatearArea(document.getElementById("area")?.value ?? ""),
        descripcion: document.getElementById("descripcion_problema")?.value ?? "",
        trabajo_realizado: document.getElementById("trabajo_realizado")?.value ?? "",
        refacciones_usadas: document.getElementById("refacciones_usadas")?.value ?? "",
        estado_antes: estadoAntesSeleccionado ? estadoAntesSeleccionado.value : "",
        tipo_mtt: tipoMttSeleccionado ? tipoMttSeleccionado.value : "",
        recibe_orden: document.getElementById("recibe_orden")?.value ?? "",
        jefe_taller: document.getElementById("jefe_taller")?.value ?? "",
        realiza_trabajo: document.getElementById("realiza_trabajo")?.value ?? "",
        recibe_trabajo: document.getElementById("recibe_trabajo")?.value ?? "",
        observaciones: document.getElementById("observaciones_inf")?.value ?? ""
    };
}

// --- VERIFICA SI UN FOLIO YA EXISTE EN OTRA ORDEN ---
// idAExcluir: cuando estamos corrigiendo una orden, no queremos que se compare contra sí misma.
async function folioYaExiste(folio, idAExcluir = null) {
    const q = query(collection(db, "ordenes"), where("folio", "==", folio));
    const snapshot = await getDocs(q);
    return snapshot.docs.some((docSnap) => docSnap.id !== idAExcluir);
}

// ---  BOTÓN GUARDAR EN FIREBASE ---
const guardarBtn = document.getElementById("guardarBtn");
if (guardarBtn) {
    guardarBtn.addEventListener("click", async () => {
        const ordenData = capturarDatosFormulario();

        if (!ordenData.folio || ordenData.folio === "IBM-") {
            alert("Por favor ingresa un número de folio válido antes de guardar.");
            return;
        }

        try {
            const yaExiste = await folioYaExiste(ordenData.folio, ordenEditandoId);
            if (yaExiste) {
                alert(`El folio "${ordenData.folio}" ya existe en otra orden. Usa un folio diferente.`);
                return;
            }

            if (ordenEditandoId) {
                // Modo corrección: actualizamos la orden existente en vez de crear una nueva
                await updateDoc(doc(db, "ordenes", ordenEditandoId), ordenData);
                alert("Orden corregida exitosamente.");
                localStorage.removeItem("ordenEditandoId");
                ordenEditandoId = null;
                guardarBtn.textContent = "Guardar";
            } else {
                await addDoc(collection(db, "ordenes"), ordenData);
                alert("Orden guardada exitosamente en la base de datos");
            }
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

// --- BOTÓN GENERAR Y DESCARGAR WORD ---
const wordBtn = document.getElementById("wordBtn");
if (wordBtn) {
    wordBtn.addEventListener("click", async () => {
        const ordenData = capturarDatosFormulario();

        if (!ordenData.folio || ordenData.folio === "IBM-") {
            alert("Por favor ingresa un número de folio válido antes de generar el Word.");
            return;
        }

        try {
            await generarWordOrden(ordenData);
        } catch (e) {
            console.error("Error al generar Word: ", e);
            alert("Hubo un error al generar el documento de Word.");
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

        // Si estábamos corrigiendo una orden, "Limpiar" también cancela ese modo
        // y regresa el formulario a "orden nueva".
        localStorage.removeItem("ordenEditandoId");
        ordenEditandoId = null;
        if (guardarBtn) guardarBtn.textContent = "Guardar";

        alert("Formulario limpiado correctamente.");
    });
}

// Función para rellenar la plantilla oficial con coordenadas fijas
// --- FUNCIÓN: GENERA UN DOCUMENTO DE WORD (.docx) CON LOS DATOS DE LA ORDEN ---
async function generarWordOrden(data) {
    // Cargar los 3 logos del encabezado (deben estar en la misma carpeta del proyecto)
    const logoGiganteBytes = await fetch("logo_gigante.png").then(res => res.arrayBuffer());
    const logoBiomedicaBytes = await fetch("logo_biomedica.png").then(res => res.arrayBuffer());
    const logoHospitalBytes = await fetch("logo_hospital.png").then(res => res.arrayBuffer());

    // --- CALCULAR EL TAMAÑO DE CADA LOGO RESPETANDO SU PROPORCIÓN REAL (sin distorsión) ---
    async function calcularTamanoLogo(bytes, alturaObjetivoPx) {
        const bitmap = await createImageBitmap(new Blob([bytes]));
        const escala = alturaObjetivoPx / bitmap.height;
        return { width: Math.round(bitmap.width * escala), height: alturaObjetivoPx };
    }

    const ALTURA_LOGO_PX = 65;
    const dimGigante = await calcularTamanoLogo(logoGiganteBytes, ALTURA_LOGO_PX);
    const dimBiomedica = await calcularTamanoLogo(logoBiomedicaBytes, ALTURA_LOGO_PX);
    const dimHospital = await calcularTamanoLogo(logoHospitalBytes, ALTURA_LOGO_PX);

    // --- TAMAÑO DE FUENTE ---
    const SIZE_NORMAL = 22;  // 11pt
    const SIZE_FIRMAS = 18;  // 9pt

    // Bordes estándar de celdas
    const bordeDelgado = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
    const bordersTabla = {
        top: bordeDelgado,
        bottom: bordeDelgado,
        left: bordeDelgado,
        right: bordeDelgado,
        insideHorizontal: bordeDelgado,
        insideVertical: bordeDelgado,
    };

    // --- ANCHOS EN TWIPS (1440 twips = 1 pulgada). Ancho útil de la hoja = 10240 twips ---
    const ANCHO_TABLA = 10240;
    const NOMBRE_DXA = 3400;
    const AREA_DXA = 3400;
    const DIA_DXA = 1000;
    const MES_DXA = 1000;
    const ANIO_DXA = 1440;
    const TITULO_DXA = NOMBRE_DXA + AREA_DXA;                 // 6800
    const FOLIO_DXA = DIA_DXA + MES_DXA + ANIO_DXA;           // 3440 (idéntico a Día+Mes+Año)
    const MITAD_DXA = ANCHO_TABLA / 2;                        // 5120
    const CUARTO_DXA = ANCHO_TABLA / 4;                       // 2560

    // Un separador "invisible" entre tablas: no deja espacio visible pero
    // es válido tenerlo entre dos tablas consecutivas en un documento Word.
    const separadorInvisible = () => new Paragraph({
        spacing: { before: 0, after: 0, line: 20 },
        children: [new TextRun({ text: "", size: 2 })],
    });

    // Bordes sin línea superior: se usan en la primera fila de cada tabla (excepto la primera
    // del documento) para que no se dibuje una doble línea justo donde termina la tabla anterior
    // y empieza la siguiente. Así todas las secciones se ven como una sola tabla continua.
    const bordeNinguno = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
    const bordersSinTop = { ...bordersTabla, top: bordeNinguno };

    // --- Helpers para celdas ---
    const celdaTitulo = (texto, anchoDxa, opciones = {}) =>
        new TableCell({
            width: { size: anchoDxa, type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            columnSpan: opciones.colSpan,
            verticalMerge: opciones.verticalMerge,
            borders: opciones.borders ?? bordersTabla,
            margins: { top: 100, bottom: 100, left: 100, right: 100 },
            children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: texto, bold: opciones.bold ?? true, size: SIZE_NORMAL, font: "Arial" })],
            })],
        });

    const celdaValor = (texto, anchoDxa, opciones = {}) =>
        new TableCell({
            width: { size: anchoDxa, type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            columnSpan: opciones.colSpan,
            verticalMerge: opciones.verticalMerge,
            borders: bordersTabla,
            margins: { top: 120, bottom: 120, left: 100, right: 100 },
            children: [new Paragraph({
                alignment: opciones.centrado ? AlignmentType.CENTER : AlignmentType.LEFT,
                children: [new TextRun({ text: texto || "", size: opciones.size ?? SIZE_NORMAL, font: "Arial" })],
            })],
        });

    // Celda de checkbox individual: cuadrito arriba (marcado o no) y la etiqueta debajo,
    // centrado en su propia mini-columna (para que el texto nunca se desborde), pero SIN
    // caja/borde propio — igual que en la plantilla original. `bordes` indica en cuáles de
    // los 4 lados sí queremos línea (por ejemplo, solo el borde exterior del bloque y la
    // línea divisoria entre "Estado" y "Tipo").
    const celdaCheckbox = (marcado, etiqueta, anchoDxa, bordes = {}) =>
        new TableCell({
            width: { size: anchoDxa, type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            borders: {
                top: bordes.top ? bordeDelgado : bordeNinguno,
                bottom: bordes.bottom ? bordeDelgado : bordeNinguno,
                left: bordes.left ? bordeDelgado : bordeNinguno,
                right: bordes.right ? bordeDelgado : bordeNinguno,
            },
            margins: { top: 80, bottom: 80, left: 40, right: 40 },
            children: [
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 20 },
                    children: [new TextRun({ text: marcado ? "☒" : "☐", size: 34, font: "Arial" })],
                }),
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: etiqueta, size: 22, font: "Arial" })],
                }),
            ],
        });

    // Celda de texto grande (descripciones/observaciones) con una altura MÍNIMA garantizada,
    // para que se vea espaciosa como en la plantilla original, tenga o no mucho texto.
    const filaTextoGrande = (texto, alturaMinimaTwips) =>
        new TableRow({
            height: { value: alturaMinimaTwips, rule: HeightRule.ATLEAST },
            children: [
                new TableCell({
                    width: { size: ANCHO_TABLA, type: WidthType.DXA },
                    borders: bordersTabla,
                    margins: { top: 120, bottom: 120, left: 100, right: 100 },
                    children: [new Paragraph({
                        alignment: AlignmentType.LEFT,
                        children: [new TextRun({ text: texto || "", size: SIZE_NORMAL, font: "Arial" })],
                    })],
                }),
            ],
        });

    // --- ENCABEZADO REAL DE WORD CON LOS 3 LOGOS ---
    const encabezadoLogos = new Header({
        children: [
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: {
                    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                },
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                width: { size: 33, type: WidthType.PERCENTAGE },
                                children: [new Paragraph({
                                    alignment: AlignmentType.LEFT,
                                    children: [new ImageRun({ data: logoGiganteBytes, transformation: dimGigante })],
                                })],
                            }),
                            new TableCell({
                                width: { size: 34, type: WidthType.PERCENTAGE },
                                children: [new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [new ImageRun({ data: logoBiomedicaBytes, transformation: dimBiomedica })],
                                })],
                            }),
                            new TableCell({
                                width: { size: 33, type: WidthType.PERCENTAGE },
                                children: [new Paragraph({
                                    alignment: AlignmentType.RIGHT,
                                    children: [new ImageRun({ data: logoHospitalBytes, transformation: dimHospital })],
                                })],
                            }),
                        ],
                    }),
                ],
            }),
        ],
    });

    // --- TABLA 1: Título + Folio + Nombre/Área/Día/Mes/Año (deben compartir cuadrícula) ---
    const tablaEncabezado = new Table({
        width: { size: ANCHO_TABLA, type: WidthType.DXA },
        layout: TableLayoutType.FIXED,
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        width: { size: TITULO_DXA, type: WidthType.DXA },
                        columnSpan: 2,
                        verticalMerge: VerticalMergeType.RESTART,
                        verticalAlign: VerticalAlign.CENTER,
                        borders: bordersTabla,
                        margins: { top: 120, bottom: 120, left: 100, right: 100 },
                        children: [
                            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "INSTITUTO GENERAL DE SALUD DEL ESTADO DE AGUASCALIENTES", bold: true, size: SIZE_NORMAL, font: "Arial" })] }),
                            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "HOSPITAL GENERAL TERCER MILENIO", bold: true, size: SIZE_NORMAL, font: "Arial" })] }),
                            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "DEPARTAMENTO DE INGENIERÍA BIOMÉDICA", bold: true, size: SIZE_NORMAL, font: "Arial" })] }),
                            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ORDEN DE SERVICIO", bold: false, size: SIZE_NORMAL, font: "Arial" })] }),
                        ],
                    }),
                    celdaTitulo("FOLIO", FOLIO_DXA, { colSpan: 3 }),
                ],
            }),
            new TableRow({
                children: [
                    new TableCell({
                        width: { size: TITULO_DXA, type: WidthType.DXA },
                        columnSpan: 2,
                        verticalMerge: VerticalMergeType.CONTINUE,
                        borders: bordersTabla,
                        children: [new Paragraph({ text: "" })],
                    }),
                    celdaValor(data.folio, FOLIO_DXA, { centrado: true, colSpan: 3 }),
                ],
            }),
            new TableRow({
                children: [
                    celdaTitulo("NOMBRE DEL SOLICITANTE", NOMBRE_DXA),
                    celdaTitulo("ÁREA", AREA_DXA),
                    celdaTitulo("DÍA", DIA_DXA),
                    celdaTitulo("MES", MES_DXA),
                    celdaTitulo("AÑO", ANIO_DXA),
                ],
            }),
            new TableRow({
                children: [
                    celdaValor(data.solicitante, NOMBRE_DXA),
                    celdaValor(typeof formatearArea === 'function' ? formatearArea(data.area) : data.area, AREA_DXA),
                    celdaValor(data.dia, DIA_DXA, { centrado: true }),
                    celdaValor(data.mes, MES_DXA, { centrado: true }),
                    celdaValor(data.anio, ANIO_DXA, { centrado: true }),
                ],
            }),
        ],
    });

    // --- TABLA 2: Descripción del trabajo requerido (independiente, con altura mínima espaciosa) ---
    const tablaDescReq = new Table({
        width: { size: ANCHO_TABLA, type: WidthType.DXA },
        layout: TableLayoutType.FIXED,
        rows: [
            new TableRow({ children: [celdaTitulo("DESCRIPCIÓN DEL TRABAJO REQUERIDO", ANCHO_TABLA, { borders: bordersSinTop })] }),
            filaTextoGrande(data.descripcion, 2000),
        ],
    });

    // --- TABLA 3: Descripción del trabajo realizado ---
    const tablaDescReal = new Table({
        width: { size: ANCHO_TABLA, type: WidthType.DXA },
        layout: TableLayoutType.FIXED,
        rows: [
            new TableRow({ children: [celdaTitulo("DESCRIPCIÓN DEL TRABAJO REALIZADO", ANCHO_TABLA, { borders: bordersSinTop })] }),
            filaTextoGrande(data.trabajo_realizado, 2000),
        ],
    });

    // --- TABLA 4: Refacciones ---
    const tablaRefacciones = new Table({
        width: { size: ANCHO_TABLA, type: WidthType.DXA },
        layout: TableLayoutType.FIXED,
        rows: [
            new TableRow({ children: [celdaTitulo("REFACCIONES, ACCESORIOS Y/O QUÍMICOS UTILIZADOS", ANCHO_TABLA, { borders: bordersSinTop })] }),
            filaTextoGrande(data.refacciones_usadas, 900),
        ],
    });

    // --- TABLA 5: Estado del equipo / Tipo de mantenimiento ---
    // CORREGIDO: cada opción va en su propia mini-columna (cuadrito arriba, etiqueta abajo)
    // en vez de un solo bloque de texto por mitad. Así nunca se desborda del borde de la tabla,
    // igual que en la plantilla de referencia.
    const ANCHOS_ESTADO = [1707, 1707, 1706];       // 3 columnas -> suman MITAD_DXA (5120)
    const ANCHOS_TIPO = [1280, 1280, 1280, 1280];   // 4 columnas -> suman MITAD_DXA (5120)

    const tablaEstadoTipo = new Table({
        width: { size: ANCHO_TABLA, type: WidthType.DXA },
        layout: TableLayoutType.FIXED,
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        width: { size: MITAD_DXA, type: WidthType.DXA },
                        columnSpan: 3,
                        borders: bordersSinTop,
                        margins: { top: 100, bottom: 100, left: 50, right: 50 },
                        children: [
                            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ESTADO DEL EQUIPO Y/O ACCESORIO", bold: true, size: SIZE_NORMAL, font: "Arial" })] }),
                            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "(ANTES DEL MTTO)", bold: true, size: SIZE_NORMAL, font: "Arial" })] }),
                        ],
                    }),
                    new TableCell({
                        width: { size: MITAD_DXA, type: WidthType.DXA },
                        columnSpan: 4,
                        borders: bordersSinTop,
                        margins: { top: 100, bottom: 100, left: 50, right: 50 },
                        children: [
                            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "TIPO DE MANTENIMIENTO", bold: true, size: SIZE_NORMAL, font: "Arial" })] }),
                            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "(EQUIPO Y/O ACCESORIO)", bold: true, size: SIZE_NORMAL, font: "Arial" })] }),
                        ],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    celdaCheckbox(data.estado_antes === "funcionamiento", "Funcionamiento", ANCHOS_ESTADO[0], { left: true, bottom: true }),
                    celdaCheckbox(data.estado_antes === "fuera_servicio", "Fuera de Servicio", ANCHOS_ESTADO[1], { bottom: true }),
                    celdaCheckbox(data.estado_antes === "NA", "No aplica", ANCHOS_ESTADO[2], { right: true, bottom: true }),
                    celdaCheckbox(data.tipo_mtt === "MP", "MP", ANCHOS_TIPO[0], { bottom: true }),
                    celdaCheckbox(data.tipo_mtt === "MC", "MC", ANCHOS_TIPO[1], { bottom: true }),
                    celdaCheckbox(data.tipo_mtt === "MC_MP", "MP/MC", ANCHOS_TIPO[2], { bottom: true }),
                    celdaCheckbox(data.tipo_mtt === "NA", "No aplica", ANCHOS_TIPO[3], { right: true, bottom: true }),
                ],
            }),
        ],
    });

    // --- TABLA 6: Observaciones ---
    const tablaObservaciones = new Table({
        width: { size: ANCHO_TABLA, type: WidthType.DXA },
        layout: TableLayoutType.FIXED,
        rows: [
            new TableRow({ children: [celdaTitulo("OBSERVACIONES", ANCHO_TABLA, { borders: bordersSinTop })] }),
            filaTextoGrande(data.observaciones, 700),
        ],
    });

    // --- TABLA 7: Firmas (independiente, 4 columnas iguales) ---
    const tablaFirmas = new Table({
        width: { size: ANCHO_TABLA, type: WidthType.DXA },
        layout: TableLayoutType.FIXED,
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        width: { size: CUARTO_DXA, type: WidthType.DXA },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: bordersTabla,
                        margins: { top: 100, bottom: 100, left: 50, right: 50 },
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "RECIBE ORDEN DE TRABAJO", bold: true, size: SIZE_NORMAL, font: "Arial" })] })],
                    }),
                    new TableCell({
                        width: { size: CUARTO_DXA, type: WidthType.DXA },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: bordersTabla,
                        margins: { top: 100, bottom: 100, left: 50, right: 50 },
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "REALIZA TRABAJO", bold: true, size: SIZE_NORMAL, font: "Arial" })] })],
                    }),
                    new TableCell({
                        width: { size: CUARTO_DXA, type: WidthType.DXA },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: bordersTabla,
                        margins: { top: 100, bottom: 100, left: 50, right: 50 },
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "RECIBE TRABAJO", bold: true, size: SIZE_NORMAL, font: "Arial" })] })],
                    }),
                    new TableCell({
                        width: { size: CUARTO_DXA, type: WidthType.DXA },
                        verticalAlign: VerticalAlign.CENTER,
                        borders: bordersTabla,
                        margins: { top: 100, bottom: 100, left: 50, right: 50 },
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "JEFE DE TALLER", bold: true, size: SIZE_NORMAL, font: "Arial" })] })],
                    }),
                ],
            }),
            new TableRow({
                children: [
                    celdaValor(data.recibe_orden, CUARTO_DXA, { centrado: true, size: SIZE_FIRMAS }),
                    celdaValor(data.realiza_trabajo, CUARTO_DXA, { centrado: true, size: SIZE_FIRMAS }),
                    celdaValor(data.recibe_trabajo, CUARTO_DXA, { centrado: true, size: SIZE_FIRMAS }),
                    celdaValor(data.jefe_taller, CUARTO_DXA, { centrado: true, size: SIZE_FIRMAS }),
                ],
            }),
        ],
    });

    const doc = new Document({
        sections: [
            {
                properties: {
                    page: {
                        margin: {
                            top: 1600,
                            bottom: 1000,
                            left: 1000,
                            right: 1000,
                            header: 400,
                        },
                    },
                },
                headers: {
                    default: encabezadoLogos,
                },
                children: [
                    tablaEncabezado,
                    separadorInvisible(),
                    tablaDescReq,
                    separadorInvisible(),
                    tablaDescReal,
                    separadorInvisible(),
                    tablaRefacciones,
                    separadorInvisible(),
                    tablaEstadoTipo,
                    separadorInvisible(),
                    tablaObservaciones,
                    new Paragraph({ text: "", spacing: { before: 150, after: 100 } }), // aquí sí un espacio visible antes de firmas
                    tablaFirmas,
                ],
            },
        ],
    });

    const blob = await Packer.toBlob(doc);
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `Orden_${data.folio}.docx`;
    link.click();
}
async function rellenarPlantillaPDF(data) {
    // 1. Cargar tu archivo "plantilla1.pdf" que debe estar en la misma carpeta raíz
    const urlPlantilla = "plantilla1.pdf";
    const existingPdfBytes = await fetch(urlPlantilla).then(res => res.arrayBuffer());

    // 2. Cargar el documento con pdf-lib
    const pdfDoc = await PDFDocument.load(existingPdfBytes);
    const pages = pdfDoc.getPages();
    const firstPage = pages[0];

    // Incrustar la fuente Carlito (debe estar en la misma carpeta que este proyecto, junto a plantilla1.pdf)
    pdfDoc.registerFontkit(fontkit);
    const fontBytes = await fetch("Carlito-Regular.ttf").then(res => res.arrayBuffer());
    const fuentePersonalizada = await pdfDoc.embedFont(fontBytes);

    // Color de texto negro estándar
    const textColor = rgb(0, 0, 0);
    const FONT_SIZE = 11;
    const FONT_SIZE_FIRMAS = 9;

    // --- FUNCIÓN AUXILIAR: escribe texto ajustándolo a varias líneas si no cabe en el ancho disponible ---
    // maxWidth: ancho máximo en puntos antes de saltar de línea
    // maxLines: número máximo de líneas permitidas (para no salirse del cuadro hacia abajo)
    // lineHeight: separación vertical entre líneas
    function drawTextWrapped(page, texto, opciones) {
        const { x, y, maxWidth, size = FONT_SIZE, color = textColor, lineHeight = size + 2, maxLines = 3 } = opciones;
        if (!texto) return;

        const palabras = String(texto).split(" ");
        let lineaActual = "";
        const lineas = [];

        // Estimación simple de ancho de texto (aprox. 0.55 * size por carácter para fuente estándar)
        const anchoAprox = (str) => str.length * size * 0.5;

        for (const palabra of palabras) {
            const pruebaLinea = lineaActual ? `${lineaActual} ${palabra}` : palabra;
            if (anchoAprox(pruebaLinea) > maxWidth && lineaActual) {
                lineas.push(lineaActual);
                lineaActual = palabra;
            } else {
                lineaActual = pruebaLinea;
            }
        }
        if (lineaActual) lineas.push(lineaActual);

        // Si el texto no cabe en maxLines, recortamos y avisamos con "..."
        let lineasFinales = lineas.slice(0, maxLines);
        if (lineas.length > maxLines) {
            const ultima = lineasFinales[maxLines - 1];
            lineasFinales[maxLines - 1] = ultima.slice(0, Math.max(0, ultima.length - 3)) + "...";
        }

        lineasFinales.forEach((linea, i) => {
            page.drawText(linea, { x, y: y - (i * lineHeight), size, font: fuentePersonalizada, color });
        });
    }

    // --- ESCRIBIR TEXTOS EN LAS COORDENADAS EXACTAS DE LA PLANTILLA ---

    // Folio
    firstPage.drawText(data.folio ?? "", { x: 476, y: 650, size: FONT_SIZE, font: fuentePersonalizada, color: textColor });

    // Nombre del Solicitante y Área
    firstPage.drawText(data.solicitante ?? "", { x: 50, y: 610, size: FONT_SIZE, font: fuentePersonalizada, color: textColor });
    firstPage.drawText(formatearArea(data.area) ?? "", { x: 225, y: 610, size: FONT_SIZE, font: fuentePersonalizada, color: textColor });

    // Día / Mes / Año
    firstPage.drawText(data.dia ?? "", { x: 476, y: 605, size: FONT_SIZE, font: fuentePersonalizada, color: textColor });
    firstPage.drawText(data.mes ?? "", { x: 511, y: 605, size: FONT_SIZE, font: fuentePersonalizada, color: textColor });
    firstPage.drawText(data.anio ?? "", { x: 543, y: 605, size: FONT_SIZE, font: fuentePersonalizada, color: textColor });

    // Descripción del Trabajo Requerido (bloque grande, ajustado a varias líneas)
    drawTextWrapped(firstPage, data.descripcion, { x: 50, y: 550, maxWidth: 500, maxLines: 7 });

    // Descripción del Trabajo Realizado
    drawTextWrapped(firstPage, data.trabajo_realizado, { x: 50, y: 460, maxWidth: 500, maxLines: 7 });

    // Refacciones, accesorios y/o químicos utilizados
    drawTextWrapped(firstPage, data.refacciones_usadas, { x: 50, y: 375, maxWidth: 500, maxLines: 4 });

    // --- MARCAR CON UNA "X" LOS CHECKBOXES SEGÚN SELECCIÓN ---
    // Estado del equipo antes del mantenimiento: "funcionamiento", "fuera_servicio", "NA"
    if (data.estado_antes === "funcionamiento") {
        firstPage.drawText("X", { x: 66, y: 250, size: FONT_SIZE, font: fuentePersonalizada, color: textColor });
    } else if (data.estado_antes === "fuera_servicio") {
        firstPage.drawText("X", { x: 177, y: 250, size: FONT_SIZE, font: fuentePersonalizada, color: textColor });
    } else if (data.estado_antes === "NA") {
        firstPage.drawText("X", { x: 267, y: 250, size: FONT_SIZE, font: fuentePersonalizada, color: textColor });
    }

    // Tipo de mantenimiento: "MP", "MC", "MC_MP", "NA"
    if (data.tipo_mtt === "MP") {
        firstPage.drawText("X", { x: 336, y: 250, size: FONT_SIZE, font: fuentePersonalizada, color: textColor });
    } else if (data.tipo_mtt === "MC") {
        firstPage.drawText("X", { x: 402, y: 250, size: FONT_SIZE, font: fuentePersonalizada, color: textColor });
    } else if (data.tipo_mtt === "MC_MP") {
        firstPage.drawText("X", { x: 457, y: 250, size: FONT_SIZE, font: fuentePersonalizada, color: textColor });
    } else if (data.tipo_mtt === "NA") {
        firstPage.drawText("X", { x: 520, y: 250, size: FONT_SIZE, font: fuentePersonalizada, color: textColor });
    }

    // Observaciones
    drawTextWrapped(firstPage, data.observaciones, { x: 50, y: 190, maxWidth: 500, maxLines: 2 });

    // Firmas y Responsables (Parte inferior de la tabla) — tamaño 9, hasta 2 líneas por columna
    drawTextWrapped(firstPage, data.recibe_orden, { x: 50, y: 125, maxWidth: 140, maxLines: 2, size: FONT_SIZE_FIRMAS, lineHeight: 11 });
    drawTextWrapped(firstPage, data.realiza_trabajo, { x: 200, y: 125, maxWidth: 110, maxLines: 2, size: FONT_SIZE_FIRMAS, lineHeight: 11 });
    drawTextWrapped(firstPage, data.recibe_trabajo, { x: 320, y: 125, maxWidth: 115, maxLines: 2, size: FONT_SIZE_FIRMAS, lineHeight: 11 });
    drawTextWrapped(firstPage, data.jefe_taller, { x: 445, y: 125, maxWidth: 140, maxLines: 2, size: FONT_SIZE_FIRMAS, lineHeight: 11 });

    // 3. Serializar y descargar el PDF resultante de forma automática
    const pdfBytes = await pdfDoc.save();
    const blob = new Blob([pdfBytes], { type: "application/pdf" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `Orden_${data.folio}.pdf`;
    link.click();
}