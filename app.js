import { initializeApp } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-app.js";
import { getAuth, signInWithEmailAndPassword, signOut } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-auth.js";
import { getFirestore, collection, addDoc, getDocs, query, orderBy } from "https://www.gstatic.com/firebasejs/12.15.0/firebase-firestore.js";
import { PDFDocument, rgb } from "https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.esm.js";
import fontkit from "https://esm.sh/@pdf-lib/fontkit@1.1.1";
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, ImageRun } from "https://esm.sh/docx@8.5.0";

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

// --- PARA GESTIÓN--- 
// --- CONVIERTE EL VALOR INTERNO DEL SELECT DE ÁREA A UN TEXTO LEGIBLE ---
// Ej: "medicina_interna" -> "Medicina Interna"
// Si agregas más áreas al <select id="area"> en el HTML, agrégalas también aquí.
function formatearArea(valor) {
    const mapaAreas = {
        medicina_interna: "Medicina Interna",
        urgencias: "Urgencias"
    };
    return mapaAreas[valor] ?? valor ?? "";
}

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
                <td>${formatearArea(data.area)}</td>
                <td>${data.descripcion ?? ""}</td>
                <td><button type="button" class="btn-pdf-fila">Descargar PDF</button></td>
            `;

            // Conectamos el botón de esta fila con los datos exactos de esta orden
            const btnPdfFila = fila.querySelector(".btn-pdf-fila");
            btnPdfFila.addEventListener("click", async () => {
                try {
                    await rellenarPlantillaPDF(data);
                } catch (e) {
                    console.error("Error al generar PDF: ", e);
                    alert("Hubo un error al generar el PDF. Asegúrate de tener 'plantilla.pdf' en la carpeta.");
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

    const filaEtiqueta = (texto) =>
        new TableCell({
            width: { size: 25, type: WidthType.PERCENTAGE },
            children: [new Paragraph({ children: [new TextRun({ text: texto, bold: true, size: 20 })] })],
        });

    const filaValor = (texto) =>
        new TableCell({
            width: { size: 75, type: WidthType.PERCENTAGE },
            children: [new Paragraph({ children: [new TextRun({ text: texto || "", size: 20 })] })],
        });

    const checkboxTexto = (marcado, etiqueta) => `${marcado ? "☒" : "☐"} ${etiqueta}`;

    const doc = new Document({
        sections: [
            {
                children: [
                    // --- Fila de logos (tabla sin bordes para que queden alineados en 3 columnas) ---
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
                                            children: [new ImageRun({ data: logoGiganteBytes, transformation: { width: 90, height: 60 } })],
                                        })],
                                    }),
                                    new TableCell({
                                        width: { size: 34, type: WidthType.PERCENTAGE },
                                        children: [new Paragraph({
                                            alignment: AlignmentType.CENTER,
                                            children: [new ImageRun({ data: logoBiomedicaBytes, transformation: { width: 90, height: 60 } })],
                                        })],
                                    }),
                                    new TableCell({
                                        width: { size: 33, type: WidthType.PERCENTAGE },
                                        children: [new Paragraph({
                                            alignment: AlignmentType.RIGHT,
                                            children: [new ImageRun({ data: logoHospitalBytes, transformation: { width: 90, height: 60 } })],
                                        })],
                                    }),
                                ],
                            }),
                        ],
                    }),

                    new Paragraph({ text: "", spacing: { after: 200 } }),

                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [new TextRun({ text: "INSTITUTO GENERAL DE SALUD DEL ESTADO DE AGUASCALIENTES", bold: true, size: 24 })],
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [new TextRun({ text: "HOSPITAL GENERAL TERCER MILENIO", bold: true, size: 22 })],
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [new TextRun({ text: "DEPARTAMENTO DE INGENIERÍA BIOMÉDICA", bold: true, size: 22 })],
                    }),
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        children: [new TextRun({ text: "ORDEN DE SERVICIO", bold: true, size: 22 })],
                        spacing: { after: 300 },
                    }),

                    // --- Datos generales ---
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({ children: [filaEtiqueta("Folio"), filaValor(data.folio)] }),
                            new TableRow({ children: [filaEtiqueta("Fecha"), filaValor(`${data.dia}/${data.mes}/${data.anio}`)] }),
                            new TableRow({ children: [filaEtiqueta("Hora"), filaValor(data.hora)] }),
                            new TableRow({ children: [filaEtiqueta("Nombre del Solicitante"), filaValor(data.solicitante)] }),
                            new TableRow({ children: [filaEtiqueta("Área"), filaValor(data.area)] }),
                        ],
                    }),

                    new Paragraph({ text: "", spacing: { after: 200 } }),

                    // --- Descripciones ---
                    new Paragraph({ children: [new TextRun({ text: "DESCRIPCIÓN DEL TRABAJO REQUERIDO", bold: true, size: 22 })], spacing: { after: 100 } }),
                    new Paragraph({ children: [new TextRun({ text: data.descripcion || "-", size: 20 })], spacing: { after: 200 } }),

                    new Paragraph({ children: [new TextRun({ text: "DESCRIPCIÓN DEL TRABAJO REALIZADO", bold: true, size: 22 })], spacing: { after: 100 } }),
                    new Paragraph({ children: [new TextRun({ text: data.trabajo_realizado || "-", size: 20 })], spacing: { after: 200 } }),

                    new Paragraph({ children: [new TextRun({ text: "REFACCIONES, ACCESORIOS Y/O QUÍMICOS UTILIZADOS", bold: true, size: 22 })], spacing: { after: 100 } }),
                    new Paragraph({ children: [new TextRun({ text: data.refacciones_usadas || "-", size: 20 })], spacing: { after: 200 } }),

                    // --- Estado y tipo de mantenimiento ---
                    new Paragraph({ children: [new TextRun({ text: "ESTADO DEL EQUIPO Y/O ACCESORIO (ANTES DEL MTTO)", bold: true, size: 22 })], spacing: { after: 100 } }),
                    new Paragraph({
                        children: [new TextRun({
                            text: [
                                checkboxTexto(data.estado_antes === "funcionamiento", "Funcionamiento"),
                                checkboxTexto(data.estado_antes === "fuera_servicio", "Fuera de Servicio"),
                                checkboxTexto(data.estado_antes === "NA", "No aplica"),
                            ].join("     "),
                            size: 20,
                        })],
                        spacing: { after: 200 },
                    }),

                    new Paragraph({ children: [new TextRun({ text: "TIPO DE MANTENIMIENTO (EQUIPO Y/O ACCESORIO)", bold: true, size: 22 })], spacing: { after: 100 } }),
                    new Paragraph({
                        children: [new TextRun({
                            text: [
                                checkboxTexto(data.tipo_mtt === "MP", "MP"),
                                checkboxTexto(data.tipo_mtt === "MC", "MC"),
                                checkboxTexto(data.tipo_mtt === "MC_MP", "MP/MC"),
                                checkboxTexto(data.tipo_mtt === "NA", "No aplica"),
                            ].join("     "),
                            size: 20,
                        })],
                        spacing: { after: 200 },
                    }),

                    new Paragraph({ children: [new TextRun({ text: "OBSERVACIONES", bold: true, size: 22 })], spacing: { after: 100 } }),
                    new Paragraph({ children: [new TextRun({ text: data.observaciones || "-", size: 20 })], spacing: { after: 300 } }),

                    // --- Firmas ---
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                                children: [
                                    filaEtiqueta("Recibe Orden"),
                                    filaValor(data.recibe_orden),
                                ],
                            }),
                            new TableRow({
                                children: [
                                    filaEtiqueta("Realiza Trabajo"),
                                    filaValor(data.realiza_trabajo),
                                ],
                            }),
                            new TableRow({
                                children: [
                                    filaEtiqueta("Recibe Trabajo"),
                                    filaValor(data.recibe_trabajo),
                                ],
                            }),
                            new TableRow({
                                children: [
                                    filaEtiqueta("Jefe de Taller"),
                                    filaValor(data.jefe_taller),
                                ],
                            }),
                        ],
                    }),
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
    drawTextWrapped(firstPage, data.descripcion, { x: 50, y: 550, maxWidth: 500, maxLines: 5 });

    // Descripción del Trabajo Realizado
    drawTextWrapped(firstPage, data.trabajo_realizado, { x: 50, y: 460, maxWidth: 500, maxLines: 5 });

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