// Parte lenta del aviso de firma completada — la llama firma-webhook.js
// (que responde rápido a Firma.dev) pasándole solo el signingRequestId. Al
// llevar "-background" en el nombre, Netlify la ejecuta en segundo plano con
// mucho más tiempo disponible (hasta 15 minutos), así que aquí sí se puede
// esperar todo lo que haga falta a que Firma.dev tenga listo el PDF firmado,
// en vez de rendirse a los pocos segundos.
//
// Variables de entorno necesarias en Netlify (las mismas que ya usan
// firma-webhook.js y firma-enviar.js):
//   FIRMA_API_KEY_TEST / FIRMA_API_KEY

const FIRMA_API = "https://api.firma.dev/functions/v1/signing-request-api";
const FIREBASE_DB_URL = "https://crmalumavel-default-rtdb.europe-west1.firebasedatabase.app";
// Clave secreta de la base de datos (variable FIREBASE_DB_SECRET en Netlify): hace falta
// para leer/escribir una vez cerradas las reglas de Firebase. Sin ella funciona igual mientras
// las reglas sigan abiertas.
const FB_AUTH = process.env.FIREBASE_DB_SECRET ? `?auth=${encodeURIComponent(process.env.FIREBASE_DB_SECRET)}` : "";
const STORAGE_BUCKET = "crmalumavel.firebasestorage.app";

const toArray = (obj) => (obj ? Object.values(obj) : []);

// Busca en qué colección (presupuestos o proyectos) y qué registro tiene este
// signingRequestId guardado en su campo "firma" — no lo sabemos de antemano
// porque el webhook solo trae el id de la solicitud de firma. OJO: el CRM
// guarda cada colección como una lista con posiciones (0,1,2...) — la clave
// real de Firebase NO es el id del registro, así que hay que devolverla aparte
// para poder escribir luego en el sitio correcto.
async function buscarRegistroPorSigningRequestId(signingRequestId) {
  for (const coleccion of ["presupuestos", "proyectos"]) {
    const res = await fetch(`${FIREBASE_DB_URL}/${coleccion}.json${FB_AUTH}`);
    const datos = await res.json();
    if (!datos) continue;
    const clave = Object.keys(datos).find((k) => datos[k]?.firma?.signingRequestId === signingRequestId);
    if (clave) return { coleccion, clave, registro: datos[clave] };
  }
  return null;
}

export const handler = async (event) => {
  try {
    const { signingRequestId } = JSON.parse(event.body || "{}");
    if (!signingRequestId) {
      console.error("firma-webhook-background sin signingRequestId");
      return;
    }

    const encontrado = await buscarRegistroPorSigningRequestId(signingRequestId);
    if (!encontrado) {
      console.error("No se encontró ningún registro con signingRequestId:", signingRequestId);
      return;
    }
    const { coleccion, clave, registro } = encontrado;

    const apiKey = process.env.FIRMA_API_KEY || process.env.FIRMA_API_KEY_TEST;
    if (!apiKey) {
      console.error("Falta FIRMA_API_KEY_TEST (o FIRMA_API_KEY) en Netlify");
      return;
    }

    // Pide la URL de descarga del PDF ya firmado. Justo al completarse la firma,
    // Firma.dev a veces todavía no tiene el documento final listo (da
    // "no_document_available" o 503) — aquí hay margen de sobra para esperar,
    // así que se reintenta bastantes veces antes de rendirse de verdad.
    let descargaData = null;
    for (let intento = 1; intento <= 12; intento++) {
      const descargaRes = await fetch(`${FIRMA_API}/signing-requests/${signingRequestId}/download`, {
        headers: { Authorization: apiKey },
      });
      descargaData = await descargaRes.json();
      if (descargaRes.ok && descargaData.download_url) break;
      console.error(`Intento ${intento}/12: PDF firmado todavía no disponible:`, JSON.stringify(descargaData));
      descargaData = null;
      if (intento < 12) await new Promise((r) => setTimeout(r, 5000));
    }

    if (!descargaData) {
      // Ni siquiera con más de un minuto de reintentos se consiguió el PDF —
      // se marca como "Firmado" (sin PDF) para no dejar el presupuesto colgado,
      // pero queda registrado en los logs para revisarlo a mano si hace falta.
      console.error(`No se pudo obtener el PDF firmado tras 12 intentos para signingRequestId ${signingRequestId}`);
      const firmaSinPdf = { ...(registro.firma || {}), estado: "firmado", firmadoEn: Date.now() };
      await fetch(`${FIREBASE_DB_URL}/${coleccion}/${clave}/firma.json${FB_AUTH}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(firmaSinPdf),
      });
      return;
    }

    // Descarga el PDF real
    const pdfRes = await fetch(descargaData.download_url);
    const pdfBuffer = Buffer.from(await pdfRes.arrayBuffer());

    // Sube el PDF a Firebase Storage (carpeta presupuestos-firmados/)
    const nombreArchivo = `presupuestos-firmados/${coleccion}-${registro.id}-${signingRequestId}.pdf`;
    const subidaRes = await fetch(
      `https://firebasestorage.googleapis.com/v0/b/${STORAGE_BUCKET}/o?uploadType=media&name=${encodeURIComponent(nombreArchivo)}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/pdf" },
        body: pdfBuffer,
      }
    );
    if (!subidaRes.ok) {
      const errorSubida = await subidaRes.text();
      console.error("Error subiendo el PDF a Firebase Storage:", errorSubida);
      // Aun sin PDF guardado, se marca como firmado para no dejarlo colgado.
      const firmaSinPdf = { ...(registro.firma || {}), estado: "firmado", firmadoEn: Date.now() };
      await fetch(`${FIREBASE_DB_URL}/${coleccion}/${clave}/firma.json${FB_AUTH}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(firmaSinPdf),
      });
      return;
    }

    const pdfUrl = `https://firebasestorage.googleapis.com/v0/b/${STORAGE_BUCKET}/o/${encodeURIComponent(nombreArchivo)}?alt=media`;

    // Actualiza el registro: estado "firmado" + URL del PDF
    const firmaActualizada = {
      ...(registro.firma || {}),
      estado: "firmado",
      firmadoEn: Date.now(),
      pdfUrl,
    };
    await fetch(`${FIREBASE_DB_URL}/${coleccion}/${clave}/firma.json${FB_AUTH}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(firmaActualizada),
    });

    // Si el presupuesto firmado pertenece a una obra con contrato ya firmado con
    // constructora (aprobación interna), pasamos el presupuesto a Proyecto
    // automáticamente en cuanto el responsable firma — sin que nadie tenga que
    // entrar a moverlo a mano. Si es un presupuesto normal (firma del cliente),
    // no se toca nada más: sigue el flujo manual de siempre ("Crear proyecto").
    if (coleccion === "presupuestos" && !registro.proyectoCreadoId) {
      const proyectosRes = await fetch(`${FIREBASE_DB_URL}/proyectos.json${FB_AUTH}`);
      const proyectos = toArray(await proyectosRes.json());
      const proyectoVinculado = proyectos.find((p) => p.id === registro.proyectoId);

      if (proyectoVinculado?.contratoConstructoraFirmado) {
        const nums = proyectos.map((p) => parseInt(String(p.numero).replace(/\D/g, ""), 10)).filter((n) => !isNaN(n));
        const siguienteNumero = String((nums.length ? Math.max(...nums) : 4189) + 1);
        const nuevoProyectoId = `${signingRequestId.slice(0, 8)}${Date.now().toString(36)}`;
        const documentosProyecto = [...(registro.documentos || [])];
        if (pdfUrl) {
          documentosProyecto.push({ id: `${nuevoProyectoId}-firma`, nombre: `Presupuesto ${registro.numero} — firmado.pdf`, url: pdfUrl, subidoEn: Date.now() });
        }

        const nuevoProyecto = {
          id: nuevoProyectoId,
          numero: siguienteNumero,
          nombre: registro.descripcion || registro.numero,
          clienteId: proyectoVinculado.clienteId || "",
          importePresupuesto: registro.importe || 0,
          estadoPresupuesto: "Presupuesto aceptado",
          estadoTrabajo: "Pendiente de aceptación",
          gastos: [],
          registroHorario: [],
          checklistMateriales: {}, // simplificado — se puede editar luego desde la ficha del proyecto
          documentos: documentosProyecto,
        };
        await fetch(`${FIREBASE_DB_URL}/proyectos/${nuevoProyectoId}.json${FB_AUTH}`, {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(nuevoProyecto),
        });
        await fetch(`${FIREBASE_DB_URL}/presupuestos/${clave}.json${FB_AUTH}`, {
          method: "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ estado: "Aceptado", proyectoCreadoId: nuevoProyectoId }),
        });
      }
    }
  } catch (err) {
    console.error("Excepción en firma-webhook-background:", err.message);
    console.error("Stack:", err.stack);
  }
};
