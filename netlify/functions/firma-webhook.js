// Recibe el aviso de Firma.dev cuando un documento se termina de firmar,
// descarga el PDF firmado y lo sube a Firebase Storage. Guarda solo la URL en
// el registro de Firebase Realtime Database (no el archivo entero), igual que
// el resto de archivos que ya maneja el CRM.
//
// CONFIGURACIÓN PENDIENTE (hacer una vez, desde el panel de Firma.dev):
//   1. Registra este webhook en Firma.dev apuntando a:
//        https://dynamic-eclair-be67a5.netlify.app/.netlify/functions/firma-webhook
//      con el evento "signing_request.completed".
//   2. Copia el "signing secret" que te da Firma.dev al crear el webhook y
//      guárdalo en Netlify como FIRMA_WEBHOOK_SECRET.
//   3. En Firebase Console > Storage, activa Storage si no lo está, y pon estas
//      reglas (Storage > Rules) para permitir que esta función suba/lea los PDFs:
//        rules_version = '2';
//        service firebase.storage {
//          match /b/{bucket}/o {
//            match /presupuestos-firmados/{allPaths=**} {
//              allow read, write: if true;
//            }
//          }
//        }
//      (El acceso real de esta carpeta queda igual de abierto que la base de
//      datos actual del CRM — es el mismo modelo de seguridad que ya usáis.)
//
// Variables de entorno necesarias en Netlify:
//   FIRMA_API_KEY_TEST / FIRMA_API_KEY = igual que en firma-enviar.js
//   FIRMA_WEBHOOK_SECRET               = el secreto que da Firma.dev al crear el webhook

import crypto from "crypto";

const FIRMA_API = "https://api.firma.dev/functions/v1/signing-request-api";
const FIREBASE_DB_URL = "https://crmalumavel-default-rtdb.europe-west1.firebasedatabase.app";
const STORAGE_BUCKET = "crmalumavel.firebasestorage.app";

const toArray = (obj) => (obj ? Object.values(obj) : []);

function verificarFirmaWebhook(rawBody, headerSignature, secret) {
  if (!secret || !headerSignature) return false;
  // Formato: "t=1707500000,v1=abc123..."
  const partes = Object.fromEntries(headerSignature.split(",").map((p) => p.split("=")));
  const { t, v1 } = partes;
  if (!t || !v1) return false;
  const payloadFirmado = `${t}.${rawBody}`;
  const esperado = crypto.createHmac("sha256", secret).update(payloadFirmado).digest("hex");
  try {
    return crypto.timingSafeEqual(Buffer.from(esperado), Buffer.from(v1));
  } catch {
    return false;
  }
}

// Busca en qué colección (presupuestos o proyectos) y qué registro tiene este
// signingRequestId guardado en su campo "firma" — no lo sabemos de antemano
// porque el webhook solo trae el id de la solicitud de firma.
async function buscarRegistroPorSigningRequestId(signingRequestId) {
  for (const coleccion of ["presupuestos", "proyectos"]) {
    const res = await fetch(`${FIREBASE_DB_URL}/${coleccion}.json`);
    const datos = toArray(await res.json());
    const encontrado = datos.find((r) => r?.firma?.signingRequestId === signingRequestId);
    if (encontrado) return { coleccion, registro: encontrado };
  }
  return null;
}

export const handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method not allowed" };
  }

  try {
    const rawBody = event.body || "";
    const secret = process.env.FIRMA_WEBHOOK_SECRET;
    const headerSignature = event.headers["x-firma-signature"] || event.headers["X-Firma-Signature"];

    if (secret) {
      const valido = verificarFirmaWebhook(rawBody, headerSignature, secret);
      if (!valido) {
        console.error("Firma de webhook inválida — posible petición falsa");
        return { statusCode: 401, body: JSON.stringify({ error: "Firma de webhook inválida" }) };
      }
    } else {
      // Sin FIRMA_WEBHOOK_SECRET configurado todavía no se puede verificar el
      // origen — se procesa igualmente para no bloquear las pruebas iniciales,
      // pero esto debe configurarse antes de pasar a producción real.
      console.warn("FIRMA_WEBHOOK_SECRET no configurado — el webhook no está verificado.");
    }

    const payload = JSON.parse(rawBody);
    const { type, data } = payload;

    if (type !== "signing_request.completed") {
      // Otros eventos (viewed, sent, recipient.signed...) se reciben pero se
      // ignoran por ahora — solo nos interesa cuando TODOS han firmado.
      return { statusCode: 200, body: JSON.stringify({ recibido: true, ignorado: type }) };
    }

    const signingRequestId = data?.signing_request?.id || data?.id;
    if (!signingRequestId) {
      console.error("Webhook signing_request.completed sin id:", JSON.stringify(payload));
      return { statusCode: 200, body: JSON.stringify({ error: "Sin signing_request.id en el payload" }) };
    }

    const encontrado = await buscarRegistroPorSigningRequestId(signingRequestId);
    if (!encontrado) {
      console.error("No se encontró ningún registro con signingRequestId:", signingRequestId);
      return { statusCode: 200, body: JSON.stringify({ error: "Registro no encontrado para este signingRequestId" }) };
    }
    const { coleccion, registro } = encontrado;

    const apiKey = process.env.FIRMA_API_KEY || process.env.FIRMA_API_KEY_TEST;
    if (!apiKey) {
      console.error("Falta FIRMA_API_KEY_TEST (o FIRMA_API_KEY) en Netlify");
      return { statusCode: 200, body: JSON.stringify({ error: "Falta FIRMA_API_KEY_TEST en Netlify" }) };
    }

    // Pide la URL de descarga del PDF ya firmado
    const descargaRes = await fetch(`${FIRMA_API}/signing-requests/${signingRequestId}/download`, {
      headers: { Authorization: apiKey },
    });
    const descargaData = await descargaRes.json();
    if (!descargaRes.ok || !descargaData.download_url) {
      console.error("No se pudo obtener el PDF firmado:", JSON.stringify(descargaData));
      return { statusCode: 200, body: JSON.stringify({ error: "No se pudo descargar el PDF firmado", detalle: descargaData }) };
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
      return { statusCode: 200, body: JSON.stringify({ error: "No se pudo subir el PDF a Storage", detalle: errorSubida }) };
    }

    const pdfUrl = `https://firebasestorage.googleapis.com/v0/b/${STORAGE_BUCKET}/o/${encodeURIComponent(nombreArchivo)}?alt=media`;

    // Actualiza el registro: estado "firmado" + URL del PDF
    const firmaActualizada = {
      ...(registro.firma || {}),
      estado: "firmado",
      firmadoEn: Date.now(),
      pdfUrl,
    };
    await fetch(`${FIREBASE_DB_URL}/${coleccion}/${registro.id}/firma.json`, {
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
      const proyectosRes = await fetch(`${FIREBASE_DB_URL}/proyectos.json`);
      const proyectos = toArray(await proyectosRes.json());
      const proyectoVinculado = proyectos.find((p) => p.id === registro.proyectoId);

      if (proyectoVinculado?.contratoConstructoraFirmado) {
        const nums = proyectos.map((p) => parseInt(String(p.numero).replace(/\D/g, ""), 10)).filter((n) => !isNaN(n));
        const siguienteNumero = String((nums.length ? Math.max(...nums) : 4189) + 1);
        const nuevoProyectoId = `${signingRequestId.slice(0, 8)}${Date.now().toString(36)}`;

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
        };
        await fetch(`${FIREBASE_DB_URL}/proyectos/${nuevoProyectoId}.json`, {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(nuevoProyecto),
        });
        await fetch(`${FIREBASE_DB_URL}/presupuestos/${registro.id}.json`, {
          method: "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ estado: "Aceptado", proyectoCreadoId: nuevoProyectoId }),
        });
      }
    }

    return { statusCode: 200, body: JSON.stringify({ ok: true, coleccion, registroId: registro.id, pdfUrl }) };
  } catch (err) {
    console.error("Excepción en firma-webhook:", err.message);
    console.error("Stack:", err.stack);
    // Devolvemos 200 igualmente para que Firma.dev no reintente indefinidamente
    // por un error nuestro — el fallo queda registrado en los logs de Netlify.
    return { statusCode: 200, body: JSON.stringify({ error: err.message }) };
  }
};
