// Envía un documento (presupuesto) a firmar a través de Firma.dev.
// Se usa tanto para la firma del cliente como para la aprobación interna del
// responsable de una obra con contrato ya firmado con la constructora — el
// mecanismo es el mismo (un enlace de firma con validez legal eIDAS), solo
// cambia a quién se le manda: al cliente, o al responsable interno de la obra.
//
// Variables de entorno necesarias en Netlify (Site configuration > Environment variables):
//   FIRMA_API_KEY_TEST = clave de prueba de Firma.dev (gratis, ilimitada, para probar)
//   FIRMA_API_KEY       = clave real (cuando se pase a producción; genera coste
//                         de ~0,029€ por documento enviado)
// Se usa la clave "live" (FIRMA_API_KEY) si existe; si no, cae a la de prueba,
// para no romper nada mientras se está probando el flujo sin gastar.
//
// Guarda el estado "enviado a firma" directamente en el registro de Firebase
// (presupuestos/<id>/firma), usando la misma base de datos REST que ya usan
// el resto de funciones del CRM (informe-semanal-instalaciones.js, etc.).

const FIRMA_API = "https://api.firma.dev/functions/v1/signing-request-api";
// PDF grandes: el navegador no puede mandarlos por aquí (las peticiones a funciones de
// Netlify tienen un tope de ~6 MB), así que los sube a Firebase Storage y manda solo la
// dirección (pdfUrl). La función lo descarga desde el servidor y se lo pasa a Firma.dev.
const STORAGE_PREFIX = "https://firebasestorage.googleapis.com/v0/b/crmalumavel.firebasestorage.app/o/";
const LIMITE_FIRMA_BYTES = 19 * 1024 * 1024; // Firma.dev admite 20 MB por documento
const FIREBASE_DB_URL = "https://crmalumavel-default-rtdb.europe-west1.firebasedatabase.app";
// Clave secreta de la base de datos (variable FIREBASE_DB_SECRET en Netlify): hace falta
// para leer/escribir una vez cerradas las reglas de Firebase. Sin ella funciona igual mientras
// las reglas sigan abiertas.
// Acceso del servidor a la base de datos. Preferido: el usuario "robot" del CRM
// (variables FIREBASE_ROBOT_EMAIL y FIREBASE_ROBOT_PASSWORD en Netlify), que entra
// como un usuario más del equipo. Si no están, se usa la clave antigua
// FIREBASE_DB_SECRET (Firebase ya no la acepta en este proyecto).
const FIREBASE_WEB_API_KEY = "AIzaSyBf51Gy2drHdUyJ4kCstNdvvV2h3uYe4RM";
let _tokenRobot = null, _tokenRobotCaduca = 0;
async function fbAuthQuery() {
  const email = (process.env.FIREBASE_ROBOT_EMAIL || "").trim();
  const password = (process.env.FIREBASE_ROBOT_PASSWORD || "").trim();
  if (email && password) {
    if (_tokenRobot && Date.now() < _tokenRobotCaduca) return `?auth=${encodeURIComponent(_tokenRobot)}`;
    const r = await fetch(`https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=${FIREBASE_WEB_API_KEY}`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, password, returnSecureToken: true }),
    });
    const d = await r.json().catch(() => ({}));
    if (r.ok && d.idToken) {
      _tokenRobot = d.idToken;
      _tokenRobotCaduca = Date.now() + 50 * 60 * 1000; // el token dura 1 h
      return `?auth=${encodeURIComponent(_tokenRobot)}`;
    }
    console.error("No se pudo entrar con el usuario robot:", JSON.stringify(d.error || d));
  }
  const secreto = (process.env.FIREBASE_DB_SECRET || "").trim();
  return secreto ? `?auth=${encodeURIComponent(secreto)}` : "";
}
let FB_AUTH = "";

export const handler = async (event) => {
  FB_AUTH = await fbAuthQuery().catch(() => "");
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: JSON.stringify({ error: "Method not allowed" }) };
  }

  try {
    const body = JSON.parse(event.body || "{}");
    const { registroTipo, registroId, pdfBase64, pdfUrl, nombreDocumento, firmante } = body;

    if (!registroTipo || !registroId || (!pdfBase64 && !pdfUrl) || !firmante?.email) {
      return { statusCode: 400, body: JSON.stringify({ error: "Faltan datos: registroTipo, registroId, pdfBase64 (o pdfUrl) o firmante.email." }) };
    }
    if (pdfUrl && !pdfUrl.startsWith(STORAGE_PREFIX)) {
      return { statusCode: 400, body: JSON.stringify({ error: "pdfUrl no permitida" }) };
    }
    if (registroTipo !== "presupuestos" && registroTipo !== "proyectos") {
      return { statusCode: 400, body: JSON.stringify({ error: "registroTipo debe ser 'presupuestos' o 'proyectos'." }) };
    }

    // Acceso a la base de datos: se usa el token del usuario que ha iniciado sesión
    // en el CRM (lo manda el navegador). Si no viene, se usa FIREBASE_DB_SECRET.
    const tokenUsuario = (event.headers?.authorization || event.headers?.Authorization || "").replace(/^Bearer\s+/i, "");
    const dbAuth = tokenUsuario ? `?auth=${encodeURIComponent(tokenUsuario)}` : FB_AUTH;

    // Se comprueba que el registro existe ANTES de crear la solicitud en Firma.dev,
    // para no gastar un envío si luego no se puede guardar. OJO: el CRM guarda
    // presupuestos/proyectos como una lista (índices 0,1,2...), no con el id como
    // clave — hay que traer la colección y buscar el índice real.
    const getColRes = await fetch(`${FIREBASE_DB_URL}/${registroTipo}.json${dbAuth}`);
    const coleccion = await getColRes.json().catch(() => null);
    if (!getColRes.ok || coleccion?.error) {
      console.error("Firebase rechazó la lectura:", getColRes.status, JSON.stringify(coleccion));
      return { statusCode: 500, body: JSON.stringify({ error: `El servidor no tiene permiso para leer la base de datos (${coleccion?.error || getColRes.status}). Cierra sesión y vuelve a entrar; si sigue, revisa FIREBASE_DB_SECRET en Netlify.` }) };
    }
    const claveReal = coleccion ? Object.keys(coleccion).find((k) => coleccion[k]?.id === registroId) : null;
    if (!claveReal) {
      return { statusCode: 404, body: JSON.stringify({ error: "No se encontró el presupuesto en la base de datos (¿está recién creado? espera unos segundos y reintenta)." }) };
    }

    const apiKey = process.env.FIRMA_API_KEY || process.env.FIRMA_API_KEY_TEST;
    if (!apiKey) {
      console.error("Falta FIRMA_API_KEY_TEST (o FIRMA_API_KEY) en Netlify");
      return { statusCode: 500, body: JSON.stringify({ error: "La firma electrónica no está configurada todavía (falta FIRMA_API_KEY_TEST en Netlify)." }) };
    }

    // PDF grande: se descarga aquí, desde el servidor (sin el tope de 6 MB del navegador).
    let documentoBase64 = pdfBase64;
    if (!documentoBase64 && pdfUrl) {
      const authStorage = tokenUsuario ? { Authorization: `Firebase ${tokenUsuario}` } : {};
      const pdfRes = await fetch(pdfUrl, { headers: authStorage });
      if (!pdfRes.ok) {
        console.error("No se pudo leer el PDF subido:", pdfRes.status);
        return { statusCode: 502, body: JSON.stringify({ error: `No se pudo leer el PDF grande de Storage (${pdfRes.status})` }) };
      }
      const buf = Buffer.from(await pdfRes.arrayBuffer());
      if (buf.length > LIMITE_FIRMA_BYTES) {
        return { statusCode: 413, body: JSON.stringify({ error: `El PDF final pesa ${(buf.length / 1048576).toFixed(1)} MB y Firma.dev admite 20 MB como máximo` }) };
      }
      documentoBase64 = buf.toString("base64");
    }

    const partesNombre = (firmante.nombre || "Firmante").trim().split(" ");
    const nombre = partesNombre[0] || "Firmante";
    const apellido = partesNombre.slice(1).join(" ") || "-";

    const respuestaFirma = await fetch(`${FIRMA_API}/signing-requests/create-and-send`, {
      method: "POST",
      headers: { Authorization: apiKey, "Content-Type": "application/json" },
      body: JSON.stringify({
        name: nombreDocumento || `Documento ${registroId}`,
        document: documentoBase64,
        recipients: [
          {
            first_name: nombre,
            last_name: apellido,
            email: firmante.email,
            designation: "Signer",
          },
        ],
      }),
    });

    // El PDF grande subido solo para este envío ya no hace falta: se borra (si falla, no pasa nada).
    if (pdfUrl) {
      try {
        const nombreTmp = decodeURIComponent(pdfUrl.slice(STORAGE_PREFIX.length).split("?")[0]);
        if (nombreTmp.startsWith("firmas-temporales/")) {
          await fetch(`${STORAGE_PREFIX}${encodeURIComponent(nombreTmp)}`, { method: "DELETE", headers: tokenUsuario ? { Authorization: `Firebase ${tokenUsuario}` } : {} });
        }
      } catch (e) { console.warn("No se pudo borrar el PDF temporal:", e.message); }
    }

    const dataFirma = await respuestaFirma.json();
    if (!respuestaFirma.ok) {
      console.error("Firma.dev devolvió error:", JSON.stringify(dataFirma));
      return { statusCode: 502, body: JSON.stringify({ error: "Firma.dev rechazó la solicitud", detalle: dataFirma }) };
    }

    const signingRequestId = dataFirma.id || dataFirma.signing_request?.id;
    if (!signingRequestId) {
      console.error("Firma.dev no devolvió un id de solicitud:", JSON.stringify(dataFirma));
      return { statusCode: 502, body: JSON.stringify({ error: "Firma.dev no devolvió un id de solicitud", detalle: dataFirma }) };
    }
    // El enlace real de firma (para poder compartirlo por WhatsApp) viene en
    // first_signer.signing_link — si Firma.dev cambia el formato, se prueban un
    // par de sitios más por si acaso, en vez de fallar en silencio.
    const signingLink = dataFirma.first_signer?.signing_link || dataFirma.signing_link || dataFirma.recipients?.[0]?.signing_link || "";

    const firmaInfo = {
      signingRequestId,
      estado: "enviado",
      enviadoEn: Date.now(),
      firmanteNombre: firmante.nombre || "",
      firmanteEmail: firmante.email,
      signingLink,
    };

    await fetch(`${FIREBASE_DB_URL}/${registroTipo}/${claveReal}/firma.json${dbAuth}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(firmaInfo),
    });

    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ok: true, signingRequestId, signingLink }),
    };
  } catch (err) {
    console.error("Excepción en firma-enviar:", err.message);
    console.error("Stack:", err.stack);
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
