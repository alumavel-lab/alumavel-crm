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
const FIREBASE_DB_URL = "https://crmalumavel-default-rtdb.europe-west1.firebasedatabase.app";
// Clave secreta de la base de datos (variable FIREBASE_DB_SECRET en Netlify): hace falta
// para leer/escribir una vez cerradas las reglas de Firebase. Sin ella funciona igual mientras
// las reglas sigan abiertas.
const FB_AUTH = process.env.FIREBASE_DB_SECRET ? `?auth=${encodeURIComponent(process.env.FIREBASE_DB_SECRET)}` : "";

export const handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: JSON.stringify({ error: "Method not allowed" }) };
  }

  try {
    const body = JSON.parse(event.body || "{}");
    const { registroTipo, registroId, pdfBase64, nombreDocumento, firmante } = body;

    if (!registroTipo || !registroId || !pdfBase64 || !firmante?.email) {
      return { statusCode: 400, body: JSON.stringify({ error: "Faltan datos: registroTipo, registroId, pdfBase64 o firmante.email." }) };
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

    const partesNombre = (firmante.nombre || "Firmante").trim().split(" ");
    const nombre = partesNombre[0] || "Firmante";
    const apellido = partesNombre.slice(1).join(" ") || "-";

    const respuestaFirma = await fetch(`${FIRMA_API}/signing-requests/create-and-send`, {
      method: "POST",
      headers: { Authorization: apiKey, "Content-Type": "application/json" },
      body: JSON.stringify({
        name: nombreDocumento || `Documento ${registroId}`,
        document: pdfBase64,
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
