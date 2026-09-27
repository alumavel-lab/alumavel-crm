// Descarga un archivo de Firebase Storage (del bucket del CRM) desde el servidor
// y lo devuelve en base64. Hace falta porque el navegador NO puede descargar
// directamente de Firebase Storage con fetch() (el bucket no tiene CORS
// configurado → "Failed to fetch"). Desde el servidor no hay CORS, así que funciona.
//
// Si el navegador manda el token del usuario (cabecera Authorization), se
// reenvía a Storage para que funcione también con reglas de Storage cerradas.
// Solo acepta URLs del bucket del CRM (no es un proxy abierto).

const STORAGE_PREFIX = "https://firebasestorage.googleapis.com/v0/b/crmalumavel.firebasestorage.app/o/";
const LIMITE_BYTES = 4.2 * 1024 * 1024; // Netlify corta las respuestas de ~6 MB (base64 ocupa un 33% más)

export const handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: JSON.stringify({ error: "Method not allowed" }) };
  }
  try {
    const { url } = JSON.parse(event.body || "{}");
    if (!url || !url.startsWith(STORAGE_PREFIX)) {
      return { statusCode: 400, body: JSON.stringify({ error: "URL no permitida" }) };
    }
    const token = event.headers?.authorization || event.headers?.Authorization || "";
    const headers = token ? { Authorization: token.replace(/^Bearer\s+/i, "Firebase ") } : {};

    const res = await fetch(url, { headers });
    if (!res.ok) {
      const texto = await res.text().catch(() => "");
      console.error("Storage devolvió", res.status, texto.slice(0, 300));
      return { statusCode: 502, body: JSON.stringify({ error: `Firebase Storage devolvió ${res.status}${res.status === 403 ? " (sin permiso)" : res.status === 404 ? " (el archivo ya no existe)" : ""}` }) };
    }
    const buffer = Buffer.from(await res.arrayBuffer());
    if (buffer.length > LIMITE_BYTES) {
      return { statusCode: 413, body: JSON.stringify({ error: `El archivo pesa ${(buffer.length / 1048576).toFixed(1)} MB y el máximo para firmar es 4 MB — comprímelo y vuelve a subirlo` }) };
    }
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ base64: buffer.toString("base64"), contentType: res.headers.get("content-type") || "" }),
    };
  } catch (err) {
    console.error("Excepción en descargar-archivo:", err.message);
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
