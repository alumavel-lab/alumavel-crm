// Descarga un archivo de Firebase Storage (del bucket del CRM) desde el servidor
// y lo devuelve en base64. Hace falta porque el navegador NO puede descargar
// directamente de Firebase Storage con fetch() (el bucket no tiene CORS
// configurado → "Failed to fetch"). Desde el servidor no hay CORS, así que funciona.
//
// Si el navegador manda el token del usuario (cabecera Authorization), se
// reenvía a Storage para que funcione también con reglas de Storage cerradas.
// Solo acepta URLs del bucket del CRM (no es un proxy abierto).
//
// Archivos grandes: las respuestas de Netlify tienen un tope de ~6 MB, así que el
// navegador pide el archivo POR TROZOS (parámetro "desde" = byte por el que empezar).
// Cada respuesta trae "total" (peso completo del archivo) para que sepa cuántos
// trozos le faltan. Sin "desde" funciona como antes (archivo entero, máximo 4 MB).

const STORAGE_PREFIX = "https://firebasestorage.googleapis.com/v0/b/crmalumavel.firebasestorage.app/o/";
const LIMITE_BYTES = 4.2 * 1024 * 1024; // Netlify corta las respuestas de ~6 MB (base64 ocupa un 33% más)

export const handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: JSON.stringify({ error: "Method not allowed" }) };
  }
  try {
    const { url, desde } = JSON.parse(event.body || "{}");
    if (!url || !url.startsWith(STORAGE_PREFIX)) {
      return { statusCode: 400, body: JSON.stringify({ error: "URL no permitida" }) };
    }
    const token = event.headers?.authorization || event.headers?.Authorization || "";
    const headers = token ? { Authorization: token.replace(/^Bearer\s+/i, "Firebase ") } : {};
    const porTrozos = Number.isFinite(desde) && desde >= 0;
    const inicio = porTrozos ? Math.floor(desde) : 0;
    const fin = inicio + Math.floor(LIMITE_BYTES) - 1;

    const res = await fetch(url, { headers: porTrozos ? { ...headers, Range: `bytes=${inicio}-${fin}` } : headers });
    if (!res.ok) {
      const texto = await res.text().catch(() => "");
      console.error("Storage devolvió", res.status, texto.slice(0, 300));
      return { statusCode: 502, body: JSON.stringify({ error: `Firebase Storage devolvió ${res.status}${res.status === 403 ? " (sin permiso)" : res.status === 404 ? " (el archivo ya no existe)" : ""}` }) };
    }
    const buffer = Buffer.from(await res.arrayBuffer());
    const contentType = res.headers.get("content-type") || "";

    if (!porTrozos) {
      if (buffer.length > LIMITE_BYTES) {
        return { statusCode: 413, body: JSON.stringify({ error: `El archivo pesa ${(buffer.length / 1048576).toFixed(1)} MB y el máximo para firmar es 4 MB — comprímelo y vuelve a subirlo` }) };
      }
      return {
        statusCode: 200,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ base64: buffer.toString("base64"), contentType }),
      };
    }

    // Por trozos. Si Storage respeta "Range" (206) el cuerpo ya es el trozo y el peso
    // total viene en Content-Range ("bytes 0-4404018/5767168"). Si lo ignora (200),
    // llega el archivo entero y se corta aquí.
    let trozo, total;
    if (res.status === 206) {
      trozo = buffer;
      total = parseInt((res.headers.get("content-range") || "").split("/")[1], 10);
      if (!Number.isFinite(total)) total = inicio + buffer.length; // sin dato fiable: se da por terminado
    } else {
      total = buffer.length;
      trozo = buffer.subarray(inicio, Math.min(total, fin + 1));
    }
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ base64: trozo.toString("base64"), contentType, total, desde: inicio, bytes: trozo.length }),
    };
  } catch (err) {
    console.error("Excepción en descargar-archivo:", err.message);
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
