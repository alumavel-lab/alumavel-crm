// Comprueba en Firma.dev si un presupuesto enviado a firmar ya está firmado y,
// si lo está, descarga el PDF firmado y lo guarda en Firebase Storage
// (carpeta presupuestos-firmados/). Devuelve la URL para que el CRM la guarde.
//
// Es la vía "manual" (botón en el CRM y comprobación al abrir el presupuesto):
// funciona aunque el webhook de Firma.dev no esté configurado o falle, porque
// no necesita tocar la base de datos — eso lo hace el propio navegador.

const FIRMA_API = "https://api.firma.dev/functions/v1/signing-request-api";
const STORAGE_BUCKET = "crmalumavel.firebasestorage.app";

export const handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: JSON.stringify({ error: "Method not allowed" }) };
  }
  try {
    const { signingRequestId, registroId } = JSON.parse(event.body || "{}");
    if (!signingRequestId) {
      return { statusCode: 400, body: JSON.stringify({ error: "Falta signingRequestId" }) };
    }
    const apiKey = process.env.FIRMA_API_KEY || process.env.FIRMA_API_KEY_TEST;
    if (!apiKey) {
      return { statusCode: 500, body: JSON.stringify({ error: "Falta FIRMA_API_KEY_TEST (o FIRMA_API_KEY) en Netlify" }) };
    }

    const r = await fetch(`${FIRMA_API}/signing-requests/${signingRequestId}/download`, { headers: { Authorization: apiKey } });
    const d = await r.json().catch(() => ({}));

    // 409 = enviado pero sin documento todavía; 503 = generándose el PDF final
    if (r.status === 409 || r.status === 503) {
      return { statusCode: 200, headers: { "Content-Type": "application/json" }, body: JSON.stringify({ estado: r.status === 503 ? "generando" : "pendiente" }) };
    }
    if (!r.ok) {
      console.error("Firma.dev download error:", r.status, JSON.stringify(d));
      return { statusCode: 502, body: JSON.stringify({ error: `Firma.dev respondió ${r.status}${d.error ? ` (${d.error})` : ""}` }) };
    }
    if (d.status !== "finished") {
      // in_progress / cancelled / declined / expired
      return { statusCode: 200, headers: { "Content-Type": "application/json" }, body: JSON.stringify({ estado: d.status === "in_progress" ? "pendiente" : d.status }) };
    }

    // Firmado: descargar el PDF final y guardarlo en Storage
    const pdfRes = await fetch(d.download_url);
    if (!pdfRes.ok) {
      return { statusCode: 200, headers: { "Content-Type": "application/json" }, body: JSON.stringify({ estado: "firmado", pdfUrl: "", aviso: "Está firmado, pero no se pudo descargar el PDF todavía — vuelve a comprobarlo en un rato" }) };
    }
    const pdfBuffer = Buffer.from(await pdfRes.arrayBuffer());
    const nombreArchivo = `presupuestos-firmados/presupuestos-${registroId || "x"}-${signingRequestId}.pdf`;
    const subida = await fetch(
      `https://firebasestorage.googleapis.com/v0/b/${STORAGE_BUCKET}/o?uploadType=media&name=${encodeURIComponent(nombreArchivo)}`,
      { method: "POST", headers: { "Content-Type": "application/pdf" }, body: pdfBuffer }
    );
    if (!subida.ok) {
      const txt = await subida.text().catch(() => "");
      console.error("Error subiendo PDF firmado a Storage:", subida.status, txt.slice(0, 300));
      return { statusCode: 200, headers: { "Content-Type": "application/json" }, body: JSON.stringify({ estado: "firmado", pdfUrl: "", aviso: `Está firmado, pero no se pudo guardar el PDF en el CRM (Storage ${subida.status})` }) };
    }
    const pdfUrl = `https://firebasestorage.googleapis.com/v0/b/${STORAGE_BUCKET}/o/${encodeURIComponent(nombreArchivo)}?alt=media`;
    return { statusCode: 200, headers: { "Content-Type": "application/json" }, body: JSON.stringify({ estado: "firmado", pdfUrl, firmadoEn: d.generated_at || null }) };
  } catch (err) {
    console.error("Excepción en firma-comprobar:", err.message);
    return { statusCode: 500, body: JSON.stringify({ error: err.message }) };
  }
};
