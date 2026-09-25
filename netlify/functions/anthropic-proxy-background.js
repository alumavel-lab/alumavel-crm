// Función de "segundo plano" (background function) de Netlify.
// A diferencia de anthropic-proxy.js (que tiene un límite de 26 segundos),
// esta puede tardar hasta 15 minutos, porque no responde directamente al
// navegador: guarda el resultado en Firebase, y el CRM va preguntando
// (sondeando) cada pocos segundos hasta que el resultado está listo.
// IMPORTANTE: el nombre del archivo debe terminar en "-background.js"
// para que Netlify la trate como función en segundo plano.

const FIREBASE_DB_URL = "https://crmalumavel-default-rtdb.europe-west1.firebasedatabase.app";
// Clave secreta de la base de datos (variable FIREBASE_DB_SECRET en Netlify): hace falta
// para leer/escribir una vez cerradas las reglas de Firebase. Sin ella funciona igual mientras
// las reglas sigan abiertas.
const FB_AUTH = process.env.FIREBASE_DB_SECRET ? `?auth=${encodeURIComponent(process.env.FIREBASE_DB_SECRET)}` : "";

async function guardarResultado(jobId, payload) {
  try {
    await fetch(`${FIREBASE_DB_URL}/packingListJobs/${jobId}.json${FB_AUTH}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ...payload, ts: Date.now() }),
    });
  } catch (e) {
    console.error("No se pudo guardar el resultado en Firebase:", e);
  }
}

// La foto/PDF (en base64) puede pesar varios MB, y las funciones de Netlify
// rechazan peticiones grandes con un error 413. Por eso el navegador ya NO
// nos manda el archivo directamente: lo deja guardado en Firebase primero
// (que no tiene ese límite), y aquí solo recibimos el jobId y vamos a
// buscarlo nosotros mismos.
async function leerEntrada(jobId) {
  const res = await fetch(`${FIREBASE_DB_URL}/packingListJobsInput/${jobId}.json${FB_AUTH}`);
  if (!res.ok) throw new Error("No se pudo leer la entrada del trabajo desde Firebase (status " + res.status + ")");
  const data = await res.json();
  if (!data) throw new Error("No se encontró la entrada del trabajo en Firebase (jobId: " + jobId + ")");
  return data;
}

export const handler = async (event) => {
  let jobId = null;
  try {
    const body = JSON.parse(event.body || "{}");
    jobId = body.jobId;
    const apiKey = process.env.ANTHROPIC_API_KEY;

    if (!apiKey) {
      await guardarResultado(jobId, { status: "error", error: "ANTHROPIC_API_KEY no configurada en Netlify" });
      return { statusCode: 200, body: "ok" };
    }

    const { model, max_tokens, contentBlock, prompt } = await leerEntrada(jobId);

    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model,
        max_tokens,
        messages: [{ role: "user", content: [contentBlock, { type: "text", text: prompt }] }],
      }),
    });

    const data = await response.json();

    if (!response.ok || data.error) {
      console.error("Anthropic API devolvió error:", response.status, JSON.stringify(data));
      await guardarResultado(jobId, { status: "error", error: (data.error && data.error.message) || ("Status " + response.status) });
      return { statusCode: 200, body: "ok" };
    }

    const texto = (data.content || []).filter((c) => c.type === "text").map((c) => c.text).join("");
    await guardarResultado(jobId, { status: "done", texto });
    return { statusCode: 200, body: "ok" };
  } catch (err) {
    console.error("Excepción en anthropic-proxy-background:", err.message);
    if (jobId) await guardarResultado(jobId, { status: "error", error: err.message });
    return { statusCode: 200, body: "ok" };
  }
};
