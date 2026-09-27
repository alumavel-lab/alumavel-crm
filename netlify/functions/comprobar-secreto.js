// Comprobación rápida de que FIREBASE_DB_SECRET funciona: abre en el navegador
//   https://dynamic-eclair-be67a5.netlify.app/.netlify/functions/comprobar-secreto
// Solo dice si la clave vale o no; no devuelve ningún dato de la base de datos
// ni la propia clave.
const FIREBASE_DB_URL = "https://crmalumavel-default-rtdb.europe-west1.firebasedatabase.app";

export const handler = async () => {
  const html = (color, titulo, texto) => ({
    statusCode: 200,
    headers: { "Content-Type": "text/html; charset=utf-8" },
    body: `<html><body style="font-family:sans-serif;padding:40px;max-width:640px"><h1 style="color:${color}">${titulo}</h1><p style="font-size:18px">${texto}</p></body></html>`,
  });
  const secreto = (process.env.FIREBASE_DB_SECRET || "").trim();
  if (!secreto) {
    return html("#c0392b", "✗ No hay clave", "La variable FIREBASE_DB_SECRET no existe en Netlify (o está vacía). Créala y vuelve a publicar (Deploys → Trigger deploy).");
  }
  if (secreto !== process.env.FIREBASE_DB_SECRET) {
    // tiene espacios o saltos de línea alrededor: se avisa, porque así no funciona
    return html("#c0392b", "✗ La clave tiene espacios", "La clave de Netlify tiene espacios o un salto de línea al principio o al final. Edítala, bórralos, guarda y vuelve a publicar.");
  }
  try {
    const r = await fetch(`${FIREBASE_DB_URL}/usuarios.json?shallow=true&auth=${encodeURIComponent(secreto)}`);
    if (r.ok) {
      return html("#2E8B57", "✓ La clave funciona", "El servidor ya puede leer la base de datos: copias de seguridad, informes, avisos y el paso a \"Firmado\" automático vuelven a funcionar.");
    }
    const d = await r.json().catch(() => ({}));
    return html("#c0392b", "✗ La clave no vale", `Firebase la rechaza (${r.status}${d.error ? ` — ${d.error}` : ""}). Lo más probable es que sea de otro proyecto (alumavelcrm en vez de crmalumavel) o que esté mal copiada.`);
  } catch (e) {
    return html("#c0392b", "✗ Error de conexión", e.message);
  }
};
