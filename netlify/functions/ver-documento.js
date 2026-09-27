// Enseña al cliente el PDF de su presupuesto con un enlace sencillo (el que va en el
// WhatsApp): /.netlify/functions/ver-documento?p=<id del presupuesto>&d=<id del documento>
// El servidor busca el documento con el usuario robot y lo descarga de Firebase Storage,
// así el enlace funciona aunque Storage no sea público. Los ids son aleatorios y largos.
const FIREBASE_DB_URL = "https://crmalumavel-default-rtdb.europe-west1.firebasedatabase.app";
const FIREBASE_WEB_API_KEY = "AIzaSyBf51Gy2drHdUyJ4kCstNdvvV2h3uYe4RM";
const STORAGE_PREFIX = "https://firebasestorage.googleapis.com/v0/b/crmalumavel.firebasestorage.app/o/";

let _token = null, _caduca = 0;
async function tokenRobot() {
  if (_token && Date.now() < _caduca) return _token;
  const email = (process.env.FIREBASE_ROBOT_EMAIL || "").trim();
  const password = (process.env.FIREBASE_ROBOT_PASSWORD || "").trim();
  if (!email || !password) return "";
  const r = await fetch(`https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=${FIREBASE_WEB_API_KEY}`, {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ email, password, returnSecureToken: true }),
  });
  const d = await r.json().catch(() => ({}));
  if (!d.idToken) return "";
  _token = d.idToken; _caduca = Date.now() + 50 * 60 * 1000;
  return _token;
}

const pagina = (texto) => ({
  statusCode: 404,
  headers: { "Content-Type": "text/html; charset=utf-8" },
  body: `<html><body style="font-family:sans-serif;padding:40px"><h2>Ecowin PVC</h2><p>${texto}</p></body></html>`,
});

export const handler = async (event) => {
  try {
    const { p, d } = event.queryStringParameters || {};
    if (!p || !d) return pagina("Enlace no válido.");
    const token = await tokenRobot();
    const q = token ? `?auth=${encodeURIComponent(token)}` : "";
    const r = await fetch(`${FIREBASE_DB_URL}/presupuestos.json${q}`);
    const lista = await r.json().catch(() => null);
    const pres = lista && Object.values(lista).find((x) => x && x.id === p);
    const doc = pres && Object.values(pres.documentos || {}).find((x) => x && x.id === d);
    if (!doc || !String(doc.url || "").startsWith(STORAGE_PREFIX)) return pagina("No se ha encontrado este documento. Pídanos que se lo volvamos a enviar.");
    const f = await fetch(doc.url, { headers: token ? { Authorization: `Firebase ${token}` } : {} });
    if (!f.ok) return pagina("No se ha podido abrir el documento ahora mismo. Inténtelo de nuevo en unos minutos.");
    const buf = Buffer.from(await f.arrayBuffer());
    const nombre = String(doc.nombre || "presupuesto.pdf").replace(/[^\w.\-]+/g, "_");
    return {
      statusCode: 200,
      headers: { "Content-Type": f.headers.get("content-type") || "application/pdf", "Content-Disposition": `inline; filename="${nombre}"`, "Cache-Control": "private, max-age=300" },
      body: buf.toString("base64"),
      isBase64Encoded: true,
    };
  } catch (e) {
    console.error("ver-documento:", e.message);
    return pagina("Ha habido un error al abrir el documento.");
  }
};
