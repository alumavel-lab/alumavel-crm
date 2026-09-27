// Comprobación de que el servidor puede entrar en la base de datos. Abre en el navegador:
//   https://dynamic-eclair-be67a5.netlify.app/.netlify/functions/comprobar-secreto
// Solo dice si funciona o no; no devuelve datos ni claves.
const FIREBASE_DB_URL = "https://crmalumavel-default-rtdb.europe-west1.firebasedatabase.app";
const FIREBASE_WEB_API_KEY = "AIzaSyBf51Gy2drHdUyJ4kCstNdvvV2h3uYe4RM";

const html = (color, titulo, texto) => ({
  statusCode: 200,
  headers: { "Content-Type": "text/html; charset=utf-8" },
  body: `<html><body style="font-family:sans-serif;padding:40px;max-width:680px"><h1 style="color:${color}">${titulo}</h1><p style="font-size:18px;line-height:1.5">${texto}</p></body></html>`,
});

export const handler = async () => {
  const email = (process.env.FIREBASE_ROBOT_EMAIL || "").trim();
  const password = (process.env.FIREBASE_ROBOT_PASSWORD || "").trim();
  if (!email || !password) {
    return html("#c0392b", "✗ Falta el usuario robot", "No están las variables FIREBASE_ROBOT_EMAIL y FIREBASE_ROBOT_PASSWORD en Netlify (o alguna está vacía). Créalas y vuelve a publicar (Deploys → Trigger deploy → Deploy site).");
  }
  try {
    const r = await fetch(`https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=${FIREBASE_WEB_API_KEY}`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, password, returnSecureToken: true }),
    });
    const d = await r.json().catch(() => ({}));
    if (!r.ok || !d.idToken) {
      const motivo = d.error?.message || r.status;
      const explicacion = /INVALID_LOGIN_CREDENTIALS|INVALID_PASSWORD|EMAIL_NOT_FOUND/.test(String(motivo))
        ? "El email o la contraseña no coinciden con el usuario creado en el CRM. Revisa que estén escritos igual en Netlify."
        : "No se pudo iniciar sesión con el usuario robot.";
      return html("#c0392b", "✗ El robot no puede entrar", `${explicacion} (${motivo})`);
    }
    const r2 = await fetch(`${FIREBASE_DB_URL}/usuarios.json?shallow=true&auth=${encodeURIComponent(d.idToken)}`);
    if (r2.ok) {
      return html("#2E8B57", "✓ Funciona", "El servidor ya puede leer la base de datos: copias de seguridad, informes, avisos y el paso a \"Firmado\" automático vuelven a funcionar.");
    }
    return html("#c0392b", "✗ El robot entra, pero no tiene permiso", `El usuario existe pero no está dado de alta en el equipo (${r2.status}). Crea el usuario desde la pantalla de Usuarios del CRM, no desde la consola de Firebase, o entra una vez con él en el CRM.`);
  } catch (e) {
    return html("#c0392b", "✗ Error de conexión", e.message);
  }
};
