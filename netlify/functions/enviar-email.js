// Envía correos desde el CRM usando el buzón de la empresa en CDmon
// (alumavel@alumavel.es, servidor smtp.alumavel.es, puerto 465), sin que el usuario tenga que abrir su propio correo.
//
// Variables de entorno necesarias en Netlify (Site settings > Environment variables):
//   EMAIL_USER     = alumavel@alumavel.es
//   EMAIL_PASSWORD = la contraseña del buzón en CDmon
// Opcionales (por defecto smtp.alumavel.es y 465): SMTP_HOST, SMTP_PORT

import nodemailer from "nodemailer";

// Solo atiende a quien ha iniciado sesión en el CRM: el navegador manda el token del
// usuario (cabecera Authorization) y aquí se comprueba con Firebase que es válido y que
// esa persona está en el equipo o es usuario del portal. Sin esto, cualquiera que
// conociera la dirección podría usar esta función.
const FIREBASE_DB_URL = "https://crmalumavel-default-rtdb.europe-west1.firebasedatabase.app";
const FIREBASE_WEB_API_KEY = "AIzaSyBf51Gy2drHdUyJ4kCstNdvvV2h3uYe4RM";
async function usuarioAutorizado(event) {
  const token = (event.headers?.authorization || event.headers?.Authorization || "").replace(/^Bearer\s+/i, "").trim();
  if (!token) return false;
  const r = await fetch(`https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=${FIREBASE_WEB_API_KEY}`, {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ idToken: token }),
  });
  const d = await r.json().catch(() => ({}));
  const uid = r.ok && d.users && d.users[0] && d.users[0].localId;
  if (!uid) return false;
  for (const ruta of [`equipo/${uid}`, `portalUsuarios/${uid}`]) {
    const x = await fetch(`${FIREBASE_DB_URL}/${ruta}.json?auth=${encodeURIComponent(token)}`);
    if (x.ok && await x.json().catch(() => null)) return true;
  }
  return false;
}

export const handler = async (event) => {
  try {
    if (!(await usuarioAutorizado(event))) {
      return { statusCode: 200, body: JSON.stringify({ error: "Sesión caducada o sin permiso. Cierra sesión y vuelve a entrar en el CRM." }) };
    }

    const body = JSON.parse(event.body || "{}");
    const { destinatario, asunto, cuerpo, replyTo, adjuntos, nombreRemitente, cc } = body;

    if (!destinatario || !asunto || !cuerpo) {
      return { statusCode: 400, body: JSON.stringify({ error: "Faltan datos: destinatario, asunto o cuerpo." }) };
    }

    const user = process.env.EMAIL_USER;
    const pass = process.env.EMAIL_PASSWORD;
    if (!user || !pass) {
      return { statusCode: 200, body: JSON.stringify({ error: "El envío de correo no está configurado todavía (faltan EMAIL_USER / EMAIL_PASSWORD en Netlify)." }) };
    }

    const transporter = nodemailer.createTransport({
      host: process.env.SMTP_HOST || "smtp.alumavel.es",
      port: Number(process.env.SMTP_PORT || 465),
      secure: Number(process.env.SMTP_PORT || 465) === 465, // 465 = SSL/TLS directo (CDmon)
      auth: { user, pass },
    });

    // adjuntos: [{ nombre, dataUrl }] — dataUrl tipo "data:application/pdf;base64,...."
    // o bien [{ nombre, url }] — un archivo ya subido (p.ej. el PDF de un pedido de
    // Uxcar); se descarga aquí, en el servidor, y se adjunta.
    const lista = [];
    for (const a of adjuntos || []) {
      if (a && a.dataUrl) { lista.push(a); continue; }
      if (a && a.url) {
        const r = await fetch(a.url);
        if (!r.ok) throw new Error("no se pudo descargar el adjunto " + (a.nombre || ""));
        const buf = Buffer.from(await r.arrayBuffer());
        const tipo = r.headers.get("content-type") || "application/octet-stream";
        lista.push({ nombre: a.nombre, dataUrl: `data:${tipo};base64,${buf.toString("base64")}` });
      }
    }
    const attachments = lista
      .map((a) => {
        const match = /^data:([^;]+);base64,(.*)$/.exec(a.dataUrl);
        return {
          filename: a.nombre || "adjunto",
          content: match ? match[2] : a.dataUrl,
          encoding: "base64",
          contentType: match ? match[1] : undefined,
        };
      });

    // cc: correo (o lista) de quien genera el envío; recibe copia. Solo se acepta si tiene forma de correo y no repite al destinatario.
    const esCorreo = (x) => /^[^\s@,;<>"]+@[^\s@,;<>"]+\.[^\s@,;<>"]+$/.test(String(x || "").trim());
    const dest = String(destinatario).split(/[;,]/).map((x) => x.trim().toLowerCase());
    const copia = [...new Set((Array.isArray(cc) ? cc : String(cc || "").split(/[;,]/)).map((x) => String(x).trim()).filter(esCorreo))]
      .filter((x) => !dest.includes(x.toLowerCase()));

    await transporter.sendMail({
      // nombreRemitente: p.ej. "Uxcar" cuando el pedido lo manda Uxcar desde su portal.
      // La dirección sigue siendo la de la empresa, pero se ve su nombre y las
      // respuestas le llegan a su correo (replyTo).
      from: `"${String(nombreRemitente || "ALUMAVEL").replace(/["\r\n]/g, "")}" <${user}>`,
      to: destinatario,
      ...(copia.length > 0 ? { cc: copia } : {}),
      replyTo: replyTo || user,
      subject: asunto,
      text: cuerpo,
      ...(attachments.length > 0 ? { attachments } : {}),
    });

    return { statusCode: 200, body: JSON.stringify({ ok: true }) };
  } catch (err) {
    console.error("Error enviando email:", err);
    return { statusCode: 200, body: JSON.stringify({ error: "No se pudo enviar el correo: " + err.message }) };
  }
};
