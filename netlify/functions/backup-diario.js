// Copia de seguridad automática del CRM.
// Se ejecuta sola cada día (ver "schedule" en netlify.toml) y manda por email
// una copia completa de toda la base de datos, en un archivo adjunto.
//
// No hace falta hacer nada para que funcione: usa las mismas variables de
// entorno que ya configuraste para el envío de correos (EMAIL_USER y
// EMAIL_PASSWORD). Si quieres que la copia llegue a un correo distinto al de
// envío, añade también BACKUP_EMAIL_TO en Netlify con esa dirección.

import nodemailer from "nodemailer";

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

export const handler = async () => {
  FB_AUTH = await fbAuthQuery().catch(() => "");
  try {
    const user = process.env.EMAIL_USER;
    const pass = process.env.EMAIL_PASSWORD;
    const destinatario = process.env.BACKUP_EMAIL_TO || user;

    if (!user || !pass) {
      console.error("Backup diario: faltan EMAIL_USER / EMAIL_PASSWORD, no se puede enviar.");
      return { statusCode: 200, body: "Faltan variables de entorno de email." };
    }

    // Descarga toda la base de datos de un golpe.
    const respuesta = await fetch(`${FIREBASE_DB_URL}/.json${FB_AUTH}`);
    if (!respuesta.ok) {
      throw new Error("No se pudo descargar la base de datos (status " + respuesta.status + ")");
    }
    const datos = await respuesta.text();

    const fecha = new Date().toISOString().slice(0, 10);
    const nombreArchivo = `backup-alumavel-crm-${fecha}.json`;

    const transporter = nodemailer.createTransport({
      host: "smtp.office365.com",
      port: 587,
      secure: false,
      auth: { user, pass },
    });

    await transporter.sendMail({
      from: `"ALUMAVEL CRM — Copias de seguridad" <${user}>`,
      to: destinatario,
      subject: `Copia de seguridad del CRM — ${fecha}`,
      text: `Copia de seguridad automática de todo el CRM ALUMAVEL, generada el ${fecha}.\n\nGuarda este correo o descarga el archivo adjunto en un sitio seguro (por ejemplo, una carpeta de Dropbox dedicada a copias de seguridad).\n\nEste correo se envía automáticamente todos los días, no hace falta hacer nada.`,
      attachments: [
        {
          filename: nombreArchivo,
          content: datos,
          contentType: "application/json",
        },
      ],
    });

    console.log("Backup diario enviado correctamente:", nombreArchivo);
    return { statusCode: 200, body: "Backup enviado: " + nombreArchivo };
  } catch (err) {
    console.error("Error en el backup diario:", err);
    return { statusCode: 200, body: "Error: " + err.message };
  }
};
