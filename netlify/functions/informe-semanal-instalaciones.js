// Informe semanal de Instalaciones.
// Se ejecuta solo cada viernes (ver "schedule" en netlify.toml) y manda por
// email un resumen de la semana: por cada obra con instalación activa,
// cuántas viviendas se han marcado como instaladas y cuántas horas se han
// registrado en los últimos 7 días.
//
// Usa las mismas variables de entorno que el resto de emails del CRM
// (EMAIL_USER y EMAIL_PASSWORD). Si quieres que llegue a un correo distinto
// al de envío, añade INFORME_SEMANAL_EMAIL_TO en Netlify.

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
const SIETE_DIAS_MS = 7 * 24 * 60 * 60 * 1000;

const toArray = (obj) => (obj ? Object.values(obj) : []);

export const handler = async () => {
  FB_AUTH = await fbAuthQuery().catch(() => "");
  try {
    const user = process.env.EMAIL_USER;
    const pass = process.env.EMAIL_PASSWORD;
    const destinatario = process.env.INFORME_SEMANAL_EMAIL_TO || "alumavel@gmail.com";

    if (!user || !pass) {
      console.error("Informe semanal: faltan EMAIL_USER / EMAIL_PASSWORD, no se puede enviar.");
      return { statusCode: 200, body: "Faltan variables de entorno de email." };
    }

    const ahora = Date.now();
    const desde = ahora - SIETE_DIAS_MS;

    const [instalacionesRes, proyectosRes, controlMontajeRes] = await Promise.all([
      fetch(`${FIREBASE_DB_URL}/instalaciones.json${FB_AUTH}`),
      fetch(`${FIREBASE_DB_URL}/proyectos.json${FB_AUTH}`),
      fetch(`${FIREBASE_DB_URL}/controlMontaje.json${FB_AUTH}`),
    ]);
    const instalaciones = toArray(await instalacionesRes.json());
    const proyectos = toArray(await proyectosRes.json());
    const controlMontaje = (await controlMontajeRes.json()) || {};

    const lineas = [];
    instalaciones.forEach((inst) => {
      const proyecto = proyectos.find((p) => p.id === inst.proyectoId);
      const nombre = proyecto ? `#${proyecto.numero} — ${proyecto.nombre}` : (inst.nombre || "Instalación sin proyecto");

      const horasSemana = (inst.registroHoras || [])
        .filter((r) => r.fecha && new Date(r.fecha).getTime() >= desde)
        .reduce((s, r) => s + (parseFloat(r.horas) || 0), 0);

      const datosMontaje = controlMontaje[inst.id];
      let viviendasCompletadasSemana = 0;
      if (datosMontaje && datosMontaje.viviendas) {
        toArray(datosMontaje.viviendas).forEach((v) => {
          toArray(v.elementos).forEach((el) => {
            if (el.instalado && el.instaladoEn && el.instaladoEn >= desde) viviendasCompletadasSemana++;
          });
        });
      }

      if (horasSemana > 0 || viviendasCompletadasSemana > 0) {
        lineas.push({ nombre, horasSemana, viviendasCompletadasSemana });
      }
    });

    const fecha = new Date().toISOString().slice(0, 10);
    let cuerpo;
    if (lineas.length === 0) {
      cuerpo = "Esta semana no se ha registrado actividad (horas ni elementos instalados) en ninguna instalación.";
    } else {
      cuerpo = lineas
        .map((l) => `• ${l.nombre}: ${l.viviendasCompletadasSemana} elemento(s) instalado(s), ${l.horasSemana.toFixed(1)} h registradas`)
        .join("\n");
    }

    const transporter = nodemailer.createTransport({
      host: "smtp.office365.com",
      port: 587,
      secure: false,
      auth: { user, pass },
    });

    await transporter.sendMail({
      from: `"ALUMAVEL CRM — Informe semanal" <${user}>`,
      to: destinatario,
      subject: `Informe semanal de instalaciones — ${fecha}`,
      text: `Resumen de la actividad de instalaciones de los últimos 7 días:\n\n${cuerpo}\n\nEste correo se envía automáticamente todos los viernes.`,
    });

    console.log("Informe semanal enviado correctamente.");
    return { statusCode: 200, body: "Informe semanal enviado." };
  } catch (err) {
    console.error("Error en el informe semanal:", err);
    return { statusCode: 200, body: "Error: " + err.message };
  }
};
