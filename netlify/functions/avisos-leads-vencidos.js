// Aviso diario de Leads vencidos.
// Se ejecuta todos los días (ver "schedule" en netlify.toml) y revisa los
// leads cuya "próxima acción" (fecha de recontacto) ya ha pasado y que
// siguen en un estado activo (no Aceptado / Rechazado / Descartado).
//
// Por cada lead vencido manda un aviso por email:
//   - al comercial asignado (si tiene email en su usuario)
//   - y también, en un correo aparte con el resumen de todos, a Miguel
//     (EMAIL_USER, o AVISOS_LEADS_EMAIL_TO si se define en Netlify)
//
// Usa las mismas variables de entorno que el resto de emails del CRM
// (EMAIL_USER y EMAIL_PASSWORD).

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
const ESTADOS_FINALES = ["Aceptado", "Rechazado", "Descartado"];

const toArray = (obj) => (obj ? Object.values(obj) : []);

export const handler = async () => {
  FB_AUTH = await fbAuthQuery().catch(() => "");
  try {
    const user = process.env.EMAIL_USER;
    const pass = process.env.EMAIL_PASSWORD;
    const destinatarioMiguel = process.env.AVISOS_LEADS_EMAIL_TO || "alumavel@gmail.com";

    if (!user || !pass) {
      console.error("Avisos de leads: faltan EMAIL_USER / EMAIL_PASSWORD, no se puede enviar.");
      return { statusCode: 200, body: "Faltan variables de entorno de email." };
    }

    const [leadsRes, usuariosRes] = await Promise.all([
      fetch(`${FIREBASE_DB_URL}/leads.json${FB_AUTH}`),
      fetch(`${FIREBASE_DB_URL}/usuarios.json${FB_AUTH}`),
    ]);
    const leads = toArray(await leadsRes.json());
    const usuarios = toArray(await usuariosRes.json());

    const hoy = new Date().toISOString().slice(0, 10);
    const vencidos = leads.filter((l) =>
      l.proximaAccionFecha && l.proximaAccionFecha < hoy && !ESTADOS_FINALES.includes(l.estado)
    );

    if (vencidos.length === 0) {
      console.log("Avisos de leads: no hay leads vencidos hoy.");
      return { statusCode: 200, body: "No hay leads vencidos." };
    }

    const nombreUsuario = (u) => `${u.nombre || ""} ${u.apellidos || ""}`.trim();
    const lineaLead = (l) => `• ${l.nombre || "Sin nombre"} — ${l.telefono || "sin teléfono"} — fecha prevista: ${l.proximaAccionFecha} — estado: ${l.estado || "Pendiente"}`;

    const transporter = nodemailer.createTransport({
      host: "smtp.office365.com",
      port: 587,
      secure: false,
      auth: { user, pass },
    });

    // Agrupar por comercial asignado (los sin asignar van aparte).
    const porComercial = {};
    const sinAsignar = [];
    vencidos.forEach((l) => {
      if (l.comercialId) {
        (porComercial[l.comercialId] = porComercial[l.comercialId] || []).push(l);
      } else {
        sinAsignar.push(l);
      }
    });

    // Email individual a cada comercial con sus leads vencidos.
    for (const [comercialId, susLeads] of Object.entries(porComercial)) {
      const u = usuarios.find((x) => x.id === comercialId);
      if (!u || !u.email) continue;
      const cuerpo = `Tienes ${susLeads.length} lead(s) con la fecha de recontacto vencida y sin seguimiento registrado:\n\n${susLeads.map(lineaLead).join("\n")}\n\nEntra en el CRM (Leads) para registrar la llamada y ponerlas al día.\n\nEste correo se envía automáticamente cada día.`;
      try {
        await transporter.sendMail({
          from: `"ALUMAVEL CRM — Leads" <${user}>`,
          to: u.email,
          subject: `${susLeads.length} lead(s) vencido(s) por recontactar — ${hoy}`,
          text: cuerpo,
        });
      } catch (err) {
        console.error(`No se pudo avisar a ${nombreUsuario(u)} (${u.email}):`, err.message);
      }
    }

    // Resumen completo (todos los vencidos, asignados o no) a Miguel.
    const resumenLineas = vencidos.map((l) => {
      const u = usuarios.find((x) => x.id === l.comercialId);
      return `${lineaLead(l)} — comercial: ${u ? nombreUsuario(u) : "sin asignar"}`;
    });
    const cuerpoResumen = `Resumen de leads vencidos hoy (${vencidos.length} en total, ${sinAsignar.length} sin comercial asignado):\n\n${resumenLineas.join("\n")}\n\nEste correo se envía automáticamente cada día.`;
    await transporter.sendMail({
      from: `"ALUMAVEL CRM — Leads" <${user}>`,
      to: destinatarioMiguel,
      subject: `Resumen: ${vencidos.length} lead(s) vencido(s) — ${hoy}`,
      text: cuerpoResumen,
    });

    console.log(`Avisos de leads enviados: ${vencidos.length} lead(s) vencido(s).`);
    return { statusCode: 200, body: `Avisos enviados: ${vencidos.length} lead(s) vencido(s).` };
  } catch (err) {
    console.error("Error en los avisos de leads vencidos:", err);
    return { statusCode: 200, body: "Error: " + err.message };
  }
};
