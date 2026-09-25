// Aviso diario de material fuera en proceso externo (pintura, anodizado...).
// Se ejecuta todos los días (ver "schedule" en netlify.toml) y revisa los
// registros de Fábrica → "Material fuera (proceso externo)" cuya fecha
// prevista de recogida ya ha pasado y que siguen sin marcarse como
// "Recogido". Avisa por email al responsable asignado (si tiene email en
// su usuario) y, con el resumen completo, a Miguel.
//
// Usa las mismas variables de entorno que el resto de emails del CRM
// (EMAIL_USER y EMAIL_PASSWORD). El destinatario de Miguel por defecto es
// alumavel@gmail.com (se puede cambiar con AVISOS_MATERIAL_EMAIL_TO).

import nodemailer from "nodemailer";

const FIREBASE_DB_URL = "https://crmalumavel-default-rtdb.europe-west1.firebasedatabase.app";
// Clave secreta de la base de datos (variable FIREBASE_DB_SECRET en Netlify): hace falta
// para leer/escribir una vez cerradas las reglas de Firebase. Sin ella funciona igual mientras
// las reglas sigan abiertas.
const FB_AUTH = process.env.FIREBASE_DB_SECRET ? `?auth=${encodeURIComponent(process.env.FIREBASE_DB_SECRET)}` : "";

const toArray = (obj) => (obj ? Object.values(obj) : []);

export const handler = async () => {
  try {
    const user = process.env.EMAIL_USER;
    const pass = process.env.EMAIL_PASSWORD;
    const destinatarioMiguel = process.env.AVISOS_MATERIAL_EMAIL_TO || "alumavel@gmail.com";

    if (!user || !pass) {
      console.error("Avisos de material fuera: faltan EMAIL_USER / EMAIL_PASSWORD, no se puede enviar.");
      return { statusCode: 200, body: "Faltan variables de entorno de email." };
    }

    const [enviosRes, proyectosRes, proveedoresRes, usuariosRes] = await Promise.all([
      fetch(`${FIREBASE_DB_URL}/enviosProceso.json${FB_AUTH}`),
      fetch(`${FIREBASE_DB_URL}/proyectos.json${FB_AUTH}`),
      fetch(`${FIREBASE_DB_URL}/proveedores.json${FB_AUTH}`),
      fetch(`${FIREBASE_DB_URL}/usuarios.json${FB_AUTH}`),
    ]);
    const envios = toArray(await enviosRes.json());
    const proyectos = toArray(await proyectosRes.json());
    const proveedores = toArray(await proveedoresRes.json());
    const usuarios = toArray(await usuariosRes.json());

    const hoy = new Date().toISOString().slice(0, 10);
    const vencidos = envios.filter((e) => e.estado === "Fuera" && e.fechaPrevistaRecogida && e.fechaPrevistaRecogida < hoy);

    if (vencidos.length === 0) {
      console.log("Avisos de material fuera: no hay ninguno vencido hoy.");
      return { statusCode: 200, body: "No hay material vencido." };
    }

    const nombreUsuario = (u) => `${u.nombre || ""} ${u.apellidos || ""}`.trim();
    const proveedorNombre = (id) => proveedores.find((p) => p.id === id)?.nombre || "—";
    const proyectoNombre = (id) => { const p = proyectos.find((x) => x.id === id); return p ? `#${p.numero} — ${p.nombre}` : "material general (sin obra)"; };
    const linea = (e) => `• ${e.descripcion || "Sin descripción"}${e.cantidad ? ` (${e.cantidad})` : ""} — en ${proveedorNombre(e.proveedorId)} — ${proyectoNombre(e.proyectoId)} — previsto recoger el ${e.fechaPrevistaRecogida}`;

    const transporter = nodemailer.createTransport({
      host: "smtp.office365.com",
      port: 587,
      secure: false,
      auth: { user, pass },
    });

    // Agrupar por responsable asignado (los sin asignar solo van en el resumen de Miguel).
    const porResponsable = {};
    vencidos.forEach((e) => {
      if (e.responsableId) (porResponsable[e.responsableId] = porResponsable[e.responsableId] || []).push(e);
    });

    for (const [responsableId, susEnvios] of Object.entries(porResponsable)) {
      const u = usuarios.find((x) => x.id === responsableId);
      if (!u || !u.email) continue;
      const cuerpo = `Tienes ${susEnvios.length} envío(s) a proceso externo con la fecha prevista de recogida ya pasada:\n\n${susEnvios.map(linea).join("\n")}\n\nEntra en el CRM (Fábrica → Material fuera) para marcarlos como recogidos o revisar qué pasa.\n\nEste correo se envía automáticamente cada día.`;
      try {
        await transporter.sendMail({
          from: `"ALUMAVEL CRM — Fábrica" <${user}>`,
          to: u.email,
          subject: `${susEnvios.length} envío(s) a proceso externo sin recoger — ${hoy}`,
          text: cuerpo,
        });
      } catch (err) {
        console.error(`No se pudo avisar a ${nombreUsuario(u)} (${u.email}):`, err.message);
      }
    }

    // Resumen completo (todos los vencidos, asignados o no) a Miguel.
    const resumenLineas = vencidos.map((e) => {
      const u = usuarios.find((x) => x.id === e.responsableId);
      return `${linea(e)} — responsable: ${u ? nombreUsuario(u) : "sin asignar"}`;
    });
    const cuerpoResumen = `Hay ${vencidos.length} envío(s) a proceso externo con la fecha prevista de recogida ya pasada y sin marcar como recogidos:\n\n${resumenLineas.join("\n")}\n\nEntra en el CRM (Fábrica → Material fuera) para marcarlos como recogidos o revisar qué pasa.\n\nEste correo se envía automáticamente cada día.`;
    await transporter.sendMail({
      from: `"ALUMAVEL CRM — Fábrica" <${user}>`,
      to: destinatarioMiguel,
      subject: `${vencidos.length} envío(s) a proceso externo sin recoger — ${hoy}`,
      text: cuerpoResumen,
    });

    console.log(`Aviso de material fuera enviado: ${vencidos.length} registro(s) vencido(s).`);
    return { statusCode: 200, body: `Aviso enviado: ${vencidos.length} registro(s) vencido(s).` };
  } catch (err) {
    console.error("Error en el aviso de material fuera:", err);
    return { statusCode: 200, body: "Error: " + err.message };
  }
};

