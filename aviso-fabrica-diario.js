// Aviso diario a fábrica: de lunes a viernes a las 7:00 (hora de España, al entrar a trabajar, ver "schedule" en
// netlify.toml) manda al correo de fábrica lo que hay que preparar para el SIGUIENTE DÍA
// DE TRABAJO (el viernes, lo del lunes) según el
// último planning confirmado en el CRM (Fábrica → Planning → "Confirmar planning").
//
// Incluye, para cada obra que se empieza ese día: perfiles, refuerzos, herrajes y
// accesorios de su listado de materiales (con almacén y estantería), las persianas
// (del listado y las que ya están en los carros) y los caballetes de cristal que tiene
// en el almacén, con su hueco. Así no se olvida nada.
//
// Envío: si existe RESEND_API_KEY en Netlify usa Resend (y RESEND_FROM como remitente);
// si no, usa la cuenta de Microsoft 365 de siempre (EMAIL_USER / EMAIL_PASSWORD).

import nodemailer from "nodemailer";

const FIREBASE_DB_URL = "https://crmalumavel-default-rtdb.europe-west1.firebasedatabase.app";
const FB_AUTH = process.env.FIREBASE_DB_SECRET ? `?auth=${encodeURIComponent(process.env.FIREBASE_DB_SECRET)}` : "";
const toArray = (obj) => (obj ? (Array.isArray(obj) ? obj.filter(Boolean) : Object.values(obj)) : []);
const nums = (txt) => String(txt || "").match(/\d{2,}/g) || [];
const leer = async (ruta) => { const r = await fetch(`${FIREBASE_DB_URL}/${ruta}.json${FB_AUTH}`); return r.ok ? r.json() : null; };

async function enviar({ to, subject, text }) {
  if (process.env.RESEND_API_KEY) {
    const r = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: { Authorization: `Bearer ${process.env.RESEND_API_KEY}`, "Content-Type": "application/json" },
      body: JSON.stringify({ from: process.env.RESEND_FROM || "Planning Ecowin PVC <onboarding@resend.dev>", to: [to], subject, text }),
    });
    if (!r.ok) throw new Error("Resend: " + (await r.text()));
    return;
  }
  const user = process.env.EMAIL_USER, pass = process.env.EMAIL_PASSWORD;
  if (!user || !pass) throw new Error("Correo sin configurar (faltan RESEND_API_KEY o EMAIL_USER/EMAIL_PASSWORD)");
  const t = nodemailer.createTransport({ host: "smtp.office365.com", port: 587, secure: false, auth: { user, pass } });
  await t.sendMail({ from: `"Planning Ecowin PVC" <${user}>`, to, subject, text });
}

// Día laborable siguiente (lunes a viernes), en hora de España
function manana() {
  const ahora = new Date(new Date().toLocaleString("en-US", { timeZone: "Europe/Madrid" }));
  const d = new Date(ahora);
  do { d.setDate(d.getDate() + 1); } while (d.getDay() === 0 || d.getDay() === 6);
  const f = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  return { f, texto: d.toLocaleDateString("es-ES", { weekday: "long", day: "numeric", month: "long" }) };
}

export const handler = async (event) => {
  try {
    // Solo a las 7 en España (se lanza a las 5 y a las 6 UTC por el cambio de hora).
    // Si se llama a mano con ?forzar=1 se envía igualmente, para probar.
    const horaEs = parseInt(new Date().toLocaleString("en-GB", { timeZone: "Europe/Madrid", hour: "2-digit", hour12: false }), 10);
    const forzar = event && event.queryStringParameters && event.queryStringParameters.forzar;
    if (horaEs !== 7 && !forzar) return { statusCode: 200, body: `No son las 7 en España (son las ${horaEs}).` };
    const plan = await leer("planningPublicado");
    if (!plan || !plan.emailFabrica) return { statusCode: 200, body: "Sin planning confirmado o sin correo de fábrica." };
    const { f, texto: diaTexto } = manana();
    const trozos = toArray(plan.dias && plan.dias[f]);
    if (trozos.length === 0) return { statusCode: 200, body: `Nada planificado para ${f}.` };

    const [cristales, persianasAlmacen, proyectos] = await Promise.all([leer("cristales"), leer("persianasAlmacen"), leer("proyectos")]);
    const obrasDia = trozos.map((t) => ({ ...t, obra: (plan.obras || {})[t.id] || { nombre: t.nombre, lineas: [], expNums: [] } }));

    let cuerpo = `PREPARAR HOY PARA EL ${diaTexto.toUpperCase()}\n\n`;
    cuerpo += `Obras del próximo día de trabajo:\n${obrasDia.map((o) => `  • ${o.obra.nombre} — ${o.horas} h${o.obra.ventanas ? ` (${o.obra.ventanas} ventanas)` : ""}${o.obra.inicio && o.obra.inicio < f ? " · (sigue de días anteriores)" : ""}`).join("\n")}\n`;

    // Solo se prepara el material de las obras que EMPIEZAN ese día (las que siguen ya lo tienen)
    const empiezan = obrasDia.filter((o) => !o.obra.inicio || o.obra.inicio === f);
    if (empiezan.length === 0) cuerpo += `\nEse día no empieza ninguna obra nueva: se sigue con las que están en marcha.\n`;

    empiezan.forEach((o) => {
      const ob = o.obra;
      cuerpo += `\n==============================\n${ob.nombre.toUpperCase()}\n==============================\n`;
      // Perfiles, refuerzos, herrajes y accesorios, por almacén y estantería
      const mat = toArray(ob.lineas).filter((l) => l.seccion !== "persianas")
        .sort((a, b) => ((a.almacen ? 0 : 1) - (b.almacen ? 0 : 1) || String(a.almacen || "").localeCompare(String(b.almacen || ""))) || String(a.estanteria || "").localeCompare(String(b.estanteria || ""), "es", { numeric: true }));
      if (mat.length) {
        let alm = null;
        cuerpo += `\nMATERIAL (perfiles, refuerzos, herrajes, accesorios):\n`;
        mat.forEach((l) => {
          const a = l.almacen || "Sin almacén asignado";
          if (a !== alm) { cuerpo += ` [${a}]\n`; alm = a; }
          cuerpo += `   ☐ ${l.estanteria ? l.estanteria + " · " : ""}${l.codigo} ${l.descripcion}${l.color ? " · " + l.color : ""} → ${l.uds}\n`;
        });
      } else cuerpo += `\n(OJO: esta obra no tiene listado de materiales en el CRM)\n`;
      // Persianas
      const persListado = toArray(ob.lineas).filter((l) => l.seccion === "persianas");
      const persAlm = toArray(persianasAlmacen).filter((x) => nums(x.expediente).some((n) => (ob.expNums || []).includes(n)));
      const proy = ob.proyectoId ? toArray(proyectos).find((p) => p.id === ob.proyectoId) : null;
      const persProy = proy ? toArray(proy.persianasControl) : [];
      if (persListado.length || persAlm.length || persProy.length) {
        cuerpo += `\nPERSIANAS:\n`;
        persAlm.forEach((x) => { cuerpo += `   ☐ ${x.estante ? `Carro ${x.estante.carro} · lado ${x.estante.lado} · estante ${x.estante.nivel}` : "sin sitio"} — ${[x.modelo, x.descripcion].filter(Boolean).join(" ") || "persiana"}${x.ancho || x.largo ? ` ${x.ancho || x.largo}${x.alto ? "×" + x.alto : ""}` : ""}\n`; });
        persProy.forEach((u) => { cuerpo += `   ☐ ${u.carro ? `Carro ${u.carro}` : "sin carro"} — ${[u.nombre, u.descripcion, u.modelo].filter(Boolean).join(" ") || "persiana"}${u.ancho ? ` ${u.ancho}×${u.alto || ""}` : ""}\n`; });
        if (!persAlm.length && !persProy.length) persListado.forEach((l) => { cuerpo += `   ☐ ${l.descripcion}${l.ancho ? ` ${l.ancho}×${l.alto}` : ""}${l.color ? " · " + l.color : ""} → ${l.uds} (del listado; comprobar que han llegado)\n`; });
      }
      // Cristales
      const cris = toArray(cristales).filter((c) => {
        const set = new Set(nums(c.expediente));
        toArray(c.piezas).forEach((p) => nums(p.expediente).forEach((n) => set.add(n)));
        return (ob.expNums || []).some((n) => set.has(n));
      });
      if (cris.length) {
        cuerpo += `\nCRISTALES (caballetes):\n`;
        cris.forEach((c) => {
          const u = c.ubicacion ? `Zona ${c.ubicacion.zona} · fila ${c.ubicacion.fila} · hueco ${c.ubicacion.hueco}` : "SIN UBICAR";
          const piezas = toArray(c.piezas).filter((p) => nums(p.expediente).some((n) => (ob.expNums || []).includes(n))).length;
          cuerpo += `   ☐ ${u} — caballete ${c.lote || c.numero || ""}${piezas ? ` (${piezas} cristales de esta obra)` : ""}\n`;
        });
      } else cuerpo += `\nCRISTALES: no encuentro caballetes de esta obra en el almacén (¿han llegado?)\n`;
    });

    cuerpo += `\n— Aviso automático del CRM, según el planning confirmado el ${new Date(plan.publicadoEn || Date.now()).toLocaleDateString("es-ES")}.`;
    await enviar({ to: plan.emailFabrica, subject: `Preparar hoy para el ${diaTexto} — ${empiezan.length} obra${empiezan.length === 1 ? "" : "s"}`, text: cuerpo });
    return { statusCode: 200, body: "Enviado" };
  } catch (e) {
    console.error("aviso-fabrica-diario:", e.message);
    return { statusCode: 200, body: "Error: " + e.message };
  }
};
