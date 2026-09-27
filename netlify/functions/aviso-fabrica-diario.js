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

// Caballetes de ventanas: los que hay que tener listos para las obras que se cargan ese día
// (las que terminan de fabricarse el día anterior según el planning, o con fecha de reparto
// ese día), con los reservados, y los que están fuera con la vuelta ya pasada (a reclamar).
function caballetesSeccion({ f, plan, proyectos, config, caballetes, expedientes }) {
  const cap = Object.assign({ capacidad: 12, puerta: 2, corredera: 1.5, fijo: 1, osciloparalela: 2 }, config.capacidadCaballetes || {});
  const pc = Object.assign({ diasDevolucion: 10 }, config.planningCaballetes || {});
  const hoy = new Date(new Date().toLocaleString("en-US", { timeZone: "Europe/Madrid" })).toISOString().slice(0, 10);
  const num = (v) => parseFloat(v) || 0;
  const huecos = (rec, tot) => {
    const r = toArray(rec);
    if (r.length) return r.reduce((a, l) => {
      if (l.tipo === "mosquitera" || l.tipo === "otro") return a;
      const k = { puerta: cap.puerta, corredera: cap.corredera, osciloparalela: cap.osciloparalela, fijo: cap.fijo }[l.tipo];
      return a + num(l.uds) * (num(k) || 1);
    }, 0);
    return tot.v + tot.p * (num(cap.puerta) || 1) + tot.op * (num(cap.osciloparalela) || 1);
  };
  const totRecuento = (rec) => { const t = { v: 0, p: 0, op: 0 }; toArray(rec).forEach((l) => { const u = num(l.uds); if (l.tipo === "puerta") t.p += u; else if (l.tipo === "osciloparalela") t.op += u; else if (l.tipo !== "mosquitera" && l.tipo !== "otro") t.v += u; }); return t; };
  const antes = (() => { const d = new Date(f + "T12:00:00"); do { d.setDate(d.getDate() - 1); } while (d.getDay() === 0 || d.getDay() === 6); return d.toISOString().slice(0, 10); })();
  const finPlan = (key) => { const o = (plan.obras || {})[key]; return o && o.fin; };
  const obras = [];
  proyectos.filter((p) => p.origen !== "portalUxcar" && !["Entregado", "Cancelado", "Albarán de carga firmado"].includes(p.estadoTrabajo)).forEach((p) => {
    const key = `p-${p.id}`;
    if (finPlan(key) === antes || (p.fechaReparto || "") === f) obras.push({ key, proyectoId: p.id, nombre: `#${p.numero} ${p.nombre}`, h: huecos(p.recuento, totRecuento(p.recuento)) });
  });
  expedientes.filter((e) => e.estado !== "entregado").forEach((e) => {
    const key = `u-${e.id}`;
    const pl = e.proyectoId ? proyectos.find((x) => x.id === e.proyectoId) : null;
    if (finPlan(key) === antes || ((pl && pl.fechaReparto) || "") === f) obras.push({ key, proyectoId: e.proyectoId || "", nombre: `Uxcar exp. ${e.numero}`, h: huecos(e.recuento, { v: num(e.ventanas), p: num(e.puertas), op: num(e.osciloParalelas) }) });
  });
  const de = (c, o) => c.obra && (c.obra.key === o.key || (o.proyectoId && c.obra.proyectoId === o.proyectoId));
  const libres = caballetes.filter((c) => (c.estado || "libre") === "libre");
  let texto = "";
  if (obras.length) {
    let falta = 0;
    texto += `CABALLETES DE VENTANAS PARA ESE DÍA (caben ${cap.capacidad} ventanas por caballete):\n`;
    obras.forEach((o) => {
      const nec = o.h > 0 ? Math.ceil(o.h / (num(cap.capacidad) || 12)) : 0;
      const ya = caballetes.filter((c) => c.estado === "cargado" && de(c, o)).map((c) => c.numero);
      const res = libres.filter((c) => c.reserva && c.reserva.key === o.key).map((c) => c.numero);
      falta += Math.max(0, nec - ya.length - res.length);
      texto += `  ☐ ${o.nombre}: ${nec ? `${nec} caballete${nec === 1 ? "" : "s"}` : "ventanas sin contar, revisar"}${res.length ? ` · reservados: ${res.join(", ")}` : ""}${ya.length ? ` · ya cargados: ${ya.join(", ")}` : ""}\n`;
    });
    const sinReserva = libres.filter((c) => !c.reserva).length;
    if (falta) texto += `  → Faltan por asignar ${falta}. Libres sin reservar: ${sinReserva}.${falta > sinReserva ? " ⚠ NO HAY SUFICIENTES: reclamad los que están fuera." : ""}\n`;
    texto += "\n";
  }
  // Fuera con la vuelta pasada (media de días de cada cliente o, si no hay, el valor por defecto)
  const medias = {};
  caballetes.forEach((c) => toArray(c.devoluciones).forEach((d) => { const k = String(d.cliente || "").trim().toLowerCase(); if (k) (medias[k] = medias[k] || []).push(num(d.dias)); }));
  const diasCli = (cl) => { const a = medias[String(cl || "").trim().toLowerCase()]; return a && a.length ? Math.max(1, Math.round(a.reduce((x, y) => x + y, 0) / a.length)) : num(pc.diasDevolucion) || 10; };
  const sumar = (fecha, n) => { const d = new Date(fecha + "T12:00:00"); d.setDate(d.getDate() + n); return d.toISOString().slice(0, 10); };
  const tarde = caballetes.filter((c) => c.estado === "fuera" && c.salida).map((c) => ({ c, v: c.salida.vuelvePrevista || (c.salida.fecha ? sumar(c.salida.fecha, diasCli(c.salida.cliente)) : "") })).filter((x) => !x.v || x.v < hoy);
  if (tarde.length) {
    texto += `CABALLETES FUERA QUE YA TENÍAN QUE HABER VUELTO (reclamar):\n`;
    tarde.forEach(({ c, v }) => { texto += `  ☐ ${c.numero} — ${c.salida.cliente || "—"} · ${c.salida.obra || ""}${c.salida.direccion ? ` · ${c.salida.direccion}` : ""}${v ? ` · tenía que volver el ${v.split("-").reverse().join("/")}` : ""}\n`; });
    texto += "\n";
  }
  return { hay: !!(obras.length || tarde.length), texto };
}

export const handler = async (event) => {
  FB_AUTH = await fbAuthQuery().catch(() => "");
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
    const [cristales, persianasAlmacen, proyectos, config, caballetesRaw, expedientesRaw] = await Promise.all([leer("cristales"), leer("persianasAlmacen"), leer("proyectos"), leer("configVentanas"), leer("caballetesVentanas"), leer("portalUxcar/expedientes")]);
    const cab = caballetesSeccion({ f, plan, proyectos: toArray(proyectos), config: config || {}, caballetes: toArray(caballetesRaw), expedientes: toArray(expedientesRaw) });
    if (trozos.length === 0 && !cab.hay) return { statusCode: 200, body: `Nada planificado ni caballetes que avisar para ${f}.` };
    const obrasDia = trozos.map((t) => ({ ...t, obra: (plan.obras || {})[t.id] || { nombre: t.nombre, lineas: [], expNums: [] } }));

    let cuerpo = `PREPARAR HOY PARA EL ${diaTexto.toUpperCase()}\n\n`;
    cuerpo += cab.texto;
    if (obrasDia.length) cuerpo += `Obras del próximo día de trabajo:\n${obrasDia.map((o) => `  • ${o.obra.nombre} — ${o.horas} h${o.obra.ventanas ? ` (${o.obra.ventanas} ventanas)` : ""}${o.obra.inicio && o.obra.inicio < f ? " · (sigue de días anteriores)" : ""}`).join("\n")}\n`;

    // Solo se prepara el material de las obras que EMPIEZAN ese día (las que siguen ya lo tienen)
    const empiezan = obrasDia.filter((o) => !o.obra.inicio || o.obra.inicio === f);
    if (obrasDia.length && empiezan.length === 0) cuerpo += `\nEse día no empieza ninguna obra nueva: se sigue con las que están en marcha.\n`;

    empiezan.forEach((o) => {
      const ob = o.obra;
      cuerpo += `\n==============================\n${ob.nombre.toUpperCase()}\n==============================\n`;
      // Perfiles, refuerzos, herrajes y accesorios, por almacén y estantería
      // (persianas y cristales van aparte, más abajo, con su sitio en su almacén)
      const mat = toArray(ob.lineas).filter((l) => l.seccion !== "persianas" && l.seccion !== "cristal")
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
      } else {
        const crisListado = toArray(ob.lineas).filter((l) => l.seccion === "cristal");
        cuerpo += `\nCRISTALES: no encuentro caballetes de esta obra en el almacén (¿han llegado?)\n`;
        crisListado.forEach((l) => { cuerpo += `   ☐ ${l.codigo}${l.ancho ? ` ${l.ancho}×${l.alto}` : ""} → ${l.uds} (del listado)\n`; });
      }
    });

    cuerpo += `\n— Aviso automático del CRM, según el planning confirmado el ${new Date(plan.publicadoEn || Date.now()).toLocaleDateString("es-ES")}.`;
    await enviar({ to: plan.emailFabrica, subject: `Preparar hoy para el ${diaTexto} — ${empiezan.length ? `${empiezan.length} obra${empiezan.length === 1 ? "" : "s"}` : "caballetes"}`, text: cuerpo });
    return { statusCode: 200, body: "Enviado" };
  } catch (e) {
    console.error("aviso-fabrica-diario:", e.message);
    return { statusCode: 200, body: "Error: " + e.message };
  }
};
