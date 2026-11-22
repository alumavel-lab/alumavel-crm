// Licitaciones públicas de carpintería (PVC, aluminio, metálica, ventanas, cerramientos)
// en toda España. La lanza cada mañana "avisos-licitaciones" (ver netlify.toml).
//
// Qué hace:
//  1. Lee los ficheros oficiales de la Plataforma de Contratación del Sector Público
//     (los del Estado y los de las plataformas autonómicas: Cataluña, Euskadi, Madrid,
//     Andalucía, Galicia...) desde la última revisión.
//  2. Se queda con las de carpintería (por código CPV o por palabras del título) y las
//     guarda en Firebase en "licitaciones/items".
//  3. Manda un correo con las NUEVAS en plazo y con las que CIERRAN en los próximos 7 días,
//     primero las de PVC y las de suministro (lo que más nos interesa), luego el resto.
//
// Destinatario: AVISOS_LICITACIONES_EMAIL_TO (puede llevar varias direcciones separadas por
// comas); si no está, EMAIL_USER. Envío: RESEND_API_KEY o la cuenta de Microsoft 365.
// Búsqueda a mano: desde el CRM (Licitaciones → «Buscar ahora»), que llama a
// /.netlify/functions/avisos-licitaciones-background?dias=N&sinCorreo=1 con la sesión del usuario.

import nodemailer from "nodemailer";

const FIREBASE_DB_URL = "https://crmalumavel-default-rtdb.europe-west1.firebasedatabase.app";
const FIREBASE_WEB_API_KEY = "AIzaSyBf51Gy2drHdUyJ4kCstNdvvV2h3uYe4RM";
const FEEDS = [
  { clave: "estatal", url: "https://contrataciondelestado.es/sindicacion/sindicacion_643/licitacionesPerfilesContratanteCompleto3.atom" },
  { clave: "autonomicas", url: "https://contrataciondelestado.es/sindicacion/sindicacion_1044/PlataformasAgregadasSinMenores.atom" },
];
const MAX_PAGINAS = 40; // tope de seguridad por fichero en la pasada diaria
const MAX_PAGINAS_MANUAL = 400; // búsqueda a mano desde el CRM ("Buscar ahora", hasta 45 días)

async function fbAuthQuery() {
  const email = (process.env.FIREBASE_ROBOT_EMAIL || "").trim();
  const password = (process.env.FIREBASE_ROBOT_PASSWORD || "").trim();
  if (email && password) {
    const r = await fetch(`https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=${FIREBASE_WEB_API_KEY}`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, password, returnSecureToken: true }),
    });
    const d = await r.json().catch(() => ({}));
    if (r.ok && d.idToken) return `?auth=${encodeURIComponent(d.idToken)}`;
    console.error("No se pudo entrar con el usuario robot:", JSON.stringify(d.error || d));
  }
  const secreto = (process.env.FIREBASE_DB_SECRET || "").trim();
  return secreto ? `?auth=${encodeURIComponent(secreto)}` : "";
}
let FB_AUTH = "";
const leer = async (ruta) => { const r = await fetch(`${FIREBASE_DB_URL}/${ruta}.json${FB_AUTH}`); return r.ok ? r.json() : null; };
const escribir = async (ruta, datos, metodo = "PUT") => {
  const r = await fetch(`${FIREBASE_DB_URL}/${ruta}.json${FB_AUTH}`, { method: metodo, headers: { "Content-Type": "application/json" }, body: JSON.stringify(datos) });
  if (!r.ok) throw new Error(`Firebase ${ruta}: ${r.status} ${await r.text()}`);
};

async function enviar({ to, subject, text }) {
  const destinos = to.split(",").map((s) => s.trim()).filter(Boolean);
  if (process.env.RESEND_API_KEY) {
    const r = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: { Authorization: `Bearer ${process.env.RESEND_API_KEY}`, "Content-Type": "application/json" },
      body: JSON.stringify({ from: process.env.RESEND_FROM || "Licitaciones ALUMAVEL <onboarding@resend.dev>", to: destinos, subject, text }),
    });
    if (!r.ok) throw new Error("Resend: " + (await r.text()));
    return;
  }
  const user = process.env.EMAIL_USER, pass = process.env.EMAIL_PASSWORD;
  if (!user || !pass) throw new Error("Correo sin configurar (faltan RESEND_API_KEY o EMAIL_USER/EMAIL_PASSWORD)");
  const t = nodemailer.createTransport({ host: "smtp.office365.com", port: 587, secure: false, auth: { user, pass } });
  await t.sendMail({ from: `"Licitaciones ALUMAVEL" <${user}>`, to: destinos.join(", "), subject, text });
}

// ---------- Lectura del fichero oficial (XML/Atom) sin librerías ----------
const desc = (s) => String(s || "").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, n) => String.fromCharCode(n)).replace(/&amp;/g, "&").trim();
const uno = (xml, tag) => { const m = xml.match(new RegExp(`<${tag}(?:\\s[^>]*)?>([^<]*)</${tag}>`)); return m ? desc(m[1]) : ""; };
const todos = (xml, tag) => [...xml.matchAll(new RegExp(`<${tag}(?:\\s[^>]*)?>([^<]*)</${tag}>`, "g"))].map((m) => desc(m[1]));

const CPV_CARP = ["45421", "44221", "44220"];
// Dentro de "carpintería" (45421...) hay códigos que no son lo nuestro: tabiques y falsos techos.
const CPV_NO = ["45421141", "45421146"];
// El título dice que el objeto ES carpintería (castellano, catalán, gallego, euskera)
const CARP = /ventan|finestr|leiho|carpinter|fusteri|tancaments? (exterior|de fa)|cerramientos? (de )?(aluminio|pvc|vidrio|exterior|de fachada)|acristal|(puertas?|portas?) (\S+ ){0,3}(de )?(aluminio|pvc|met[aá]l|exterior|antip[aá]nico|cortafuego|autom[aá]tic)|persiana|mampara|perfiler[ií]a de aluminio|cristaler/i;
const SUMI = /suministr|subministr|hornidura|adquisici/i;
// Cosas que se cuelan por las palabras pero no son carpintería (tuberías de PVC, "sobres con ventana"...)
const FALSO = /tuber|tubos?\b|manguer|mangueir|desag[üu]e|bajante|geomembran|l[aá]mina impermeab|suelo|pavimento|c[eé]sped|canonad|llaut[oó]|lat[oó]n|saneamient|abastecimiento|xarxa|red de agua|riego|sulfato|policloruro|hidr[oó]xido|ventanilla|latas? de aluminio|papel de aluminio|envases?|bandejas?|ventanas? (de )?(tramitaci|atenci)|tabiquer|falsos? techos?|nichos/i;
const NUNCA = /sobres|furgoneta|veh[ií]culo|redacci[oó]n|proyecto y asistencia|direcci[oó]n (facultativa|de (la )?obra)|asistencia t[eé]cnica|coordinaci[oó]n de seguridad/i;
const TIPOS = { 1: "Suministro", 2: "Servicio", 3: "Obra", 21: "Gestión serv.", 31: "Concesión obra", 7: "Admin. especial", 8: "Privado" };
// Grupos: A = suministro (o suministro y colocación) de carpintería · B = obra/servicio cuyo
// objeto es carpintería · C = obra general o lote de materiales donde la carpintería es una parte.
const GRUPOS = {
  A: "SUMINISTRO DE CARPINTERÍA (lo que más nos encaja)",
  B: "OBRAS / SERVICIOS DE CARPINTERÍA",
  C: "OBRAS O LOTES DONDE LA CARPINTERÍA ES UNA PARTE (subcontrata o lote)",
};

function analizar(entrada) {
  const titulo = uno(entrada, "title").replace(/\s+/g, " ");
  const cpvs = [...new Set(todos(entrada, "cbc:ItemClassificationCode"))];
  const cpvCarp = cpvs.some((c) => CPV_CARP.some((p) => c.startsWith(p)) && !CPV_NO.includes(c));
  const carp = CARP.test(titulo);
  if (NUNCA.test(titulo) || (FALSO.test(titulo) && !carp)) return null;
  if (!(carp || cpvCarp)) return null;
  const tipo = uno(entrada, "cbc:TypeCode");
  const sumi = tipo === "1" || SUMI.test(titulo);
  const enlace = (entrada.match(/<link href="([^"]+)"/) || [])[1] || "";
  const importe = parseFloat(uno(entrada, "cbc:TotalAmount") || uno(entrada, "cbc:EstimatedOverallContractAmount")) || null;
  return {
    id: uno(entrada, "id").split("/").pop(), titulo, cpvs,
    pvc: /pvc/i.test(titulo), grupo: carp ? (sumi ? "A" : "B") : "C",
    estado: uno(entrada, "cbc-place-ext:ContractFolderStatusCode"),
    expediente: uno(entrada, "cbc:ContractFolderID"),
    organo: uno((entrada.match(/<cac:PartyName>[\s\S]*?<\/cac:PartyName>/) || [""])[0], "cbc:Name"),
    ciudad: uno(entrada, "cbc:CityName"), provincia: uno(entrada, "cbc:CountrySubentity"),
    // Lugar de ejecución (código NUTS de la provincia) y código postal del organismo: el CRM los usa
    // para calcular la distancia desde la fábrica (María, Almería).
    nuts: uno(entrada, "cbc:CountrySubentityCode"), cp: uno(entrada, "cbc:PostalZone"),
    tipo: TIPOS[tipo] || tipo, importe,
    fin: uno(entrada, "cbc:EndDate"), hora: uno(entrada, "cbc:EndTime").slice(0, 5),
    enlace: desc(enlace), actualizado: uno(entrada, "updated"),
  };
}

async function recorrer(feed, desde, maxPaginas, guardarTanda) {
  const encontradas = {};
  let url = feed.url, paginas = 0, masAntigua = "9999", pendientes = {};
  while (url && paginas < maxPaginas) {
    const r = await fetch(url);
    if (!r.ok) throw new Error(`${feed.clave}: ${r.status} en ${url}`);
    const xml = await r.text();
    paginas++;
    const siguiente = (xml.match(/<link href="([^"]+)" rel="next"\/>/) || [])[1];
    for (const m of xml.matchAll(/<entry>([\s\S]*?)<\/entry>/g)) {
      const act = uno(m[1], "updated");
      if (act < masAntigua) masAntigua = act;
      if (act < desde) continue;
      const l = analizar(m[1]);
      if (l && (!encontradas[l.id] || encontradas[l.id].actualizado < l.actualizado)) pendientes[l.id] = encontradas[l.id] = { ...l, origen: feed.clave };
    }
    // Se guarda cada 10 páginas: si una búsqueda larga se corta, lo encontrado no se pierde.
    if (paginas % 10 === 0 && Object.keys(pendientes).length) { await guardarTanda(pendientes); pendientes = {}; }
    url = masAntigua >= desde ? siguiente : null;
  }
  if (Object.keys(pendientes).length) await guardarTanda(pendientes);
  return { encontradas, paginas };
}

const hoyEs = () => new Date().toLocaleDateString("sv-SE", { timeZone: "Europe/Madrid" }); // AAAA-MM-DD
const diasHasta = (f) => Math.round((new Date(f + "T12:00:00") - new Date(hoyEs() + "T12:00:00")) / 864e5);
const eur = (n) => (n ? n.toLocaleString("es-ES", { maximumFractionDigits: 0 }) + " €" : "importe sin indicar");
const linea = (l) => `  ☐ [${l.fin ? `cierra ${l.fin.split("-").reverse().join("/")}${l.hora ? " " + l.hora : ""} · ${diasHasta(l.fin)} días` : "sin fecha de cierre"}] ${l.pvc ? "PVC · " : ""}${l.tipo} · ${eur(l.importe)}\n     ${l.titulo.replace(/\s+/g, " ").slice(0, 220)}\n     ${l.organo}${l.ciudad || l.provincia ? " — " + (l.ciudad || l.provincia) : ""} · Exp. ${l.expediente}\n     ${l.enlace}\n`;
const ordenar = (a, b) => String(a.grupo).localeCompare(String(b.grupo)) || b.pvc - a.pvc || String(a.fin).localeCompare(String(b.fin));
const porGrupos = (arr) => ["A", "B", "C"].map((g) => { const x = arr.filter((l) => (l.grupo || "C") === g); return x.length ? `\n-- ${GRUPOS[g]} (${x.length}) --\n${x.map(linea).join("\n")}` : ""; }).join("");

// Solo quien ha iniciado sesión en el CRM puede pedir una búsqueda a mano (como en enviar-email).
async function usuarioAutorizado(event) {
  const token = (event.headers?.authorization || event.headers?.Authorization || "").replace(/^Bearer\s+/i, "").trim();
  if (!token) return false;
  const r = await fetch(`https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=${FIREBASE_WEB_API_KEY}`, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ idToken: token }),
  });
  const d = await r.json().catch(() => ({}));
  const uid = r.ok && d.users && d.users[0] && d.users[0].localId;
  if (!uid) return false;
  const x = await fetch(`${FIREBASE_DB_URL}/equipo/${uid}.json?auth=${encodeURIComponent(token)}`);
  return x.ok && !!(await x.json().catch(() => null));
}

export const handler = async (event) => {
  try {
    const q = (event && event.queryStringParameters) || {};
    const manual = !!(q.dias || q.sinCorreo);
    if (manual && !(await usuarioAutorizado(event))) return { statusCode: 200, body: "Sin permiso: hay que entrar en el CRM." };
    FB_AUTH = await fbAuthQuery();
    const estado = (await leer("licitaciones/estado")) || {};
    // La pasada automática no se repite si ya se hizo hace menos de 30 minutos (evita correos repetidos).
    if (!manual && estado.ultimaRevision && Date.now() - new Date(estado.ultimaRevision).getTime() < 30 * 6e4) return { statusCode: 200, body: "Ya revisado hace poco." };
    if (estado.enCurso && Date.now() - new Date(estado.enCurso).getTime() < 16 * 6e4) return { statusCode: 200, body: "Ya hay una búsqueda en marcha." };
    const inicio = new Date().toISOString();
    await escribir("licitaciones/estado/enCurso", inicio);
    // Desde cuándo mirar: la última revisión (menos 2 h de margen); la primera vez, 3 días. A mano: los días pedidos (máx. 45).
    const dias = Math.min(45, parseInt(q.dias, 10) || 0);
    const desde = dias > 0 ? new Date(Date.now() - dias * 864e5).toISOString()
      : estado.ultimaRevision ? new Date(new Date(estado.ultimaRevision).getTime() - 2 * 36e5).toISOString()
      : new Date(Date.now() - 3 * 864e5).toISOString();

    const guardadas = (await leer("licitaciones/items")) || {};
    const hoy = hoyEs();
    const avisarNuevas = [];
    const yaContadas = new Set();
    // Se escribe campo a campo ("id/campo"): así no se pisa lo que el equipo haya marcado
    // en el CRM (seguimiento, notas, responsable...).
    const guardarTanda = async (tanda) => {
      const cambios = {};
      for (const l of Object.values(tanda)) {
        for (const [k, v] of Object.entries(l)) cambios[`${l.id}/${k}`] = v;
        if (!guardadas[l.id]) {
          cambios[`${l.id}/primeraVez`] = inicio;
          if (!yaContadas.has(l.id) && l.estado === "PUB" && l.fin && l.fin >= hoy) { avisarNuevas.push(l); yaContadas.add(l.id); }
        }
      }
      if (Object.keys(cambios).length) await escribir("licitaciones/items", cambios, "PATCH");
    };
    const maxPag = manual ? MAX_PAGINAS_MANUAL : MAX_PAGINAS;
    const resultados = await Promise.all(FEEDS.map((f) => recorrer(f, desde, maxPag, guardarTanda)));
    const nuevas = Object.assign({}, ...resultados.map((r) => r.encontradas));

    const todasItems = { ...guardadas };
    for (const l of Object.values(nuevas)) todasItems[l.id] = { ...(guardadas[l.id] || {}), ...l };
    const idsNuevas = new Set(avisarNuevas.map((l) => l.id));
    const cierran = Object.values(todasItems).filter((l) => l.estado === "PUB" && l.fin && l.fin >= hoy && diasHasta(l.fin) <= 7 && !idsNuevas.has(l.id) && l.seguimiento !== "descartada");

    await escribir("licitaciones/estado", { ultimaRevision: inicio, terminada: new Date().toISOString(), paginas: resultados.map((r) => r.paginas), encontradas: Object.keys(nuevas).length, nuevas: avisarNuevas.length, desde, manual });

    if (q.sinCorreo) return { statusCode: 200, body: `Guardadas ${Object.keys(nuevas).length} (${avisarNuevas.length} nuevas en plazo). Sin correo.` };
    if (!avisarNuevas.length && !cierran.length) return { statusCode: 200, body: "Nada nuevo ni que cierre esta semana." };
    avisarNuevas.sort(ordenar); cierran.sort(ordenar);
    let cuerpo = `LICITACIONES DE CARPINTERÍA — ${hoy.split("-").reverse().join("/")}\n(también las tenéis en el CRM, sección «Licitaciones»)\n`;
    if (avisarNuevas.length) cuerpo += `\n=== NUEVAS EN PLAZO (${avisarNuevas.length}) ===\n${porGrupos(avisarNuevas)}`;
    if (cierran.length) cuerpo += `\n=== CIERRAN EN LOS PRÓXIMOS 7 DÍAS (${cierran.length}) ===\n${porGrupos(cierran)}`;
    cuerpo += `\n— Aviso automático del CRM. Fuente: Plataforma de Contratación del Sector Público (Estado y plataformas autonómicas). Comprobad siempre plazo y pliego en el enlace.`;
    const to = process.env.AVISOS_LICITACIONES_EMAIL_TO || process.env.EMAIL_USER;
    if (!to) throw new Error("Sin destinatario (AVISOS_LICITACIONES_EMAIL_TO o EMAIL_USER)");
    await enviar({ to, subject: `Licitaciones carpintería: ${avisarNuevas.length} nuevas${cierran.length ? `, ${cierran.length} cierran esta semana` : ""}`, text: cuerpo });
    return { statusCode: 200, body: "Enviado" };
  } catch (e) {
    console.error("avisos-licitaciones:", e.message);
    await escribir("licitaciones/estado/error", `${new Date().toISOString()} · ${e.message}`).catch(() => {});
    return { statusCode: 200, body: "Error: " + e.message };
  } finally {
    await escribir("licitaciones/estado/enCurso", null).catch(() => {});
  }
};
