// Aviso diario de materiales por debajo de su stock mínimo.
// Se ejecuta todos los días (ver "schedule" en netlify.toml) y revisa todos los
// materiales cuyo stockReal ha caído a su stockMinimo o por debajo — lo mismo que
// muestra la pestaña "A reponer" de Stock, pero avisado por email sin tener que
// entrar al CRM a mirarlo.
//
// Manda UN solo correo resumen, agrupado por proveedor, a Miguel (EMAIL_USER, o
// AVISOS_STOCK_EMAIL_TO si se define en Netlify). Usa las mismas variables de
// entorno que el resto de emails del CRM (EMAIL_USER y EMAIL_PASSWORD).

import nodemailer from "nodemailer";

const FIREBASE_DB_URL = "https://crmalumavel-default-rtdb.europe-west1.firebasedatabase.app";

const toArray = (obj) => (obj ? Object.values(obj) : []);

export const handler = async () => {
  try {
    const user = process.env.EMAIL_USER;
    const pass = process.env.EMAIL_PASSWORD;
    const destinatario = process.env.AVISOS_STOCK_EMAIL_TO || "alumavel@gmail.com";

    if (!user || !pass) {
      console.error("Aviso de stock mínimo: faltan EMAIL_USER / EMAIL_PASSWORD, no se puede enviar.");
      return { statusCode: 200, body: "Faltan variables de entorno de email." };
    }

    const [materialesRes, proveedoresRes] = await Promise.all([
      fetch(`${FIREBASE_DB_URL}/materiales.json`),
      fetch(`${FIREBASE_DB_URL}/proveedores.json`),
    ]);
    const materiales = toArray(await materialesRes.json());
    const proveedores = toArray(await proveedoresRes.json());

    const bajoMinimo = materiales.filter((m) => (m.stockReal ?? 0) <= (m.stockMinimo ?? 0));

    if (bajoMinimo.length === 0) {
      console.log("Aviso de stock mínimo: ningún material por debajo de su mínimo hoy.");
      return { statusCode: 200, body: "Ningún material por debajo de su mínimo." };
    }

    const nombreProveedor = (id) => proveedores.find((p) => p.id === id)?.nombre || "Sin proveedor";
    const hoy = new Date().toISOString().slice(0, 10);

    const porProveedor = {};
    bajoMinimo.forEach((m) => {
      const clave = nombreProveedor(m.proveedorId);
      (porProveedor[clave] = porProveedor[clave] || []).push(m);
    });

    const bloques = Object.entries(porProveedor).map(([proveedor, mats]) => {
      const lineas = mats.map((m) =>
        `   • ${m.codigo || "s/cod"} — ${m.descripcion} — tienes ${m.stockReal ?? 0}, mínimo ${m.stockMinimo ?? 0}`
      ).join("\n");
      return `${proveedor}:\n${lineas}`;
    }).join("\n\n");

    const cuerpo = `${bajoMinimo.length} material(es) por debajo de su stock mínimo hoy (${hoy}):\n\n${bloques}\n\nEntra en el CRM (Stock → "A reponer") para generar el pedido.\n\nEste correo se envía automáticamente cada día mientras haya materiales por debajo del mínimo.`;

    const transporter = nodemailer.createTransport({
      host: "smtp.office365.com",
      port: 587,
      secure: false,
      auth: { user, pass },
    });

    await transporter.sendMail({
      from: `"ALUMAVEL CRM — Stock" <${user}>`,
      to: destinatario,
      subject: `${bajoMinimo.length} material(es) por debajo de stock mínimo — ${hoy}`,
      text: cuerpo,
    });

    console.log(`Aviso de stock mínimo enviado: ${bajoMinimo.length} material(es).`);
    return { statusCode: 200, body: `Aviso enviado: ${bajoMinimo.length} material(es).` };
  } catch (err) {
    console.error("Error en el aviso de stock mínimo:", err);
    return { statusCode: 200, body: "Error: " + err.message };
  }
};
