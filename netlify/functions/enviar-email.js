// Envía correos desde el CRM usando la cuenta de Microsoft 365 de la empresa
// (alumavel@alumavel.es), sin que el usuario tenga que abrir su propio correo.
//
// Variables de entorno necesarias en Netlify (Site settings > Environment variables):
//   EMAIL_USER     = alumavel@alumavel.es
//   EMAIL_PASSWORD = la contraseña de la cuenta (o una "contraseña de aplicación"
//                    si la cuenta tiene verificación en dos pasos activada)
//
// Si el envío falla con un error de autenticación, lo más probable es que la
// cuenta tenga desactivado el "SMTP autenticado" — hay que activarlo desde el
// Centro de administración de Microsoft 365 (Exchange admin center > buzones >
// esta cuenta > administrar aplicaciones de correo > activar "SMTP autenticado").

import nodemailer from "nodemailer";

export const handler = async (event) => {
  try {
    const body = JSON.parse(event.body || "{}");
    const { destinatario, asunto, cuerpo, replyTo, adjuntos } = body;

    if (!destinatario || !asunto || !cuerpo) {
      return { statusCode: 400, body: JSON.stringify({ error: "Faltan datos: destinatario, asunto o cuerpo." }) };
    }

    const user = process.env.EMAIL_USER;
    const pass = process.env.EMAIL_PASSWORD;
    if (!user || !pass) {
      return { statusCode: 200, body: JSON.stringify({ error: "El envío de correo no está configurado todavía (faltan EMAIL_USER / EMAIL_PASSWORD en Netlify)." }) };
    }

    const transporter = nodemailer.createTransport({
      host: "smtp.office365.com",
      port: 587,
      secure: false, // usa STARTTLS en el puerto 587
      auth: { user, pass },
    });

    // adjuntos: [{ nombre, dataUrl }] — dataUrl tipo "data:application/pdf;base64,...."
    const attachments = (adjuntos || [])
      .filter((a) => a && a.dataUrl)
      .map((a) => {
        const match = /^data:([^;]+);base64,(.*)$/.exec(a.dataUrl);
        return {
          filename: a.nombre || "adjunto",
          content: match ? match[2] : a.dataUrl,
          encoding: "base64",
          contentType: match ? match[1] : undefined,
        };
      });

    await transporter.sendMail({
      from: `"ALUMAVEL" <${user}>`,
      to: destinatario,
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
