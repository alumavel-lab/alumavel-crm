// Recibe el aviso de Firma.dev cuando un documento se termina de firmar.
// Esta función solo verifica la firma del webhook y responde rápido — la
// parte lenta (descargar el PDF firmado, que a veces tarda unos segundos en
// estar listo en Firma.dev, subirlo a Storage y actualizar el registro) se
// delega a firma-webhook-background.js, que puede tardar mucho más sin que
// Firma.dev llegue a darlo por fallido o a reintentar por su cuenta.
//
// CONFIGURACIÓN PENDIENTE (hacer una vez, desde el panel de Firma.dev):
//   1. Registra este webhook en Firma.dev apuntando a:
//        https://dynamic-eclair-be67a5.netlify.app/.netlify/functions/firma-webhook
//      con el evento "signing_request.completed".
//   2. Copia el "signing secret" que te da Firma.dev al crear el webhook y
//      guárdalo en Netlify como FIRMA_WEBHOOK_SECRET.
//   3. En Firebase Console > Storage, activa Storage si no lo está, y pon estas
//      reglas (Storage > Rules) para permitir que esta función suba/lea los PDFs:
//        rules_version = '2';
//        service firebase.storage {
//          match /b/{bucket}/o {
//            match /presupuestos-firmados/{allPaths=**} {
//              allow read, write: if true;
//            }
//          }
//        }
//      (El acceso real de esta carpeta queda igual de abierto que la base de
//      datos actual del CRM — es el mismo modelo de seguridad que ya usáis.)
//
// Variables de entorno necesarias en Netlify:
//   FIRMA_API_KEY_TEST / FIRMA_API_KEY = igual que en firma-enviar.js
//   FIRMA_WEBHOOK_SECRET               = el secreto que da Firma.dev al crear el webhook

import crypto from "crypto";

function verificarFirmaWebhook(rawBody, headerSignature, secret) {
  if (!secret || !headerSignature) return false;
  // Formato: "t=1707500000,v1=abc123..."
  const partes = Object.fromEntries(headerSignature.split(",").map((p) => p.split("=")));
  const { t, v1 } = partes;
  if (!t || !v1) return false;
  const payloadFirmado = `${t}.${rawBody}`;
  const esperado = crypto.createHmac("sha256", secret).update(payloadFirmado).digest("hex");
  try {
    return crypto.timingSafeEqual(Buffer.from(esperado), Buffer.from(v1));
  } catch {
    return false;
  }
}

export const handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method not allowed" };
  }

  try {
    const rawBody = event.body || "";
    const secret = process.env.FIRMA_WEBHOOK_SECRET;
    const headerSignature = event.headers["x-firma-signature"] || event.headers["X-Firma-Signature"];

    if (secret) {
      const valido = verificarFirmaWebhook(rawBody, headerSignature, secret);
      if (!valido) {
        console.error("Firma de webhook inválida — posible petición falsa");
        return { statusCode: 401, body: JSON.stringify({ error: "Firma de webhook inválida" }) };
      }
    } else {
      // Sin FIRMA_WEBHOOK_SECRET configurado todavía no se puede verificar el
      // origen — se procesa igualmente para no bloquear las pruebas iniciales,
      // pero esto debe configurarse antes de pasar a producción real.
      console.warn("FIRMA_WEBHOOK_SECRET no configurado — el webhook no está verificado.");
    }

    const payload = JSON.parse(rawBody);
    const { type, data } = payload;

    if (type !== "signing_request.completed") {
      // Otros eventos (viewed, sent, recipient.signed...) se reciben pero se
      // ignoran por ahora — solo nos interesa cuando TODOS han firmado.
      return { statusCode: 200, body: JSON.stringify({ recibido: true, ignorado: type }) };
    }

    const signingRequestId = data?.signing_request?.id || data?.id;
    if (!signingRequestId) {
      console.error("Webhook signing_request.completed sin id:", JSON.stringify(payload));
      return { statusCode: 200, body: JSON.stringify({ error: "Sin signing_request.id en el payload" }) };
    }

    // Lanza la función de fondo y no espera a que termine (puede tardar bastante
    // más de lo que Firma.dev espera para dar el webhook por recibido) — solo
    // espera a que Netlify confirme que la ha aceptado para ejecutarla.
    try {
      await fetch(`${process.env.URL || "https://dynamic-eclair-be67a5.netlify.app"}/.netlify/functions/firma-webhook-background`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ signingRequestId }),
      });
    } catch (err) {
      // Si ni siquiera se pudo lanzar la función de fondo, se registra pero se
      // responde 200 igual — no tiene sentido que Firma.dev lo reintente solo
      // por esto, y queda constancia en los logs para revisarlo a mano.
      console.error("No se pudo lanzar firma-webhook-background:", err.message);
    }

    return { statusCode: 200, body: JSON.stringify({ recibido: true, signingRequestId }) };
  } catch (err) {
    console.error("Excepción en firma-webhook:", err.message);
    console.error("Stack:", err.stack);
    // Devolvemos 200 igualmente para que Firma.dev no reintente indefinidamente
    // por un error nuestro — el fallo queda registrado en los logs de Netlify.
    return { statusCode: 200, body: JSON.stringify({ error: err.message }) };
  }
};
