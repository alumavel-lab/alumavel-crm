// Solo atiende a quien ha iniciado sesión en el CRM: el navegador manda el token del
// usuario (cabecera Authorization) y aquí se comprueba con Firebase que es válido y que
// esa persona está en el equipo o es usuario del portal. Sin esto, cualquiera que
// conociera la dirección podría usar esta función.
const FIREBASE_DB_URL = "https://crmalumavel-default-rtdb.europe-west1.firebasedatabase.app";
const FIREBASE_WEB_API_KEY = "AIzaSyBf51Gy2drHdUyJ4kCstNdvvV2h3uYe4RM";
async function usuarioAutorizado(event) {
  const token = (event.headers?.authorization || event.headers?.Authorization || "").replace(/^Bearer\s+/i, "").trim();
  if (!token) return false;
  const r = await fetch(`https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=${FIREBASE_WEB_API_KEY}`, {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ idToken: token }),
  });
  const d = await r.json().catch(() => ({}));
  const uid = r.ok && d.users && d.users[0] && d.users[0].localId;
  if (!uid) return false;
  for (const ruta of [`equipo/${uid}`, `portalUsuarios/${uid}`]) {
    const x = await fetch(`${FIREBASE_DB_URL}/${ruta}.json?auth=${encodeURIComponent(token)}`);
    if (x.ok && await x.json().catch(() => null)) return true;
  }
  return false;
}

export const handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      body: JSON.stringify({ error: 'Method not allowed' }),
    };
  }

  try {
    if (!(await usuarioAutorizado(event))) {
      return {
        statusCode: 401,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ error: { message: 'Sesión caducada o sin permiso. Cierra sesión y vuelve a entrar en el CRM.' } }),
      };
    }

    const apiKey = process.env.ANTHROPIC_API_KEY;

    if (!apiKey) {
      console.error('FALTA ANTHROPIC_API_KEY');
      return {
        statusCode: 500,
        body: JSON.stringify({ error: 'ANTHROPIC_API_KEY no configurada en Netlify' }),
      };
    }

    const requestBody = JSON.parse(event.body);
    console.log('Modelo solicitado:', requestBody.model);
    console.log('Tamano del body recibido (KB):', Math.round(event.body.length / 1024));

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify(requestBody),
    });

    const data = await response.json();

    if (!response.ok) {
      console.error('Anthropic API devolvio error. Status:', response.status);
      console.error('Detalle del error:', JSON.stringify(data));
    } else {
      console.log('Respuesta OK de Anthropic. Status:', response.status);
    }

    return {
      statusCode: response.status,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
      },
      body: JSON.stringify(data),
    };
  } catch (error) {
    console.error('Excepcion en anthropic-proxy:', error.message);
    console.error('Stack:', error.stack);
    return {
      statusCode: 500,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
      },
      body: JSON.stringify({ error: error.message }),
    };
  }
};
