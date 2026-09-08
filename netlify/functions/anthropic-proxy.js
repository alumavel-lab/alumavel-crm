// Función de servidor (Netlify Function) que hace de intermediaria segura entre el
// CRM y la API de Anthropic. La clave de API vive SOLO aquí, como variable de entorno
// en Netlify — nunca en el código del navegador, donde cualquiera podría verla.
//
// El frontend llama a /.netlify/functions/anthropic-proxy en vez de llamar
// directamente a api.anthropic.com.

exports.handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method not allowed" };
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: "Falta configurar ANTHROPIC_API_KEY en las variables de entorno de Netlify." }),
    };
  }

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: event.body,
    });

    const data = await response.text();
    return {
      statusCode: response.status,
      headers: { "Content-Type": "application/json" },
      body: data,
    };
  } catch (err) {
    return {
      statusCode: 500,
      body: JSON.stringify({ error: "Error al conectar con la IA: " + err.message }),
    };
  }
};
