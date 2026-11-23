// Lanzador programado (ver netlify.toml) del aviso de licitaciones de carpintería.
// Las funciones programadas solo pueden durar 30 segundos y leer los ficheros oficiales
// lleva más, así que esta solo arranca "avisos-licitaciones-background" (hasta 15 min).
export const handler = async () => {
  const base = process.env.URL || process.env.DEPLOY_PRIME_URL;
  if (!base) return { statusCode: 200, body: "Sin URL del sitio (variable URL de Netlify)." };
  const r = await fetch(`${base}/.netlify/functions/avisos-licitaciones-background`, { method: "POST" });
  return { statusCode: 200, body: `Lanzado (${r.status})` };
};
