const GAS_URL =
  "https://script.google.com/macros/s/AKfycbwQyMc_ZbQQiWfwg4pKbncAS1L4so0ikWLm7O0_gfuPf8F8eNf5p-PG7HCceXG8OL8ihA/exec";

const CORS = {
  "Access-Control-Allow-Origin":  "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

export async function onRequest(context) {
  const { request } = context;

  if (request.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: CORS });
  }

  try {
    const url  = new URL(request.url);
    const isGet = request.method === "GET";

    const resp = await fetch(isGet ? GAS_URL + url.search : GAS_URL, {
      method:  request.method,
      headers: { "Content-Type": "application/json" },
      body:    isGet ? undefined : await request.text(),
      redirect: "follow",
    });

    const text = await resp.text();
    return new Response(text, {
      status: 200,
      headers: { ...CORS, "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: err.message }), {
      status: 500,
      headers: CORS,
    });
  }
}
