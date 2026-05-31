const GAS_URL =
  "https://script.google.com/macros/s/AKfycbyNpjpZzIEhxO0H-xn-24BIZPRK8lb4gDzho_IBTcOi6y6fuVOD-FXF4ZR5-PRD4x1Iww/exec";

const CORS = {
  "Access-Control-Allow-Origin":  "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

exports.handler = async (event) => {
  // Preflight
  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 204, headers: CORS, body: "" };
  }

  try {
    const isGet = event.httpMethod === "GET";
    const qs    = event.rawQuery ? "?" + event.rawQuery : "";

    const resp = await fetch(isGet ? GAS_URL + qs : GAS_URL, {
      method:  event.httpMethod,
      headers: { "Content-Type": "application/json" },
      body:    isGet ? undefined : event.body,
      redirect: "follow",
    });

    const text = await resp.text();
    return {
      statusCode: 200,
      headers: { ...CORS, "Content-Type": "application/json" },
      body: text,
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: CORS,
      body: JSON.stringify({ error: err.message }),
    };
  }
};
