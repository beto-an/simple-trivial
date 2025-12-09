// functions/sheet.js
const GOOGLE_CSV_URL =
  "https://docs.google.com/spreadsheets/d/e/2PACX-1vTnohmkk48wDzR1-ZgoJtJoJyhRZCyQBchOY28hN5F2e-P4ZloIuqBHZlk3-HIJ_OvaPLaHjaucv64P/pub?output=csv";

export async function onRequest(context) {
  const { request } = context;

  // Handle preflight (if any)
  if (request.method === "OPTIONS") {
    return new Response(null, {
      status: 204,
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET,OPTIONS",
        "Access-Control-Allow-Headers": "*",
      },
    });
  }

  try {
    const upstream = await fetch(GOOGLE_CSV_URL, {
      // Optional: you can tweak cf options if needed
    });

    if (!upstream.ok) {
      return new Response("Upstream error from Google Sheets", {
        status: 502,
      });
    }

    const body = await upstream.text();

    return new Response(body, {
      status: 200,
      headers: {
        "Content-Type": "text/csv; charset=utf-8",
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET,OPTIONS",
        "Access-Control-Allow-Headers": "*",
      },
    });
  } catch (err) {
    return new Response("Error fetching Google Sheet", { status: 500 });
  }
}
