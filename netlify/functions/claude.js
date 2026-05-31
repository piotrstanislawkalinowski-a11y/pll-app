exports.handler = async function (event) {
  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 200, headers: corsHeaders(), body: "" };
  }
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, headers: corsHeaders(), body: JSON.stringify({ error: "Method not allowed" }) };
  }
  let body;
  try { body = JSON.parse(event.body); }
  catch { return { statusCode: 400, headers: corsHeaders(), body: JSON.stringify({ error: "Nieprawidłowy JSON" }) }; }

  const { text, mode, apiKey } = body;
  if (!text || !apiKey) {
    return { statusCode: 400, headers: corsHeaders(), body: JSON.stringify({ error: "Brak tekstu lub klucza API" }) };
  }

  const audience = mode === "zarzad" ? "Zarządu" : "Rady Nadzorczej";
  const audienceFull = mode === "zarzad" ? "Zarządu Port Lotniczy Lublin S.A." : "Rady Nadzorczej Port Lotniczy Lublin S.A.";

  const systemPrompt = `Jesteś ekspertem od przygotowywania profesjonalnych materiałów korporacyjnych dla ${audienceFull}. Zwróć WYŁĄCZNIE sformatowany HTML bez DOCTYPE/html/head/body. Używaj: h1,h2,h3,p,ul,ol,li,table,thead,tbody,tr,th,td,blockquote,strong,em. Struktura: 1) h1=tytuł 2) blockquote=executive summary 3) h2/h3=sekcje 4) tabele dla danych 5) blockquote dla wniosków. Zacznij od <h1>.`;

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01"
      },
      body: JSON.stringify({
        model: "claude-haiku-4-5",
        max_tokens: 4096,
        system: systemPrompt,
        messages: [{ role: "user", content: `Sformatuj jako materiał dla ${audience}:\n\n${text}` }]
      })
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      return { statusCode: response.status, headers: corsHeaders(), body: JSON.stringify({ error: err?.error?.message || response.statusText }) };
    }

    const data = await response.json();
    const content = data?.content?.[0]?.text;
    if (!content) return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: "Brak odpowiedzi z Claude" }) };

    return { statusCode: 200, headers: corsHeaders(), body: JSON.stringify({ result: content }) };
  } catch (err) {
    return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: err.message || "Błąd serwera" }) };
  }
};

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Content-Type": "application/json"
  };
}
