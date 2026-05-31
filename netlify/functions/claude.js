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

  const systemPrompt = `Jesteś ekspertem od przygotowywania profesjonalnych materiałów korporacyjnych dla ${audienceFull}.

Zwróć WYŁĄCZNIE czysty HTML — bez żadnych znaczników markdown, bez backtick, bez \`\`\`html, bez żadnego tekstu przed ani po HTML.
Zacznij BEZPOŚREDNIO od <h1> i zakończ ostatnim tagiem HTML.

Używaj tagów: h1, h2, h3, p, ul, ol, li, table, thead, tbody, tr, th, td, blockquote, strong, em.

Struktura:
1. <h1> — konkretny tytuł dokumentu
2. <blockquote> — Executive Summary (3-5 zdań)
3. <h2>Spis treści</h2> jeśli dokument jest długi
4. Sekcje jako <h2>, podsekcje jako <h3>
5. Dane liczbowe w <table>
6. Wnioski w <blockquote>

Nie używaj CSS, atrybutów style, DOCTYPE, html, head, body.`;

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
    let content = data?.content?.[0]?.text || "";

    // Usuń markdown code fences jeśli Claude je dodał
    content = content.replace(/^```html\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/i, "").trim();

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
