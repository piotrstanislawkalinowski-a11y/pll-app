const https = require('https');

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

  const { action } = body;

  if (action === "send_email") {
    return await sendEmail(body);
  } else {
    return await claudeFormat(body);
  }
};

// ── CLAUDE FORMAT / CORRECT ──
async function claudeFormat(body) {
  const { text, imageBase64, type, mode, apiKey } = body;
  if (!apiKey) return { statusCode: 400, headers: corsHeaders(), body: JSON.stringify({ error: "Brak klucza API" }) };

  const audience = mode === "zarzad" ? "Zarządu" : "Rady Nadzorczej";
  const audienceFull = mode === "zarzad" ? "Zarządu Port Lotniczy Lublin S.A." : "Rady Nadzorczej Port Lotniczy Lublin S.A.";

  let systemPrompt, messages;

  if (type === "correct") {
    systemPrompt = `Jesteś ekspertem od redakcji profesjonalnych materiałów korporacyjnych dla ${audienceFull}.
Otrzymasz aktualny HTML dokumentu oraz zdjęcie wydruku z odręcznymi poprawkami.
Przeanalizuj zdjęcie, odczytaj wszystkie adnotacje i poprawki napisane długopisem, następnie wprowadź je do dokumentu HTML.
Zwróć WYŁĄCZNIE poprawiony HTML — bez markdown, bez backtick, bez komentarzy. Zacznij bezpośrednio od pierwszego tagu HTML.`;

    messages = [{
      role: "user",
      content: [
        { type: "image", source: { type: "base64", media_type: "image/jpeg", data: imageBase64 } },
        { type: "text", text: `Oto aktualny HTML dokumentu. Wprowadź poprawki widoczne na zdjęciu:\n\n${text}` }
      ]
    }];
  } else {
    systemPrompt = `Jesteś ekspertem od przygotowywania profesjonalnych materiałów korporacyjnych dla ${audienceFull}.
Zwróć WYŁĄCZNIE czysty HTML — bez markdown, bez backtick, bez \`\`\`html, bez żadnego tekstu przed ani po HTML.
Zacznij BEZPOŚREDNIO od <h1>.
Używaj tagów: h1, h2, h3, p, ul, ol, li, table, thead, tbody, tr, th, td, blockquote, strong, em.
Struktura: 1) h1=tytuł 2) blockquote=executive summary 3) h2/h3=sekcje 4) tabele dla danych 5) blockquote dla wniosków.
Nie używaj CSS, atrybutów style, DOCTYPE, html, head, body.`;

    messages = [{ role: "user", content: `Sformatuj jako materiał dla ${audience}:\n\n${text}` }];
  }

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01"
      },
      body: JSON.stringify({ model: "claude-sonnet-4-5", max_tokens: 4096, system: systemPrompt, messages })
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      return { statusCode: response.status, headers: corsHeaders(), body: JSON.stringify({ error: err?.error?.message || response.statusText }) };
    }

    const data = await response.json();
    let content = data?.content?.[0]?.text || "";
    content = content.replace(/^```html\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/i, "").trim();
    if (!content) return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: "Brak odpowiedzi z Claude" }) };

    return { statusCode: 200, headers: corsHeaders(), body: JSON.stringify({ result: content }) };
  } catch (err) {
    return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: err.message || "Błąd serwera" }) };
  }
}

// ── SEND EMAIL VIA SENDGRID ──
async function sendEmail(body) {
  const { to, subject, pdfBase64, docxBase64, pdfName, docxName } = body;
  const sgKey = process.env.SENDGRID_API_KEY;

  if (!sgKey) return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: "Brak klucza SendGrid na serwerze" }) };
  if (!to) return { statusCode: 400, headers: corsHeaders(), body: JSON.stringify({ error: "Brak adresu e-mail" }) };

  const payload = {
    personalizations: [{ to: [{ email: to }] }],
    from: { email: "piotr.stanislaw.kalinowski@gmail.com", name: "Port Lotniczy Lublin" },
    subject: subject || "Materiał korporacyjny",
    content: [{ type: "text/plain", value: `W załączeniu przesyłam materiał korporacyjny: ${subject}\n\nDokument wygenerowany przez system Port Lotniczy Lublin.` }],
    attachments: []
  };

  if (pdfBase64) {
    payload.attachments.push({
      content: pdfBase64,
      filename: pdfName || "dokument.pdf",
      type: "application/pdf",
      disposition: "attachment"
    });
  }

  if (docxBase64) {
    payload.attachments.push({
      content: docxBase64,
      filename: docxName || "dokument.docx",
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      disposition: "attachment"
    });
  }

  try {
    const sgResp = await fetch("https://api.sendgrid.com/v3/mail/send", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${sgKey}`
      },
      body: JSON.stringify(payload)
    });

    if (sgResp.status === 202) {
      return { statusCode: 200, headers: corsHeaders(), body: JSON.stringify({ ok: true }) };
    }

    const errData = await sgResp.json().catch(() => ({}));
    const errMsg = errData?.errors?.[0]?.message || `HTTP ${sgResp.status}`;
    return { statusCode: sgResp.status, headers: corsHeaders(), body: JSON.stringify({ error: errMsg }) };
  } catch (err) {
    return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: err.message }) };
  }
}

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Content-Type": "application/json"
  };
}
