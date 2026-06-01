const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, PageNumber,
  Footer, Header, LineRuleType, convertInchesToTwip } = require('docx');
const { JSDOM } = require('jsdom');

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
  if (action === "send_email") return await sendEmail(body);
  return await claudeFormat(body);
};

// ── CLAUDE FORMAT ──
async function claudeFormat(body) {
  const { text, imageBase64, type, mode, apiKey } = body;
  if (!apiKey) return { statusCode: 400, headers: corsHeaders(), body: JSON.stringify({ error: "Brak klucza API" }) };

  const audience = mode === "zarzad" ? "Zarządu" : "Rady Nadzorczej";
  const audienceFull = mode === "zarzad" ? "Zarządu Port Lotniczy Lublin S.A." : "Rady Nadzorczej Port Lotniczy Lublin S.A.";

  let systemPrompt, messages;
  if (type === "correct") {
    systemPrompt = `Jesteś ekspertem od redakcji profesjonalnych materiałów korporacyjnych dla ${audienceFull}. Otrzymasz aktualny HTML dokumentu oraz zdjęcie wydruku z odręcznymi poprawkami. Przeanalizuj zdjęcie, odczytaj wszystkie adnotacje i poprawki napisane długopisem, następnie wprowadź je do dokumentu HTML. Zwróć WYŁĄCZNIE poprawiony HTML — bez markdown, bez backtick, bez komentarzy. Zacznij bezpośrednio od pierwszego tagu HTML.`;
    messages = [{ role: "user", content: [
      { type: "image", source: { type: "base64", media_type: "image/jpeg", data: imageBase64 } },
      { type: "text", text: `Oto aktualny HTML dokumentu. Wprowadź poprawki widoczne na zdjęciu:\n\n${text}` }
    ]}];
  } else {
    systemPrompt = `Jesteś ekspertem od przygotowywania profesjonalnych materiałów korporacyjnych dla ${audienceFull}. Zwróć WYŁĄCZNIE czysty HTML — bez markdown, bez backtick, bez \`\`\`html. Zacznij BEZPOŚREDNIO od <h1>. Używaj: h1,h2,h3,p,ul,ol,li,table,thead,tbody,tr,th,td,blockquote,strong,em. Nie używaj CSS, style, DOCTYPE, html, head, body.`;
    messages = [{ role: "user", content: `Sformatuj jako materiał dla ${audience}:\n\n${text}` }];
  }

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
      body: JSON.stringify({ model: "claude-haiku-4-5", max_tokens: 4096, system: systemPrompt, messages })
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
    return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: err.message }) };
  }
}

// ── SEND EMAIL ──
async function sendEmail(body) {
  const { to, subject, htmlContent, title, meta } = body;
  const resendKey = process.env.RESEND_API_KEY;
  if (!resendKey) return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: "Brak klucza Resend" }) };
  if (!to) return { statusCode: 400, headers: corsHeaders(), body: JSON.stringify({ error: "Brak adresu e-mail" }) };

  try {
    // Generuj DOCX po stronie serwera
    const docxBuffer = await generateDOCX(title || 'Dokument', meta || '', htmlContent || '');
    const docxBase64 = docxBuffer.toString('base64');
    const safeName = (title || 'dokument').replace(/[^a-zA-Z0-9_\- ]/g, '').substring(0, 50);

    const payload = {
      from: "Port Lotniczy Lublin <onboarding@resend.dev>",
      to: [to],
      subject: subject || "Materiał korporacyjny",
      text: `W załączeniu przesyłam materiał korporacyjny: ${subject}`,
      attachments: [{
        filename: safeName + '.docx',
        content: docxBase64
      }]
    };

    const resp = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${resendKey}` },
      body: JSON.stringify(payload)
    });
    const data = await resp.json();
    if (resp.ok && data.id) return { statusCode: 200, headers: corsHeaders(), body: JSON.stringify({ ok: true }) };
    return { statusCode: resp.status, headers: corsHeaders(), body: JSON.stringify({ error: data?.message || "Błąd wysyłki" }) };
  } catch (err) {
    return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: err.message }) };
  }
}

// ── GENERATE DOCX ──
async function generateDOCX(title, meta, htmlContent) {
  const twip = (inch) => convertInchesToTwip(inch);
  const AUTO = LineRuleType.AUTO;
  const NAVY = '0A1628', GOLD = 'C9A84C', NAVY_LIGHT = '1E3160';
  const GOLD_PALE = 'F5E9C8', TEXT = '1A2540', MUTED = '5A6A8A', CREAM = 'FAF8F3';

  const dom = new JSDOM('<div>' + htmlContent + '</div>');
  const doc = dom.window.document;
  const children = [];

  // Tytuł
  children.push(new Paragraph({
    children: [new TextRun({ text: title, bold: true, size: 32, color: NAVY, font: 'Calibri' })],
    spacing: { before: 0, after: 200 },
    border: { bottom: { color: GOLD, size: 16, style: BorderStyle.SINGLE, space: 6 } },
  }));
  children.push(new Paragraph({ text: '', spacing: { after: 100 } }));

  function isNum(text) {
    return /^[\d\s\-–.,+%złzłzl\/():]+$/.test(text.trim()) && /\d/.test(text);
  }

  function makeTable(node) {
    const rows = node.querySelectorAll('tr');
    if (!rows.length) return null;
    const dataRows = Array.from(rows).filter(r => r.querySelector('td'));
    const numericCols = new Set();
    dataRows.forEach(row => {
      Array.from(row.querySelectorAll('td')).forEach((cell, idx) => {
        if (isNum(cell.textContent)) numericCols.add(idx);
      });
    });
    const tableRows = [];
    rows.forEach((row, rIdx) => {
      const cells = row.querySelectorAll('th, td');
      const isHeaderRow = row.querySelector('th') !== null;
      tableRows.push(new TableRow({
        tableHeader: isHeaderRow,
        children: Array.from(cells).map((cell, cIdx) => {
          const isH = cell.tagName.toLowerCase() === 'th';
          const cellText = cell.textContent.trim();
          const isNumCol = !isH && numericCols.has(cIdx);
          return new TableCell({
            children: [new Paragraph({
              children: [new TextRun({ text: cellText, bold: isH, color: isH ? GOLD : TEXT, size: 20, font: 'Calibri' })],
              alignment: isH ? AlignmentType.CENTER : (isNumCol ? AlignmentType.RIGHT : AlignmentType.LEFT),
              spacing: { before: 60, after: 60 },
            })],
            shading: isH ? { fill: NAVY, type: ShadingType.CLEAR }
              : rIdx % 2 === 0 ? { fill: CREAM, type: ShadingType.CLEAR }
              : { fill: 'FFFFFF', type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 4, color: 'D8D0BC' },
              bottom: { style: BorderStyle.SINGLE, size: 4, color: 'D8D0BC' },
              left: { style: BorderStyle.SINGLE, size: 4, color: 'D8D0BC' },
              right: { style: BorderStyle.SINGLE, size: 4, color: 'D8D0BC' },
            },
          });
        }),
      }));
    });
    return new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } });
  }

  function parseNode(node) {
    const tag = node.tagName ? node.tagName.toLowerCase() : '';
    const text = (node.textContent || '').replace(/\s+/g, ' ').trim();
    if (tag === 'h1') return; // pomijamy - tytuł już dodany
    if (tag === 'h2') {
      if (text) children.push(new Paragraph({ children: [new TextRun({ text, bold: true, size: 24, color: NAVY_LIGHT, font: 'Calibri' })], spacing: { before: 220, after: 100 } }));
    } else if (tag === 'h3') {
      if (text) children.push(new Paragraph({ children: [new TextRun({ text, bold: true, size: 22, color: NAVY, font: 'Calibri' })], spacing: { before: 160, after: 80 } }));
    } else if (tag === 'p') {
      if (text) children.push(new Paragraph({ children: [new TextRun({ text, size: 22, color: TEXT, font: 'Calibri' })], alignment: AlignmentType.JUSTIFIED, spacing: { before: 60, after: 100, line: 276, lineRule: AUTO } }));
    } else if (tag === 'blockquote') {
      if (text) children.push(new Paragraph({ children: [new TextRun({ text, italics: true, size: 21, color: NAVY, font: 'Calibri' })], indent: { left: twip(0.3), right: twip(0.3) }, shading: { fill: GOLD_PALE, type: ShadingType.CLEAR }, border: { left: { color: GOLD, size: 24, style: BorderStyle.SINGLE, space: 8 } }, spacing: { before: 140, after: 140, line: 276, lineRule: AUTO } }));
    } else if (tag === 'ul' || tag === 'ol') {
      node.querySelectorAll('li').forEach((li, idx) => {
        const liText = li.textContent.replace(/\s+/g, ' ').trim();
        if (liText) children.push(new Paragraph({ children: [new TextRun({ text: (tag === 'ol' ? (idx+1)+'. ' : '• ') + liText, size: 22, color: TEXT, font: 'Calibri' })], indent: { left: twip(0.35) }, spacing: { before: 40, after: 60 } }));
      });
      children.push(new Paragraph({ text: '', spacing: { after: 80 } }));
    } else if (tag === 'table') {
      const tbl = makeTable(node);
      if (tbl) { children.push(tbl); children.push(new Paragraph({ text: '', spacing: { after: 120 } })); }
    } else if (tag === 'strong' || tag === 'em' || tag === 'b' || tag === 'i') {
      // Ignoruj — tekst obsługiwany przez rodzica
    } else {
      Array.from(node.childNodes).forEach(child => parseNode(child));
    }
  }

  Array.from(doc.body.firstChild.childNodes).forEach(el => parseNode(el));

  const header = new Header({
    children: [new Paragraph({
      children: [
        new TextRun({ text: 'PORT LOTNICZY LUBLIN S.A.  ·  ', size: 16, color: GOLD, bold: true, font: 'Calibri' }),
        new TextRun({ text: meta.replace('MATERIAŁ DLA ', '').replace('MATERIAA DLA ', '').replace(' · PORT LOTNICZY LUBLIN S.A.', '').replace(' · PORT LOTNICZY LUBLIN S.A.', ''), size: 16, color: MUTED, font: 'Calibri' }),
      ],
      alignment: AlignmentType.RIGHT,
      border: { bottom: { color: GOLD, size: 6, style: BorderStyle.SINGLE, space: 4 } },
      spacing: { after: 0 },
    })],
  });

  const footer = new Footer({
    children: [new Paragraph({
      children: [
        new TextRun({ text: 'Port Lotniczy Lublin S.A.          Strona ', size: 16, color: MUTED, font: 'Calibri' }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, color: MUTED, font: 'Calibri' }),
        new TextRun({ text: ' / ', size: 16, color: MUTED, font: 'Calibri' }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: MUTED, font: 'Calibri' }),
      ],
      alignment: AlignmentType.CENTER,
      border: { top: { color: 'D8D0BC', size: 4, style: BorderStyle.SINGLE, space: 4 } },
    })],
  });

  const document = new Document({
    creator: '', description: '', title: '', subject: '', keywords: '', lastModifiedBy: '', revision: 1,
    sections: [{
      properties: { page: { margin: { top: twip(1.0), right: twip(1.1), bottom: twip(1.0), left: twip(1.1), header: twip(0.4), footer: twip(0.4) } } },
      headers: { default: header },
      footers: { default: footer },
      children,
    }],
  });

  return await Packer.toBuffer(document);
}

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Content-Type": "application/json"
  };
}
