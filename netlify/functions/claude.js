const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, PageNumber,
  Footer, Header, LineRuleType, convertInchesToTwip } = require('docx');

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

async function claudeFormat(body) {
  const { text, imageBase64, type, mode, apiKey } = body;
  if (!apiKey) return { statusCode: 400, headers: corsHeaders(), body: JSON.stringify({ error: "Brak klucza API" }) };

  const audience = mode === "zarzad" ? "Zarządu" : "Rady Nadzorczej";
  const audienceFull = mode === "zarzad" ? "Zarządu Port Lotniczy Lublin S.A." : "Rady Nadzorczej Port Lotniczy Lublin S.A.";

  let systemPrompt, messages;
  if (type === "correct") {
    systemPrompt = `Jesteś ekspertem od redakcji profesjonalnych materiałów korporacyjnych dla ${audienceFull}. Otrzymasz aktualny HTML dokumentu oraz zdjęcie wydruku z odręcznymi poprawkami. Przeanalizuj zdjęcie, odczytaj wszystkie adnotacje i poprawki napisane długopisem, następnie wprowadź je do dokumentu HTML. Zwróć WYŁĄCZNIE poprawiony HTML bez markdown, bez backtick. Zacznij od pierwszego tagu HTML.`;
    messages = [{ role: "user", content: [
      { type: "image", source: { type: "base64", media_type: "image/jpeg", data: imageBase64 } },
      { type: "text", text: `HTML dokumentu. Wprowadź poprawki ze zdjęcia:\n\n${text}` }
    ]}];
  } else {
    systemPrompt = `Jesteś ekspertem od przygotowywania profesjonalnych materiałów korporacyjnych dla ${audienceFull}.
Zwróć WYŁĄCZNIE czysty HTML bez markdown, bez backtick, bez \`\`\`html.
Zacznij BEZPOŚREDNIO od <h1>.
Używaj TYLKO: h1,h2,h3,p,ul,ol,li,table,thead,tbody,tr,th,td,blockquote,strong,em,hr.
NIE używaj CSS, style, class, DOCTYPE, html, head, body, div, span.
WAŻNE: Przetwórz CAŁY tekst do końca, nie ucinaj. Każdy element z oryginału musi być w HTML.`;
    messages = [{ role: "user", content: `Sformatuj CAŁY tekst jako materiał dla ${audience}. Nie pomijaj żadnego fragmentu:\n\n${text}` }];
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

async function sendEmail(body) {
  const { to, subject, htmlContent, title, meta } = body;
  const resendKey = process.env.RESEND_API_KEY;
  if (!resendKey) return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: "Brak klucza Resend" }) };
  if (!to) return { statusCode: 400, headers: corsHeaders(), body: JSON.stringify({ error: "Brak adresu e-mail" }) };

  try {
    const docxBuffer = await generateDOCX(title || 'Dokument', meta || '', htmlContent || '');
    const docxBase64 = docxBuffer.toString('base64');
    const safeName = (title || 'dokument').replace(/[^a-zA-Z0-9_\- ]/g, '').substring(0, 50);

    const payload = {
      from: "Port Lotniczy Lublin <onboarding@resend.dev>",
      to: [to],
      subject: subject || "Materiał korporacyjny",
      text: `W załączeniu przesyłam materiał korporacyjny: ${subject}`,
      attachments: [{ filename: safeName + '.docx', content: docxBase64 }]
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

async function generateDOCX(title, meta, htmlContent) {
  const twip = convertInchesToTwip;
  const AUTO = LineRuleType.AUTO;
  const NAVY = '0A1628', GOLD = 'C9A84C', NAVY_LIGHT = '1E3160';
  const GOLD_PALE = 'F5E9C8', TEXT = '1A2540', MUTED = '5A6A8A', CREAM = 'FAF8F3';

  // Parse HTML bez jsdom - używamy prostego parsera
  function stripTags(html) {
    return html.replace(/<[^>]+>/g, '').replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&nbsp;/g,' ').replace(/&#\d+;/g,'').trim();
  }

  function isNum(text) {
    return /^[\d\s\-–.,+%złzłzl\/():★☆]+$/.test(text.trim()) && /\d/.test(text);
  }

  const children = [];

  // Tytuł
  children.push(new Paragraph({
    children: [new TextRun({ text: title, bold: true, size: 32, color: NAVY, font: 'Calibri' })],
    spacing: { before: 0, after: 200 },
    border: { bottom: { color: GOLD, size: 16, style: BorderStyle.SINGLE, space: 6 } },
  }));
  children.push(new Paragraph({ text: '', spacing: { after: 100 } }));

  // Parsuj HTML jako tekst z tagami
  const segments = htmlContent.split(/(<\/?(?:h[1-6]|p|ul|ol|li|table|tr|th|td|thead|tbody|blockquote|strong|em|hr|br)[^>]*>)/gi).filter(s => s);

  let inList = false, listOrdered = false, listIdx = 0;
  let inTable = false, tableRows = [], currentRow = [], currentCells = [];
  let inBlockquote = false, blockquoteText = '';
  let skipUntil = null;

  for (let i = 0; i < segments.length; i++) {
    const seg = segments[i];
    if (!seg.trim() && !seg.includes('<')) continue;

    const tagMatch = seg.match(/^<(\/?)(h[1-6]|p|ul|ol|li|table|tr|th|td|thead|tbody|blockquote|strong|em|hr|br)([^>]*)>$/i);

    if (!tagMatch) {
      // Tekst
      const text = seg.replace(/<[^>]+>/g,'').replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&nbsp;/g,' ').replace(/&#\d+;/g,'').trim();
      if (!text) continue;

      if (inTable) {
        currentCells.push(text);
      } else if (inBlockquote) {
        blockquoteText += ' ' + text;
      } else if (inList) {
        // tekst wewnątrz li - handled by li tag
      }
      continue;
    }

    const [, closing, tag, attrs] = tagMatch;
    const tagLower = tag.toLowerCase();
    const isClose = closing === '/';

    if (tagLower === 'table') {
      if (!isClose) { inTable = true; tableRows = []; currentRow = []; currentCells = []; }
      else {
        inTable = false;
        if (tableRows.length > 0) {
          const numericCols = new Set();
          tableRows.slice(1).forEach(row => {
            row.forEach((cell, idx) => { if (isNum(cell.text)) numericCols.add(idx); });
          });
          const docRows = tableRows.map((row, rIdx) => new TableRow({
            tableHeader: row[0]?.isHeader,
            children: row.map((cell, cIdx) => new TableCell({
              children: [new Paragraph({
                children: [new TextRun({ text: cell.text, bold: cell.isHeader, color: cell.isHeader ? GOLD : TEXT, size: 20, font: 'Calibri' })],
                alignment: cell.isHeader ? AlignmentType.CENTER : (numericCols.has(cIdx) ? AlignmentType.RIGHT : AlignmentType.LEFT),
                spacing: { before: 60, after: 60 },
              })],
              shading: cell.isHeader ? { fill: NAVY, type: ShadingType.CLEAR } : rIdx % 2 === 0 ? { fill: CREAM, type: ShadingType.CLEAR } : { fill: 'FFFFFF', type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              borders: { top:{style:BorderStyle.SINGLE,size:4,color:'D8D0BC'}, bottom:{style:BorderStyle.SINGLE,size:4,color:'D8D0BC'}, left:{style:BorderStyle.SINGLE,size:4,color:'D8D0BC'}, right:{style:BorderStyle.SINGLE,size:4,color:'D8D0BC'} },
            })),
          }));
          children.push(new Table({ rows: docRows, width: { size: 100, type: WidthType.PERCENTAGE } }));
          children.push(new Paragraph({ text: '', spacing: { after: 120 } }));
        }
      }
    } else if (tagLower === 'tr') {
      if (!isClose) { currentRow = []; }
      else { if (currentRow.length > 0) tableRows.push(currentRow); currentRow = []; }
    } else if (tagLower === 'th' || tagLower === 'td') {
      if (isClose) {
        const cellText = currentCells.join(' ').trim();
        currentRow.push({ text: cellText, isHeader: tagLower === 'th' });
        currentCells = [];
      }
    } else if (tagLower === 'blockquote') {
      if (!isClose) { inBlockquote = true; blockquoteText = ''; }
      else {
        inBlockquote = false;
        const t = blockquoteText.trim();
        if (t) children.push(new Paragraph({
          children: [new TextRun({ text: t, italics: true, size: 21, color: NAVY, font: 'Calibri' })],
          indent: { left: twip(0.3), right: twip(0.3) },
          shading: { fill: GOLD_PALE, type: ShadingType.CLEAR },
          border: { left: { color: GOLD, size: 24, style: BorderStyle.SINGLE, space: 8 } },
          spacing: { before: 140, after: 140, line: 276, lineRule: AUTO },
        }));
      }
    } else if (tagLower === 'ul' || tagLower === 'ol') {
      if (!isClose) { inList = true; listOrdered = tagLower === 'ol'; listIdx = 0; }
      else { inList = false; children.push(new Paragraph({ text: '', spacing: { after: 80 } })); }
    } else if (tagLower === 'li') {
      if (!isClose) { /* start li */ }
      else {
        // Zbierz tekst li z segmentów
        const liText = currentCells.join(' ').trim();
        currentCells = [];
        if (liText) {
          children.push(new Paragraph({
            children: [new TextRun({ text: (listOrdered ? (++listIdx)+'. ' : '• ') + liText, size: 22, color: TEXT, font: 'Calibri' })],
            indent: { left: twip(0.35) },
            spacing: { before: 40, after: 60 },
          }));
        }
      }
    } else if (tagLower.match(/^h[1-6]$/)) {
      if (isClose) {
        const t = currentCells.join(' ').trim();
        currentCells = [];
        const level = parseInt(tagLower[1]);
        if (t) {
          if (level === 1) {
            // pomijamy h1 — tytuł już dodany
          } else if (level === 2) {
            children.push(new Paragraph({ children: [new TextRun({ text: t, bold: true, size: 26, color: NAVY_LIGHT, font: 'Calibri' })], spacing: { before: 240, after: 120 } }));
          } else if (level === 3) {
            children.push(new Paragraph({ children: [new TextRun({ text: t, bold: true, size: 23, color: NAVY, font: 'Calibri' })], spacing: { before: 180, after: 80 } }));
          } else {
            children.push(new Paragraph({ children: [new TextRun({ text: t, bold: true, size: 22, color: NAVY, font: 'Calibri' })], spacing: { before: 140, after: 60 } }));
          }
        }
      } else { currentCells = []; }
    } else if (tagLower === 'p') {
      if (isClose) {
        const t = currentCells.join(' ').trim();
        currentCells = [];
        if (t) children.push(new Paragraph({
          children: [new TextRun({ text: t, size: 22, color: TEXT, font: 'Calibri' })],
          alignment: AlignmentType.JUSTIFIED,
          spacing: { before: 60, after: 100, line: 276, lineRule: AUTO },
        }));
      } else { currentCells = []; }
    } else if (tagLower === 'hr') {
      children.push(new Paragraph({
        children: [new TextRun({ text: '' })],
        border: { bottom: { color: 'D8D0BC', size: 6, style: BorderStyle.SINGLE, space: 4 } },
        spacing: { before: 120, after: 120 },
      }));
    }

    // Zbierz tekst wewnątrz tagów
    if (!isClose && (tagLower === 'h1' || tagLower === 'h2' || tagLower === 'h3' || tagLower === 'h4' || tagLower === 'h5' || tagLower === 'h6' || tagLower === 'p' || tagLower === 'li' || tagLower === 'th' || tagLower === 'td')) {
      currentCells = [];
      // Zbierz kolejne segmenty aż do zamknięcia tagu
      let depth = 1;
      let collected = [];
      let j = i + 1;
      while (j < segments.length && depth > 0) {
        const s = segments[j];
        const m = s.match(/^<(\/?)(h[1-6]|p|li|th|td)([^>]*)>$/i);
        if (m) {
          if (m[2].toLowerCase() === tagLower) {
            if (m[1] === '/') depth--;
            else depth++;
          }
          if (depth === 0) break;
        }
        if (!s.match(/^<[^>]+>$/)) {
          collected.push(s.replace(/<[^>]+>/g,'').replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&nbsp;/g,' ').replace(/&#\d+;/g,''));
        }
        j++;
      }
      currentCells = collected;
      i = j; // skip to closing tag
    }
  }

  const metaClean = meta.replace(/MATERIAA? DLA /i,'').replace(/MATERIAŁ DLA /i,'').replace(/ · PORT LOTNICZY LUBLIN S\.A\./i,'').trim();

  const header = new Header({ children: [new Paragraph({
    children: [
      new TextRun({ text: 'PORT LOTNICZY LUBLIN S.A.  ·  ', size: 16, color: GOLD, bold: true, font: 'Calibri' }),
      new TextRun({ text: metaClean, size: 16, color: MUTED, font: 'Calibri' }),
    ],
    alignment: AlignmentType.RIGHT,
    border: { bottom: { color: GOLD, size: 6, style: BorderStyle.SINGLE, space: 4 } },
    spacing: { after: 0 },
  })] });

  const footer = new Footer({ children: [new Paragraph({
    children: [
      new TextRun({ text: 'Port Lotniczy Lublin S.A.          Strona ', size: 16, color: MUTED, font: 'Calibri' }),
      new TextRun({ children: [PageNumber.CURRENT], size: 16, color: MUTED, font: 'Calibri' }),
      new TextRun({ text: ' / ', size: 16, color: MUTED, font: 'Calibri' }),
      new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: MUTED, font: 'Calibri' }),
    ],
    alignment: AlignmentType.CENTER,
    border: { top: { color: 'D8D0BC', size: 4, style: BorderStyle.SINGLE, space: 4 } },
  })] });

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
