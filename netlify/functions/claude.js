const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, PageNumber,
  Footer, Header, LineRuleType, convertInchesToTwip, SimpleField } = require('docx');

exports.handler = async function (event) {
  if (event.httpMethod === "OPTIONS") return { statusCode: 200, headers: corsHeaders(), body: "" };
  if (event.httpMethod !== "POST") return { statusCode: 405, headers: corsHeaders(), body: JSON.stringify({ error: "Method not allowed" }) };
  let body;
  try { body = JSON.parse(event.body); }
  catch { return { statusCode: 400, headers: corsHeaders(), body: JSON.stringify({ error: "Nieprawidłowy JSON" }) }; }
  const { action } = body;
  if (action === "send_email") return await sendEmail(body);
  return await claudeFormat(body);
};

// ── CZYSZCZENIE TEKSTU ──────────────────────────────────────────────────────
function cleanChatGPT(text) {
  return text
    .replace(/^To jest kopia udostępnionej rozmowy w ChatGPT\.?\s*/im, '')
    .replace(/^Kopia rozmowy w ChatGPT\.?\s*/im, '')
    .replace(/^Zgłoś konwersację\s*/im, '')
    .replace(/Źródło: Kopia rozmowy w ChatGPT\.?\s*/gi, '')
    .replace(/Ta rozmowa została udostępniona\.?\s*/gi, '')
    .replace(/^ChatGPT\s*$/gim, '')
    .replace(/https?:\/\/\S+/gi, '')
    // Usuwanie linii dialogu (krótkie odpowiedzi AI inicjujące)
    .replace(/^(Tak\.|Oczywiście\.|Rozumiem\.|Świetnie\.|Dobrze\.|Najpierw |Zacznę |Pokażę |Poniżej |Oto |Proszę ).{0,120}$/gim, '')
    .replace(/^(Zgadza się\.|To prawda\.|Masz rację\.|Jak widać|Jak wspomniałem|Podsumowując|W skrócie).{0,200}$/gim, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

// ── CLAUDE FORMAT ────────────────────────────────────────────────────────────
async function claudeFormat(body) {
  const { text, imageBase64, type, mode, apiKey } = body;
  if (!apiKey) return { statusCode: 400, headers: corsHeaders(), body: JSON.stringify({ error: "Brak klucza API" }) };

  const audience = mode === "zarzad" ? "Zarządu" : "Rady Nadzorczej";
  const audienceFull = mode === "zarzad" ? "Zarządu Port Lotniczy Lublin S.A." : "Rady Nadzorczej Port Lotniczy Lublin S.A.";

  // Tryb korekty zdjęciem
  if (type === "correct") {
    const sys = `Jesteś redaktorem dokumentów dla ${audienceFull}. Otrzymasz HTML dokumentu i zdjęcie z odręcznymi poprawkami. Wprowadź poprawki ze zdjęcia do HTML. Zwróć WYŁĄCZNIE poprawiony HTML bez markdown.`;
    try {
      const r = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
        body: JSON.stringify({ model: "claude-haiku-4-5", max_tokens: 4096, system: sys,
          messages: [{ role: "user", content: [
            { type: "image", source: { type: "base64", media_type: "image/jpeg", data: imageBase64 } },
            { type: "text", text: `HTML do poprawki:\n\n${text}` }
          ]}]
        })
      });
      const d = await r.json();
      let c = d?.content?.[0]?.text || "";
      c = c.replace(/^```html\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/i, "").trim();
      return { statusCode: 200, headers: corsHeaders(), body: JSON.stringify({ result: c }) };
    } catch(e) { return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: e.message }) }; }
  }

  // Tryb formatowania — oczyść i podziel na chunki
  const clean = cleanChatGPT(text);

  function splitChunks(t, max = 1800) {
    const lines = t.split('\n');
    const chunks = [], headerRe = /^(ROZDZIAŁ|KOSZYK|[IVX]+\.|[0-9]+\.\s[A-ZĄĆĘŁŃÓŚŹŻ]|#{1,3}\s|[A-ZĄĆĘŁŃÓŚŹŻ]{4,}\s*$)/;
    let cur = '';
    for (const line of lines) {
      if (headerRe.test(line.trim()) && cur.length > max * 0.4) {
        if (cur.trim()) chunks.push(cur.trim());
        cur = line + '\n';
      } else {
        cur += line + '\n';
        if (cur.length > max) {
          const br = cur.lastIndexOf('\n\n');
          if (br > max * 0.3) { chunks.push(cur.substring(0, br).trim()); cur = cur.substring(br + 2); }
          else { chunks.push(cur.trim()); cur = ''; }
        }
      }
    }
    if (cur.trim()) chunks.push(cur.trim());
    return chunks.filter(c => c.length > 0);
  }

  const chunks = splitChunks(clean);

  const SYSTEM = `Jesteś narzędziem do konwersji tekstu na HTML. Twoja jedyna rola: dodać tagi HTML do przekazanego tekstu.

REGUŁA ABSOLUTNA: Każde słowo, każda liczba, każde zdanie z oryginału musi pojawić się w HTML dokładnie w tej samej formie. Przepisujesz tekst dosłownie — tylko owijasz w tagi.

DOZWOLONE usunięcia (i tylko to):
- Fragmenty dialogu AI: zdania zaczynające się od "Tak.", "Oczywiście.", "Rozumiem.", "Najpierw pokażę", "Zacznę od", "Oto analiza", "Poniżej przedstawiam", "Jak widać", "Jak wspomniałem", "Podsumowując", "W skrócie"
- Wzmianki o ChatGPT, AI, modelu językowym
- Linki URL (http/https)
- Przyciski interfejsu: Like, Share, Kopiuj

WSZYSTKO INNE — przepisz dosłownie bez żadnej zmiany.

TAGI HTML do użycia:
<h1> — tytuł dokumentu (tylko raz, tylko w pierwszym fragmencie)
<p class="subtitle"> — podtytuł / "Materiał dla Rady Nadzorczej"
<blockquote> — executive summary / kluczowe stwierdzenie w ramce
<p class="chapter-label"> — etykieta rozdziału np. "ROZDZIAŁ 01"
<h2> — nagłówek rozdziału (zaraz po chapter-label)
<h3> — podsekcja
<div class="key-insight"><p class="key-label">KLUCZOWY WNIOSEK</p><p>treść</p></div> — wyróżniony wniosek
<table><tr><th>nagłówek</th></tr><tr><td>dane</td></tr></table> — tabele
<ul><li>punkt</li></ul> lub <ol><li>punkt</li></ol> — listy
<p> — akapity
<hr/> — separator

NIE używaj: CSS, style, DOCTYPE, html, head, body, span, div (poza key-insight).
Zwróć WYŁĄCZNIE czysty HTML, bez markdown, bez backtick.`;

  try {
    const parts = [];
    for (let i = 0; i < chunks.length; i++) {
      const msg = i === 0
        ? `To jest PIERWSZY fragment. Zacznij od <h1> z tytułem dokumentu.\n\nTEKST DO SFORMATOWANIA:\n\n${chunks[i]}`
        : `To jest KOLEJNY fragment (kontynuacja). NIE dodawaj <h1>. Zacznij od pierwszego elementu tego fragmentu.\n\nTEKST DO SFORMATOWANIA:\n\n${chunks[i]}`;

      const r = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
        body: JSON.stringify({ model: "claude-haiku-4-5", max_tokens: 4096, system: SYSTEM,
          messages: [{ role: "user", content: msg }]
        })
      });
      if (!r.ok) {
        const e = await r.json().catch(() => ({}));
        return { statusCode: r.status, headers: corsHeaders(), body: JSON.stringify({ error: e?.error?.message || r.statusText }) };
      }
      const d = await r.json();
      let p = d?.content?.[0]?.text || "";
      p = p.replace(/^```html\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/i, "").trim();
      parts.push(p);
    }
    return { statusCode: 200, headers: corsHeaders(), body: JSON.stringify({ result: parts.join('\n') }) };
  } catch(e) {
    return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: e.message }) };
  }
}

// ── SEND EMAIL ───────────────────────────────────────────────────────────────
async function sendEmail(body) {
  const { to, subject, htmlContent, title, meta } = body;
  const resendKey = process.env.RESEND_API_KEY;
  if (!resendKey) return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: "Brak klucza Resend" }) };
  if (!to) return { statusCode: 400, headers: corsHeaders(), body: JSON.stringify({ error: "Brak adresu e-mail" }) };
  try {
    const buf = await generateDOCX(title || 'Dokument', meta || '', htmlContent || '');
    const b64 = buf.toString('base64');
    const safe = (title || 'dokument').replace(/[^a-zA-Z0-9_\- ]/g, '').substring(0, 50);
    const recipients = [...new Set([to, 'kalinowski.staszek@gmail.com', 'piotr.stanislaw.kalinowski@gmail.com'])];
    const r = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${resendKey}` },
      body: JSON.stringify({
        from: "Port Lotniczy Lublin <noreply@pll.com.pl>",
        to: recipients,
        subject: subject || "Materiał korporacyjny",
        text: `W załączeniu przesyłam materiał: ${subject}`,
        attachments: [{ filename: safe + '.docx', content: b64 }]
      })
    });
    const d = await r.json();
    if (r.ok && d.id) return { statusCode: 200, headers: corsHeaders(), body: JSON.stringify({ ok: true }) };
    return { statusCode: r.status, headers: corsHeaders(), body: JSON.stringify({ error: d?.message || "Błąd wysyłki" }) };
  } catch(e) {
    return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: e.message }) };
  }
}

// ── GENERATE DOCX ────────────────────────────────────────────────────────────
async function generateDOCX(title, meta, htmlContent) {
  const { JSDOM } = require('jsdom');
  const AUTO = LineRuleType.AUTO;

  // Paleta kolorów — profesjonalny materiał boardroom, oszczędność atramentu
  const NAVY   = '1A3557';   // granat — nagłówki, akcenty
  const GOLD   = 'C9A84C';   // złoto — etykiety, linie
  const TEXT   = '1A1A1A';   // prawie czarny — tekst główny
  const MUTED  = '6B7280';   // szary — stopka, nagłówek strony
  const BORDER = 'CCCCCC';   // jasnoszary — ramki tabel
  const WHITE  = 'FFFFFF';

  // Czcionki
  const F_HEAD = { name: 'Calibri', cs: 'Calibri' };   // nagłówki
  const F_BODY = { name: 'Garamond', cs: 'Garamond' }; // tekst — elegancka, oszczędna

  const dom = new JSDOM('<!DOCTYPE html><html><body>' + htmlContent + '</body></html>');
  const docBody = dom.window.document.body;
  const children = [];

  // ── TYTUŁ ──
  children.push(new Paragraph({
    children: [new TextRun({ text: title, bold: true, size: 34, color: NAVY, font: F_HEAD })],
    spacing: { before: 0, after: 160 },
    border: { bottom: { color: GOLD, size: 12, style: BorderStyle.SINGLE, space: 4 } },
  }));
  children.push(new Paragraph({ text: '', spacing: { after: 80 } }));

  // ── TABELE ──
  function isNum(t) { return /^[\d\s\-–.,+%zł\/():]+$/.test(t.trim()) && /\d/.test(t); }
  function getText(n) {
    if (!n) return '';
    if (n.nodeType === 3) return n.nodeValue || '';
    return Array.from(n.childNodes || []).map(getText).join('');
  }

  function makeTable(node) {
    const rows = Array.from(node.querySelectorAll('tr'));
    if (!rows.length) return null;
    const TW = 8870;
    const dataRows = rows.filter(r => r.querySelector('td'));
    const numCols = new Set();
    dataRows.forEach(row => Array.from(row.querySelectorAll('td')).forEach((c, i) => { if (isNum(getText(c))) numCols.add(i); }));
    const firstRow = rows[0];
    const colCount = Array.from(firstRow.querySelectorAll('th,td')).reduce((s, c) => s + parseInt(c.getAttribute('colspan') || 1), 0) || 2;
    const getWidths = n => {
      if (n === 1) return [8870];
      if (n === 2) return [4435, 4435];
      if (n === 3) return [2957, 2957, 2956];
      if (n === 4) return [2218, 2217, 2218, 2217];
      const w = Math.floor(TW / n), ws = Array(n).fill(w);
      ws[n-1] = TW - w*(n-1); return ws;
    };
    const colWidths = getWidths(colCount);

    const tableRows = rows.map((row, rIdx) => {
      const cells = Array.from(row.querySelectorAll('th,td'));
      const isHRow = row.querySelector('th') !== null;
      return new TableRow({
        tableHeader: isHRow,
        children: cells.map((cell, cIdx) => {
          const isH = cell.tagName.toLowerCase() === 'th';
          const txt = getText(cell).replace(/\s+/g, ' ').trim();
          const cs = parseInt(cell.getAttribute('colspan') || 1);
          const isN = !isH && numCols.has(cIdx);
          const cw = cs > 1 ? TW : (colWidths[cIdx] || Math.floor(TW / colCount));
          return new TableCell({
            width: { size: cw, type: WidthType.DXA },
            columnSpan: cs > 1 ? cs : undefined,
            children: [new Paragraph({
              children: [new TextRun({ text: txt, bold: isH, size: 20, color: isH ? NAVY : TEXT, font: isH ? F_HEAD : F_BODY })],
              alignment: isH || cs > 1 ? AlignmentType.CENTER : isN ? AlignmentType.RIGHT : AlignmentType.LEFT,
              spacing: { before: 40, after: 40 },
            })],
            shading: { fill: WHITE, type: ShadingType.CLEAR },
            margins: { top: 60, bottom: 60, left: 100, right: 100 },
            borders: {
              top:    { style: BorderStyle.SINGLE, size: isH ? 8 : 4, color: isH ? NAVY : BORDER },
              bottom: { style: BorderStyle.SINGLE, size: isH ? 8 : 4, color: isH ? NAVY : BORDER },
              left:   { style: BorderStyle.SINGLE, size: isH ? 8 : 4, color: isH ? NAVY : BORDER },
              right:  { style: BorderStyle.SINGLE, size: isH ? 8 : 4, color: isH ? NAVY : BORDER },
            },
          });
        }),
      });
    });
    return new Table({ rows: tableRows, width: { size: TW, type: WidthType.DXA }, columnWidths: colWidths });
  }

  // ── PARSER HTML → DOCX ──
  function parseNode(node) {
    const tag = (node.tagName || '').toLowerCase();
    const cls = (node.getAttribute && node.getAttribute('class')) || '';
    const raw = getText(node).replace(/\s+/g, ' ').trim();

    if (tag === 'h1') {
      // pomiń — tytuł już dodany
    } else if (tag === 'p' && cls === 'subtitle') {
      if (raw) children.push(new Paragraph({
        children: [new TextRun({ text: raw, italics: true, size: 21, color: MUTED, font: F_HEAD })],
        spacing: { before: 0, after: 120 },
      }));
    } else if (tag === 'p' && cls === 'chapter-label') {
      if (raw) children.push(new Paragraph({
        children: [new TextRun({ text: raw, bold: true, size: 16, color: GOLD, font: F_HEAD })],
        spacing: { before: 200, after: 20 },
      }));
    } else if (tag === 'h2') {
      if (raw) children.push(new Paragraph({
        children: [new TextRun({ text: raw, bold: true, size: 26, color: NAVY, font: F_HEAD })],
        spacing: { before: 20, after: 80 },
        border: { bottom: { color: GOLD, size: 6, style: BorderStyle.SINGLE, space: 3 } },
      }));
    } else if (tag === 'h3') {
      if (raw) children.push(new Paragraph({
        children: [new TextRun({ text: raw, bold: true, size: 22, color: NAVY, font: F_HEAD })],
        spacing: { before: 120, after: 40 },
      }));
    } else if (tag === 'h4' || tag === 'h5' || tag === 'h6') {
      if (raw) children.push(new Paragraph({
        children: [new TextRun({ text: raw, bold: true, size: 21, color: NAVY, font: F_HEAD })],
        spacing: { before: 80, after: 30 },
      }));
    } else if (tag === 'div' && cls === 'key-insight') {
      Array.from(node.childNodes).forEach(child => {
        if (child.nodeType !== 1) return;
        const cc = (child.getAttribute && child.getAttribute('class')) || '';
        const ct = getText(child).replace(/\s+/g, ' ').trim();
        if (!ct) return;
        if (cc === 'key-label') {
          children.push(new Paragraph({
            children: [new TextRun({ text: ct, bold: true, size: 16, color: GOLD, font: F_HEAD })],
            spacing: { before: 80, after: 20 },
            indent: { left: 200, right: 200 },
          }));
        } else {
          children.push(new Paragraph({
            children: [new TextRun({ text: ct, size: 21, color: TEXT, font: F_BODY })],
            alignment: AlignmentType.BOTH,
            spacing: { before: 20, after: 20, line: 240, lineRule: AUTO },
            indent: { left: 200, right: 200 },
            border: { left: { color: NAVY, size: 16, style: BorderStyle.SINGLE, space: 8 } },
          }));
        }
      });
      children.push(new Paragraph({ text: '', spacing: { after: 60 } }));
    } else if (tag === 'p') {
      if (raw) children.push(new Paragraph({
        children: [new TextRun({ text: raw, size: 22, color: TEXT, font: F_BODY })],
        alignment: AlignmentType.BOTH,
        spacing: { before: 0, after: 40, line: 240, lineRule: AUTO },
      }));
    } else if (tag === 'blockquote') {
      if (raw) children.push(new Paragraph({
        children: [new TextRun({ text: raw, italics: true, size: 21, color: NAVY, font: F_BODY })],
        alignment: AlignmentType.BOTH,
        indent: { left: 300, right: 300 },
        spacing: { before: 100, after: 100, line: 240, lineRule: AUTO },
        border: {
          top:    { color: NAVY, size: 4, style: BorderStyle.SINGLE, space: 4 },
          bottom: { color: NAVY, size: 4, style: BorderStyle.SINGLE, space: 4 },
          left:   { color: GOLD, size: 20, style: BorderStyle.SINGLE, space: 8 },
        },
      }));
    } else if (tag === 'ul' || tag === 'ol') {
      Array.from(node.querySelectorAll(':scope > li')).forEach((li, idx) => {
        const lt = getText(li).replace(/\s+/g, ' ').trim();
        if (lt) children.push(new Paragraph({
          children: [new TextRun({ text: (tag === 'ol' ? (idx+1)+'. ' : '•  ') + lt, size: 22, color: TEXT, font: F_BODY })],
          indent: { left: 400, hanging: 200 },
          spacing: { before: 20, after: 20, line: 240, lineRule: AUTO },
        }));
      });
      children.push(new Paragraph({ text: '', spacing: { after: 40 } }));
    } else if (tag === 'table') {
      const tbl = makeTable(node);
      if (tbl) { children.push(tbl); children.push(new Paragraph({ text: '', spacing: { after: 80 } })); }
    } else if (tag === 'hr') {
      children.push(new Paragraph({
        children: [new TextRun({ text: '' })],
        border: { bottom: { color: BORDER, size: 4, style: BorderStyle.SINGLE, space: 3 } },
        spacing: { before: 80, after: 80 },
      }));
    } else {
      Array.from(node.childNodes || []).forEach(c => { if (c.nodeType === 1) parseNode(c); });
    }
  }

  Array.from(docBody.childNodes).forEach(n => { if (n.nodeType === 1) parseNode(n); });

  // ── NAGŁÓWEK STRONY ──
  const metaClean = meta.replace(/MATERIAA? DLA /i, '').replace(/MATERIAŁ DLA /i, '').replace(/ · PORT LOTNICZY LUBLIN S\.A\./gi, '').trim();
  const header = new Header({ children: [new Paragraph({
    children: [
      new TextRun({ text: 'PORT LOTNICZY LUBLIN S.A.', bold: true, size: 16, color: GOLD, font: F_HEAD }),
      new TextRun({ text: metaClean ? '  ·  ' + metaClean : '', size: 16, color: MUTED, font: F_HEAD }),
    ],
    alignment: AlignmentType.RIGHT,
    border: { bottom: { color: GOLD, size: 6, style: BorderStyle.SINGLE, space: 3 } },
    spacing: { after: 0 },
  })] });

  // ── STOPKA Z NUMERAMI STRON ──
  const shortTitle = title.length > 45 ? title.substring(0, 45) + '…' : title;
  const footer = new Footer({ children: [new Paragraph({
    children: [
      new TextRun({ text: shortTitle + '  |  Port Lotniczy Lublin  |  Strona ', size: 16, color: MUTED, font: F_HEAD }),
      new SimpleField('PAGE', { size: 16, color: MUTED, font: F_HEAD }),
      new TextRun({ text: ' / ', size: 16, color: MUTED, font: F_HEAD }),
      new SimpleField('NUMPAGES', { size: 16, color: MUTED, font: F_HEAD }),
    ],
    alignment: AlignmentType.CENTER,
    border: { top: { color: BORDER, size: 4, style: BorderStyle.SINGLE, space: 3 } },
  })] });

  // ── DOKUMENT ──
  const document = new Document({
    creator: '', description: '', title: '', subject: '', keywords: '', lastModifiedBy: '', revision: 1,
    sections: [{
      properties: {
        page: {
          margin: { top: 1440, right: 1584, bottom: 1440, left: 1584, header: 576, footer: 576, gutter: 0 },
          size: { width: 11906, height: 16838, orientation: 'portrait' },
        },
      },
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
