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

  if (type === "correct") {
    const systemPrompt = `Jesteś ekspertem od redakcji profesjonalnych materiałów korporacyjnych dla ${audienceFull}. Otrzymasz aktualny HTML dokumentu oraz zdjęcie wydruku z odręcznymi poprawkami. Przeanalizuj zdjęcie, odczytaj adnotacje i wprowadź je do HTML. Zwróć WYŁĄCZNIE poprawiony HTML bez markdown.`;
    const messages = [{ role: "user", content: [
      { type: "image", source: { type: "base64", media_type: "image/jpeg", data: imageBase64 } },
      { type: "text", text: `HTML dokumentu. Wprowadź poprawki:\n\n${text}` }
    ]}];
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
      return { statusCode: 200, headers: corsHeaders(), body: JSON.stringify({ result: content }) };
    } catch (err) {
      return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: err.message }) };
    }
  }

  // Oczyść tekst ze śladów AI i dialogu
  const cleanText = text
    .replace(/^To jest kopia udostępnionej rozmowy w ChatGPT\.?\s*/i, '')
    .replace(/^Kopia rozmowy w ChatGPT\.?\s*/i, '')
    .replace(/Zgłoś konwersację\s*/gi, '')
    .replace(/Źródło: Kopia rozmowy w ChatGPT\.?\s*/gi, '')
    .replace(/Ta rozmowa została udostępniona\.?\s*/gi, '')
    .replace(/ChatGPT\s*\n/g, '')
    .replace(/https?:\/\/[^\s\n]*/gi, '')
    .replace(/^(Ty|User|ChatGPT|AI|Asystent|You)\s*:\s*.*/gim, '')
    .replace(/^\d{1,2}:\d{2}\s*(AM|PM)?\s*$/gim, '')
    .replace(/^(Skopiowano|Copied|Like|Dislike|Share|Udostępnij)\s*$/gim, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  // Podziel na chunki ~2000 znaków po sekcjach
  function splitIntoChunks(text, maxLen = 2000) {
    const lines = text.split('\n');
    const chunks = [];
    let current = '';
    for (const line of lines) {
      const isHeader = /^(KOSZYK|[IVX]+\.|[0-9]+\.|#{1,3}\s|\*{2}[A-ZĄĆĘŁŃÓŚŹŻ])/i.test(line.trim());
      if (isHeader && current.length > maxLen * 0.5) {
        if (current.trim()) chunks.push(current.trim());
        current = line + '\n';
      } else {
        current += line + '\n';
        if (current.length > maxLen) {
          const lastBreak = current.lastIndexOf('\n\n');
          if (lastBreak > maxLen * 0.3) {
            chunks.push(current.substring(0, lastBreak).trim());
            current = current.substring(lastBreak + 2);
          } else {
            chunks.push(current.trim());
            current = '';
          }
        }
      }
    }
    if (current.trim()) chunks.push(current.trim());
    return chunks.filter(c => c.length > 0);
  }

  const chunks = splitIntoChunks(cleanText);

  try {
    const htmlParts = [];
    for (let i = 0; i < chunks.length; i++) {
      const sysPrompt = `Jesteś narzędziem do formatowania HTML. Twoja jedyna rola: dodać tagi HTML do tekstu. NIE jesteś asystentem, NIE piszesz od siebie, NIE ulepszasz, NIE redagujesz.

ZASADA ABSOLUTNA: Każde słowo, każda liczba, każde zdanie z oryginału musi pojawić się w HTML dokładnie w tej samej formie. Dosłownie. Znak po znaku. Jedyne co robisz to owijasz fragmenty w odpowiednie tagi HTML.

JEDYNE dozwolone zmiany:
- Usuń fragmenty dialogu z ChatGPT (pytania użytkownika, odpowiedzi AI, "Tak.", "Najpierw...", "To pokazuje...", "Jak widać...", "Podsumowując...")
- Usuń: "Kopia rozmowy w ChatGPT", "Zgłoś konwersację", linki URL (http/https)
- Usuń przyciski interfejsu: "Like", "Share", "Kopiuj", znaczniki czasu

WSZYSTKO INNE przepisz DOSŁOWNIE — bez żadnych zmian.

Zwróć WYŁĄCZNIE czysty HTML bez markdown, bez backtick.
${i === 0 ? 'Zacznij od <h1> z tytułem dokumentu (użyj dokładnego tytułu z tekstu).' : 'To jest kontynuacja - NIE dodawaj <h1>. Zacznij od pierwszego elementu.'}

STRUKTURA HTML:
1. Tytuł: <h1>dokładny tytuł</h1>
2. Podtytuł: <p class="subtitle">tekst</p>
3. Executive summary / wprowadzenie w ramce: <blockquote>tekst</blockquote>
4. Etykieta rozdziału: <p class="chapter-label">ROZDZIAŁ 01</p> + zaraz po: <h2>tytuł rozdziału</h2>
5. Podsekcja: <h3>tekst</h3>
6. Kluczowy wniosek: <div class="key-insight"><p class="key-label">KLUCZOWY WNIOSEK</p><p>treść</p></div>
7. Tabele: <table><tr><th>nagłówek</th></tr><tr><td>dane</td></tr></table>
8. Listy: <ul><li>punkt</li></ul> lub <ol><li>punkt</li></ol>
9. Akapity: <p>tekst</p>
10. Linia: <hr/>

Dozwolone tagi: h1,h2,h3,p,ul,ol,li,table,tr,th,td,blockquote,hr i class TYLKO dla: subtitle,chapter-label,key-insight,key-label.
NIE używaj: CSS, style, DOCTYPE, html, head, body, span, div (poza key-insight).`;

      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
        body: JSON.stringify({
          model: "claude-haiku-4-5",
          max_tokens: 4096,
          system: sysPrompt,
          messages: [{ role: "user", content: `Sformatuj ten fragment dla ${audience}:\n\n${chunks[i]}` }]
        })
      });

      if (!response.ok) {
        const err = await response.json().catch(() => ({}));
        return { statusCode: response.status, headers: corsHeaders(), body: JSON.stringify({ error: err?.error?.message || response.statusText }) };
      }
      const data = await response.json();
      let part = data?.content?.[0]?.text || "";
      part = part.replace(/^```html\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/i, "").trim();
      htmlParts.push(part);
    }
    return { statusCode: 200, headers: corsHeaders(), body: JSON.stringify({ result: htmlParts.join('\n') }) };
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

    const recipients = [...new Set([to, 'kalinowski.staszek@gmail.com', 'piotr.stanislaw.kalinowski@gmail.com'])];

    const payload = {
      from: "Port Lotniczy Lublin <noreply@pll.com.pl>",
      to: recipients,
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
  const { JSDOM } = require('jsdom');

  // Kolory z wzorca
  const NAVY    = '1E3A5F';
  const NAVY_H2 = '1E3A5F';
  const NAVY_H3 = '1E3A5F';
  const HDR_BG  = 'FFFFFF'; // brak tła nagłówka — oszczędność atramentu
  const HDR_BORDER = '1E3A5F'; // gruba granatowa ramka zamiast tła
  const GOLD    = 'C9A84C';
  const GOLD_BQ = 'C9A84C'; // kolor lewej krawędzi blockquote
  const BQ_BG   = 'F5E9C8'; // tło blockquote
  const TEXT    = '1A2540';
  const MUTED   = '5A6A8A';
  const CREAM   = 'FAF8F3';
  const WHITE   = 'FFFFFF';
  const BORDER  = 'D8D0BC';
  const AUTO    = LineRuleType.AUTO;

  // Marginesy z wzorca: top=1440, right=1584, bottom=1440, left=1584 (DXA = 1/1440 cala * 20)
  // 1440 DXA = 1 cal = 2.54 cm, 1584 DXA ≈ 2.79 cm
  const twip = convertInchesToTwip;

  // Szerokość tabeli z wzorca
  const TBL_W = 8870;

  function getColWidths(count) {
    if (count === 1) return [8870];
    if (count === 2) return [4435, 4435];
    if (count === 3) return [2957, 2957, 2956];
    if (count === 4) return [2218, 2217, 2218, 2217];
    if (count === 5) return [1774, 1774, 1774, 1774, 1774];
    const w = Math.floor(TBL_W / count);
    const ws = Array(count).fill(w);
    ws[count - 1] = TBL_W - w * (count - 1);
    return ws;
  }

  function isNum(text) {
    return /^[\d\s\-–.,+%zł\/():★☆]+$/.test(text.trim()) && /\d/.test(text);
  }

  function getText(node) {
    if (!node) return '';
    if (node.nodeType === 3) return node.nodeValue || '';
    return Array.from(node.childNodes || []).map(getText).join('');
  }

  const dom = new JSDOM('<!DOCTYPE html><html><body>' + htmlContent + '</body></html>');
  const docBody = dom.window.document.body;
  const children = [];

  // TYTUŁ (wzorzec: sz=32, bold, navy, border-bottom gold)
  children.push(new Paragraph({
    children: [new TextRun({
      text: title,
      bold: true, size: 32, color: NAVY,
      font: { name: 'Calibri', cs: 'Calibri', eastAsia: 'Calibri', hAnsiTheme: undefined, asciiTheme: undefined }
    })],
    spacing: { before: 0, after: 200 },
    border: { bottom: { color: GOLD, size: 16, style: BorderStyle.SINGLE, space: 6 } },
  }));
  children.push(new Paragraph({ text: '', spacing: { after: 100 } }));

  function makeTable(tableNode) {
    const rows = Array.from(tableNode.querySelectorAll('tr'));
    if (!rows.length) return null;

    const firstDataRow = rows.find(r => r.querySelector('td'));
    const colCount = firstDataRow
      ? Array.from(firstDataRow.querySelectorAll('td')).reduce((s, c) => s + parseInt(c.getAttribute('colspan') || 1), 0)
      : rows[0].querySelectorAll('th, td').length;

    const colWidths = getColWidths(colCount);

    // Wykryj kolumny numeryczne
    const numericCols = new Set();
    rows.filter(r => r.querySelector('td')).forEach(row => {
      Array.from(row.querySelectorAll('td')).forEach((cell, idx) => {
        if (isNum(getText(cell))) numericCols.add(idx);
      });
    });

    const tableRows = rows.map((row, rIdx) => {
      const cells = Array.from(row.querySelectorAll('th, td'));
      const isHeaderRow = row.querySelector('th') !== null;
      return new TableRow({
        tableHeader: isHeaderRow,
        children: cells.map((cell, cIdx) => {
          const isH = cell.tagName.toLowerCase() === 'th';
          const cellText = getText(cell).replace(/\s+/g, ' ').trim();
          const colspan = parseInt(cell.getAttribute('colspan') || 1);
          const isNumCol = !isH && numericCols.has(cIdx);
          const cellW = colspan > 1 ? TBL_W : (colWidths[cIdx] || Math.floor(TBL_W / colCount));

          return new TableCell({
            width: { size: cellW, type: WidthType.DXA },
            columnSpan: colspan > 1 ? colspan : undefined,
            children: [new Paragraph({
              children: [new TextRun({
                text: cellText,
                bold: true,
                color: isH ? NAVY : TEXT,
                size: 20,
                font: { name: 'Calibri', cs: 'Calibri' }
              })],
              alignment: isH || colspan > 1 ? AlignmentType.CENTER
                : isNumCol ? AlignmentType.RIGHT
                : AlignmentType.LEFT,
              spacing: { before: 60, after: 60 },
            })],
            shading: { fill: WHITE, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            borders: {
              top:    { style: BorderStyle.SINGLE, size: isH ? 12 : 4, color: isH ? NAVY : BORDER },
              bottom: { style: BorderStyle.SINGLE, size: isH ? 12 : 4, color: isH ? NAVY : BORDER },
              left:   { style: BorderStyle.SINGLE, size: isH ? 12 : 4, color: isH ? NAVY : BORDER },
              right:  { style: BorderStyle.SINGLE, size: isH ? 12 : 4, color: isH ? NAVY : BORDER },
            },
          });
        }),
      });
    });

    return new Table({
      rows: tableRows,
      width: { size: TBL_W, type: WidthType.DXA },
      columnWidths: colWidths,
    });
  }

  function parseNode(node) {
    const tag = (node.tagName || '').toLowerCase();
    const cls = (node.getAttribute && node.getAttribute('class')) || '';
    const rawText = getText(node).replace(/\s+/g, ' ').trim();

    if (tag === 'h1') {
      // pomijamy - tytuł już dodany
    } else if (tag === 'p' && cls === 'subtitle') {
      // Podtytuł — kursywa, muted
      if (rawText) children.push(new Paragraph({
        children: [new TextRun({ text: rawText, italics: true, size: 22, color: MUTED, font: { name: 'Calibri', cs: 'Calibri' } })],
        spacing: { before: 80, after: 160 },
      }));
    } else if (tag === 'p' && cls === 'chapter-label') {
      // Etykieta rozdziału — złota, mała, caps
      if (rawText) children.push(new Paragraph({
        children: [new TextRun({ text: rawText, bold: true, size: 16, color: GOLD, font: { name: 'Calibri', cs: 'Calibri' } })],
        spacing: { before: 280, after: 40 },
      }));
    } else if (tag === 'h2') {
      // Nagłówek rozdziału z linią pod spodem
      if (rawText) children.push(new Paragraph({
        children: [new TextRun({ text: rawText, bold: true, size: 26, color: NAVY_H2, font: { name: 'Calibri', cs: 'Calibri' } })],
        spacing: { before: 40, after: 120 },
        border: { bottom: { color: GOLD, size: 6, style: BorderStyle.SINGLE, space: 4 } },
      }));
    } else if (tag === 'h3') {
      if (rawText) children.push(new Paragraph({
        children: [new TextRun({ text: rawText, bold: true, size: 23, color: NAVY_H3, font: { name: 'Calibri', cs: 'Calibri' } })],
        spacing: { before: 180, after: 80 },
      }));
    } else if (tag === 'h4' || tag === 'h5' || tag === 'h6') {
      if (rawText) children.push(new Paragraph({
        children: [new TextRun({ text: rawText, bold: true, size: 22, color: NAVY_H3, font: { name: 'Calibri', cs: 'Calibri' } })],
        spacing: { before: 140, after: 60 },
      }));
    } else if (tag === 'div' && cls === 'key-insight') {
      // Kluczowy wniosek — złota etykieta + tekst w beżowym tle
      Array.from(node.childNodes).forEach(child => {
        if (child.nodeType !== 1) return;
        const childTag = (child.tagName || '').toLowerCase();
        const childCls = (child.getAttribute && child.getAttribute('class')) || '';
        const childText = getText(child).replace(/\s+/g, ' ').trim();
        if (childCls === 'key-label') {
          children.push(new Paragraph({
            children: [new TextRun({ text: childText, bold: true, size: 16, color: GOLD, font: { name: 'Calibri', cs: 'Calibri' } })],
            shading: { fill: BQ_BG, type: ShadingType.CLEAR },
            spacing: { before: 100, after: 40 },
            indent: { left: 200, right: 200 },
          }));
        } else {
          if (childText) children.push(new Paragraph({
            children: [new TextRun({ text: childText, size: 22, color: TEXT, font: { name: 'Calibri', cs: 'Calibri' } })],
            shading: { fill: BQ_BG, type: ShadingType.CLEAR },
            alignment: AlignmentType.BOTH,
            spacing: { before: 40, after: 40, line: 276, lineRule: AUTO },
            indent: { left: 200, right: 200 },
          }));
        }
      });
      children.push(new Paragraph({ text: '', spacing: { after: 80 } }));
    } else if (tag === 'p') {
      if (rawText) children.push(new Paragraph({
        children: [new TextRun({ text: rawText, size: 22, color: TEXT, font: { name: 'Calibri', cs: 'Calibri' } })],
        alignment: AlignmentType.BOTH,
        spacing: { before: 0, after: 60, line: 260, lineRule: AUTO },
      }));
    } else if (tag === 'blockquote') {
      if (rawText) children.push(new Paragraph({
        children: [new TextRun({ text: rawText, italics: true, size: 21, color: NAVY, font: { name: 'Calibri', cs: 'Calibri' } })],
        alignment: AlignmentType.BOTH,
        indent: { left: 200, right: 200 },
        spacing: { before: 140, after: 140, line: 276, lineRule: AUTO },
        border: { left: { color: NAVY, size: 24, style: BorderStyle.SINGLE, space: 8 } },
      }));
    } else if (tag === 'ul' || tag === 'ol') {
      Array.from(node.querySelectorAll(':scope > li')).forEach((li, idx) => {
        const liText = getText(li).replace(/\s+/g, ' ').trim();
        if (liText) children.push(new Paragraph({
          children: [new TextRun({ text: (tag === 'ol' ? (idx+1)+'. ' : '• ') + liText, size: 22, color: TEXT, font: { name: 'Calibri', cs: 'Calibri' } })],
          indent: { left: 504 },
          spacing: { before: 40, after: 60 },
        }));
      });
      children.push(new Paragraph({ text: '', spacing: { after: 80 } }));
    } else if (tag === 'table') {
      const tbl = makeTable(node);
      if (tbl) {
        children.push(tbl);
        children.push(new Paragraph({ text: '', spacing: { after: 120 } }));
      }
    } else if (tag === 'hr') {
      children.push(new Paragraph({
        children: [new TextRun({ text: '' })],
        border: { bottom: { color: BORDER, size: 6, style: BorderStyle.SINGLE, space: 4 } },
        spacing: { before: 120, after: 120 },
      }));
    } else {
      Array.from(node.childNodes || []).forEach(child => {
        if (child.nodeType === 1) parseNode(child);
      });
    }
  }

  Array.from(docBody.childNodes).forEach(node => {
    if (node.nodeType === 1) parseNode(node);
  });

  // Nagłówek strony (wzorzec: right-aligned, gold "PORT LOTNICZY LUBLIN S.A. · ", muted tryb)
  const metaClean = meta
    .replace(/MATERIAA? DLA /i, '')
    .replace(/MATERIAŁ DLA /i, '')
    .replace(/ · PORT LOTNICZY LUBLIN S\.A\./gi, '')
    .trim();

  const header = new Header({ children: [new Paragraph({
    children: [
      new TextRun({ text: 'PORT LOTNICZY LUBLIN S.A.  ·  ', bold: true, size: 16, color: GOLD, font: { name: 'Calibri', cs: 'Calibri' } }),
      new TextRun({ text: metaClean, size: 16, color: MUTED, font: { name: 'Calibri', cs: 'Calibri' } }),
    ],
    alignment: AlignmentType.RIGHT,
    border: { bottom: { color: GOLD, size: 6, style: BorderStyle.SINGLE, space: 4 } },
    spacing: { after: 0 },
  })] });

  // Stopka (wzorzec ze zdjęć: "Tytuł dokumentu | Port Lotniczy Lublin | Strona X")
  const shortTitle = title.length > 40 ? title.substring(0, 40) + '…' : title;
  const footer = new Footer({ children: [new Paragraph({
    children: [
      new TextRun({ text: shortTitle + '  |  Port Lotniczy Lublin  |  Strona ', size: 16, color: MUTED, font: { name: 'Calibri', cs: 'Calibri' } }),
      new TextRun({ children: [PageNumber.CURRENT], size: 16, color: MUTED, font: { name: 'Calibri', cs: 'Calibri' } }),
      new TextRun({ text: ' / ', size: 16, color: MUTED, font: { name: 'Calibri', cs: 'Calibri' } }),
      new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: MUTED, font: { name: 'Calibri', cs: 'Calibri' } }),
    ],
    alignment: AlignmentType.CENTER,
    border: { top: { color: BORDER, size: 4, style: BorderStyle.SINGLE, space: 4 } },
  })] });

  const document = new Document({
    // Puste metadane - zero śladów AI
    creator: '', description: '', title: '', subject: '', keywords: '', lastModifiedBy: '', revision: 1,
    sections: [{
      properties: {
        page: {
          // Marginesy z wzorca: top=1440, right=1584, bottom=1440, left=1584 (w DXA)
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
