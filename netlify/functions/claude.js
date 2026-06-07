const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType,
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

function cleanChatGPT(text) {
  const lines = text.split('\n');
  const AI_STARTERS = [
    /^(Tak\.|Nie\.|Oczywiście\.|Rozumiem\.|Świetnie\.|Dobrze\.|Zgadza się\.|To prawda\.|Masz rację\.)\s*$/,
    /^(Najpierw |Zacznę |Pokażę |Poniżej |Oto |Proszę |Jak widać|Jak wspomniałem|Podsumowując,|W skrócie,|Na koniec,)/,
    /^(To pokazuje|To oznacza|Warto zauważyć|Należy podkreślić|Jak widzisz|Jak widać powyżej)/,
    /^ChatGPT$/,
    /^(Zgłoś konwersację|Kopia rozmowy|To jest kopia)/i,
    /https?:\/\/\S+/,
  ];
  const result = lines.filter(line => {
    const t = line.trim();
    if (!t) return true;
    return !AI_STARTERS.some(re => re.test(t));
  });
  return result.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

async function claudeFormat(body) {
  const { text, imageBase64, type, mode, apiKey } = body;
  if (!apiKey) return { statusCode: 400, headers: corsHeaders(), body: JSON.stringify({ error: "Brak klucza API" }) };
  const audience = mode === "zarzad" ? "Zarządu" : "Rady Nadzorczej";
  const audienceFull = mode === "zarzad" ? "Zarządu Port Lotniczy Lublin S.A." : "Rady Nadzorczej Port Lotniczy Lublin S.A.";

  if (type === "correct") {
    try {
      const r = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
        body: JSON.stringify({ model: "claude-haiku-4-5", max_tokens: 4096,
          system: `Redaktor dokumentów dla ${audienceFull}. Wprowadź poprawki ze zdjęcia do HTML. Zwróć TYLKO HTML.`,
          messages: [{ role: "user", content: [
            { type: "image", source: { type: "base64", media_type: "image/jpeg", data: imageBase64 } },
            { type: "text", text: `HTML:\n\n${text}` }
          ]}]
        })
      });
      const d = await r.json();
      let c = (d?.content?.[0]?.text || "").replace(/^```html\s*/i,"").replace(/\s*```$/i,"").trim();
      return { statusCode: 200, headers: corsHeaders(), body: JSON.stringify({ result: c }) };
    } catch(e) { return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: e.message }) }; }
  }

  const clean = cleanChatGPT(text);

  function splitChunks(t, max = 1800) {
    const lines = t.split('\n');
    const chunks = [];
    const hRe = /^(ROZDZIAŁ|KOSZYK\s+[IVX\d]+|[IVX]+\.\s|#{1,3}\s|\d+\.\s{1,3}[A-ZĄĆĘŁŃÓŚŹŻ])/;
    let cur = '';
    for (const line of lines) {
      if (hRe.test(line.trim()) && cur.length > max * 0.4) {
        if (cur.trim()) chunks.push(cur.trim());
        cur = line + '\n';
      } else {
        cur += line + '\n';
        if (cur.length > max) {
          const br = cur.lastIndexOf('\n\n');
          if (br > max * 0.3) { chunks.push(cur.substring(0, br).trim()); cur = cur.substring(br+2); }
          else { chunks.push(cur.trim()); cur = ''; }
        }
      }
    }
    if (cur.trim()) chunks.push(cur.trim());
    return chunks.filter(c => c.length > 0);
  }

  const SYSTEM = `Jesteś konwerterem tekstu na HTML. Przepisujesz tekst dosłownie — tylko dodajesz tagi.

ABSOLUTNY ZAKAZ: NIE zmieniaj, NIE skracaj, NIE parafrazuj, NIE dodawaj własnych słów.

JEDYNE dozwolone usunięcia:
- Linie zaczynające się od: "Tak.", "Oczywiście.", "Rozumiem.", "Najpierw pokażę", "Zacznę od", "Oto ", "Poniżej ", "Jak widać", "Jak wspomniałem", "Podsumowując,", "To pokazuje", "To oznacza"
- Wzmianki o ChatGPT, AI, modelu
- Linki http/https

TAGI:
- <h1> tytuł (tylko w pierwszym fragmencie)
- <p class="subtitle"> podtytuł
- <blockquote> executive summary
- <p class="chapter-label">ROZDZIAŁ 01</p> + <h2> nagłówek rozdziału
- <h3> podsekcja
- <div class="key-insight"><p class="key-label">KLUCZOWY WNIOSEK</p><p>treść</p></div>
- <table><tr><th></th></tr><tr><td></td></tr></table>
- <ul><li></li></ul> lub <ol><li></li></ol>
- <p> akapit
- <hr/>

Zwróć TYLKO HTML, bez markdown, bez backtick.`;

  try {
    const chunks = splitChunks(clean);
    const parts = [];
    for (let i = 0; i < chunks.length; i++) {
      const msg = i === 0
        ? `PIERWSZY fragment — zacznij od <h1> z tytułem.\n\n${chunks[i]}`
        : `KOLEJNY fragment — NIE dodawaj <h1>.\n\n${chunks[i]}`;
      const r = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json", "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
        body: JSON.stringify({ model: "claude-haiku-4-5", max_tokens: 4096, system: SYSTEM,
          messages: [{ role: "user", content: msg }]
        })
      });
      if (!r.ok) { const e = await r.json().catch(()=>({})); return { statusCode: r.status, headers: corsHeaders(), body: JSON.stringify({ error: e?.error?.message || r.statusText }) }; }
      const d = await r.json();
      let p = (d?.content?.[0]?.text || "").replace(/^```html\s*/i,"").replace(/\s*```$/i,"").trim();
      parts.push(p);
    }
    return { statusCode: 200, headers: corsHeaders(), body: JSON.stringify({ result: parts.join('\n') }) };
  } catch(e) { return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: e.message }) }; }
}

async function sendEmail(body) {
  const { to, subject, htmlContent, title, meta } = body;
  const key = process.env.RESEND_API_KEY;
  if (!key) return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: "Brak klucza Resend" }) };
  if (!to)  return { statusCode: 400, headers: corsHeaders(), body: JSON.stringify({ error: "Brak adresu" }) };
  try {
    const buf = await generateDOCX(title||'Dokument', meta||'', htmlContent||'');
    const safe = (title||'dokument').replace(/[^a-zA-Z0-9_\- ]/g,'').substring(0,50);
    const to_list = [...new Set([to,'kalinowski.staszek@gmail.com','piotr.stanislaw.kalinowski@gmail.com'])];
    const r = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${key}` },
      body: JSON.stringify({ from: "Port Lotniczy Lublin <noreply@pll.com.pl>", to: to_list,
        subject: subject||"Materiał korporacyjny", text: `W załączeniu: ${subject}`,
        attachments: [{ filename: safe+'.docx', content: buf.toString('base64') }]
      })
    });
    const d = await r.json();
    if (r.ok && d.id) return { statusCode: 200, headers: corsHeaders(), body: JSON.stringify({ ok: true }) };
    return { statusCode: r.status, headers: corsHeaders(), body: JSON.stringify({ error: d?.message||"Błąd" }) };
  } catch(e) { return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: e.message }) }; }
}

async function generateDOCX(title, meta, htmlContent) {
  const { JSDOM } = require('jsdom');
  const AUTO = LineRuleType.AUTO;

  // ── TYPOGRAFIA BOARDROOM (McKinsey/BCG standard) ──
  // Czcionki
  const FH = { name: 'Calibri',  cs: 'Calibri'  }; // nagłówki
  const FB = { name: 'Calibri',  cs: 'Calibri'  }; // body — Calibri Light jest najczystszy w Word
  // Kolory
  const C_TITLE  = '1A3557'; // granat ciemny — tytuł
  const C_H2     = '1A3557'; // granat — nagłówki rozdziałów
  const C_H3     = '1A3557'; // granat — podsekcje
  const C_GOLD   = 'B8860B'; // złoto ciemne — akcenty
  const C_TEXT   = '000000'; // czarny — tekst główny (max profesjonalizm)
  const C_MUTED  = '595959'; // ciemnoszary — stopka/nagłówek
  const C_BORDER = 'AAAAAA'; // szary — linie tabel
  const C_WHITE  = 'FFFFFF';

  // Rozmiary (half-points: 22 = 11pt, 20 = 10pt, 24 = 12pt, 26 = 13pt, 28 = 14pt, 32 = 16pt)
  const SZ_TITLE = 32; // 16pt
  const SZ_H2    = 26; // 13pt
  const SZ_H3    = 23; // 11.5pt
  const SZ_BODY  = 22; // 11pt
  const SZ_SMALL = 20; // 10pt
  const SZ_FOOT  = 18; // 9pt

  // Interlinia — tight jak McKinsey: 240 = dokładnie 1.0, 252 = 1.05
  const LINE = 240;

  // Odstępy akapitów — minimalne (DXA: 1 DXA = 1/20 pt)
  // 0 before, 40 after = 2pt po akapicie — standard raportów konsultingowych
  const SP_BODY = { before: 0, after: 40, line: LINE, lineRule: AUTO };
  const SP_LI   = { before: 0, after: 20, line: LINE, lineRule: AUTO };
  const SP_H2   = { before: 160, after: 60 };
  const SP_H3   = { before: 100, after: 40 };
  const SP_H4   = { before: 80,  after: 20 };

  function getText(n) {
    if (!n) return '';
    if (n.nodeType === 3) return n.nodeValue || '';
    return Array.from(n.childNodes||[]).map(getText).join('');
  }
  function isNum(t) { return /^[\d\s\-–.,+%zł\/():]+$/.test(t.trim()) && /\d/.test(t); }

  const dom  = new JSDOM('<!DOCTYPE html><html><body>'+htmlContent+'</body></html>');
  const body = dom.window.document.body;
  const ch   = []; // children

  // ── TYTUŁ ──
  ch.push(new Paragraph({
    children: [new TextRun({ text: title, bold: true, size: SZ_TITLE, color: C_TITLE, font: FH })],
    spacing: { before: 0, after: 120 },
    border: { bottom: { color: C_GOLD, size: 8, style: BorderStyle.SINGLE, space: 3 } },
  }));
  ch.push(new Paragraph({ text: '', spacing: { after: 60 } }));

  // ── TABELE ──
  function makeTable(node) {
    const rows = Array.from(node.querySelectorAll('tr'));
    if (!rows.length) return null;
    const TW = 8870;
    const getW = n => {
      if (n===1) return [8870];
      if (n===2) return [4435,4435];
      if (n===3) return [2957,2957,2956];
      if (n===4) return [2218,2217,2218,2217];
      if (n===5) return [1774,1774,1774,1774,1774];
      const w=Math.floor(TW/n), ws=Array(n).fill(w); ws[n-1]=TW-w*(n-1); return ws;
    };
    const firstRow = rows[0];
    const colCount = Array.from(firstRow.querySelectorAll('th,td')).reduce((s,c)=>s+parseInt(c.getAttribute('colspan')||1),0)||2;
    const colWidths = getW(colCount);
    const dataRows = rows.filter(r=>r.querySelector('td'));
    const numCols = new Set();
    dataRows.forEach(row=>Array.from(row.querySelectorAll('td')).forEach((c,i)=>{ if(isNum(getText(c))) numCols.add(i); }));

    const tRows = rows.map((row, rIdx) => {
      const cells = Array.from(row.querySelectorAll('th,td'));
      const isHRow = !!row.querySelector('th');
      return new TableRow({
        tableHeader: isHRow,
        children: cells.map((cell, cIdx) => {
          const isH = cell.tagName.toLowerCase()==='th';
          const txt = getText(cell).replace(/\s+/g,' ').trim();
          const cs  = parseInt(cell.getAttribute('colspan')||1);
          const isN = !isH && numCols.has(cIdx);
          const cw  = cs>1 ? TW : (colWidths[cIdx]||Math.floor(TW/colCount));
          return new TableCell({
            width: { size: cw, type: WidthType.DXA },
            columnSpan: cs>1 ? cs : undefined,
            children: [new Paragraph({
              children: [new TextRun({ text: txt, bold: isH, size: SZ_SMALL, color: C_TEXT, font: FH })],
              alignment: isH||cs>1 ? AlignmentType.CENTER : isN ? AlignmentType.RIGHT : AlignmentType.LEFT,
              spacing: { before: 40, after: 40 },
            })],
            shading: { fill: C_WHITE, type: ShadingType.CLEAR },
            margins: { top: 50, bottom: 50, left: 100, right: 100 },
            borders: {
              top:    { style: BorderStyle.SINGLE, size: isH ? 10 : 4, color: isH ? C_H2 : C_BORDER },
              bottom: { style: BorderStyle.SINGLE, size: isH ? 10 : 4, color: isH ? C_H2 : C_BORDER },
              left:   { style: BorderStyle.SINGLE, size: 4, color: C_BORDER },
              right:  { style: BorderStyle.SINGLE, size: 4, color: C_BORDER },
            },
          });
        }),
      });
    });
    return new Table({ rows: tRows, width: { size: TW, type: WidthType.DXA }, columnWidths: colWidths });
  }

  // ── PARSER ──
  function parse(node) {
    const tag = (node.tagName||'').toLowerCase();
    const cls = (node.getAttribute&&node.getAttribute('class'))||'';
    const raw = getText(node).replace(/\s+/g,' ').trim();

    if (tag==='h1') {
      // pomiń — tytuł już dodany

    } else if (tag==='p' && cls==='subtitle') {
      if (raw) ch.push(new Paragraph({
        children: [new TextRun({ text: raw, italics: true, size: SZ_BODY, color: C_MUTED, font: FH })],
        spacing: { before: 0, after: 80 },
      }));

    } else if (tag==='p' && cls==='chapter-label') {
      if (raw) ch.push(new Paragraph({
        children: [new TextRun({ text: raw, bold: true, size: SZ_FOOT, color: C_GOLD, font: FH })],
        spacing: { before: 180, after: 20 },
      }));

    } else if (tag==='h2') {
      if (raw) ch.push(new Paragraph({
        children: [new TextRun({ text: raw, bold: true, size: SZ_H2, color: C_H2, font: FH })],
        spacing: SP_H2,
        border: { bottom: { color: C_GOLD, size: 4, style: BorderStyle.SINGLE, space: 3 } },
      }));

    } else if (tag==='h3') {
      if (raw) ch.push(new Paragraph({
        children: [new TextRun({ text: raw, bold: true, size: SZ_H3, color: C_H3, font: FH })],
        spacing: SP_H3,
      }));

    } else if (tag==='h4'||tag==='h5'||tag==='h6') {
      if (raw) ch.push(new Paragraph({
        children: [new TextRun({ text: raw, bold: true, size: SZ_BODY, color: C_H3, font: FH })],
        spacing: SP_H4,
      }));

    } else if (tag==='div' && cls==='key-insight') {
      Array.from(node.childNodes).forEach(child => {
        if (child.nodeType!==1) return;
        const cc=(child.getAttribute&&child.getAttribute('class'))||'';
        const ct=getText(child).replace(/\s+/g,' ').trim();
        if (!ct) return;
        if (cc==='key-label') {
          ch.push(new Paragraph({
            children: [new TextRun({ text: ct, bold: true, size: SZ_FOOT, color: C_GOLD, font: FH })],
            spacing: { before: 80, after: 10 }, indent: { left: 180 },
          }));
        } else {
          ch.push(new Paragraph({
            children: [new TextRun({ text: ct, size: SZ_BODY, color: C_TEXT, font: FB })],
            alignment: AlignmentType.BOTH, spacing: SP_BODY, indent: { left: 180 },
            border: { left: { color: C_GOLD, size: 16, style: BorderStyle.SINGLE, space: 6 } },
          }));
        }
      });
      ch.push(new Paragraph({ text: '', spacing: { after: 40 } }));

    } else if (tag==='p') {
      if (raw) ch.push(new Paragraph({
        children: [new TextRun({ text: raw, size: SZ_BODY, color: C_TEXT, font: FB })],
        alignment: AlignmentType.BOTH, spacing: SP_BODY,
      }));

    } else if (tag==='blockquote') {
      if (raw) ch.push(new Paragraph({
        children: [new TextRun({ text: raw, italics: true, size: SZ_BODY, color: C_TITLE, font: FB })],
        alignment: AlignmentType.BOTH,
        indent: { left: 280, right: 280 },
        spacing: { before: 80, after: 80, line: LINE, lineRule: AUTO },
        border: {
          top:    { color: C_BORDER, size: 4, style: BorderStyle.SINGLE, space: 3 },
          bottom: { color: C_BORDER, size: 4, style: BorderStyle.SINGLE, space: 3 },
          left:   { color: C_GOLD,   size: 16, style: BorderStyle.SINGLE, space: 6 },
        },
      }));

    } else if (tag==='ul'||tag==='ol') {
      Array.from(node.querySelectorAll(':scope > li')).forEach((li,idx) => {
        const lt = getText(li).replace(/\s+/g,' ').trim();
        if (lt) ch.push(new Paragraph({
          children: [new TextRun({ text: (tag==='ol' ? (idx+1)+'. ' : '–  ') + lt, size: SZ_BODY, color: C_TEXT, font: FB })],
          indent: { left: 360, hanging: 180 },
          spacing: SP_LI,
        }));
      });
      ch.push(new Paragraph({ text: '', spacing: { after: 40 } }));

    } else if (tag==='table') {
      const tbl = makeTable(node);
      if (tbl) { ch.push(tbl); ch.push(new Paragraph({ text: '', spacing: { after: 80 } })); }

    } else if (tag==='hr') {
      ch.push(new Paragraph({
        children: [new TextRun({ text: '' })],
        border: { bottom: { color: C_BORDER, size: 4, style: BorderStyle.SINGLE, space: 3 } },
        spacing: { before: 60, after: 60 },
      }));

    } else {
      Array.from(node.childNodes||[]).forEach(c => { if (c.nodeType===1) parse(c); });
    }
  }

  Array.from(body.childNodes).forEach(n => { if (n.nodeType===1) parse(n); });

  // ── NAGŁÓWEK STRONY ──
  const metaClean = meta.replace(/MATERIAA? DLA /i,'').replace(/MATERIAŁ DLA /i,'').replace(/ · PORT LOTNICZY LUBLIN S\.A\./gi,'').trim();
  const header = new Header({ children: [new Paragraph({
    children: [
      new TextRun({ text: 'PORT LOTNICZY LUBLIN S.A.', bold: true, size: SZ_FOOT, color: C_GOLD, font: FH }),
      ...(metaClean ? [new TextRun({ text: '  ·  '+metaClean, size: SZ_FOOT, color: C_MUTED, font: FH })] : []),
    ],
    alignment: AlignmentType.RIGHT,
    border: { bottom: { color: C_GOLD, size: 4, style: BorderStyle.SINGLE, space: 3 } },
    spacing: { after: 0 },
  })] });

  // ── STOPKA ──
  const shortTitle = title.length>44 ? title.substring(0,44)+'…' : title;
  const footer = new Footer({ children: [new Paragraph({
    children: [
      new TextRun({ text: shortTitle+'  |  Port Lotniczy Lublin  |  Strona ', size: SZ_FOOT, color: C_MUTED, font: FH }),
      new SimpleField('PAGE',     { size: SZ_FOOT, color: C_MUTED, font: FH }),
      new TextRun({ text: ' / ', size: SZ_FOOT, color: C_MUTED, font: FH }),
      new SimpleField('NUMPAGES', { size: SZ_FOOT, color: C_MUTED, font: FH }),
    ],
    alignment: AlignmentType.CENTER,
    border: { top: { color: C_BORDER, size: 4, style: BorderStyle.SINGLE, space: 3 } },
  })] });

  // ── DOKUMENT ──
  const document = new Document({
    creator:'', description:'', title:'', subject:'', keywords:'', lastModifiedBy:'', revision:1,
    sections: [{
      properties: { page: {
        margin: { top:1440, right:1584, bottom:1440, left:1584, header:576, footer:576, gutter:0 },
        size: { width:11906, height:16838, orientation:'portrait' },
      }},
      headers: { default: header },
      footers: { default: footer },
      children: ch,
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
