const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType,
  Footer, Header, LineRuleType, convertInchesToTwip, PageNumber, SimpleField } = require('docx');

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

function cleanInput(text) {
  const AI_LINE = /^(Tak\.|Nie\.|Oczywiście\.|Rozumiem\.|Świetnie\.|Dobrze\.|Zgadza się\.|To prawda\.|Masz rację\.|Najpierw |Zacznę |Pokażę |Poniżej |Oto |Proszę |Jak widać|Jak wspomniałem|Podsumowując,|W skrócie,|To pokazuje|To oznacza|ChatGPT$|Zgłoś konwersację|Kopia rozmowy|To jest kopia)/;
  return text.split('\n')
    .filter(l => !AI_LINE.test(l.trim()) && !/^https?:\/\//.test(l.trim()))
    .join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

async function claudeFormat(body) {
  const { text, imageBase64, type, mode, apiKey } = body;
  if (!apiKey) return { statusCode: 400, headers: corsHeaders(), body: JSON.stringify({ error: "Brak klucza API" }) };
  const audience = mode === "zarzad" ? "Zarządu" : "Rady Nadzorczej";

  if (type === "correct") {
    try {
      const r = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST", headers: { "Content-Type": "application/json", "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
        body: JSON.stringify({ model: "claude-haiku-4-5", max_tokens: 4096,
          system: "Redaktor dokumentów. Wprowadź poprawki ze zdjęcia do HTML. Zwróć TYLKO HTML.",
          messages: [{ role: "user", content: [
            { type: "image", source: { type: "base64", media_type: "image/jpeg", data: imageBase64 } },
            { type: "text", text: `HTML:\n\n${text}` }
          ]}]
        })
      });
      const d = await r.json();
      let c = (d?.content?.[0]?.text||"").replace(/^```html\s*/i,"").replace(/\s*```$/i,"").trim();
      return { statusCode: 200, headers: corsHeaders(), body: JSON.stringify({ result: c }) };
    } catch(e) { return { statusCode: 500, headers: corsHeaders(), body: JSON.stringify({ error: e.message }) }; }
  }

  const clean = cleanInput(text);

  function splitChunks(t, max=1800) {
    const lines = t.split('\n'), chunks = [];
    const hRe = /^(ROZDZIAŁ|KOSZYK\s+[IVX\d]+|[IVX]+\.\s|#{1,3}\s|\d+\.\s[A-ZĄĆĘŁŃÓŚŹŻ])/;
    let cur = '';
    for (const l of lines) {
      if (hRe.test(l.trim()) && cur.length > max*0.4) { if (cur.trim()) chunks.push(cur.trim()); cur = l+'\n'; }
      else {
        cur += l+'\n';
        if (cur.length > max) {
          const br = cur.lastIndexOf('\n\n');
          if (br > max*0.3) { chunks.push(cur.substring(0,br).trim()); cur = cur.substring(br+2); }
          else { chunks.push(cur.trim()); cur=''; }
        }
      }
    }
    if (cur.trim()) chunks.push(cur.trim());
    return chunks.filter(c=>c.length>0);
  }

  const SYSTEM = `Konwerter tekstu na HTML dla materiałów boardroom. Przepisujesz tekst DOSŁOWNIE — tylko dodajesz tagi HTML.

ZAKAZ: nie zmieniaj treści merytorycznej, nie skracaj, nie dodawaj własnych słów.

KOREKTA STYLU (obowiązkowa):
— Formy osobowe → bezosobowe: "rekomendowałbym"→"rekomenduje się", "proponuję"→"proponuje się", "uważam"→"ocenia się", "moim zdaniem"→"w świetle analizy", "sugerowałbym"→"sugeruje się", "podkreśliłbym"→"należy podkreślić", "po naszych analizach"→"na podstawie przeprowadzonych analiz"
— Listy ≤4 elementów → zdanie narracyjne w <p>: np. [a, b, c] → <p>...obejmuje: a, b oraz c.</p>
— Listy ≥5 lub KPI/harmonogramy → zachowaj jako <ul><li>

STRUKTURA (stosuj konsekwentnie dla analogicznych treści):
<h1>tytuł</h1>  — tylko raz, tylko w pierwszym fragmencie
<p class="subtitle">podtytuł</p>
<blockquote>executive summary</blockquote>
<p class="chapter-label">ROZDZIAŁ 01</p>  — ZAWSZE przed h2, numeruj 01 02 03...
<h2>nagłówek rozdziału</h2>  — ZAWSZE zaraz po chapter-label
<h3>podsekcja</h3>
<div class="key-insight"><p class="key-label">KLUCZOWY WNIOSEK</p><p>treść</p></div>
<table><tr><th>kol</th></tr><tr><td>dane</td></tr></table>
<ul><li>punkt</li></ul>  — tylko ≥5 elementów lub KPI
<ol><li>krok</li></ol>  — tylko sekwencje kroków
<p>akapit</p>
<hr/>

Zwróć TYLKO HTML bez markdown bez backtick.`;

  try {
    const chunks = splitChunks(clean);
    const parts = [];
    let lastChapterNum = 0;
    for (let i=0; i<chunks.length; i++) {
      let ctx = '';
      if (i===0) ctx = 'PIERWSZY fragment — zacznij od <h1>. Numeracja rozdziałów startuje od 01.';
      else ctx = `KOLEJNY fragment — NIE dodawaj <h1>. Ostatni numer rozdziału w poprzednim fragmencie: ${String(lastChapterNum).padStart(2,'0')}. Kontynuuj numerację od ${String(lastChapterNum+1).padStart(2,'0')} jeśli pojawi się nowy rozdział; jeśli fragment nie zawiera nowych rozdziałów — nie dodawaj chapter-label.`;
      const msg = `${ctx}\n\n${chunks[i]}`;
      const r = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST", headers: { "Content-Type": "application/json", "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
        body: JSON.stringify({ model: "claude-haiku-4-5", max_tokens: 4096, system: SYSTEM, messages: [{ role: "user", content: msg }] })
      });
      if (!r.ok) { const e=await r.json().catch(()=>({})); return { statusCode: r.status, headers: corsHeaders(), body: JSON.stringify({ error: e?.error?.message||r.statusText }) }; }
      const d = await r.json();
      let p = (d?.content?.[0]?.text||"").replace(/^```html\s*/i,"").replace(/\s*```$/i,"").trim();
      // wyciągnij ostatni numer rozdziału z tego fragmentu
      const chNums = [...p.matchAll(/ROZDZIAŁ\s+(\d{1,2})/gi)].map(m=>parseInt(m[1]));
      if (chNums.length) lastChapterNum = Math.max(...chNums);
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
    const toList = [...new Set([to,'kalinowski.staszek@gmail.com','piotr.stanislaw.kalinowski@gmail.com'])];
    const r = await fetch("https://api.resend.com/emails", {
      method: "POST", headers: { "Content-Type": "application/json", "Authorization": `Bearer ${key}` },
      body: JSON.stringify({ from: "Port Lotniczy Lublin <noreply@pll.com.pl>", to: toList,
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

  // ── PALETA KOLORÓW ──
  const NAVY  = '1A3557';  // granat — tytuł, nagłówki
  const GOLD  = 'C9A84C';  // złoto — akcenty, etykiety rozdziałów
  const BLACK = '000000';  // czarny — tekst główny
  const GREY  = '595959';  // szary — stopka, nagłówek strony
  const LINE_C= 'BBBBBB';  // jasny szary — linie tabel i separatory

  // ── CZCIONKI ──
  const FH = { name: 'Calibri', cs: 'Calibri' };  // nagłówki
  const FB = { name: 'Calibri', cs: 'Calibri' };  // body

  // ── TYPOGRAFIA (McKinsey/BCG standard) ──
  // Interlinia 1.0 = 240, Interlinia 1.08 = 259 (Word default)
  const LINE_H = 240;  // interlinia body — tight
  const A      = AUTO;

  // Rozmiary (half-points)
  const T  = 34;  // tytuł         17pt
  const H2 = 26;  // rozdział      13pt
  const H3 = 22;  // podsekcja     11pt
  const B  = 22;  // body          11pt
  const S  = 20;  // tabela/small  10pt
  const F  = 18;  // footer        9pt

  function getText(n) {
    if (!n) return '';
    if (n.nodeType===3) return n.nodeValue||'';
    return Array.from(n.childNodes||[]).map(getText).join('');
  }
  function isNum(t) { return /^[\d\s\-–.,+%zł\/():]+$/.test(t.trim()) && /\d/.test(t); }

  const dom  = new JSDOM('<!DOCTYPE html><html><body>'+htmlContent+'</body></html>');
  const body = dom.window.document.body;
  const ch   = [];

  // ── TYTUŁ ──
  ch.push(new Paragraph({
    children: [new TextRun({ text: title, bold: true, size: T, color: NAVY, font: FH })],
    spacing: { before: 0, after: 160 },
    border: { bottom: { color: GOLD, size: 6, style: BorderStyle.SINGLE, space: 3 } },
  }));

  // ── TABELE ──
  function makeTable(node) {
    const rows = Array.from(node.querySelectorAll('tr'));
    if (!rows.length) return null;
    const TW = 8870;
    const getW = n => {
      if (n===1) return [8870]; if (n===2) return [4435,4435];
      if (n===3) return [2957,2957,2956]; if (n===4) return [2218,2217,2218,2217];
      if (n===5) return [1774,1774,1774,1774,1774];
      const w=Math.floor(TW/n),ws=Array(n).fill(w); ws[n-1]=TW-w*(n-1); return ws;
    };
    const colCount = Array.from(rows[0].querySelectorAll('th,td')).reduce((s,c)=>s+parseInt(c.getAttribute('colspan')||1),0)||2;
    const colWidths = getW(colCount);
    const numCols = new Set();
    rows.filter(r=>r.querySelector('td')).forEach(row=>
      Array.from(row.querySelectorAll('td')).forEach((c,i)=>{ if(isNum(getText(c))) numCols.add(i); }));

    const dataRows = rows.filter(r=>r.querySelector('td'));
    const tRows = rows.map((row,rIdx)=>{
      const cells = Array.from(row.querySelectorAll('th,td'));
      const isHRow = !!row.querySelector('th');
      const dataIdx = dataRows.indexOf(row);
      const isLastData = dataIdx === dataRows.length - 1;
      const isEven = dataIdx % 2 === 1;
      return new TableRow({ tableHeader: isHRow, children: cells.map((cell,cIdx)=>{
        const isH = cell.tagName.toLowerCase()==='th';
        const txt = getText(cell).replace(/\s+/g,' ').trim();
        const cs  = parseInt(cell.getAttribute('colspan')||1);
        const isN = !isH && numCols.has(cIdx);
        const cw  = cs>1 ? TW : (colWidths[cIdx]||Math.floor(TW/colCount));
        // tło: nagłówek=granat, ostatni wiersz=lekki złoty, parzyste=lekki szary
        const fill = isH ? 'D6E0EC' : isLastData ? 'FDF6E7' : isEven ? 'F5F5F5' : 'FFFFFF';
        const textColor = isH ? NAVY : BLACK;
        return new TableCell({
          width: { size: cw, type: WidthType.DXA },
          columnSpan: cs>1 ? cs : undefined,
          children: [new Paragraph({
            children: [new TextRun({ text: txt, bold: isH||isLastData, size: S, color: textColor, font: FH })],
            alignment: isH||cs>1 ? AlignmentType.CENTER : isN ? AlignmentType.RIGHT : AlignmentType.LEFT,
            spacing: { before: 40, after: 40 },
          })],
          shading: { fill, type: ShadingType.CLEAR },
          margins: { top: 50, bottom: 50, left: 100, right: 100 },
          borders: {
            top:    { style: BorderStyle.SINGLE, size: isH?10:4, color: isH?NAVY:LINE_C },
            bottom: { style: BorderStyle.SINGLE, size: isH?10:4, color: isH?NAVY:LINE_C },
            left:   { style: BorderStyle.SINGLE, size: 4, color: LINE_C },
            right:  { style: BorderStyle.SINGLE, size: 4, color: LINE_C },
          },
        });
      })});
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
        children: [new TextRun({ text: raw, italics: true, size: B, color: GREY, font: FH })],
        spacing: { before: 0, after: 60 },
      }));
    } else if (tag==='p' && cls==='chapter-label') {
      if (raw) ch.push(new Paragraph({
        children: [new TextRun({ text: raw.toUpperCase(), bold: true, size: 24, color: GOLD, font: FH, characterSpacing: 80 })],
        spacing: { before: 200, after: 20 },
      }));
    } else if (tag==='h2') {
      if (raw) ch.push(new Paragraph({
        children: [new TextRun({ text: raw, bold: true, size: H2, color: NAVY, font: FH })],
        spacing: { before: 16, after: 80 },
        border: { bottom: { color: GOLD, size: 4, style: BorderStyle.SINGLE, space: 3 } },
      }));
    } else if (tag==='h3') {
      if (raw) ch.push(new Paragraph({
        children: [new TextRun({ text: raw, bold: true, size: H3, color: NAVY, font: FH })],
        spacing: { before: 120, after: 40 },
      }));
    } else if (tag==='h4'||tag==='h5'||tag==='h6') {
      if (raw) ch.push(new Paragraph({
        children: [new TextRun({ text: raw, bold: true, size: B, color: NAVY, font: FH })],
        spacing: { before: 80, after: 20 },
      }));
    } else if (tag==='div' && cls==='key-insight') {
      Array.from(node.childNodes).forEach(child=>{
        if (child.nodeType!==1) return;
        const cc=(child.getAttribute&&child.getAttribute('class'))||'';
        const ct=getText(child).replace(/\s+/g,' ').trim();
        if (!ct) return;
        if (cc==='key-label') {
          ch.push(new Paragraph({
            children: [new TextRun({ text: ct.toUpperCase(), bold: true, size: 18, color: GOLD, font: FH })],
            spacing: { before: 80, after: 10 }, indent: { left: 200 },
          }));
        } else {
          ch.push(new Paragraph({
            children: [new TextRun({ text: ct, size: B, color: BLACK, font: FB })],
            alignment: AlignmentType.BOTH,
            spacing: { before: 0, after: 20, line: LINE_H, lineRule: A },
            indent: { left: 200 },
            border: { left: { color: GOLD, size: 16, style: BorderStyle.SINGLE, space: 6 } },
          }));
        }
      });
      // brak pustego spacera — after:80 w ostatnim akapicie bloku
    } else if (tag==='p') {
      if (raw) ch.push(new Paragraph({
        children: [new TextRun({ text: raw, size: B, color: BLACK, font: FB })],
        alignment: AlignmentType.BOTH,
        spacing: { before: 0, after: 40, line: LINE_H, lineRule: A },
      }));
    } else if (tag==='blockquote') {
      if (raw) ch.push(new Paragraph({
        children: [new TextRun({ text: raw, italics: true, size: B, color: NAVY, font: FB })],
        alignment: AlignmentType.BOTH,
        indent: { left: 280, right: 280 },
        spacing: { before: 80, after: 80, line: LINE_H, lineRule: A },
        shading: { fill: 'F7F9FC', type: ShadingType.CLEAR },
        border: {
          top:    { color: LINE_C, size: 4, style: BorderStyle.SINGLE, space: 6 },
          bottom: { color: LINE_C, size: 4, style: BorderStyle.SINGLE, space: 6 },
          left:   { color: GOLD,  size: 16, style: BorderStyle.SINGLE, space: 8 },
          right:  { color: LINE_C, size: 4, style: BorderStyle.SINGLE, space: 6 },
        },
      }));
    } else if (tag==='ul'||tag==='ol') {
      const items = Array.from(node.querySelectorAll(':scope > li'));
      items.forEach((li,idx)=>{
        const lt = getText(li).replace(/\s+/g,' ').trim();
        const isLast = idx===items.length-1;
        if (lt) ch.push(new Paragraph({
          children: [new TextRun({ text: (tag==='ol'?(idx+1)+'. ':'–  ')+lt, size: B, color: BLACK, font: FB })],
          indent: { left: 360, hanging: 180 },
          spacing: { before: 0, after: isLast?60:20, line: LINE_H, lineRule: A },
        }));
      });
      // brak pustego spacera — ostatni li ma after:60
    } else if (tag==='table') {
      const tbl = makeTable(node);
      if (tbl) { ch.push(tbl); ch.push(new Paragraph({ text: '', spacing: { after: 60 } })); }
    } else if (tag==='hr') {
      ch.push(new Paragraph({
        children: [new TextRun({ text: '' })],
        border: { bottom: { color: LINE_C, size: 4, style: BorderStyle.SINGLE, space: 3 } },
        spacing: { before: 60, after: 60 },
      }));
    } else {
      Array.from(node.childNodes||[]).forEach(c=>{ if (c.nodeType===1) parse(c); });
    }
  }

  Array.from(body.childNodes).forEach(n=>{ if (n.nodeType===1) parse(n); });

  // ── NAGŁÓWEK STRONY ──
  const metaClean = meta
    .replace(/ · PORT LOTNICZY LUBLIN S\.A\./gi,'')
    .replace(/PORT LOTNICZY LUBLIN S\.A\. · /gi,'')
    .trim();

  const header = new Header({ children: [
    new Paragraph({
      children: [
        new TextRun({ text: 'PORT LOTNICZY LUBLIN S.A.', bold: true, size: F, color: GOLD, font: FH }),
        ...(metaClean ? [new TextRun({ text: '  ·  '+metaClean, size: F, color: GREY, font: FH })] : []),
      ],
      alignment: AlignmentType.RIGHT,
      spacing: { before: 0, after: 60 },
    }),
    new Paragraph({
      children: [new TextRun({ text: '' })],
      border: { bottom: { color: GOLD, size: 4, style: BorderStyle.SINGLE, space: 0 } },
      spacing: { before: 0, after: 0 },
    }),
  ]});

  // ── STOPKA Z NUMERAMI STRON ──
  const footer = new Footer({ children: [
    new Paragraph({
      children: [new TextRun({ text: '' })],
      border: { top: { color: LINE_C, size: 4, style: BorderStyle.SINGLE, space: 0 } },
      spacing: { before: 0, after: 16 },
    }),
    new Paragraph({
      children: [
        new SimpleField('PAGE', '1'),
        new TextRun({ text: '/', size: F, color: GREY, font: FH }),
        new SimpleField('NUMPAGES', '1'),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 10 },
    }),
    new Paragraph({
      children: [new TextRun({ text: 'Dokument zawiera informacje poufne w rozumieniu art. 11 ustawy z dnia 16 kwietnia 1993 r. o zwalczaniu nieuczciwej konkurencji. Przeznaczony wyłącznie dla adresata. Nieuprawnione ujawnienie, kopiowanie lub rozpowszechnianie jest zabronione i może stanowić podstawę odpowiedzialności prawnej.', size: 14, color: 'CC0000', italics: true, font: FH })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0 },
    }),
  ]});

  // ── DOKUMENT ──
  return await Packer.toBuffer(new Document({
    creator:'', description:'', title:'', subject:'', keywords:'', lastModifiedBy:'', revision:1,
    features: { updateFields: true },
    sections: [{
      properties: { page: {
        margin: { top:1440, right:1584, bottom:1440, left:1584, header:576, footer:576, gutter:0 },
        size: { width:11906, height:16838, orientation:'portrait' },
      }},
      headers: { default: header },
      footers: { default: footer },
      children: ch,
    }],
  }));
}

function corsHeaders() {
  return { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "Content-Type", "Access-Control-Allow-Methods": "POST, OPTIONS", "Content-Type": "application/json" };
}
