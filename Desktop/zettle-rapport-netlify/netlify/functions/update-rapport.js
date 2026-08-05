'use strict';

// ─── Minimale XLSX builder (geen externe dependencies) ───────────────────────
const zlib = require('zlib');

function xlsxBuild(sheets) {
  // sheets = [{ name, rows }]
  // rows = [[cel, ...], ...] where cel = { v, t } or primitive
  // t: 's'=string, 'n'=number, 'b'=bool

  const files = {};

  // Shared strings
  const sharedStrings = [];
  const ssMap = {};
  function addSS(s) {
    const key = String(s);
    if (ssMap[key] === undefined) { ssMap[key] = sharedStrings.length; sharedStrings.push(key); }
    return ssMap[key];
  }

  function escXml(s) {
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  function colName(n) {
    let s = '';
    while (n >= 0) { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; }
    return s;
  }

  // Nummerformaten
  const NUM_FMTS = {
    'int':  { id: 164, fmt: '0' },
    'eur':  { id: 165, fmt: '#,##0.00' },
  };

  // Werkbladen
  const sheetXmls = [];
  const sheetRels = [];

  sheets.forEach(({ name, rows, colWidths }, si) => {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';

    // Kolombreedtes
    if (colWidths && colWidths.length) {
      xml += '<cols>';
      colWidths.forEach((w, i) => {
        if (w) xml += `<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`;
      });
      xml += '</cols>';
    }

    xml += '<sheetData>';
    rows.forEach((row, ri) => {
      if (!row || row.length === 0) { xml += `<row r="${ri+1}"/>`; return; }
      xml += `<row r="${ri+1}">`;
      row.forEach((cel, ci) => {
        if (cel === null || cel === undefined) return;
        const ref = colName(ci) + (ri + 1);
        let v = cel, t = null, fmt = null;
        if (typeof cel === 'object' && cel !== null && 'v' in cel) {
          v = cel.v; t = cel.t; fmt = cel.fmt;
        }
        if (v === null || v === undefined || v === '') return;

        if (fmt === 'eur' || (typeof v === 'number' && fmt !== 'int')) {
          const fmtId = fmt === 'int' ? NUM_FMTS.int.id : NUM_FMTS.eur.id;
          xml += `<c r="${ref}" s="${fmtId === 164 ? 1 : 2}"><v>${escXml(v)}</v></c>`;
        } else if (typeof v === 'number') {
          xml += `<c r="${ref}" s="1"><v>${escXml(v)}</v></c>`;
        } else {
          const si2 = addSS(v);
          xml += `<c r="${ref}" t="s"><v>${si2}</v></c>`;
        }
      });
      xml += '</row>';
    });
    xml += '</sheetData></worksheet>';

    files[`xl/worksheets/sheet${si+1}.xml`] = Buffer.from(xml, 'utf8');
    sheetXmls.push({ name, idx: si+1 });
    sheetRels.push(`<Relationship Id="rId${si+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${si+1}.xml"/>`);
  });

  // Shared strings XML
  const ssXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">
${sharedStrings.map(s => `<si><t xml:space="preserve">${escXml(s)}</t></si>`).join('\n')}
</sst>`;
  files['xl/sharedStrings.xml'] = Buffer.from(ssXml, 'utf8');

  // Styles (minimaal — int en eur formaten)
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2">
<numFmt numFmtId="164" formatCode="0"/>
<numFmt numFmtId="165" formatCode="#,##0.00"/>
</numFmts>
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
</styleSheet>`;
  files['xl/styles.xml'] = Buffer.from(stylesXml, 'utf8');

  // Workbook
  const wbXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
${sheetXmls.map((s,i) => `<sheet name="${escXml(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('\n')}
</sheets>
</workbook>`;
  files['xl/workbook.xml'] = Buffer.from(wbXml, 'utf8');

  // Relationships
  files['xl/_rels/workbook.xml.rels'] = Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheetRels.join('\n')}
<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
<Relationship Id="rId${sheets.length+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`, 'utf8');

  files['_rels/.rels'] = Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`, 'utf8');

  files['[Content_Types].xml'] = Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
${sheetXmls.map((_,i) => `<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>`).join('\n')}
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`, 'utf8');

  // ZIP bouwen
  return buildZip(files);
}

function buildZip(files) {
  const parts = [];
  const centralDir = [];
  let offset = 0;

  function crc32(buf) {
    let crc = 0xFFFFFFFF;
    const table = crc32.table || (crc32.table = (() => {
      const t = new Uint32Array(256);
      for (let i = 0; i < 256; i++) {
        let c = i;
        for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
        t[i] = c;
      }
      return t;
    })());
    for (let i = 0; i < buf.length; i++) crc = table[(crc ^ buf[i]) & 0xFF] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  function w16(n) { const b = Buffer.alloc(2); b.writeUInt16LE(n, 0); return b; }
  function w32(n) { const b = Buffer.alloc(4); b.writeUInt32LE(n >>> 0, 0); return b; }

  for (const [name, data] of Object.entries(files)) {
    const compressed = zlib.deflateRawSync(data, { level: 6 });
    const crc = crc32(data);
    const nameBytes = Buffer.from(name, 'utf8');

    const localHeader = Buffer.concat([
      Buffer.from([0x50,0x4B,0x03,0x04]),
      w16(20), w16(0), w16(8),
      w16(0), w16(0),
      w32(crc), w32(compressed.length), w32(data.length),
      w16(nameBytes.length), w16(0),
      nameBytes, compressed
    ]);

    centralDir.push(Buffer.concat([
      Buffer.from([0x50,0x4B,0x01,0x02]),
      w16(20), w16(20), w16(0), w16(8),
      w16(0), w16(0),
      w32(crc), w32(compressed.length), w32(data.length),
      w16(nameBytes.length), w16(0), w16(0), w16(0), w16(0),
      w32(0), w32(offset),
      nameBytes
    ]));

    parts.push(localHeader);
    offset += localHeader.length;
  }

  const cdBuf = Buffer.concat(centralDir);
  const eocd = Buffer.concat([
    Buffer.from([0x50,0x4B,0x05,0x06]),
    w16(0), w16(0),
    w16(centralDir.length), w16(centralDir.length),
    w32(cdBuf.length), w32(offset),
    w16(0)
  ]);

  return Buffer.concat([...parts, cdBuf, eocd]);
}


// ─── Dependencies ─────────────────────────────────────────────────────────────

// ─── Graph auth — directe HTTP, geen @azure/identity nodig ───────────────────
async function getGraphToken() {
  const url = `https://login.microsoftonline.com/${process.env.GRAPH_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type:    'client_credentials',
    client_id:     process.env.GRAPH_CLIENT_ID,
    client_secret: process.env.GRAPH_CLIENT_SECRET,
    scope:         'https://graph.microsoft.com/.default',
  });
  const res = await fetch(url, { method: 'POST', body });
  if (!res.ok) throw new Error(`Graph token fout: ${res.status} ${await res.text()}`);
  const data = await res.json();
  return data.access_token;
}

async function graphGet(token, path) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!res.ok) throw new Error(`Graph GET ${path}: ${res.status} ${await res.text()}`);
  return res.json();
}

async function graphPost(token, path, body) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error(`Graph POST ${path}: ${res.status} ${await res.text()}`);
  return res.status === 202 ? null : res.json().catch(() => null);
}

async function graphPut(token, path, buffer, contentType) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: 'PUT',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': contentType },
    body: buffer
  });
  if (!res.ok) throw new Error(`Graph PUT ${path}: ${res.status} ${await res.text()}`);
  return res.json();
}

// ─── Zettle helpers ───────────────────────────────────────────────────────────
async function getZettleToken() {
  // Haal verse token op via JWT assertion (geldig 2 uur)
  const jwt = process.env.ZETTLE_JWT_ASSERTION;
  const clientId = process.env.ZETTLE_CLIENT_ID;
  if (jwt && clientId) {
    const res = await fetch('https://oauth.zettle.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        client_id: clientId,
        assertion: jwt
      })
    });
    if (res.ok) {
      const data = await res.json();
      if (data.access_token) return data.access_token;
    }
  }
  // Fallback op handmatig ingestelde token
  return process.env.ZETTLE_API_TOKEN;
}

async function zettleGet(path) {
  const token = await getZettleToken();
  const res = await fetch(`https://purchase.izettle.com${path}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!res.ok) throw new Error(`Zettle ${path} -> ${res.status}: ${await res.text()}`);
  return res.json();
}

async function fetchAllPurchases(startDate, endDate) {
  // Zettle API gebruikt exclusieve einddatum — voeg 1 dag toe
  const endDateExcl = new Date(new Date(endDate).getTime() + 86400000).toISOString().slice(0, 10);
  const purchases = [];
  let lastHash = null;
  while (true) {
    const qs = new URLSearchParams({ startDate, endDate: endDateExcl, limit: '1000' });
    if (lastHash) qs.set('lastPurchaseHash', lastHash);
    const data = await zettleGet(`/purchases/v2?${qs}`);
    const batch = data.purchases || [];
    purchases.push(...batch);
    if (batch.length < 1000 || !data.lastPurchaseHash) break;
    lastHash = data.lastPurchaseHash;
  }
  return purchases;
}

// ─── Categorieën (exacte Zettle namen) ───────────────────────────────────────
const CATEGORIEEN = [
  'Reguliere vaart',
  'Reguliere vaart (fiets)',
  'VVV',
  'Kaasvaart',
  'Speciale vaart',
  'Lustrumboek',
  'Vlaggetje',
  'Fooi',
];

// ─── Bereken rapportdata ──────────────────────────────────────────────────────
function buildReportData(purchases, startDate, endDate) {
  const MINOR = 100;
  let totaleOmzet = 0, aantalVerkopen = 0;
  let cardTotal = 0, cardFee = 0, cashTotal = 0;

  const perCategorie = {};
  CATEGORIEEN.forEach(c => perCategorie[c] = { aantal: 0, omzet: 0, retouren: 0 });
  perCategorie['Overig'] = { aantal: 0, omzet: 0, retouren: 0 };

  const perLijn = {};
  const perHalte = {}; // halteplaats → { aantal, omzet }
  const details = [];
  const detailsMap = {};

  for (const p of purchases) {
    const isRefund = p.amount < 0;
    const amountEur = p.amount / MINOR;
    totaleOmzet += amountEur;
    if (!isRefund) aantalVerkopen++;

    for (const pay of (p.payments || [])) {
      const payEur = pay.amount / MINOR;
      if (pay.type === 'IZETTLE_CARD' || pay.type === 'IZETTLE_CARD_ONLINE') {
        cardTotal += payEur;
        cardFee += (pay.attributes?.serviceFeeAmount || 0) / MINOR;
      } else if (pay.type === 'IZETTLE_CASH') {
        cashTotal += payEur;
      }
    }

    const lijnRaw = p.userDisplayName || p.userId || 'Onbekend';
    const lijn = lijnRaw.includes('Loosdrecht') ? 'Lijndienst Loosdrecht'
               : lijnRaw.includes('Vecht')      ? 'Lijndienst Vecht'
               : String(lijnRaw);
    if (!perLijn[lijn]) perLijn[lijn] = { omzet: 0, aantal: 0, perCat: {} };
    perLijn[lijn].omzet += amountEur;
    perLijn[lijn].aantal += (isRefund ? -1 : 1);

    for (const item of (p.products || [])) {
      const qty = parseInt(item.quantity, 10) || 1;
      const lineEur = (item.unitPrice * qty) / MINOR;
      const cat = item.category?.name || 'Overig';
      // Vertaal oude Zettle categorienamen naar de nieuwe namen
      // Map oude categorie + productnamen naar nieuwe categorieën
      const CATEGORIE_MAP = {
        'Kaartverkoop': 'Reguliere vaart',
        'Artikelen':    'Lustrumboek',
      };
      // Uitzonderingen op productnaam (overschrijven de categorie mapping)
      const PRODUCT_MAP = {
        'Fiets':               'Reguliere vaart (fiets)',
        'Fietsboot vlag (FB1)':'Vlaggetje',
        'Fooi Vechtlijn':      'Fooi',
        'Fooi Loosdrecht':     'Fooi',
        'specials':            'Fooi',
        'Speciale Vaart':      'Speciale vaart',
        'Vroege vogeltocht':   'VVV',
        'Kaasvaart':           'Kaasvaart',
      };
      const catKey = PRODUCT_MAP[item.name]
        || (CATEGORIEEN.includes(cat) ? cat : (CATEGORIE_MAP[cat] || 'Overig'));

      if (qty < 0) {
        perCategorie[catKey].retouren += Math.abs(qty);
        perCategorie[catKey].omzet += lineEur;
      } else {
        perCategorie[catKey].aantal += qty;
        perCategorie[catKey].omzet += lineEur;
      }

      // Per lijn per categorie
      if (!perLijn[lijn].perCat[catKey]) perLijn[lijn].perCat[catKey] = { aantal: 0, omzet: 0 };
      if (qty < 0) { perLijn[lijn].perCat[catKey].omzet += lineEur; }
      else { perLijn[lijn].perCat[catKey].aantal += qty; perLijn[lijn].perCat[catKey].omzet += lineEur; }


      // Halteplaats: productnaam binnen Reguliere vaart
      if (catKey === 'Reguliere vaart' && item.name) {
        if (!perHalte[item.name]) perHalte[item.name] = { aantal: 0, omzet: 0 };
        if (qty > 0) { perHalte[item.name].aantal += qty; perHalte[item.name].omzet += lineEur; }
        else { perHalte[item.name].omzet += lineEur; }
      }
      const detKey = `${catKey}||${item.name}||${item.variantName || ''}`;
      if (!detailsMap[detKey]) {
        detailsMap[detKey] = details.length;
        details.push({ categorie: catKey, product: item.name, variant: item.variantName || '-', aantal: 0, retouren: 0, omzet: 0 });
      }
      const d = details[detailsMap[detKey]];
      if (qty < 0) { d.retouren += Math.abs(qty); d.omzet += lineEur; }
      else { d.aantal += qty; d.omzet += lineEur; }
    }
  }

  const catOrder = [...CATEGORIEEN, 'Overig'];
  details.sort((a, b) => {
    const ci = catOrder.indexOf(a.categorie) - catOrder.indexOf(b.categorie);
    return ci !== 0 ? ci : b.omzet - a.omzet;
  });


  // Sorteer haltes op omzet (aflopend)
  const perHalteSorted = Object.entries(perHalte)
    .sort((a, b) => b[1].omzet - a[1].omzet)
    .map(([naam, d]) => ({ naam, aantal: d.aantal, omzet: d.omzet, gemOmzet: d.aantal > 0 ? d.omzet / d.aantal : 0 }));
  return {
    periode: { start: startDate, end: endDate },
    CATEGORIEEN_VOLGORDE: CATEGORIEEN,
    totaleOmzet, aantalVerkopen,
    cardTotal, cardFee, cardNetto: cardTotal + cardFee,
    cashTotal, cashNetto: cashTotal,
    perCategorie,
    perHalte: perHalteSorted,
    perLijn: Object.entries(perLijn).map(([naam, d]) => ({
      naam, omzet: d.omzet, aantal: d.aantal,
      gemBon: d.aantal > 0 ? d.omzet / d.aantal : 0,
      perCat: d.perCat || {}
    })),
    details,
  };
}

// ─── Excel ────────────────────────────────────────────────────────────────────
const KVK = '65284445';
const DONKER   = '006EA8';
const MIDDEL   = '008AD1';
const LICHT    = 'CCE9F7';
const GRIJS1   = 'F2F2F2';
const GRIJS2   = 'FFFFFF';
const GRIJSTOT = 'D9D9D9';
const ROOD     = 'C00000';

function nlDate(s) {
  return new Date(s).toLocaleDateString('nl-NL', { day: '2-digit', month: '2-digit', year: 'numeric' });
}
function maandNaam(s) {
  return new Date(s).toLocaleDateString('nl-NL', { month: 'long', year: 'numeric' }).replace(/^\w/, c => c.toUpperCase());
}
function eur(n) { return Math.round(n * 100) / 100; }
// Formatteer als getal met altijd 2 decimalen (als JS number)
function eur(n) { return parseFloat((Math.round(n * 100) / 100).toFixed(2)); }
function int(n) { return parseInt(n, 10) || 0; }

function cs(fg, bg, bold, sz, align, numFmt) {
  return {
    font: { name: 'Calibri', sz: sz || 10, bold: !!bold, color: { rgb: fg || '000000' } },
    fill: bg ? { patternType: 'solid', fgColor: { rgb: bg } } : undefined,
    alignment: { horizontal: align || 'left', vertical: 'center' },
    numFmt: numFmt || 'General'
  };
}

function applyStyles(ws, styles) {
  Object.entries(styles).forEach(([key, style]) => {
    const [row, col] = key.split(',').map(Number);
    const ref = XLSX.utils.encode_cell({ r: row, c: col });
    if (!ws[ref]) ws[ref] = { v: '', t: 's' };
    ws[ref].s = style;
    if (typeof ws[ref].v === 'number') {
      ws[ref].t = 'n';
      if (style.numFmt) ws[ref].z = style.numFmt;
    }
  });
}

function buildXlsx(data) {
  const periodeStr = `Periode: ${nlDate(data.periode.start)} t/m ${nlDate(data.periode.end)}  |  KvK: ${KVK}`;
  const maand = maandNaam(data.periode.start);

  function n(v) { return { v, t: 'n', fmt: 'eur' }; }
  function i(v) { return { v: parseInt(v)||0, t: 'n', fmt: 'int' }; }

  // ── Sheet 1: Samenvatting ──
  const s1rows = [
    ['', 'Verkooprapport - de Fietsboot'],
    ['', periodeStr],
    [],
    ['', 'Totale omzet', '# Transacties', 'Kaartbetalingen', 'Contant'],
    ['', n(eur(data.totaleOmzet)), i(data.aantalVerkopen), n(eur(data.cardTotal)), n(eur(data.cashTotal))],
    [],
    [],
    ['', 'Verkopen per categorie'],
    ['', 'Categorie', 'Aantal/Passagiers', 'Retouren', 'Omzet (euro)'],
  ];

  let catTotAantal = 0, catTotRet = 0, catTotOmzet = 0;
  CATEGORIEEN.forEach(cat => {
    const d = data.perCategorie[cat] || { aantal: 0, retouren: 0, omzet: 0 };
    s1rows.push(['', cat, i(d.aantal), i(d.retouren || 0), n(eur(d.omzet))]);
    catTotAantal += d.aantal; catTotRet += d.retouren; catTotOmzet += d.omzet;
  });
  s1rows.push(['', 'Totaal', i(catTotAantal), i(catTotRet), n(eur(catTotOmzet))]);
  s1rows.push([], []);
  s1rows.push(['', 'Betalingen en kosten']);
  s1rows.push(['', 'Methode', 'Bedrag', 'Toeslag', 'Netto']);
  s1rows.push(['', 'Kaart (reader)', n(eur(data.cardTotal)), n(eur(data.cardFee)), n(eur(data.cardNetto))]);
  s1rows.push(['', 'Contant', n(eur(data.cashTotal)), n(0), n(eur(data.cashNetto))]);

  // ── Haltetabel (Reguliere vaart) ──
  if (data.perHalte && data.perHalte.length > 0) {
    s1rows.push([], []);
    s1rows.push(['', 'Passagiers per halteplaats (Reguliere vaart)']);
    s1rows.push(['', 'Halteplaats', 'Aantal/Passagiers', 'Omzet (euro)', 'Gem. omzet/passagier']);
    let hTotAantal = 0, hTotOmzet = 0;
    data.perHalte.forEach(h => {
      s1rows.push(['', h.naam, i(h.aantal), n(eur(h.omzet)), n(eur(h.gemOmzet))]);
      hTotAantal += h.aantal; hTotOmzet += h.omzet;
    });
    const hGem = hTotAantal > 0 ? hTotOmzet / hTotAantal : 0;
    s1rows.push(['', 'Totaal', i(hTotAantal), n(eur(hTotOmzet)), n(eur(hGem))]);
  }

  // ── Sheet 2: Lijnen analyse ──
  const s2rows = [
    ['', `Analyse per lijndienst - ${maand}`],
    ['', periodeStr],
    [],
    ['', 'Categorie', 'Lijndienst Vecht', '', 'Lijndienst Loosdrecht', ''],
    ['', '',          'Aantal/Passagiers', 'Omzet', 'Aantal/Passagiers', 'Omzet'],
  ];
  const lijnData = {};
  ['Lijndienst Vecht','Lijndienst Loosdrecht'].forEach(l => {
    lijnData[l] = (data.perLijn.find(x => x.naam === l) || { perCat: {} }).perCat;
  });
  let tVa=0, tVo=0, tLa=0, tLo=0;
  [...CATEGORIEEN, 'Overig'].forEach(cat => {
    const v = (lijnData['Lijndienst Vecht']      || {})[cat] || { aantal: 0, omzet: 0 };
    const l = (lijnData['Lijndienst Loosdrecht'] || {})[cat] || { aantal: 0, omzet: 0 };
    if (v.aantal===0 && v.omzet===0 && l.aantal===0 && l.omzet===0) return;
    s2rows.push(['', cat, i(v.aantal), n(eur(v.omzet)), i(l.aantal), n(eur(l.omzet))]);
    tVa+=v.aantal; tVo+=v.omzet; tLa+=l.aantal; tLo+=l.omzet;
  });
  s2rows.push(['', 'Totaal', i(tVa), n(eur(tVo)), i(tLa), n(eur(tLo))]);

  // ── Sheet 3: Details ──
  const merged = {};
  data.details.forEach(d => {
    const key = `${d.categorie}||${d.product}`;
    if (!merged[key]) merged[key] = { categorie: d.categorie, product: d.product, aantal: 0, omzet: 0 };
    merged[key].aantal += d.aantal; merged[key].omzet += d.omzet;
  });
  const mergedRows = Object.values(merged).sort((a,b) => {
    const co = [...CATEGORIEEN,'Overig'];
    const ci = co.indexOf(a.categorie) - co.indexOf(b.categorie);
    return ci !== 0 ? ci : b.omzet - a.omzet;
  });
  const s3rows = [
    ['', `Verkoopdetails - ${maand}`],
    ['', periodeStr],
    [],
    ['', 'Product', 'Categorie', 'Aantal/Passagiers', 'Omzet (euro)'],
  ];
  mergedRows.forEach(d => {
    s3rows.push(['', d.product, d.categorie, i(d.aantal), n(eur(d.omzet))]);
  });

  return xlsxBuild([
    { name: 'Samenvatting',   rows: s1rows, colWidths: [2,28,18,16,20] },
    { name: 'Lijnen analyse', rows: s2rows, colWidths: [2,26,18,14,18,14] },
    { name: 'Details',        rows: s3rows, colWidths: [2,28,24,18,14] },
  ]);
}


// ─── OneDrive upload ──────────────────────────────────────────────────────────
async function uploadToOneDrive(token, buffer, filePath) {
  const user = process.env.GRAPH_ONEDRIVE_USER;
  const encoded = filePath.split('/').map(encodeURIComponent).join('/');
  const uploaded = await graphPut(
    token,
    `/users/${user}/drive/root:/${encoded}:/content`,
    buffer,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );
  try {
    const share = await graphPost(token, `/users/${user}/drive/items/${uploaded.id}/createLink`,
      { type: 'view', scope: 'organization' });
    return share?.link?.webUrl || null;
  } catch { return uploaded?.webUrl || null; }
}


// ─── PNG tabel generator via htmlcsstoimage.com ───────────────────────────────
function buildTabelHtml(data, startDate, endDate) {
  function nlDatum(s) {
    const d = new Date(s);
    return `${String(d.getDate()).padStart(2,'0')}-${String(d.getMonth()+1).padStart(2,'0')}-${d.getFullYear()}`;
  }
  const cat = data.CATEGORIEEN_VOLGORDE || [
    'Reguliere vaart','Reguliere vaart (fiets)','VVV','Kaasvaart',
    'Speciale vaart','Lustrumboek','Vlaggetje','Extra'
  ];
  const rijen = cat.map((naam, i) => {
    const d = data.perCategorie[naam] || { aantal: 0, omzet: 0 };
    const bg = i % 2 === 0 ? '#f5f8fa' : '#ffffff';
    return `<tr bgcolor="${bg}">
      <td style="padding:8px 14px;color:#1a1d23;border-bottom:1px solid #e4e7ed;font-size:14px">${naam}</td>
      <td align="right" style="padding:8px 28px 8px 0;color:#1a1d23;border-bottom:1px solid #e4e7ed;font-size:14px">${d.aantal}</td>
      <td align="right" style="padding:8px 14px;color:#1a1d23;border-bottom:1px solid #e4e7ed;font-size:14px">&euro; ${(Math.round(d.omzet*100)/100).toFixed(2)}</td>
    </tr>`;
  }).join('');
  const totaal = cat.reduce((a, naam) => a + (data.perCategorie[naam]?.omzet || 0), 0);
  return `<html><body style="margin:0;padding:0;background:white;font-family:Arial,Helvetica,sans-serif">
  <table width="520" cellpadding="0" cellspacing="0" style="border-collapse:collapse">
    <tr bgcolor="#008AD1">
      <th width="45%" align="left" style="padding:11px 14px;color:#fff;font-size:13px;font-weight:600;vertical-align:top">&nbsp;</th>
      <th width="27%" align="right" style="padding:11px 14px;color:#fff;font-size:13px;font-weight:600;vertical-align:top;line-height:1.4">Passagiers<br><span style="font-weight:400;font-size:11px;opacity:0.85">Aantal</span></th>
      <th width="28%" align="right" style="padding:11px 14px;color:#fff;font-size:13px;font-weight:600;vertical-align:top">Omzet</th>
    </tr>
    ${rijen}
    <tr><td colspan="3" style="padding:3px 0;font-size:1px;line-height:1px;background:#fff">&nbsp;</td></tr>
    <tr bgcolor="#008AD1">
      <td style="padding:10px 14px;color:#fff;font-weight:700;font-size:14px">Totaal</td>
      <td style="padding:10px 14px"></td>
      <td align="right" style="padding:10px 14px;color:#fff;font-weight:700;font-size:14px">&euro; ${(Math.round(totaal*100)/100).toFixed(2)}</td>
    </tr>
  </table>
  <p style="margin:6px 0 0;font-size:11px;color:#9aa0ad;text-align:center;font-family:Arial">Verkoopcijfers de Fietsboot &ndash; ${nlDatum(startDate)} t/m ${nlDatum(endDate)}</p>
  </body></html>`;
}

async function buildPngBijlage(data, startDate, endDate) {
  const userId  = process.env.HCTI_USER_ID;
  const apiKey  = process.env.HCTI_API_KEY;
  if (!userId || !apiKey) return null;

  try {
    const res = await fetch('https://hcti.io/v1/image', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Basic ' + Buffer.from(`${userId}:${apiKey}`).toString('base64')
      },
      body: JSON.stringify({
        html: buildTabelHtml(data, startDate, endDate),
        viewport_width: 520,
        device_scale_factor: 2
      })
    });
    if (!res.ok) { console.error('HCTI fout:', res.status); return null; }
    const { url } = await res.json();
    // Download de PNG
    const imgRes = await fetch(url);
    if (!imgRes.ok) return null;
    const buf = await imgRes.arrayBuffer();
    return Buffer.from(buf);
  } catch (err) {
    console.error('PNG generatie mislukt:', err.message);
    return null;
  }
}

// ─── E-mail ───────────────────────────────────────────────────────────────────
async function stuurEmail(token, { ontvanger, startDate, endDate, data, buffer, onedriveUrl }) {
  const maand = maandNaam(startDate);
  const bestandsnaam = `Zettle_Verkooprapport_${startDate}_tm_${endDate}.xlsx`;
  const dagen = (new Date(endDate) - new Date(startDate)) / 86400000;
  const type = dagen >= 27 ? 'Maandrapport' : 'Weekrapport';
  const datumRegel = `${nlDate(startDate)} t/m ${nlDate(endDate)} · Bijgewerkt ${new Date().toLocaleDateString('nl-NL', { day: 'numeric', month: 'long', year: 'numeric' })}`;
  const totaalOmzet = CATEGORIEEN.reduce((a, cat) => a + ((data.perCategorie[cat] || {}).omzet || 0), 0);
  const catRows = [...CATEGORIEEN, 'Overig'].map((cat, i) => {
    const d = data.perCategorie[cat] || { aantal: 0, retouren: 0, omzet: 0 };
    if (cat === 'Overig') return ''; // Overig altijd verbergen in email, wel meegerekend in totaal
    const bg = i % 2 === 0 ? '#f5f7fa' : '#ffffff';
    return `<tr bgcolor="${bg}"><td style="padding:9px 16px;font-family:Arial,sans-serif;font-size:14px;color:#1a1d23;border-bottom:1px solid #e8eaed">${cat}</td><td align="right" style="padding:9px 16px;font-family:Arial,sans-serif;font-size:14px;color:#1a1d23;border-bottom:1px solid #e8eaed">${d.aantal}</td><td align="right" style="padding:9px 16px;font-family:Arial,sans-serif;font-size:14px;color:#1a1d23;border-bottom:1px solid #e8eaed">&euro; ${eur(d.omzet).toFixed(2)}</td></tr>`;
  }).join('');
  const html = `<!DOCTYPE html><html><body style="margin:0;padding:0;background:#fff;font-family:Arial,sans-serif"><table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td align="center" style="padding:24px 16px"><table width="560" cellpadding="0" cellspacing="0" border="0"><tr><td><table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse"><tr bgcolor="#008AD1"><th align="left" style="padding:11px 16px;font-family:Arial,sans-serif;font-size:14px;color:#fff;font-weight:700">Categorie</th><th align="right" style="padding:11px 16px;font-family:Arial,sans-serif;font-size:14px;color:#fff;font-weight:700">Aantal/Passagiers</th><th align="right" style="padding:11px 16px;font-family:Arial,sans-serif;font-size:14px;color:#fff;font-weight:700">Omzet</th></tr>${catRows}<tr bgcolor="#CCE9F7"><td style="padding:10px 16px;font-family:Arial,sans-serif;font-size:14px;color:#006EA8;font-weight:700">Totaal</td><td></td><td align="right" style="padding:10px 16px;font-family:Arial,sans-serif;font-size:14px;color:#006EA8;font-weight:700">&euro; ${eur(totaalOmzet).toFixed(2)}</td></tr></table><p style="margin:8px 0 0;font-family:Arial,sans-serif;font-size:12px;color:#999;text-align:center">${datumRegel}</p></td></tr><tr><td style="padding:20px 0 0 0"><p style="margin:0;font-family:Arial,sans-serif;font-size:13px;color:#444">Het volledige rapport is bijgevoegd als Excel-bestand.${onedriveUrl ? ` Je kunt het ook <a href="${onedriveUrl}" style="color:#006EA8">openen via OneDrive</a>.` : ''}</p></td></tr></table></td></tr></table></body></html>`;
  await graphPost(token, `/users/${process.env.GRAPH_MAIL_FROM}/sendMail`, {
    message: {
      subject: `Zettle ${type} ${maand} (${nlDate(startDate)} - ${nlDate(endDate)})`,
      body: { contentType: 'HTML', content: html },
      toRecipients: [{ emailAddress: { address: ontvanger } }],
      attachments: [{ '@odata.type': '#microsoft.graph.fileAttachment', name: bestandsnaam, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', contentBytes: Buffer.from(buffer).toString('base64') }]
    },
    saveToSentItems: true
  });
}

// ─── Netlify Function handler ─────────────────────────────────────────────────
exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: JSON.stringify({ ok: false, fout: 'Method not allowed' }) };
  }
  let body;
  try { body = JSON.parse(event.body || '{}'); }
  catch { return { statusCode: 400, body: JSON.stringify({ ok: false, fout: 'Ongeldige JSON' }) }; }

  const { startDate, endDate, ontvanger } = body;
  if (!startDate || !endDate) return { statusCode: 400, body: JSON.stringify({ ok: false, fout: 'startDate en endDate zijn verplicht' }) };
  if (!ontvanger) return { statusCode: 400, body: JSON.stringify({ ok: false, fout: 'ontvanger is verplicht' }) };

  try {
    const purchases = await fetchAllPurchases(startDate, endDate);
    const reportData = buildReportData(purchases, startDate, endDate);
    const buffer = buildXlsx(reportData);
    const token = await getGraphToken();
    const targetPath = process.env.ZETTLE_RAPPORT_PATH || 'MS365/Zettle Rapporten/Zettle_Verkooprapport_Actueel.xlsx';
    const onedriveUrl = await uploadToOneDrive(token, buffer, targetPath);
    await stuurEmail(token, { ontvanger, startDate, endDate, data: reportData, buffer, onedriveUrl });

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        ok: true,
        periode: { startDate, endDate },
        aantalAankopen: purchases.length,
        totaleOmzet: reportData.totaleOmzet,
        ontvangerEmail: ontvanger,
        onedriveUrl,
        bericht: `Rapport bijgewerkt en verzonden naar ${ontvanger}`,
        perCategorie: reportData.perCategorie
      })
    };
  } catch (err) {
    console.error('Fout:', err.message, err.stack);
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ok: false, fout: err.message })
    };
  }
};
