'use strict';
const { getGraphClient } = require('./shared/graph');
const ExcelJS = require('exceljs');

// ─── Zettle helpers ───────────────────────────────────────────────────────────

async function zettleGet(path) {
  const token = process.env.ZETTLE_API_TOKEN;
  const res = await fetch(`https://purchase.izettle.com${path}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!res.ok) throw new Error(`Zettle ${path} → ${res.status}: ${await res.text()}`);
  return res.json();
}

async function fetchAllPurchases(startDate, endDate) {
  const purchases = [];
  let lastHash = null;
  while (true) {
    const qs = new URLSearchParams({ startDate, endDate, limit: '1000' });
    if (lastHash) qs.set('lastPurchaseHash', lastHash);
    const data = await zettleGet(`/purchases/v2?${qs}`);
    const batch = data.purchases || [];
    purchases.push(...batch);
    if (batch.length < 1000 || !data.lastPurchaseHash) break;
    lastHash = data.lastPurchaseHash;
  }
  return purchases;
}

// ─── Bereken rapportdata ──────────────────────────────────────────────────────

function buildReportData(purchases, startDate, endDate) {
  const MINOR = 100;
  let totalIncl = 0, totalExcl = 0, totalVat = 0, totalRefunds = 0;
  let transactionCount = 0;
  let cardTotal = 0, cardFee = 0, cashTotal = 0;
  const perStaff = {};
  const perProduct = {};
  const cadeaubonnen = { aantal: 0, bedrag: 0 };

  for (const p of purchases) {
    const isRefund = p.amount < 0;
    const amountEur = p.amount / MINOR;
    const vatEur = (p.vatAmount || 0) / MINOR;

    if (isRefund) {
      totalRefunds += amountEur;
    } else {
      totalExcl += (amountEur - vatEur);
      totalVat += vatEur;
      totalIncl += amountEur;
      transactionCount++;
    }

    for (const pay of (p.payments || [])) {
      const payEur = pay.amount / MINOR;
      if (pay.type === 'IZETTLE_CARD' || pay.type === 'IZETTLE_CARD_ONLINE') {
        cardTotal += payEur;
        cardFee += (pay.attributes?.serviceFeeAmount || 0) / MINOR;
      } else if (pay.type === 'IZETTLE_CASH') {
        cashTotal += payEur;
      }
    }

    const staffName = p.userDisplayName || p.userId || 'Onbekend';
    if (!perStaff[staffName]) perStaff[staffName] = { omzet: 0, aantal: 0 };
    perStaff[staffName].omzet += amountEur;
    perStaff[staffName].aantal += (isRefund ? -1 : 1);

    for (const item of (p.products || [])) {
      const key = `${item.name}||${item.variantName || '—'}`;
      if (!perProduct[key]) {
        perProduct[key] = { product: item.name, variant: item.variantName || '—', verkocht: 0, geretourneerd: 0, omzet: 0 };
      }
      const qty = item.quantity || 1;
      const lineEur = (item.unitPrice * qty) / MINOR;
      if (qty < 0) perProduct[key].geretourneerd += qty;
      else         perProduct[key].verkocht += qty;
      perProduct[key].omzet += lineEur;

      if (item.name?.toLowerCase().includes('cadeaubon') || item.name?.toLowerCase().includes('gift')) {
        cadeaubonnen.aantal += Math.abs(qty);
        cadeaubonnen.bedrag += Math.abs(lineEur);
      }
    }
  }

  const totaleOmzet = totalIncl + totalRefunds;
  const gemBon = transactionCount > 0 ? totaleOmzet / transactionCount : 0;
  const staffRows = Object.entries(perStaff).map(([naam, d]) => ({
    naam, omzet: d.omzet, aantal: d.aantal,
    gemBon: d.aantal > 0 ? d.omzet / d.aantal : 0
  }));
  const productRows = Object.values(perProduct).sort((a, b) => b.omzet - a.omzet);

  return {
    periode: { start: startDate, end: endDate },
    samenvatting: {
      totaleOmzet, aantalVerkopen: transactionCount, gemBon, kaartbetalingen: cardTotal,
      totalExcl, totalVat, totalIncl, totalRefunds,
      cardTotal, cardFee, cardNetto: cardTotal + cardFee,
      cashTotal, cashNetto: cashTotal
    },
    perStaff: staffRows,
    perProduct: productRows,
    extras: { cadeaubonnen }
  };
}

// ─── Excel opbouwen — exacte huisstijl origineel ─────────────────────────────

const KVK = '65284445';

// Precieze kleuren uit het originele rapport
const C = {
  DONKER:    'FF006EA8',  // donkerblauw sectie-headers
  MIDDEL:    'FF008AD1',  // middelblauw periode-balk
  LICHT:     'FFCCE9F7',  // lichtblauw tabel-headers
  GRIJS1:    'FFF2F2F2',  // lichtgrijs afwisselende rijen
  GRIJS2:    'FFFFFFFF',  // wit afwisselende rijen
  GRIJSTOT:  'FFD9D9D9',  // grijs totaalrij
  ROOD:      'FFC00000',  // negatieve getallen
  ZWART:     'FF000000',
  WIT:       'FFFFFFFF',
  BLAUW:     'FF006EA8',
};

const EUR2 = '#,##0.00" €"';
const INT0 = '#,##0';

function exFill(argb) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}
function exFont(bold, size, color) {
  return { name: 'Calibri', bold: !!bold, size: size || 10, color: { argb: color || C.ZWART } };
}
function exAlign(h, v) {
  return { horizontal: h || 'left', vertical: v || 'middle' };
}

function setCell(ws, row, col, value, { bold, size, fc, bg, fmt, h } = {}) {
  const cell = ws.getCell(row, col);
  cell.value = value;
  cell.font = exFont(bold, size, fc || C.ZWART);
  if (bg) cell.fill = exFill(bg);
  if (fmt) cell.numFmt = fmt;
  cell.alignment = exAlign(h || 'left');
  return cell;
}

function hdrDark(ws, row, col, value, size) {
  return setCell(ws, row, col, value, { bold: true, size: size || 11, fc: C.WIT, bg: C.DONKER });
}
function hdrMid(ws, row, col, value, size) {
  return setCell(ws, row, col, value, { bold: false, size: size || 10, fc: C.WIT, bg: C.MIDDEL });
}
function hdrLight(ws, row, col, value, h) {
  return setCell(ws, row, col, value, { bold: true, size: 9, fc: C.BLAUW, bg: C.LICHT, h: h || 'left' });
}
function dataCell(ws, row, col, value, even, fmt, h) {
  return setCell(ws, row, col, value, { size: 10, fc: C.ZWART, bg: even ? C.GRIJS1 : C.GRIJS2, fmt, h: h || 'left' });
}
function totCell(ws, row, col, value, fmt, h) {
  return setCell(ws, row, col, value, { bold: true, size: 10, fc: C.ZWART, bg: C.GRIJSTOT, fmt, h: h || 'left' });
}
function rh(ws, row, height) { ws.getRow(row).height = height; }

function nlDate(s) {
  return new Date(s).toLocaleDateString('nl-NL', { day: '2-digit', month: '2-digit', year: 'numeric' });
}
function maandNaam(s) {
  return new Date(s).toLocaleDateString('nl-NL', { month: 'long', year: 'numeric' })
    .replace(/^\w/, c => c.toUpperCase());
}

async function buildXlsx(data) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'de Fietsboot – Zettle integratie';

  const s     = data.samenvatting;
  const start = data.periode.start;
  const end   = data.periode.end;
  const periodeStr = `Periode: ${nlDate(start)} t/m ${nlDate(end)}  |  KvK: ${KVK}`;
  const maand = maandNaam(start);

  // ── Sheet 1: Samenvatting ─────────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('Samenvatting');
    ws.getColumn(1).width = 2;  ws.getColumn(2).width = 32;
    ws.getColumn(3).width = 18; ws.getColumn(4).width = 18;
    ws.getColumn(5).width = 18; ws.getColumn(6).width = 2;

    rh(ws, 1, 7.5);
    rh(ws, 2, 31.5); hdrDark(ws, 2, 2, 'Verkooprapport – de Fietsboot', 14);
    ws.mergeCells('B2:E2');
    rh(ws, 3, 19.5); hdrMid(ws, 3, 2, periodeStr);
    ws.mergeCells('B3:E3');
    rh(ws, 4, 9.75);

    // KPI balk
    rh(ws, 5, 15.75);
    ['Totale omzet','Aantal verkopen','Gem. bon','Kaartbetalingen']
      .forEach((lbl, i) => hdrLight(ws, 5, 2+i, lbl, 'center'));
    rh(ws, 6, 27.75);
    [
      [2, s.totaleOmzet,    EUR2],
      [3, s.aantalVerkopen, INT0],
      [4, s.gemBon,         EUR2],
      [5, s.kaartbetalingen,EUR2],
    ].forEach(([col, val, fmt]) =>
      setCell(ws, 6, col, val, { bold: true, size: 16, fc: C.BLAUW, bg: C.GRIJS1, fmt, h: 'center' })
    );

    rh(ws, 7, 15.75); rh(ws, 8, 7.5); rh(ws, 9, 7.5);

    // Totale verkoopoverzicht
    rh(ws, 10, 21.75); hdrDark(ws, 10, 2, 'Totale verkoopoverzicht'); ws.mergeCells('B10:E10');
    rh(ws, 11, 18);
    ws.getCell(11, 2).fill = exFill(C.GRIJS1);
    ['Excl. btw (€)','Btw (€)','Incl. btw (€)'].forEach((lbl, i) => hdrLight(ws, 11, 3+i, lbl, 'right'));

    const overzichtRijen = [
      ['Verkopen',        s.totalExcl,                   0, s.totalIncl,   true],
      ['Terugbetalingen', s.totalRefunds,                 0, s.totalRefunds,false],
    ];
    overzichtRijen.forEach(([lbl, excl, btw, incl, even], i) => {
      const r = 12 + i; rh(ws, r, 18);
      const bg = even ? C.GRIJS1 : C.GRIJS2;
      setCell(ws, r, 2, lbl, { size: 10, bg });
      setCell(ws, r, 3, excl, { size: 10, fc: C.ZWART, bg, fmt: EUR2, h: 'right' });
      setCell(ws, r, 4, btw,  { size: 10, fc: C.ZWART, bg, fmt: EUR2, h: 'right' });
      setCell(ws, r, 5, incl, { size: 10, fc: C.ZWART, bg, fmt: EUR2, h: 'right' });
    });
    rh(ws, 14, 18);
    setCell(ws, 14, 2, 'Totaal', { bold: true, size: 10, bg: C.GRIJSTOT });
    setCell(ws, 14, 3, s.totalExcl + s.totalRefunds, { bold: true, size: 10, fc: C.ZWART, bg: C.GRIJSTOT, fmt: EUR2, h: 'right' });
    setCell(ws, 14, 4, 0, { bold: true, size: 10, fc: C.ZWART, bg: C.GRIJSTOT, fmt: EUR2, h: 'right' });
    setCell(ws, 14, 5, s.totaleOmzet, { bold: true, size: 10, fc: C.ZWART, bg: C.GRIJSTOT, fmt: EUR2, h: 'right' });

    rh(ws, 15, 7.5); rh(ws, 16, 7.5);

    // Per personeelslid
    rh(ws, 17, 21.75); hdrDark(ws, 17, 2, 'Verkopen per personeelslid'); ws.mergeCells('B17:E17');
    rh(ws, 18, 18);
    hdrLight(ws, 18, 2, 'Naam');
    ['Incl. btw (€)','Aantal','Gem. bon (€)'].forEach((lbl, i) => hdrLight(ws, 18, 3+i, lbl, 'right'));

    let r = 19;
    data.perStaff.forEach((st, i) => {
      rh(ws, r, 18);
      const bg = i % 2 === 0 ? C.GRIJS1 : C.GRIJS2;
      setCell(ws, r, 2, st.naam, { size: 10, bg });
      setCell(ws, r, 3, st.omzet,  { size: 10, fc: C.ZWART, bg, fmt: EUR2, h: 'right' });
      setCell(ws, r, 4, st.aantal, { size: 10, fc: C.ZWART, bg, fmt: INT0, h: 'right' });
      setCell(ws, r, 5, st.gemBon, { size: 10, fc: C.ZWART, bg, fmt: EUR2, h: 'right' });
      r++;
    });
    rh(ws, r, 7.5); r++;
    const tO1 = data.perStaff.reduce((a, b) => a + b.omzet, 0);
    const tA1 = data.perStaff.reduce((a, b) => a + b.aantal, 0);
    rh(ws, r, 18);
    totCell(ws, r, 2, 'Totaal');
    totCell(ws, r, 3, tO1, EUR2, 'right');
    totCell(ws, r, 4, tA1, INT0, 'right');
    totCell(ws, r, 5, tA1 > 0 ? tO1/tA1 : 0, EUR2, 'right');
    r += 2;

    // Betalingen & kosten
    rh(ws, r, 21.75); hdrDark(ws, r, 2, 'Betalingen & kosten'); ws.mergeCells(`B${r}:E${r}`);
    r++;
    rh(ws, r, 18);
    hdrLight(ws, r, 2, 'Methode');
    ['Bedrag (€)','Toeslag (€)','Netto (€)'].forEach((lbl, i) => hdrLight(ws, r, 3+i, lbl, 'right'));
    r++;

    [
      ['Kaart (reader)', s.cardTotal, s.cardFee,  s.cardNetto,  true],
      ['Contant',        s.cashTotal, 0,           s.cashNetto,  false],
    ].forEach(([lbl, bedrag, toesl, netto, even]) => {
      rh(ws, r, 18);
      const bg = even ? C.GRIJS1 : C.GRIJS2;
      setCell(ws, r, 2, lbl,    { size: 10, bg });
      setCell(ws, r, 3, bedrag, { size: 10, fc: C.ZWART, bg, fmt: EUR2, h: 'right' });
      setCell(ws, r, 4, toesl,  { size: 10, fc: C.ZWART, bg, fmt: EUR2, h: 'right' });
      setCell(ws, r, 5, netto,  { size: 10, fc: C.ZWART, bg, fmt: EUR2, h: 'right' });
      r++;
    });
  }

  // ── Sheet 2: Verkopen per product ─────────────────────────────────────────
  {
    const ws = wb.addWorksheet('Verkopen per product');
    ws.getColumn(1).width = 2;  ws.getColumn(2).width = 22;
    ws.getColumn(3).width = 36; ws.getColumn(4).width = 14;
    ws.getColumn(5).width = 16; ws.getColumn(6).width = 14;
    ws.getColumn(7).width = 16; ws.getColumn(8).width = 2;

    rh(ws, 1, 7.5);
    rh(ws, 2, 30); hdrDark(ws, 2, 2, `Verkopen per product – de Fietsboot  |  ${maand}`, 13);
    ws.mergeCells('B2:G2');
    rh(ws, 3, 18); hdrMid(ws, 3, 2, `Periode: ${nlDate(start)} t/m ${nlDate(end)}`, 9);
    ws.mergeCells('B3:G3');
    rh(ws, 4, 9.75);
    rh(ws, 5, 19.5);
    hdrLight(ws, 5, 2, 'Product');
    hdrLight(ws, 5, 3, 'Variant / Route');
    ['Verkocht','Geretourneerd','Totaal','Omzet (€)'].forEach((lbl, i) => hdrLight(ws, 5, 4+i, lbl, 'right'));

    let r = 6;
    data.perProduct.forEach((p, i) => {
      rh(ws, r, 16.5);
      const bg = i % 2 === 0 ? C.GRIJS1 : C.GRIJS2;
      setCell(ws, r, 2, p.product, { size: 10, bg });
      setCell(ws, r, 3, p.variant, { size: 10, bg });
      setCell(ws, r, 4, p.verkocht,  { size: 10, fc: C.ZWART, bg, fmt: INT0, h: 'right' });
      if (p.geretourneerd) {
        setCell(ws, r, 5, p.geretourneerd, { size: 10, fc: C.ROOD, bg, fmt: INT0, h: 'right' });
      } else {
        ws.getCell(r, 5).fill = exFill(bg);
      }
      setCell(ws, r, 6, p.verkocht + p.geretourneerd, { size: 10, fc: C.ZWART, bg, fmt: INT0, h: 'right' });
      setCell(ws, r, 7, p.omzet, { size: 10, fc: C.ZWART, bg, fmt: EUR2, h: 'right' });
      r++;
    });

    // Totaal met SUM-formules
    rh(ws, r, 19.5);
    const dataStart = 6, dataEnd = r - 1;
    totCell(ws, r, 2, 'Totaal');
    ws.getCell(r, 3).fill = exFill(C.GRIJSTOT);
    setCell(ws, r, 4, `=SUM(D${dataStart}:D${dataEnd})`, { bold: true, size: 10, fc: C.ZWART, bg: C.GRIJSTOT, fmt: INT0, h: 'right' });
    setCell(ws, r, 5, `=SUM(E${dataStart}:E${dataEnd})`, { bold: true, size: 10, fc: C.ZWART, bg: C.GRIJSTOT, fmt: INT0, h: 'right' });
    setCell(ws, r, 6, `=SUM(F${dataStart}:F${dataEnd})`, { bold: true, size: 10, fc: C.ZWART, bg: C.GRIJSTOT, fmt: INT0, h: 'right' });
    setCell(ws, r, 7, `=SUM(G${dataStart}:G${dataEnd})`, { bold: true, size: 10, fc: C.ZWART, bg: C.GRIJSTOT, fmt: EUR2, h: 'right' });
  }

  // ── Sheet 3: Lijnen analyse ───────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('Lijnen analyse');
    ws.getColumn(1).width = 2;  ws.getColumn(2).width = 22;
    ws.getColumn(3).width = 16; ws.getColumn(4).width = 14;
    ws.getColumn(5).width = 16; ws.getColumn(6).width = 2;

    rh(ws, 1, 7.5);
    rh(ws, 2, 30); hdrDark(ws, 2, 2, `Analyse per veerbootlijn – ${maand}`, 13);
    ws.mergeCells('B2:E2');
    rh(ws, 3, 18); hdrMid(ws, 3, 2, 'Gebaseerd op verkoopdata Zettle POS', 9);
    ws.mergeCells('B3:E3');
    rh(ws, 4, 9.75);
    rh(ws, 5, 21.75); hdrDark(ws, 5, 2, 'Omzet per lijn (personeelslid)'); ws.mergeCells('B5:E5');
    rh(ws, 6, 18);
    hdrLight(ws, 6, 2, 'Lijn');
    ['Omzet (€)','Transacties','Gem. bon (€)'].forEach((lbl, i) => hdrLight(ws, 6, 3+i, lbl, 'right'));

    let r = 7;
    data.perStaff.forEach((st, i) => {
      rh(ws, r, 18);
      const bg = i % 2 === 0 ? C.GRIJS1 : C.GRIJS2;
      setCell(ws, r, 2, st.naam,   { size: 10, bg });
      setCell(ws, r, 3, st.omzet,  { size: 10, fc: C.ZWART, bg, fmt: EUR2, h: 'right' });
      setCell(ws, r, 4, st.aantal, { size: 10, fc: C.ZWART, bg, fmt: INT0, h: 'right' });
      setCell(ws, r, 5, st.gemBon, { size: 10, fc: C.ZWART, bg, fmt: EUR2, h: 'right' });
      r++;
    });

    const ds = 7, de = r - 1;
    rh(ws, r, 18);
    totCell(ws, r, 2, 'Totaal');
    setCell(ws, r, 3, `=SUM(C${ds}:C${de})`, { bold: true, size: 10, fc: C.ZWART, bg: C.GRIJSTOT, fmt: EUR2, h: 'right' });
    setCell(ws, r, 4, `=SUM(D${ds}:D${de})`, { bold: true, size: 10, fc: C.ZWART, bg: C.GRIJSTOT, fmt: INT0, h: 'right' });
    setCell(ws, r, 5, `=C${r}/D${r}`,        { bold: true, size: 10, fc: C.ZWART, bg: C.GRIJSTOT, fmt: EUR2, h: 'right' });
    r += 2;

    rh(ws, r, 7.5); r++;
    rh(ws, r, 21.75); hdrDark(ws, r, 2, "Extra's"); ws.mergeCells(`B${r}:D${r}`);
    r++;
    rh(ws, r, 18);
    hdrLight(ws, r, 2, 'Omschrijving');
    hdrLight(ws, r, 3, 'Aantal', 'right');
    hdrLight(ws, r, 4, 'Bedrag (€)', 'right');
    r++;
    rh(ws, r, 18);
    setCell(ws, r, 2, 'Verkochte cadeaubonnen', { size: 10, bg: C.GRIJS1 });
    setCell(ws, r, 3, data.extras.cadeaubonnen.aantal,  { size: 10, fc: C.ZWART, bg: C.GRIJS1, fmt: INT0, h: 'right' });
    setCell(ws, r, 4, data.extras.cadeaubonnen.bedrag,  { size: 10, fc: C.ZWART, bg: C.GRIJS1, fmt: EUR2, h: 'right' });
  }

  return wb.xlsx.writeBuffer();
}

// ─── OneDrive upload + geef download-URL terug ────────────────────────────────

async function uploadToOneDrive(client, buffer, filePath) {
  const user = process.env.GRAPH_ONEDRIVE_USER;
  const encoded = filePath.split('/').map(encodeURIComponent).join('/');

  // Upload (overschrijft bestaand bestand)
  const uploaded = await client
    .api(`/users/${user}/drive/root:/${encoded}:/content`)
    .header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    .put(buffer);

  // Maak een deelbare link (view-only, iedereen met de link)
  try {
    const share = await client
      .api(`/users/${user}/drive/items/${uploaded.id}/createLink`)
      .post({ type: 'view', scope: 'organization' });
    return share.link?.webUrl || null;
  } catch {
    return uploaded.webUrl || null;
  }
}

// ─── E-mail via Microsoft Graph ───────────────────────────────────────────────

function nlDate(s) {
  return new Date(s).toLocaleDateString('nl-NL', { day: 'numeric', month: 'long', year: 'numeric' });
}

async function stuurEmail(client, { ontvanger, startDate, endDate, data, buffer, onedrivePath, onedriveUrl }) {
  const maand = fmtMaand(startDate);
  const s = data.samenvatting;
  const bestandsnaam = `Zettle_Verkooprapport_${startDate}_tm_${endDate}.xlsx`;

  // Bouw nette HTML-tabel voor de samenvatting
  const perStafHtml = data.perStaff.map(st =>
    `<tr><td>${st.naam}</td><td>€ ${fmt(st.omzet).toFixed(2)}</td><td>${st.aantal}</td><td>€ ${fmt(st.gemBon).toFixed(2)}</td></tr>`
  ).join('');

  const html = `
<div style="font-family:Arial,sans-serif;max-width:600px;color:#1a1a1a">
  <div style="background:#1A3A5C;padding:20px 28px;border-radius:8px 8px 0 0">
    <h2 style="color:white;margin:0;font-size:1.1rem">⛵ Zettle Verkooprapport</h2>
    <p style="color:rgba(255,255,255,0.75);margin:4px 0 0;font-size:0.85rem">
      ${nlDate(startDate)} t/m ${nlDate(endDate)}
    </p>
  </div>
  <div style="background:#f5f7fa;padding:24px 28px;border:1px solid #e4e7ed;border-top:none;border-radius:0 0 8px 8px">

    <h3 style="font-size:0.9rem;text-transform:uppercase;letter-spacing:0.05em;color:#555;margin:0 0 12px">
      Samenvatting
    </h3>
    <table style="width:100%;border-collapse:collapse;background:white;border-radius:6px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.06)">
      <tr style="background:#E8F0F7">
        <td style="padding:10px 14px;font-weight:600;font-size:0.85rem">Totale omzet</td>
        <td style="padding:10px 14px;text-align:right;font-weight:700;font-size:1.1rem;color:#1A3A5C">€ ${fmt(s.totaleOmzet).toFixed(2)}</td>
      </tr>
      <tr><td style="padding:8px 14px;font-size:0.85rem;color:#555">Aantal transacties</td>
          <td style="padding:8px 14px;text-align:right;font-size:0.85rem">${s.aantalVerkopen}</td></tr>
      <tr style="background:#f9fafb">
          <td style="padding:8px 14px;font-size:0.85rem;color:#555">Gemiddelde bon</td>
          <td style="padding:8px 14px;text-align:right;font-size:0.85rem">€ ${fmt(s.gemBon).toFixed(2)}</td></tr>
      <tr><td style="padding:8px 14px;font-size:0.85rem;color:#555">Kaartbetalingen</td>
          <td style="padding:8px 14px;text-align:right;font-size:0.85rem">€ ${fmt(s.kaartbetalingen).toFixed(2)}</td></tr>
      <tr style="background:#f9fafb">
          <td style="padding:8px 14px;font-size:0.85rem;color:#555">Contant</td>
          <td style="padding:8px 14px;text-align:right;font-size:0.85rem">€ ${fmt(s.cashTotal).toFixed(2)}</td></tr>
    </table>

    <h3 style="font-size:0.9rem;text-transform:uppercase;letter-spacing:0.05em;color:#555;margin:20px 0 12px">
      Per veerbootlijn
    </h3>
    <table style="width:100%;border-collapse:collapse;background:white;border-radius:6px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.06)">
      <tr style="background:#E8F0F7">
        <th style="padding:8px 14px;text-align:left;font-size:0.82rem">Lijn</th>
        <th style="padding:8px 14px;text-align:right;font-size:0.82rem">Omzet</th>
        <th style="padding:8px 14px;text-align:right;font-size:0.82rem">Aantal</th>
        <th style="padding:8px 14px;text-align:right;font-size:0.82rem">Gem. bon</th>
      </tr>
      ${perStafHtml}
    </table>

    <p style="margin:20px 0 6px;font-size:0.85rem;color:#555">
      Het volledige rapport is bijgevoegd als Excel-bestand.
      ${onedriveUrl ? `Je kunt het ook <a href="${onedriveUrl}" style="color:#1A3A5C">openen via OneDrive</a>.` : ''}
    </p>
    <p style="font-size:0.75rem;color:#999;margin:16px 0 0">
      Automatisch gegenereerd op ${new Date().toLocaleDateString('nl-NL')} via Zettle &amp; de Fietsboot rapportage
    </p>
  </div>
</div>`;

  const message = {
    subject: `Zettle Verkooprapport ${maand} (${nlDate(startDate)} – ${nlDate(endDate)})`,
    body: { contentType: 'HTML', content: html },
    toRecipients: [{ emailAddress: { address: ontvanger } }],
    attachments: [{
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: bestandsnaam,
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      contentBytes: Buffer.from(buffer).toString('base64')
    }]
  };

  await client
    .api(`/users/${process.env.GRAPH_MAIL_FROM}/sendMail`)
    .post({ message, saveToSentItems: true });
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
  if (!startDate || !endDate) {
    return { statusCode: 400, body: JSON.stringify({ ok: false, fout: 'startDate en endDate zijn verplicht' }) };
  }
  if (!ontvanger) {
    return { statusCode: 400, body: JSON.stringify({ ok: false, fout: 'ontvanger is verplicht' }) };
  }

  try {
    // 1. Haal aankopen op van Zettle
    const purchases = await fetchAllPurchases(startDate, endDate);

    // 2. Bereken rapportdata
    const reportData = buildReportData(purchases, startDate, endDate);

    // 3. Bouw Excel-bestand
    const buffer = await buildXlsx(reportData);

    // 4. Upload naar OneDrive
    const client = getGraphClient();
    const targetPath = process.env.ZETTLE_RAPPORT_PATH ||
      'MS365/Zettle Rapporten/Zettle_Verkooprapport_Actueel.xlsx';
    const onedriveUrl = await uploadToOneDrive(client, buffer, targetPath);

    // 5. Stuur e-mail
    await stuurEmail(client, {
      ontvanger,
      startDate,
      endDate,
      data: reportData,
      buffer,
      onedrivePath: targetPath,
      onedriveUrl
    });

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        ok: true,
        periode: { startDate, endDate },
        aantalAankopen: purchases.length,
        totaleOmzet: reportData.samenvatting.totaleOmzet,
        ontvangerEmail: ontvanger,
        onedriveUrl,
        bericht: `Rapport bijgewerkt en verzonden naar ${ontvanger}`
      })
    };
  } catch (err) {
    console.error('update-rapport fout:', err);
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ok: false, fout: err.message })
    };
  }
};
