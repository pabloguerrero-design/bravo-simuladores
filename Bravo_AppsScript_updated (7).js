// ============================================================
//  BRAVO — Propuesta de Crédito | Apps Script
//  Extensiones → Apps Script → pegar → guardar → recargar hoja
//  Celdas verificadas sobre Rosario_Elizabeth_Murillo_Zuñiga.xlsx
// ============================================================

const SIMULADOR_URL = 'https://peppy-llama-4a5548.netlify.app';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔵 Bravo')
    .addItem('📄 Generar Propuesta de Crédito', 'generarPropuesta')
    .addItem('🔍 Debug — ver datos leídos',      'debugDatos')
    .addToUi();
}

// ── Helpers ──────────────────────────────────────────────────
function safe(fn) {
  try {
    const v = fn();
    if (v === null || v === undefined) return null;
    const s = String(v).trim();
    return (s === '' || s.startsWith('#')) ? null : v;
  } catch(e) { return null; }
}

function num(v) {
  if (v === null || v === undefined) return 0;
  // Google Sheets returns raw numbers - use as-is with full precision
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function numCell(sheet, cell) {
  try {
    // Use raw value first (most reliable)
    const v = sheet.getRange(cell).getValue();
    if (v === null || v === undefined || v === '') return 0;
    const n = Number(v);
    if (!isNaN(n)) return n;
    // Fallback: parse display value handling European format (1.234,56 → 1234.56)
    const d = sheet.getRange(cell).getDisplayValue();
    return parseEU(d);
  } catch(e) { return 0; }
}

function parseEU(s) {
  if (!s) return 0;
  // Remove currency symbols and spaces
  s = String(s).replace(/[€$\s]/g,'').trim();
  // European: 1.234,56 → check if comma is decimal
  if (s.match(/^\d{1,3}(\.\d{3})*(,\d+)?$/)) {
    s = s.replace(/\./g,'').replace(',','.');
  } else if (s.match(/^\d{1,3}(,\d{3})*(\.\d+)?$/)) {
    s = s.replace(/,/g,'');
  } else {
    s = s.replace(',','.');
  }
  return parseFloat(s) || 0;
}

function fdate(v) {
  if (!v) return '';
  try {
    const d = new Date(v);
    if (isNaN(d.getTime())) return '';
    return d.getFullYear() + '-' +
      String(d.getMonth() + 1).padStart(2, '0') + '-' +
      String(d.getDate()).padStart(2, '0');
  } catch(e) { return ''; }
}

// ── Leer datos ───────────────────────────────────────────────
function leerDatos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dc = ss.getSheetByName('Datos crédito');
  const ta = ss.getSheetByName('TA');
  const dg = ss.getSheetByName('Datos Generador');
  const ct = ss.getSheetByName('Contratos');

  if (!dc) throw new Error('No se encontró la hoja "Datos crédito"');
  if (!ta) throw new Error('No se encontró la hoja "TA"');

  // ── NOMBRE ─────────────────────────────────────────────────
  // Verificado: Contratos!C3 = nombre del cliente
  const CABECERAS = new Set([
    'id deuda','nombre','cliente','nombre cliente','datos cliente',
    'spv','spv1','spv2','spv3','respaldo','vehiculo','check',
    'id','ref','referencia','dato','datos'
  ]);
  function esNombreValido(v) {
    if (!v) return false;
    const s = String(v).trim();
    if (s.length < 5) return false;
    if (/^[A-Z]{2,5}\d+/i.test(s)) return false; // UES0085716, CN69207
    if (CABECERAS.has(s.toLowerCase())) return false;
    if (!s.includes(' ')) return false; // mínimo nombre + apellido
    return true;
  }
  let nombre = null;
  const ng = ss.getSheetByName('Nego');
  const fuentesNombre = [
    () => ng ? safe(() => ng.getRange('D5').getValue()) : null,  // ✓ PRIMARIO Nego D5
    () => ct ? safe(() => ct.getRange('C3').getValue()) : null,
    () => ct ? safe(() => ct.getRange('D3').getValue()) : null,
    () => ct ? safe(() => ct.getRange('B3').getValue()) : null,
    () =>      safe(() => ta.getRange('D2').getValue()),
    () =>      safe(() => dc.getRange('B2').getValue()),
    () =>      safe(() => dc.getRange('C2').getValue()),
  ];
  for (const fn of fuentesNombre) {
    const v = fn();
    if (esNombreValido(v)) { nombre = String(v).trim(); break; }
  }
  // Último recurso: título del fichero
  if (!nombre) {
    const titulo = ss.getName()
      .replace(/copia de /gi, '')
      .replace(/[_\-]/g, ' ')
      .replace(/\b(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre|\d{4})\b/gi, '')
      .replace(/\s+/g, ' ').trim();
    if (esNombreValido(titulo)) nombre = titulo;
  }
  nombre = nombre || 'Cliente';

  // ── REFERENCIA ─────────────────────────────────────────────
  const ref = String(safe(() => dc.getRange('C18').getValue()) || '').trim();

  // ── N° CRÉDITO ─────────────────────────────────────────────
  // Verificado: Contratos!C8 = CN del nuevo crédito
  function esCN(v) {
    if (!v) return false;
    const s = String(v).trim();
    // CN válido: empieza por CN seguido de números, o es alfanumérico corto
    return /^CN\d+/i.test(s) || (s.length > 3 && s.length < 20 && /^[A-Z]{1,5}\d+/i.test(s));
  }
  let nCred = '';
  const fuentesNCred = [
    () => ct ? safe(() => ct.getRange('C27').getValue()) : null,  // ✓ PRIMARIO Contratos C27
    () => ct ? safe(() => ct.getRange('C8').getValue()) : null,
    () => ct ? safe(() => ct.getRange('I8').getValue()) : null,
    () => dg ? safe(() => dg.getRange('Q2').getValue()) : null,
  ];
  for (const fn of fuentesNCred) {
    const v = fn();
    if (esCN(v)) { nCred = String(v).trim(); break; }
  }

  // ── TIN ────────────────────────────────────────────────────
  let tin = 19;
  const tinRaw = safe(() => ta.getRange('D4').getValue());
  if (tinRaw !== null) {
    const t = num(tinRaw);
    if (t > 0) tin = t > 1 ? Math.round(t) : Math.round(t * 100);
  }

  // ── CUOTA ──────────────────────────────────────────────────
  // Buscar en muchas celdas posibles hasta encontrar un valor > 0
  let cuota = 0;
  const fuentesCuota = [
    () => Number(dc.getRange('G23').getValue()) || 0,   // ✓ PRIMARIO DC G23
    () => Number(dc.getRange('I23').getValue()) || 0,   // DC I23
    () => Number(dc.getRange('H23').getValue()) || 0,   // DC H23
    () => dg ? (Number(dg.getRange('G2').getValue()) || 0) : 0,
    () => dg ? (Number(dg.getRange('H2').getValue()) || 0) : 0,
    () => Number(ta.getRange('E13').getValue()) || 0,
  ];
  for (const fn of fuentesCuota) {
    try { const v = fn(); if (v > 0 && v < 10000) { cuota = v; break; } } catch(e) {}
  }

  // ── DESGLOSE SALDO (Datos crédito col G) ──────────────────
  // G29=Mensuales, G31=Diferimiento, G32=Refi, G33=Estructurado, G35=Ahorro
  const mensuales    = num(safe(() => dc.getRange('G29').getValue()));
  const diferimiento = num(safe(() => dc.getRange('G31').getValue()));
  const refi         = num(safe(() => dc.getRange('G32').getValue()));
  const estructurado = num(safe(() => dc.getRange('G33').getValue()));
  const ahorro       = Number(dc.getRange('I35').getValue()) || 0;

  // ── FECHAS ─────────────────────────────────────────────────
  const apertura   = fdate(safe(() => dc.getRange('I41').getValue()));
  const primerPago = fdate(safe(() => dc.getRange('I42').getValue()));

  // ── DEUDAS ─────────────────────────────────────────────────
  const SALTAR = new Set([
    '', 'Diferimiento', 'Refi', 'Condonación Exito',
    'Estructurado', 'Cierre', 'Datos cliente',
    'Nota: El cierre incluye los intereses'
  ]);
  const deudas = [];
  for (let f = 3; f <= 20; f++) {
    const entRaw = safe(() => dc.getRange(f, 2).getValue());
    if (!entRaw) break;
    const ent = String(entRaw).trim();
    if (!ent || SALTAR.has(ent) || ent.startsWith('Nota:') || ent.startsWith('Datos ')) break;
    const drRaw  = dc.getRange(f, 4).getValue();
    const pabRaw = dc.getRange(f, 5).getValue();
    const dr  = (!isNaN(Number(drRaw))  && Number(drRaw)  > 0) ? Number(drRaw)  : parseEU(dc.getRange(f, 4).getDisplayValue());
    const pab = (!isNaN(Number(pabRaw)) && Number(pabRaw) > 0) ? Number(pabRaw) : parseEU(dc.getRange(f, 5).getDisplayValue());
    if (dr > 0 || pab > 0) {
      deudas.push({ entidad: ent, dr: dr.toFixed(2), pab: pab.toFixed(2) });
    }
  }

  // ── REVOLVING ──────────────────────────────────────────────
  let revolving = false, cnAnt = '', saldoAnt = 0;
  if (ct) {
    const rv = String(safe(() => ct.getRange('P9').getValue()) || '').toLowerCase().trim();
    revolving = rv === 'sí' || rv === 'si' || rv === 'yes' || rv === '1';
    if (revolving) {
      // CN anterior y saldo desde Datos crédito C15 y E15
      cnAnt    = String(safe(() => dc.getRange('C15').getValue()) || '').trim();
      const saldoRaw = dc.getRange('E15').getValue();
      saldoAnt = (!isNaN(Number(saldoRaw)) && Number(saldoRaw) > 0)
        ? Number(saldoRaw)
        : parseEU(dc.getRange('E15').getDisplayValue());
      // Fallbacks
      if (!cnAnt) {
        const cr = ss.getSheetByName('Cierre revolving');
        if (cr) cnAnt = String(safe(() => cr.getRange('B2').getValue()) || '').trim();
      }
      if (!saldoAnt) saldoAnt = num(safe(() => ct.getRange('C8').getValue()));
    }
  }

  // Saldo Inicial real desde TA D3
  const saldoInicial = Number(ta.getRange('D3').getValue()) || 0;

  return { nombre, ref, nCred, tin, cuota, apertura, primerPago,
           mensuales, diferimiento, refi, estructurado, ahorro,
           saldoInicial, deudas, revolving, cnAnt, saldoAnt };
}

// ── Generar propuesta ────────────────────────────────────────
function generarPropuesta() {
  let d;
  try { d = leerDatos(); }
  catch(e) {
    SpreadsheetApp.getUi().alert('❌ Error leyendo datos:\n\n' + e.message);
    return;
  }

  const params = {
    nombre:        d.nombre,
    ref:           d.ref,
    ncred:         d.nCred,
    tin:           d.tin,
    apertura:      d.apertura,
    pago:          d.primerPago,
    cuota:         d.cuota.toFixed(2),
    mensuales:     d.mensuales.toFixed(2),
    ahorro:        d.ahorro.toFixed(2),
    diferimiento:  d.diferimiento.toFixed(2),
    refi:          d.refi.toFixed(2),
    estructurado:  d.estructurado.toFixed(2),
    saldoForzado:  d.saldoInicial > 0 ? d.saldoInicial.toFixed(2) : '',
    deudas:        JSON.stringify(d.deudas),
    revolving:     d.revolving ? '1' : '0',
    cnAnt:         d.cnAnt,
    saldoAnt:      d.saldoAnt.toFixed(2),
    autogenerar:   '1'
  };

  const qs  = Object.entries(params)
    .map(([k, v]) => encodeURIComponent(k) + '=' + encodeURIComponent(v))
    .join('&');
  const url = SIMULADOR_URL + '?' + qs;
  const esc = url.replace(/\\/g, '\\\\').replace(/'/g, "\\'");

  // Abrir directamente sin modal
  const html = HtmlService.createHtmlOutput(
    `<script>window.open('${esc}','_blank');google.script.host.close();<\/script>`
  ).setWidth(1).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, 'Abriendo...');
}

// ── Debug ────────────────────────────────────────────────────
function debugDatos() {
  let d;
  try { d = leerDatos(); }
  catch(e) { SpreadsheetApp.getUi().alert('❌ Error:\n\n' + e.message); return; }

  const ok = v => (v && String(v).length > 0) ? '✓' : '⚠️';
  const msg = [
    '─── CLIENTE ─────────────────────',
    `${ok(d.nombre)}  Nombre:   ${d.nombre}`,
    `${ok(d.ref)}  Ref:      ${d.ref}`,
    `${ok(d.nCred)}  CN nuevo: ${d.nCred}`,
    '',
    '─── CRÉDITO ─────────────────────',
    `${ok(d.tin)}  TIN:      ${d.tin}%`,
    `${ok(d.cuota)}  Cuota:    ${d.cuota.toFixed(2)} €`,
    `${ok(d.apertura)}  Apertura: ${d.apertura}`,
    `${ok(d.saldoInicial)}  Saldo TA D3: ${d.saldoInicial ? d.saldoInicial.toFixed(2) : '⚠️'}`,
    `${ok(d.primerPago)}  1er pago: ${d.primerPago}`,
    '',
    `─── DEUDAS (${d.deudas.length}) ──────────────────`,
    ...d.deudas.map(x => `  · ${x.entidad}  DR:${x.dr} EUR  PAB:${x.pab} EUR`),
    d.deudas.length === 0 ? '  ⚠️ Ninguna detectada' : '',
    '',
    '─── REVOLVING ───────────────────',
    d.revolving
      ? `✓  SÍ → CN anterior: ${d.cnAnt} / Saldo: ${d.saldoAnt.toFixed(2)} €`
      : '✓  NO (sin cierre de crédito anterior)',
    '',
    '─── URL GENERADA ────────────────',
  ].join('\n');

  // Construir URL para diagnóstico
  const params = {
    nombre: d.nombre, ref: d.ref, ncred: d.nCred,
    tin: d.tin, apertura: d.apertura, pago: d.primerPago,
    cuota: d.cuota.toFixed(2), deudas: JSON.stringify(d.deudas),
    revolving: d.revolving ? '1' : '0',
    cnAnt: d.cnAnt, saldoAnt: d.saldoAnt.toFixed(2), autogenerar: '1'
  };
  const qs = Object.entries(params)
    .map(([k,v]) => encodeURIComponent(k)+'='+encodeURIComponent(v)).join('&');
  const url = SIMULADOR_URL + '?' + qs;

  SpreadsheetApp.getUi().alert(
    '📊 DATOS DETECTADOS\n\n' + msg +
    '\n' + url.substring(0, 400) +
    (url.length > 400 ? '\n...(truncada)' : '')
  );
}

// ── Recibir PDF desde HTML y guardar en Drive ─────────────
function doPost(e) {
  try {
    const data     = JSON.parse(e.postData.contents);
    const base64   = data.pdfBase64;
    const filename = data.filename  || 'PropuestaCredito_Bravo.pdf';
    const folderId = data.folderId  || '1PNJQ43QMO6-qDt8Q8Hx5wlm8lc51-3_E';

    const bytes  = Utilities.base64Decode(base64);
    const blob   = Utilities.newBlob(bytes, 'application/pdf', filename);
    const folder = DriveApp.getFolderById(folderId);
    folder.createFile(blob);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, filename }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
