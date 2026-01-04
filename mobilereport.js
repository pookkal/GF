/**
 * ==============================================================================
 * FULL MOBILE REPORT GENERATOR (ALL FUNCTIONS) — BLOOMBERG-STYLE + VALUE COLORS
 * ==============================================================================
 */


/* =============================================================================
 * PUBLIC ENTRYPOINT
 * ============================================================================= */
function generateMobileReport() {
  Logger.clear();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const INPUT = ss.getSheetByName('INPUT');
  const CALC  = ss.getSheetByName('CALCULATIONS');

  if (!INPUT) throw new Error('INPUT sheet not found');
  if (!CALC) throw new Error('CALCULATIONS sheet not found');

  let REPORT = ss.getSheetByName('REPORT');
  if (!REPORT) REPORT = ss.insertSheet('REPORT');

  // A1 dropdown from INPUT!A3:A
  setupReportTickerDropdown_(REPORT, INPUT);

  // Resolve ticker: REPORT!A1 > DASHBOARD!H1
  let ticker = String(REPORT.getRange('A1').getDisplayValue() || '').trim();
  if (!ticker) {
      SpreadsheetApp.getUi().alert('No ticker found. Select one in REPORT!A1.');
      return;
  }

  // Read CALCULATIONS row
  const rowObj = getCalcRowByTicker___(CALC, ticker);
  if (!rowObj || !rowObj.dataMap) {
    renderNoDataReport___(REPORT, INPUT, ticker, 'Ticker not found in CALCULATIONS.');
    return;
  }

  // Extract + interpret
  const f = extractMasterFields___(rowObj.dataMap, ticker);
  const interp = computeAPLInterpretations___(f);

  // Protect against sheet errors leaking into headline
  if (isSheetError___(f.SIGNAL)) f.SIGNAL = interp.fallbackSignal;
  if (isSheetError___(f.DECISION)) f.DECISION = interp.fallbackDecision;

  // Render
  renderReportSheet___(REPORT, INPUT, f, interp);

  SpreadsheetApp.flush();
}

/* =============================================================================
 * DROPDOWN: REPORT!A1 from INPUT!A3:A
 * ============================================================================= */
function setupReportTickerDropdown_(reportSheet, inputSheet) {
  const last = inputSheet.getLastRow();
  const height = Math.max(1, last - 2);
  const rng = inputSheet.getRange(3, 1, height, 1);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(rng, true)
    .setAllowInvalid(false)
    .build();

  const a1 = reportSheet.getRange('A1');
  a1.setDataValidation(rule);
  a1.setFontWeight('bold');
  a1.setHorizontalAlignment('left');
}


/* =============================================================================
 * FETCH CALCULATIONS ROW BY TICKER
 * - Prefers row 2 headers; falls back to row 1 if row 2 is empty
 * - Data begins at row 3
 * ============================================================================= */
function getCalcRowByTicker___(calcSheet, ticker) {
  const lastRow = calcSheet.getLastRow();
  const lastCol = calcSheet.getLastColumn();
  if (lastRow < 3 || lastCol < 2) return null;

  const headersRow2 = calcSheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const headersRow1 = calcSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headers = headersRow2.some(h => String(h).trim()) ? headersRow2 : headersRow1;

  const startRow = 3;
  const height = Math.max(0, lastRow - startRow + 1);
  if (height <= 0) return null;

  const tickers = calcSheet.getRange(startRow, 1, height, 1)
    .getDisplayValues()
    .map(r => String(r[0]).trim().toUpperCase());

  const t = String(ticker || '').trim().toUpperCase();
  const idx = tickers.findIndex(x => x === t);
  if (idx < 0) return null;

  const rowIndex = startRow + idx;
  const rowValues = calcSheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

  const dataMap = {};
  headers.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) dataMap[key] = rowValues[i];
  });

  return { rowIndex, dataMap };
}


/* =============================================================================
 * NORMALIZATION: percent-like to 0..1
 * - Accepts 0..1, 0..100, "95.7%", "0.957", etc.
 * ============================================================================= */
function toPct01___(v) {
  if (v === null || v === undefined || v === '') return null;

  if (typeof v === 'string') {
    const s = v.trim();
    const m = s.match(/-?\d+(\.\d+)?/);
    if (!m) return null;
    const n = parseFloat(m[0]);
    if (isNaN(n)) return null;
    if (/%/.test(s)) return n / 100;
    v = n;
  }

  if (typeof v !== 'number' || isNaN(v)) return null;

  // Heuristic: >1.5 is almost always 0..100 scale
  if (v > 1.5) return v / 100;
  return v;
}


/* =============================================================================
 * FIELD EXTRACTION FROM CALCULATIONS MAP
 * - CRITICAL: StochK and BBpct normalized to 0..1
 * ============================================================================= */
function extractMasterFields___(data, ticker) {
  const get = (k, fb = '') => (Object.prototype.hasOwnProperty.call(data, k) ? data[k] : fb);

  return {
    Ticker: ticker,

    SIGNAL: String(get('SIGNAL', '') || ''),
    FUNDAMENTAL: String(get('FUNDAMENTAL', '') || ''),
    DECISION: String(get('DECISION', '') || ''),

    Price: toNum___(get('Price', get('CURRENT PRICE', ''))),
    ChangePct: get('Change %', get('CHANGE%', '')),

    RVOL: toNum___(get('Vol Trend', get('RVOL', ''))),
    TrendState: String(get('Trend State', get('TREND STATE', '')) || ''),
    TrendScore: String(get('Trend Score', '') || ''),
    RR: toNum___(get('R:R Quality', get('R:R', ''))),

    SMA20: toNum___(get('SMA 20', '')),
    SMA50: toNum___(get('SMA 50', '')),
    SMA200: toNum___(get('SMA 200', '')),

    RSI: toNum___(get('RSI', get('RSI (14)', ''))),
    MACDHist: toNum___(get('MACD Hist', '')),
    Divergence: String(get('Divergence', '') || ''),

    ADX: toNum___(get('ADX (14)', get('ADX', ''))),

    // FIXED: normalized to 0..1 (format as % in REPORT)
    StochK: toPct01___(get('Stoch %K (14)', get('Stoch %K', ''))),
    BBpct:  toPct01___(get('Bollinger %B', '')),

    Support: toNum___(get('Support', '')),
    Resistance: toNum___(get('Resistance', '')),
    Target: toNum___(get('Target (3:1)', get('Target', ''))),

    ATR: toNum___(get('ATR (14)', get('ATR', ''))),

    ATH: toNum___(get('ATH (TRUE)', get('ATH', ''))),
    ATHDiffPct: get('ATH Diff %', '')
  };
}


/* =============================================================================
 * INTERPRETATIONS (stable, value-based)
 * ============================================================================= */
function computeAPLInterpretations___(f) {
  const regime =
    (f.Price !== null && f.SMA200 !== null)
      ? (f.Price >= f.SMA200 ? 'RISK-ON (Above SMA200)' : 'RISK-OFF (Below SMA200)')
      : 'REGIME: —';

  let trendLabel = '—';
  if (f.TrendState) trendLabel = `${f.TrendState}${f.TrendScore ? ` | Score: ${f.TrendScore}` : ''}`;
  else if (f.Price !== null && f.SMA200 !== null) trendLabel = (f.Price >= f.SMA200) ? 'BULL | Score: —' : 'BEAR | Score: —';

  let stretchLabel = '—';
  if (f.Price !== null && f.SMA20 !== null && f.ATR !== null && f.ATR !== 0) {
    const stretch = (f.Price - f.SMA20) / f.ATR;
    const sign = (stretch >= 0 ? '+' : '');
    stretchLabel = `${sign}${stretch.toFixed(1)}x ATR vs SMA20`;
  }

  let athDiff = String(f.ATHDiffPct || '').trim();
  if (!athDiff && f.ATH !== null && f.Price !== null && f.ATH !== 0) {
    const diff = (f.Price / f.ATH - 1) * 100;
    athDiff = (diff > 0 ? '+' : '') + diff.toFixed(2) + '%';
  }

  const fallbackSignal = computeFallbackSignal___(f);
  const fallbackDecision = computeFallbackDecision___(f, fallbackSignal);

  const why = buildWhyText___(f, regime);
  const whyNot = buildWhyNotText___(f);

  return { regime, trendLabel, stretchLabel, athDiff, fallbackSignal, fallbackDecision, why, whyNot };
}

function computeFallbackSignal___(f) {
  if (f.Price === null) return 'LOADING';
  if (f.Support !== null && f.Price < f.Support) return 'Stop-Out';
  if (f.SMA200 !== null && f.Price < f.SMA200) return 'Risk-Off (Below SMA200)';
  return 'RISK-ON (Above SMA200)';
}

function computeFallbackDecision___(f, fallbackSignal) {
  if (fallbackSignal === 'LOADING') return 'LOADING';
  const weakParticipation = (f.RVOL !== null && f.RVOL < 1.0);
  const noTrend = (f.ADX !== null && f.ADX < 20);
  const weakRR = (f.RR !== null && f.RR < 1.8);
  if (weakParticipation || noTrend || weakRR) return 'WAIT / MONITOR';
  return 'HOLD / ACCUMULATE';
}

function buildWhyText___(f, regime) {
  const lines = [];

  if (f.Price !== null && f.SMA200 !== null) {
    lines.push(`Price (${money___(f.Price)}) is ${f.Price >= f.SMA200 ? 'above' : 'below'} SMA200 (${money___(f.SMA200)}) → ${regime}.`);
  } else {
    lines.push(`Regime: ${regime}.`);
  }

  if (f.RVOL !== null) lines.push(`RVOL ${f.RVOL.toFixed(2)}x → ${rvolComment___(f.RVOL)}.`);
  if (f.RSI !== null) lines.push(`RSI ${f.RSI.toFixed(2)} → ${rsiComment___(f.RSI)}.`);
  if (f.ADX !== null) lines.push(`ADX ${f.ADX.toFixed(2)} → ${adxComment___(f.ADX)}.`);
  if (f.StochK !== null) lines.push(`Stoch %K ${(f.StochK * 100).toFixed(1)}% → ${stochComment01___(f.StochK)}.`);
  if (f.BBpct !== null) lines.push(`Boll %B ${(f.BBpct * 100).toFixed(1)}% → ${bbComment01___(f.BBpct)}.`);
  if (f.RR !== null) lines.push(`R:R ${f.RR.toFixed(2)}x → ${rrComment___(f.RR)}.`);

  return lines.join('\n');
}

function buildWhyNotText___(f) {
  const lines = [];

  if (f.Price !== null && f.Support !== null && f.Price >= f.Support) lines.push('Stop-Out not triggered: price is not below Support.');
  else if (f.Price !== null && f.Support !== null) lines.push('Stop-Out triggered: price is below Support.');
  else lines.push('Stop-Out check inconclusive: Support/Price unavailable.');

  if (f.RVOL !== null && f.RVOL < 1.5) lines.push('Breakout confirmation weaker: RVOL < 1.5x.');
  else if (f.RVOL !== null) lines.push('Breakout confirmation stronger: RVOL ≥ 1.5x.');
  else lines.push('Breakout confirmation: RVOL unavailable.');

  return lines.join('\n');
}


/* =============================================================================
 * REPORT RENDERER (MOBILE FRIENDLY) — WITH VALUE COLORS
 * ============================================================================= */
function renderReportSheet___(REPORT, INPUT, f, interp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');

  // Preserve A1 + validation
  const ticker = String(REPORT.getRange('A1').getDisplayValue() || '').trim();
  const a1Rule = REPORT.getRange('A1').getDataValidation();

  REPORT.clear({ contentsOnly: true });
  REPORT.clearFormats();

  REPORT.getRange('A1').setValue(ticker);
  if (a1Rule) REPORT.getRange('A1').setDataValidation(a1Rule);

  setReportColumnWidthsAndWrap___(REPORT);

  const P = reportPalette___();

  // Split zone: row 9..34 has column C narrative
  const SPLIT_START = 9;
  const SPLIT_END   = 34;

  // --- Header row 1 ---
  REPORT.getRange('B1').setValue('MASTER REPORT');
  REPORT.getRange('B1:C1').merge();

  REPORT.getRange('A1:C1')
    .setBackground(P.BG_TOP)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('left');

  REPORT.getRange('A1')
    .setBackground(P.PANEL)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  // --- Rows 2–4 (yellow) ---
  REPORT.getRange('B2:C2').merge();
  REPORT.getRange('B3:C3').merge();
  REPORT.getRange('B4:C4').merge();

  REPORT.getRange('A2').setValue('SIGNAL');
  REPORT.getRange('B2').setValue(f.SIGNAL || '—');
  REPORT.getRange('A3').setValue('FUNDAMENTAL');
  REPORT.getRange('B3').setValue(f.FUNDAMENTAL || '—');
  REPORT.getRange('A4').setValue('DECISION');
  REPORT.getRange('B4').setValue(f.DECISION || '—');

  REPORT.getRange('A2:C4')
    .setBackground(P.YELLOW)
    .setFontColor(P.BLACK)
    .setFontWeight('bold')
    .setFontSize(11)
    .setBorder(true, true, true, true, true, true, P.GRID, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setHorizontalAlignment('left');

  styleChipCell___(REPORT.getRange('B2'), f.SIGNAL);
  styleChipCell___(REPORT.getRange('B3'), f.FUNDAMENTAL);
  styleChipCell___(REPORT.getRange('B4'), f.DECISION);

  // --- Row 5 meta ---
  REPORT.getRange('A5').setValue(`Generated: ${now}`);
  REPORT.getRange('A5:C5').merge()
    .setBackground(P.BG_TOP)
    .setFontColor(P.MUTED)
    .setFontSize(9)
    .setHorizontalAlignment('left');

  // Start main
  let r = 6;

  // Regime headline
  REPORT.getRange(r, 1).setValue(interp.regime);
  REPORT.getRange(r, 1, 1, 3).merge()
    .setBackground('#1F2937')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('left');
  r += 2;

  // Section header
  const writeSectionHeader = (title) => {
    REPORT.getRange(r, 1).setValue(title);
    REPORT.getRange(r, 1, 1, 3).merge()
      .setBackground(P.PANEL)
      .setFontColor(P.TEXT)
      .setFontWeight('bold')
      .setFontSize(11)
      .setBorder(true, true, true, true, false, false, P.GRID, SpreadsheetApp.BorderStyle.SOLID)
      .setHorizontalAlignment('left');
    REPORT.setRowHeight(r, 22);
    r++;
  };

  // One KV row (with auto color rules)
  const writeKV = (label, valueOrFormula, fmtType, reasonText, isFormula) => {
    const isSplit = (r >= SPLIT_START && r <= SPLIT_END);

    // label
    REPORT.getRange(r, 1).setValue(label);

    // value
    if (isFormula) REPORT.getRange(r, 2).setFormula(valueOrFormula);
    else REPORT.getRange(r, 2).setValue(valueOrFormula);

    // narrative
    if (isSplit) {
      REPORT.getRange(r, 3).setValue(reasonText || '—');
    } else {
      REPORT.getRange(r, 2, 1, 2).merge();
    }

    // base row styling
    const bg = ((r % 2) === 0) ? P.BG_ROW_A : P.BG_ROW_B;
    REPORT.getRange(r, 1, 1, 3).setBackground(bg);

    REPORT.getRange(r, 1).setFontColor(P.MUTED).setFontWeight('bold').setHorizontalAlignment('left');
    REPORT.getRange(r, 2).setFontColor(P.TEXT).setFontWeight('bold').setHorizontalAlignment('left');

    if (isSplit) {
      REPORT.getRange(r, 3).setFontColor(P.TEXT).setFontWeight('normal').setHorizontalAlignment('left').setWrap(true);
    }

    REPORT.getRange(r, 1, 1, 3)
      .setBorder(false, false, true, false, false, false, P.GRID, SpreadsheetApp.BorderStyle.SOLID);

    // number formats (may reset font color)
    if (fmtType) applyNumberFormatByType___(REPORT.getRange(r, 2), fmtType);

    // VALUE-BASED COLORING (must come AFTER formatting)
    styleIndicatorValueCell___(REPORT.getRange(r, 2), label, f, bg);

    // merged cells revert to black → force readable text
    if (!isSplit) REPORT.getRange(r, 2, 1, 2).setFontColor(P.TEXT);

    // reason cell coloring
    if (isSplit) styleReasonCell___(REPORT.getRange(r, 3), label, f);
    if (isSplit) REPORT.getRange(r, 3).setFontColor(P.TEXT);

    // row height
    REPORT.setRowHeight(r, isSplit ? 34 : 18);

    r++;
  };

  // Narrative blocks
  const writeNarrative = (title, text) => {
    writeSectionHeader(title);
    REPORT.getRange(r, 1).setValue(text || '—');
    REPORT.getRange(r, 1, 1, 3).merge()
      .setBackground('#0E1624')
      .setFontColor(P.TEXT)
      .setVerticalAlignment('top')
      .setFontSize(10)
      .setWrap(true)
      .setHorizontalAlignment('left');
    REPORT.setRowHeight(r, 120);
    r += 2;
  };

  // SNAPSHOT
  writeSectionHeader('SNAPSHOT');
  writeKV('PRICE', f.Price === null ? '—' : f.Price, 'currency', reasonFor___('PRICE', f, interp), false);
  writeKV('CHG%', parsePercent___(f.ChangePct), 'percent', reasonFor___('CHG%', f, interp), false);
  writeKV('RVOL', f.RVOL === null ? '—' : f.RVOL, 'rvolx', reasonFor___('RVOL', f, interp), false);
  writeKV('TREND', interp.trendLabel || '—', 'text', reasonFor___('TREND', f, interp), false);
  writeKV('R:R', f.RR === null ? '—' : f.RR, 'rrx', reasonFor___('R:R', f, interp), false);

  // GOOGLEFINANCE formulas anchored to REPORT!A1
  writeKV('P/E', '=IFERROR(GOOGLEFINANCE($A$1,"pe"),"")', 'pe', '', true);
  writeKV('EPS', '=IFERROR(GOOGLEFINANCE($A$1,"eps"),"")', 'eps', '', true);

  // TREND & STRUCTURE
  writeSectionHeader('TREND & STRUCTURE');
  writeKV('SMA20', f.SMA20 === null ? '—' : f.SMA20, 'currency', reasonFor___('SMA20', f, interp), false);
  writeKV('SMA50', f.SMA50 === null ? '—' : f.SMA50, 'currency', reasonFor___('SMA50', f, interp), false);
  writeKV('SMA200', f.SMA200 === null ? '—' : f.SMA200, 'currency', reasonFor___('SMA200', f, interp), false);
  writeKV('ADX', f.ADX === null ? '—' : f.ADX, 'adx', reasonFor___('ADX', f, interp), false);
  writeKV('Stretch', interp.stretchLabel || '—', 'text', reasonFor___('Stretch', f, interp), false);
  writeKV('ATR', f.ATR === null ? '—' : f.ATR, 'atr', reasonFor___('ATR', f, interp), false);

  // MOMENTUM & TIMING
  writeSectionHeader('MOMENTUM & TIMING');
  writeKV('RSI', f.RSI === null ? '—' : f.RSI, 'rsi', reasonFor___('RSI', f, interp), false);
  writeKV('MACD Hist', f.MACDHist === null ? '—' : f.MACDHist, 'macd', reasonFor___('MACD Hist', f, interp), false);
  writeKV('Divergence', (String(f.Divergence || '').trim() ? f.Divergence : '—'), 'text', reasonFor___('Divergence', f, interp), false);

  // CRITICAL: values are 0..1, formatting handles % (no /100)
  writeKV('Stoch %K', f.StochK === null ? '—' : f.StochK, 'stoch', reasonFor___('Stoch %K', f, interp), false);
  writeKV('Bollinger %B', f.BBpct === null ? '—' : f.BBpct, 'bb', reasonFor___('Bollinger %B', f, interp), false);

  // LEVELS
  writeSectionHeader('LEVELS & PLANNING');
  writeKV('Support', f.Support === null ? '—' : f.Support, 'currency', reasonFor___('Support', f, interp), false);
  writeKV('Resistance', f.Resistance === null ? '—' : f.Resistance, 'currency', reasonFor___('Resistance', f, interp), false);
  writeKV('Target', f.Target === null ? '—' : f.Target, 'currency', reasonFor___('Target', f, interp), false);
  writeKV('ATH', f.ATH === null ? '—' : f.ATH, 'currency', reasonFor___('ATH', f, interp), false);
  writeKV('ATH Diff', parsePercent___(interp.athDiff), 'percent', reasonFor___('ATH Diff', f, interp), false);

  // WHY blocks
  writeNarrative('WHY THIS DECISION', interp.why);
  writeNarrative('WHY NOT THE ALTERNATIVES', interp.whyNot);

  // Borders + cosmetics
  const lastRow = REPORT.getLastRow();
  REPORT.getRange(1, 1, lastRow, 3)
    .setBorder(true, true, true, true, true, true, P.GRID, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  REPORT.setHiddenGridlines(true);
  REPORT.getRange(1, 1, lastRow, 3).setHorizontalAlignment('left');
}


/* =============================================================================
 * COLUMN WIDTHS + WRAP (your latest)
 * - A: ~8 chars
 * - B: ~8 chars
 * - C: ~4 words length (approx 200-250px)
 * ============================================================================= */
function setReportColumnWidthsAndWrap___(REPORT) {
  const pxPerChar = 8;

  const colA = Math.max(70, Math.round(8 * pxPerChar + 15));
  const colB = Math.max(70, Math.round(8 * pxPerChar + 15));
  const colC = Math.max(170, Math.round(27 * pxPerChar + 20));

  REPORT.setColumnWidth(1, colA);
  REPORT.setColumnWidth(2, colB);
  REPORT.setColumnWidth(3, colC);

  const lastRow = Math.max(1, REPORT.getLastRow());
  REPORT.getRange(1, 1, lastRow, 3).setWrap(true);
}


/* =============================================================================
 * NUMBER FORMATS
 * ============================================================================= */
function applyNumberFormatByType___(range, type) {
  switch (type) {
    case 'currency': range.setNumberFormat('$#,##0.00'); break;
    case 'percent':  range.setNumberFormat('0.00%'); break;
    case 'rvolx':    range.setNumberFormat('0.00"x"'); break;
    case 'rrx':      range.setNumberFormat('0.00"x"'); break;
    case 'adx':      range.setNumberFormat('0.00'); break;
    case 'rsi':      range.setNumberFormat('0.00'); break;
    case 'macd':     range.setNumberFormat('0.000'); break;

    // CRITICAL: expect 0..1; formatting shows 0..100%
    case 'stoch':    range.setNumberFormat('0.0%'); break;
    case 'bb':       range.setNumberFormat('0.0%'); break;

    case 'atr':      range.setNumberFormat('0.00'); break;
    case 'pe':       range.setNumberFormat('0.00'); break;
    case 'eps':      range.setNumberFormat('0.00'); break;
    case 'text':     break;
    default: break;
  }
}


/* =============================================================================
 * PALETTE + VALUE COLORING
 * ============================================================================= */
function reportPalette___() {
  return {
    BG_TOP: '#0B0F14',
    PANEL:  '#111827',
    BG_ROW_A: '#0F172A',
    BG_ROW_B: '#111827',
    GRID: '#374151',

    TEXT: '#E5E7EB',
    MUTED: '#9CA3AF',

    // Accent text (readable on dark chips)
    POS_TXT:  '#34D399',
    NEG_TXT:  '#F87171',
    WARN_TXT: '#FBBF24',

    // Dark chips (no white fills)
    CHIP_POS:  '#06281F',
    CHIP_NEG:  '#2A0B0B',
    CHIP_WARN: '#2A1E05',
    CHIP_NEU:  '#0B1220',

    YELLOW: '#FDE047',
    BLACK:  '#111827'
  };
}


function styleChipCell___(cell, rawText) {
  const P = reportPalette___();
  const t = String(rawText || cell.getDisplayValue() || '').trim().toUpperCase();

  // Default: dark neutral chip
  let bg = P.CHIP_NEU;
  let fg = P.TEXT;

  if (/BUY|BULL|ACCUM|RISK-ON|ADD|BREAKOUT|TRADE LONG/.test(t)) { bg = P.CHIP_POS; fg = P.POS_TXT; }
  else if (/SELL|BEAR|EXIT|STOP|RISK-OFF|REDUCE|AVOID/.test(t)) { bg = P.CHIP_NEG; fg = P.NEG_TXT; }
  else if (/EXPENSIVE|OVERVALUED|PRICED FOR PERFECTION/.test(t)) { bg = P.CHIP_WARN; fg = P.WARN_TXT; }
  else if (/FAIR|NEUTRAL|WAIT|MONITOR|HOLD/.test(t)) { bg = P.CHIP_NEU; fg = P.TEXT; }

  cell
    .setBackground(bg)
    .setFontColor(fg)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');
}


/**
 * Apply indicator-aware coloring to the VALUE cell (Column B)
 */
function styleIndicatorValueCell___(cell, label, f, rowBg) {
  const P = reportPalette___();
  const L = String(label || '').trim().toUpperCase();

  // IMPORTANT: Always anchor to a dark base background (no "null" backgrounds).
  const baseBg = rowBg || P.BG_ROW_A;
  cell.setBackground(baseBg).setFontColor(P.TEXT).setFontWeight('bold');

  const setChip = (chipBg, chipFg) => {
    cell.setBackground(chipBg || baseBg);
    cell.setFontColor(chipFg || P.TEXT);
  };

  const has = (v) => v !== null && v !== undefined && !isNaN(v);
  const num = (v) => (has(v) ? Number(v) : null);

  if (L === 'CHG%' || L === 'CHANGE %') {
    const p = parsePercent___(f.ChangePct);
    if (typeof p !== 'number' || isNaN(p)) return;
    if (p > 0) return setChip(P.CHIP_POS, P.POS_TXT);
    if (p < 0) return setChip(P.CHIP_NEG, P.NEG_TXT);
    return setChip(P.CHIP_NEU, P.MUTED);
  }

  if (L === 'RVOL') {
    const r = num(f.RVOL);
    if (!has(r)) return;
    if (r >= 1.5) return setChip(P.CHIP_POS, P.POS_TXT);
    if (r >= 1.0) return setChip(P.CHIP_NEU, P.TEXT);
    return setChip(P.CHIP_WARN, P.WARN_TXT);
  }

  if (L === 'R:R' || L === 'R:R QUALITY') {
    const rr = num(f.RR);
    if (!has(rr)) return;
    if (rr >= 3) return setChip(P.CHIP_POS, P.POS_TXT);
    if (rr >= 1.5) return setChip(P.CHIP_WARN, P.WARN_TXT);
    return setChip(P.CHIP_NEG, P.NEG_TXT);
  }

  if (L === 'RSI') {
    const rsi = num(f.RSI);
    if (!has(rsi)) return;
    if (rsi >= 70) return setChip(P.CHIP_WARN, P.WARN_TXT); // overbought
    if (rsi <= 30) return setChip(P.CHIP_POS, P.POS_TXT);   // oversold
    if (rsi >= 55) return setChip(P.CHIP_POS, P.POS_TXT);
    if (rsi <= 45) return setChip(P.CHIP_WARN, P.WARN_TXT);
    return setChip(P.CHIP_NEU, P.TEXT);
  }

  if (L === 'ADX') {
    const adx = num(f.ADX);
    if (!has(adx)) return;
    if (adx >= 25) return setChip(P.CHIP_POS, P.POS_TXT);
    if (adx >= 20) return setChip(P.CHIP_WARN, P.WARN_TXT);
    if (adx >= 15) return setChip(P.CHIP_NEU, P.TEXT);
    return setChip(P.CHIP_NEU, P.MUTED);
  }

  if (L === 'MACD HIST') {
    const m = num(f.MACDHist);
    if (!has(m)) return;
    if (m > 0) return setChip(P.CHIP_POS, P.POS_TXT);
    if (m < 0) return setChip(P.CHIP_NEG, P.NEG_TXT);
    return setChip(P.CHIP_NEU, P.MUTED);
  }

  if (L === 'STOCH %K' || L === 'STOCH %K (14)') {
    const k = num(f.StochK); // 0..1
    if (!has(k)) return;
    if (k >= 0.8) return setChip(P.CHIP_WARN, P.WARN_TXT);
    if (k <= 0.2) return setChip(P.CHIP_POS, P.POS_TXT);
    return setChip(P.CHIP_NEU, P.TEXT);
  }

  if (L === 'BOLLINGER %B') {
    const b = num(f.BBpct); // 0..1 (can exceed)
    if (!has(b)) return;
    if (b > 1 || b >= 0.8) return setChip(P.CHIP_WARN, P.WARN_TXT);
    if (b < 0 || b <= 0.2) return setChip(P.CHIP_POS, P.POS_TXT);
    return setChip(P.CHIP_NEU, P.TEXT);
  }

  if (L === 'SMA20' || L === 'SMA50' || L === 'SMA200') {
    const price = num(f.Price);
    const ma = (L === 'SMA20') ? num(f.SMA20) : (L === 'SMA50') ? num(f.SMA50) : num(f.SMA200);
    if (!has(price) || !has(ma)) return;
    if (price >= ma) return setChip(P.CHIP_POS, P.POS_TXT);
    return setChip(P.CHIP_NEG, P.NEG_TXT);
  }

  if (L === 'SUPPORT') {
    const price = num(f.Price), sup = num(f.Support);
    if (!has(price) || !has(sup)) return;
    if (price < sup) return setChip(P.CHIP_NEG, P.NEG_TXT);
    if ((price / sup - 1) <= 0.01) return setChip(P.CHIP_WARN, P.WARN_TXT);
    return setChip(P.CHIP_NEU, P.TEXT);
  }

  if (L === 'RESISTANCE') {
    const price = num(f.Price), res = num(f.Resistance);
    if (!has(price) || !has(res)) return;
    if (Math.abs(price / res - 1) <= 0.01) return setChip(P.CHIP_WARN, P.WARN_TXT);
    if (price > res) return setChip(P.CHIP_POS, P.POS_TXT);
    return setChip(P.CHIP_NEU, P.TEXT);
  }

  if (L === 'TARGET') {
    const price = num(f.Price), tgt = num(f.Target);
    if (!has(price) || !has(tgt)) return;
    if (price >= tgt) return setChip(P.CHIP_POS, P.POS_TXT);
    return setChip(P.CHIP_NEU, P.TEXT);
  }

  if (L === 'ATH' || L === 'ATH DIFF') {
    const ath = num(f.ATH), price = num(f.Price);
    if (!has(ath) || !has(price) || ath === 0) return;
    const diff = price / ath - 1;
    if (diff >= 0) return setChip(P.CHIP_POS, P.POS_TXT);
    if (diff >= -0.05) return setChip(P.CHIP_WARN, P.WARN_TXT);
    return setChip(P.CHIP_NEG, P.NEG_TXT);
  }

  // Default: keep dark base background, no bright fills
}


/**
 * Optional: tint the REASON cell (Column C) based on the same indicator signal
 */
function styleReasonCell___(cell, label, f) {
  const P = reportPalette___();
  const L = String(label || '').trim().toUpperCase();
  cell.setFontColor(P.TEXT);

  if (L === 'CHG%' || L === 'CHANGE %') {
    const p = parsePercent___(f.ChangePct);
    if (typeof p === 'number' && !isNaN(p)) {
      if (p > 0) cell.setFontColor(P.POS);
      else if (p < 0) cell.setFontColor(P.NEG);
    }
    return;
  }

  if (L === 'STOCH %K' || L === 'STOCH %K (14)') {
    const k = f.StochK;
    if (k !== null && !isNaN(k)) {
      if (k >= 0.8) cell.setFontColor(P.WARN);
      else if (k <= 0.2) cell.setFontColor(P.POS);
    }
    return;
  }

  if (L === 'BOLLINGER %B') {
    const b = f.BBpct;
    if (b !== null && !isNaN(b)) {
      if (b > 1 || b >= 0.8) cell.setFontColor(P.WARN);
      else if (b < 0 || b <= 0.2) cell.setFontColor(P.POS);
    }
  }
}


/* =============================================================================
 * VALUE-BASED NARRATIVES (NO GENERIC DEFINITIONS)
 * ============================================================================= */
function reasonFor___(label, f, interp) {
  const has = (n) => n !== null && n !== undefined && !isNaN(n);
  const price = has(f.Price) ? f.Price : null;

  switch (String(label || '').trim()) {
    case 'PRICE':
      return has(f.Price) ? `Last price ${money___(f.Price)}.` : '—';

    case 'CHG%': {
      const p = parsePercent___(f.ChangePct);
      if (typeof p !== 'number' || isNaN(p)) return '—';
      const pct = p * 100;
      if (pct > 0) return `Up ${pct.toFixed(2)}% today.`;
      if (pct < 0) return `Down ${Math.abs(pct).toFixed(2)}% today.`;
      return 'Flat today.';
    }

    case 'RVOL': {
      if (!has(f.RVOL)) return '—';
      const r = f.RVOL;
      if (r >= 1.5) return `strong participation.`;
      if (r >= 1.0) return `average participation.`;
      return `low participation (drift/chop risk).`;
    }

    case 'TREND': {
      const t = String(interp.trendLabel || '').trim();
      return t || '—';
    }

    case 'R:R': {
      if (!has(f.RR)) return '—';
      const rr = f.RR;
      if (rr >= 3) return `elite asymmetry.`;
      if (rr >= 1.5) return `acceptable asymmetry.`;
      return `weak asymmetry.`;
    }

    case 'SMA20':
      return movingAverageReason___('SMA20', price, f.SMA20);
    case 'SMA50':
      return movingAverageReason___('SMA50', price, f.SMA50);
    case 'SMA200':
      return movingAverageReason___('SMA200', price, f.SMA200, true);

    case 'ADX': {
      if (!has(f.ADX)) return '—';
      const a = f.ADX;
      if (a >= 25) return `strong trend.`;
      if (a >= 20) return `trend developing.`;
      if (a >= 15) return `weak trend.`;
      return `range-bound.`;
    }

    case 'Stretch': {
      const s = String(interp.stretchLabel || '').trim();
      return s ? s : '—';
    }

    case 'ATR': {
      if (!has(f.ATR)) return '—';
      const atr = f.ATR;
      const atrPct = (has(f.Price) && f.Price !== 0) ? (atr / f.Price) * 100 : null;
      const adx = has(f.ADX) ? f.ADX : null;
      const rvol = has(f.RVOL) ? f.RVOL : null;

      if (adx !== null && adx < 15 && atrPct !== null && atrPct < 1.5) return `volatility compressed; range-bound.`;
      if (adx !== null && adx >= 20 && rvol !== null && rvol >= 1.5) return `volatility expanding with confirmation.`;
      if (atrPct !== null && atrPct >= 2.5 && (rvol === null || rvol < 1.0)) return `elevated volatility; weak confirmation (false moves risk).`;
      if (adx !== null && adx >= 15 && adx < 20) return `volatility picking up; trend may be forming.`;
      return `normal volatility; no active expansion.`;
    }

    case 'RSI': {
      if (!has(f.RSI)) return '—';
      const rsi = f.RSI;
      if (rsi >= 70) return `overbought.`;
      if (rsi <= 30) return `oversold.`;
      if (rsi >= 55) return `positive momentum.`;
      if (rsi <= 45) return `weak momentum.`;
      return `neutral.`;
    }

    case 'MACD Hist': {
      if (!has(f.MACDHist)) return '—';
      const m = f.MACDHist;
      if (m > 0) return `positive impulse.`;
      if (m < 0) return `negative impulse.`;
      return 'Hist ~0: flat impulse.';
    }

    case 'Divergence': {
      const d = String(f.Divergence || '').trim();
      return d ? d : '—';
    }

    // FIXED: 0..1 thresholds, printed as percent
    case 'Stoch %K': {
      if (!has(f.StochK)) return '—';
      const k = f.StochK;
      const kPct = k * 100;
      if (k >= 0.8) return `overbought timing.`;
      if (k <= 0.2) return `oversold timing.`;
      return `neutral timing.`;
    }

    // FIXED: 0..1 with overshoots allowed
    case 'Bollinger %B': {
      if (!has(f.BBpct)) return '—';
      const b = f.BBpct;
      const bPct = b * 100;

      if (b > 1) return `above upper band (expansion / overextension).`;
      if (b >= 0.8) return `upper-band zone.`;
      if (b < 0) return `below lower band (statistical extreme).`;
      if (b <= 0.2) return `lower-band zone.`;
      return `mid-band zone.`;
    }

    case 'Support': {
      if (!has(f.Support) || !has(price)) return '—';

      // % downside vs current price (industry standard)
      const diffPct = ((price / f.Support) - 1) * 100;
      const sign = diffPct >= 0 ? '+' : '';

      return `${sign}${diffPct.toFixed(2)}% vs price ${money___(f.Price)}.`;
    }


   case 'Resistance': {
      if (!has(f.Resistance) || !has(price)) return '—';

      // % distance vs current price (industry standard)
      const diffPct = ((f.Resistance / price) - 1) * 100;
      const sign = diffPct >= 0 ? '+' : '';

      return `${sign}${diffPct.toFixed(2)}% vs price ${money___(f.Price)}.`;
    }


    case 'Target': {
      if (!has(f.Target) || !has(price)) return '—';
      const diffPct = ((f.Target / price) - 1) * 100;
      return `Target ${money___(f.Target)} (${diffPct >= 0 ? '+' : ''}${diffPct.toFixed(2)}%).`;
    }

    case 'ATH': {
      if (!has(f.ATH) || !has(price)) return '—';
      const diffPct = ((price / f.ATH) - 1) * 100;
      return `ATH ${money___(f.ATH)} (${diffPct >= 0 ? '+' : ''}${diffPct.toFixed(2)}%).`;
    }

    case 'ATH Diff': {
      const p = parsePercent___(interp.athDiff);
      if (typeof p !== 'number' || isNaN(p)) return '—';
      const pct = p * 100;
      return `${pct >= 0 ? '+' : ''}${pct.toFixed(2)}% vs ATH.`;
    }

    default:
      return '—';
  }
}

function movingAverageReason___(name, price, ma, isRegime) {
  if (price === null || price === undefined || isNaN(price)) return '—';
  if (ma === null || ma === undefined || isNaN(ma)) return '—';

  const side = (price >= ma) ? '>' : '<';
  const diffPct = ((price / ma) - 1) * 100;

  if (isRegime) {
    const regime = (price >= ma) ? 'Risk-On' : 'Risk-Off';
    return `${regime}: price ${price} ${side} ${name} , ${diffPct >= 0 ? '+' : ''}${diffPct.toFixed(2)}%.`;
  }
  return `Price ${price} ${side} ${name} , ${diffPct >= 0 ? '+' : ''}${diffPct.toFixed(2)}%.`;
}


/* =============================================================================
 * NO-DATA FALLBACK
 * ============================================================================= */
function renderNoDataReport___(REPORT, INPUT, ticker, msg) {
  REPORT.clear({ contentsOnly: true });
  REPORT.clearFormats();

  setupReportTickerDropdown_(REPORT, INPUT);
  REPORT.getRange('A1').setValue(ticker || '');

  setReportColumnWidthsAndWrap___(REPORT);

  REPORT.getRange('A1:C1').setFontWeight('bold').setHorizontalAlignment('left');
  REPORT.getRange('A3').setValue('STATUS');
  REPORT.getRange('B3').setValue('NO DATA');
  REPORT.getRange('B3:C3').merge();
  REPORT.getRange('A4').setValue('DETAIL');
  REPORT.getRange('B4').setValue(msg);
  REPORT.getRange('B4:C4').merge();
}


/* =============================================================================
 * GENERIC HELPERS
 * ============================================================================= */
function toNum___(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number') return isNaN(v) ? null : v;
  const s = String(v).trim();
  if (!s) return null;
  const cleaned = s.replace(/[%,$\s]/g, '').replace(/x$/i, '');
  const n = parseFloat(cleaned);
  return isNaN(n) ? null : n;
}

function money___(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return '$' + Number(n).toFixed(2);
}

/**
 * Returns decimal percent:
 * -0.0031 for -0.31%
 */
function parsePercent___(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return (Math.abs(v) <= 1) ? v : (v / 100);

  const s = String(v).trim();
  if (!s) return '';
  const m = s.match(/-?\d+(\.\d+)?/);
  if (!m) return '';
  const num = parseFloat(m[0]);
  if (/%/.test(s)) return num / 100;
  if (Math.abs(num) > 1) return num / 100;
  return num;
}

function isSheetError___(s) {
  const t = String(s || '').trim().toUpperCase();
  return t === '#ERROR!' || t === '#N/A' || t === '#REF!' || t === '#VALUE!' || t === '#DIV/0!' || t === '#NAME?';
}

function rvolComment___(r) {
  if (r === null) return 'participation unknown';
  if (r >= 1.5) return 'high participation (confirmation supportive)';
  if (r >= 1.0) return 'average participation';
  return 'low participation (drift/chop risk)';
}

function adxComment___(a) {
  if (a === null) return 'trend strength unknown';
  if (a >= 25) return 'strong trend';
  if (a >= 20) return 'trend developing';
  if (a >= 15) return 'weak trend';
  return 'range-bound';
}

function rsiComment___(r) {
  if (r === null) return 'momentum unknown';
  if (r >= 70) return 'overbought';
  if (r <= 30) return 'oversold';
  return 'neutral';
}

function stochComment01___(k01) {
  if (k01 === null) return 'timing unknown';
  if (k01 >= 0.8) return 'overbought timing';
  if (k01 <= 0.2) return 'oversold timing';
  return 'neutral timing';
}

function bbComment01___(b01) {
  if (b01 === null) return 'position unknown';
  if (b01 > 1) return 'above upper band';
  if (b01 >= 0.8) return 'upper-band zone';
  if (b01 < 0) return 'below lower band';
  if (b01 <= 0.2) return 'lower-band zone';
  return 'mid-band zone';
}

function rrComment___(rr) {
  if (rr === null) return 'asymmetry unknown';
  if (rr >= 3) return 'elite';
  if (rr >= 1.5) return 'acceptable';
  return 'weak';
}

/*
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== "REPORT") return;

  if (range.getA1Notation() === "A1") {
    SpreadsheetApp.getActive().toast("Refreshing Mobile Dashboard...", "⚙️ TERMINAL", 3);
    try {
      generateMobileReport();
      SpreadsheetApp.flush();
    } catch (err) {
      const msg = err?.message || String(err);
      SpreadsheetApp.getActive().toast("Mobile report failed: " + msg, "⚠️ onEdit", 8);
      console.error(err);
    }
  }
}
*/
