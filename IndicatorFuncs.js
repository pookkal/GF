/*
  Replacements for: LIVERSI, LIVEMACD, LIVEADX (hardening), plus new LIVEATR
  Drop these into Code.js to replace the originals.
*/

function LIVERSI(history, currentPrice, period = 14) {
  // Wilder RSI (smoothed) with defensive parsing.
  if (!history) return 50;
  const raw = history.flat().map(v => {
    const n = Number(v);
    return isFinite(n) ? n : NaN;
  }).filter(v => !Number.isNaN(v) && v > 0); // keep >0 closes

  // Append current price if provided and valid
  const cp = Number(currentPrice);
  if (isFinite(cp) && cp > 0) raw.push(cp);

  if (raw.length < period + 1) return 50; // not enough data

  // Ensure chronological order (oldest -> newest). The caller uses slices where row 5.. are oldest first.
  // We'll work with raw as chronological.
  const closes = raw.slice(-Math.max(60, period * 5)); // reasonable cap

  // Initial avg gain/loss = simple average of first 'period' changes
  let gains = 0, losses = 0;
  for (let i = 1; i <= period; i++) {
    const change = closes[i] - closes[i - 1];
    if (change > 0) gains += change; else losses += Math.abs(change);
  }
  let avgGain = gains / period;
  let avgLoss = losses / period;

  // Wilder smoothing for subsequent values
  for (let i = period + 1; i < closes.length; i++) {
    const change = closes[i] - closes[i - 1];
    const gain = change > 0 ? change : 0;
    const loss = change < 0 ? Math.abs(change) : 0;
    avgGain = ((avgGain * (period - 1)) + gain) / period;
    avgLoss = ((avgLoss * (period - 1)) + loss) / period;
  }

  if (avgLoss <= 0) return 100;
  const rs = avgGain / avgLoss;
  const rsi = 100 - (100 / (1 + rs));
  return Number(rsi.toFixed(2));
}

function LIVEMACD(history, currentPrice) {
  // MACD (12,26,9) with SMA seed for EMA (industry standard)
  if (!history && !currentPrice) return 0;

  const raw = (history || []).flat().map(v => {
    const n = Number(v);
    return isFinite(n) ? n : NaN;
  }).filter(v => !Number.isNaN(v) && v > 0);

  const cp = Number(currentPrice);
  if (isFinite(cp) && cp > 0) raw.push(cp);

  // need at least 26 points to compute EMA26
  if (raw.length < 26) return 0;

  function calcEMAWithSMASeed(data, period) {
    // data: chronological oldest..newest
    const out = [];
    // seed: SMA of first 'period' values
    const seedSlice = data.slice(0, period);
    const seed = seedSlice.reduce((s, x) => s + x, 0) / period;
    out[period - 1] = seed;
    const k = 2 / (period + 1);
    // compute EMA starting from index=period
    for (let i = period; i < data.length; i++) {
      const prev = out[i - 1];
      const ema = (data[i] * k) + (prev * (1 - k));
      out[i] = ema;
    }
    // For indices < period-1, fill with nulls to keep indexing simple
    for (let i = 0; i < period - 1; i++) out[i] = null;
    return out;
  }

  const ema12 = calcEMAWithSMASeed(raw, 12);
  const ema26 = calcEMAWithSMASeed(raw, 26);

  // macdLine only valid where both ema12 and ema26 exist
  const macdLine = raw.map((_, i) => {
    const a = ema12[i], b = ema26[i];
    return (a !== null && b !== null) ? (a - b) : null;
  }).filter(v => v !== null);

  if (macdLine.length < 9) return 0;

  // Signal: EMA(9) of macdLine with SMA seed
  const signalArray = (function calcSignal(arr, period = 9) {
    // arr is macdLine chronological
    const out = [];
    const seed = arr.slice(0, period).reduce((s, x) => s + x, 0) / period;
    out[period - 1] = seed;
    const k = 2 / (period + 1);
    for (let i = period; i < arr.length; i++) {
      const prev = out[i - 1];
      out[i] = (arr[i] * k) + (prev * (1 - k));
    }
    // return array aligned to arr indices (indices < period-1 are null)
    for (let i = 0; i < period - 1; i++) out[i] = null;
    return out;
  })(macdLine, 9);

  // final indices: use last valid macd and last valid signal
  const lastMacd = macdLine[macdLine.length - 1];
  const lastSignal = signalArray[signalArray.length - 1];

  if (!isFinite(lastMacd) || !isFinite(lastSignal)) return 0;
  return Number((lastMacd - lastSignal).toFixed(3));
}

function LIVEATR(highHist, lowHist, closeHist, currentPrice, period = 14) {
  // True Range ATR (Wilder) — returns numeric ATR (not percentage)
  if (!highHist || !lowHist || !closeHist) return 0;

  const Hraw = highHist.flat().map(v => Number(v)).filter(n => isFinite(n) && n > 0);
  const Lraw = lowHist.flat().map(v => Number(v)).filter(n => isFinite(n) && n > 0);
  const Craw = closeHist.flat().map(v => Number(v)).filter(n => isFinite(n) && n > 0);

  const m = Math.min(Hraw.length, Lraw.length, Craw.length);
  if (m < period + 1) return 0;

  const h = [], l = [], c = [];
  for (let i = 0; i < m; i++) {
    const hi = Hraw[i], lo = Lraw[i], cl = Craw[i];
    if (isFinite(hi) && isFinite(lo) && isFinite(cl) && hi > 0 && lo > 0 && cl > 0 && hi >= lo) {
      h.push(hi); l.push(lo); c.push(cl);
    }
  }

  const n = h.length;
  if (n < period + 1) return 0;

  // limit for speed
  const take = Math.min(n, 260);
  const H = h.slice(n - take);
  const L = l.slice(n - take);
  const C = c.slice(n - take);

  const live = Number(currentPrice);
  if (isFinite(live) && live > 0) C[C.length - 1] = live;

  const tr = [];
  for (let i = 1; i < C.length; i++) {
    const r1 = H[i] - L[i];
    const r2 = Math.abs(H[i] - C[i - 1]);
    const r3 = Math.abs(L[i] - C[i - 1]);
    const trValue = Math.max(r1, r2, r3);
    if (isFinite(trValue) && trValue > 0) {
      tr.push(trValue);
    }
  }

  if (tr.length < period) return 0;

  // initial ATR = average of first 'period' TRs
  let atr = tr.slice(0, period).reduce((a, b) => a + b, 0) / period;

  // Wilder smoothing
  for (let i = period; i < tr.length; i++) {
    atr = ((atr * (period - 1)) + tr[i]) / period;
  }

  return Number(atr.toFixed(2));
}

// Replacement LIVESTOCHK — robust, optional smoothing
function LIVESTOCHK(highHist, lowHist, closeHist, currentPrice, period = 14, smoothK = 1) {
  try {
    // Basic validation
    if (!highHist || !lowHist || !closeHist) return 0.5;

    // Helper: coerce to finite numbers
    const toNum = (v) => {
      if (typeof v === 'number') return isFinite(v) ? v : NaN;
      if (v === null || v === undefined) return NaN;
      const s = String(v).trim();
      if (s === "") return NaN;
      const n = Number(s);
      return isFinite(n) ? n : NaN;
    };

    // Flatten and coerce arrays; preserve chronological order (oldest -> newest)
    const Hraw = highHist.flat().map(toNum).filter(n => !Number.isNaN(n));
    const Lraw = lowHist.flat().map(toNum).filter(n => !Number.isNaN(n));
    const Craw = closeHist.flat().map(toNum).filter(n => !Number.isNaN(n));

    const n = Math.min(Hraw.length, Lraw.length, Craw.length);
    if (n < Math.max(period, 5)) return 0.5; // not enough data

    // Align arrays to the same length window
    const start = Math.max(0, n - Math.max(period, smoothK));
    const H = Hraw.slice(Hraw.length - n); // last n
    const L = Lraw.slice(Lraw.length - n);
    const C = Craw.slice(Craw.length - n);

    // Determine close to use: prefer provided currentPrice (live), else last close
    const cp = toNum(currentPrice);
    const close = (isFinite(cp) ? cp : (C.length ? C[C.length - 1] : NaN));
    if (!isFinite(close)) return 0.5;

    // Compute %K over the last `period` bars (use last `period` highs/lows)
    const useCount = Math.min(period, H.length, L.length);
    const hWindow = H.slice(H.length - useCount, H.length);
    const lWindow = L.slice(L.length - useCount, L.length);

    const hh = Math.max(...hWindow);
    const ll = Math.min(...lWindow);

    if (!isFinite(hh) || !isFinite(ll) || hh === ll) return 0.5;

    let k = (close - ll) / (hh - ll);
    k = Math.max(0, Math.min(1, k));

    // Optional smoothing of %K (simple SMA over last smoothK periods).
    if (smoothK > 1) {
      // Build an array of raw %K values for the last `smoothK` offsets (if available)
      const kVals = [];
      // For each offset (from smoothK-1 back to 0), compute %K using window ending at that offset
      for (let offset = smoothK - 1; offset >= 0; offset--) {
        const idxEnd = H.length - 1 - offset;
        const idxStart = idxEnd - useCount + 1;
        if (idxStart < 0) continue;
        const subH = H.slice(idxStart, idxEnd + 1);
        const subL = L.slice(idxStart, idxEnd + 1);
        if (subH.length < useCount || subL.length < useCount) continue;
        const hhSub = Math.max(...subH);
        const llSub = Math.min(...subL);
        if (!isFinite(hhSub) || !isFinite(llSub) || hhSub === llSub) continue;
        // choose close value for that offset: use historical close if available, else current close
        const cVal = (C.length > idxEnd && isFinite(C[idxEnd])) ? C[idxEnd] : close;
        const kSub = Math.max(0, Math.min(1, (cVal - llSub) / (hhSub - llSub)));
        kVals.push(kSub);
      }
      if (kVals.length > 0) {
        const sum = kVals.reduce((s, x) => s + x, 0);
        k = sum / kVals.length;
      }
    }

    return Number(k.toFixed(4));
  } catch (e) {
    return 0.5;
  }
}


// Robust replacement LIVEADX (defensive, lower thresholds)
function LIVEADX(highHist, lowHist, closeHist, currentPrice) {
  try {
    // Basic validation
    if (!highHist || !lowHist || !closeHist) return 0;

    const toNum = (v) => {
      if (typeof v === "number") return isFinite(v) ? v : NaN;
      if (v === null || v === undefined) return NaN;
      const s = String(v).trim();
      if (s === "" || s === "No Data") return NaN;
      const n = Number(s);
      return isFinite(n) ? n : NaN;
    };

    // Flatten & coerce; expect chronological order (oldest->newest)
    const Hraw = highHist.flat().map(toNum).filter(n => !Number.isNaN(n));
    const Lraw = lowHist.flat().map(toNum).filter(n => !Number.isNaN(n));
    const Craw = closeHist.flat().map(toNum).filter(n => !Number.isNaN(n));

    const m = Math.min(Hraw.length, Lraw.length, Craw.length);
    // Relaxed minimum requirement: need at least ~20 bars to compute ADX usefully
    if (m < 20) return 0;

    // Keep last window for performance & stability
    const take = Math.min(m, 260);
    const H = Hraw.slice(Hraw.length - take);
    const L = Lraw.slice(Lraw.length - take);
    const C = Craw.slice(Craw.length - take);

    // Use currentPrice if provided & finite to overwrite last close
    const liveClose = toNum(currentPrice);
    if (isFinite(liveClose) && liveClose > 0) {
      C[C.length - 1] = liveClose;
    }

    const period = 14;
    const tr = [], pdm = [], ndm = [];

    for (let i = 1; i < C.length; i++) {
      const upMove = H[i] - H[i - 1];
      const downMove = L[i - 1] - L[i];
      const plusDM = (upMove > downMove && upMove > 0) ? upMove : 0;
      const minusDM = (downMove > upMove && downMove > 0) ? downMove : 0;

      const r1 = H[i] - L[i];
      const r2 = Math.abs(H[i] - C[i - 1]);
      const r3 = Math.abs(L[i] - C[i - 1]);
      const trueRange = Math.max(r1, r2, r3);

      if (!isFinite(trueRange) || trueRange <= 0) {
        // skip this bar instead of failing everything
        continue;
      }
      tr.push(trueRange);
      pdm.push(plusDM);
      ndm.push(minusDM);
    }

    if (tr.length < period) return 0;

    const safeDiv = (num, den) => (den > 1e-12 ? (num / den) : 0);

    // initial sums (first 'period' TR/DM)
    let atr = tr.slice(0, period).reduce((a, b) => a + b, 0);
    let pDM14 = pdm.slice(0, period).reduce((a, b) => a + b, 0);
    let nDM14 = ndm.slice(0, period).reduce((a, b) => a + b, 0);

    let pDI = 100 * safeDiv(pDM14, atr);
    let nDI = 100 * safeDiv(nDM14, atr);

    const dxArr = [];
    dxArr.push((pDI + nDI > 1e-12) ? (100 * Math.abs(pDI - nDI) / (pDI + nDI)) : 0);

    // Wilder smoothing for subsequent bars
    for (let i = period; i < tr.length; i++) {
      atr = atr - (atr / period) + tr[i];
      pDM14 = pDM14 - (pDM14 / period) + (pdm[i] || 0);
      nDM14 = nDM14 - (nDM14 / period) + (ndm[i] || 0);

      if (!isFinite(atr) || atr <= 0) continue;

      pDI = 100 * safeDiv(pDM14, atr);
      nDI = 100 * safeDiv(nDM14, atr);

      const dx = (pDI + nDI > 1e-12) ? (100 * Math.abs(pDI - nDI) / (pDI + nDI)) : 0;
      dxArr.push(isFinite(dx) ? dx : 0);
    }

    if (dxArr.length < period) return 0;

    // ADX initial average, then Wilder smoothing
    let adx = dxArr.slice(0, period).reduce((a, b) => a + b, 0) / period;
    for (let i = period; i < dxArr.length; i++) {
      adx = ((adx * (period - 1)) + dxArr[i]) / period;
    }

    return Number((isFinite(adx) ? adx : 0).toFixed(2));
  } catch (e) {
    return 0;
  }
}
