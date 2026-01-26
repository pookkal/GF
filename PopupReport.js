

/**
 * Core Terminal Function:
 * Finds the ticker in C1, pulls data from STOCK_ANALYZER_TERMINAL_BASE (GOLDEN),
 * and generates a high-detail Intelligence Report.
 */
function runAnalysisFromInput() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("CHART");
  
  // 1. CRITICAL: Update this name to match your GOLDEN baseline tab exactly
  const DATA_TAB_NAME = "CALCULATIONS"; 
  const dataSheet = ss.getSheetByName("CALCULATIONS");
  
  // 2. SAFETY CHECKS: Ensure sheets exist
  if (!inputSheet) {
    ss.toast("❌ Error: Sheet 'CHART' not found.", "TERMINAL FAIL", 5);
    return;
  }
  if (!dataSheet) {
    ss.toast(`❌ Error: Sheet '${DATA_TAB_NAME}' not found. Check tab name.`, "TERMINAL FAIL", 5);
    return;
  }

  // 3. GET TICKER: Pull from A1
  const ticker = inputSheet.getRange("A1").getValue();
  if (!ticker || ticker === "") {
    inputSheet.getRange("C1").setValue("⚠️ STANDBY: Enter a Ticker in A1 to begin analysis.");
    return;
  }

  // 4. FIND DATA: Search column A for the Ticker
  const data = dataSheet.getDataRange().getValues();
  let tickerRow = -1;
  
  for (let i = 0; i < data.length; i++) {
    // data[i][0] corresponds to Column A
    if (data[i][0].toString().toUpperCase() === ticker.toString().toUpperCase()) {
      tickerRow = i; 
      break;
    }
  }

  if (tickerRow === -1) {
    inputSheet.getRange("C1").setValue(`❌ ERROR: Ticker '${ticker}' not found in ${DATA_TAB_NAME}.`);
    return;
  }

  // 5. EXTRACT INDICATORS: Map rowData to variables (0-based indexing)
  // Mapping based on: A=0, B=1, C=2, D=3, E=4, F=5, G=6...
  const rowData = data[tickerRow];
  const d = {
    ticker:     rowData[0],  // Col A
    price:      rowData[4],  // Col E
    volRatio:   rowData[6],  // Col G
    sma50:      rowData[12], // Col M
    sma20:      rowData[13], // Col N
    sma200:     rowData[14], // Col O
    rsi:        rowData[15], // Col P
    macdHist:   rowData[16], // Col Q
    adx:        rowData[18], // Col S
    stochK:     rowData[19], // Col T
    support:    rowData[21], // Col V
    resistance: rowData[22], // Col W
    target:     rowData[23], // Col X
    atr:        rowData[24], // Col Y
    bolB:       rowData[27], // Col AB
    status:     rowData[3],  // Col D - DECISION (swapped)
    valuation:  rowData[2],  // Col C - FUNDAMENTAL (swapped)
    stopLoss:   rowData[21]  // Mapping Support to StopLoss logic
  };

  // 6. RUN ANALYSIS: Calculate Signal and Recommendation
  try {

    const d = getTickerDataFromBaseline(ticker);
    if (!d) return;

    const narrative = MasterAnalysisEngine.analyze(d);

    // 7. OUTPUT: Write back to INPUT sheet
    const outputRange = inputSheet.getRange("C1");
    outputRange.setValue(narrative);
    
    // Formatting for terminal look
    outputRange.setWrap(true);
    outputRange.setVerticalAlignment("top");
    
    ss.toast(`Intelligence Report for ${ticker} Complete.`, "✅ SYNCED", 3);
    
  } catch (err) {
    inputSheet.getRange("C1").setValue(`⚙️ SCRIPT ERROR: ${err.message}`);
    console.error(err.stack);
  }
}

/**
 * MASTER INSTITUTIONAL INTELLIGENCE ENGINE v8.0
 * Integrated Volume Intelligence & Smart Money Analysis
 */
class MasterAnalysisEngineText {
  static analyze(d) {
    const volRatio = parseFloat(d.volRatio) || 0;
    const priceChange = parseFloat(d.changePct) || 0;
    const rsiVal = parseFloat(d.rsi) || 0;
    const adxVal = parseFloat(d.adx) || 0;

    // --- VOLUME INTELLIGENCE LOGIC ---
    let volumeNarrative = "";
    if (volRatio > 2.0 && priceChange < 0) {
      volumeNarrative = "⚠️ CAPITULATION VOLUME: Massive institutional selling detected. This is a 'Panic Flush'.";
    } else if (volRatio > 2.0 && priceChange > 0) {
      volumeNarrative = "🔥 ABSORPTION VOLUME: Institutional buyers are aggressively soaking up supply.";
    } else if (volRatio < 1.0) {
      volumeNarrative = "⚪ LOW CONVICTION: Trading on 'Retail Churn'. Large institutions are sitting on the sidelines.";
    } else {
      volumeNarrative = "✅ HEALTHY PARTICIPATION: Volume is tracking the 10-day average.";
    }

    let report = `◈━━━━━━━━ MASTER INSTITUTIONAL INTELLIGENCE ━━━━━━━━◈\n\n`;

    // SECTION 1: SYSTEM SENTIMENT
    report += `[🧭 SYSTEM SENTIMENT]\n`;
    report += `Status: ${d.price > d.sma200 ? "📈 STRUCTURAL BULLISH" : "📉 STRUCTURAL BEARISH"}\n`;
    report += `Signal: ${d.signal} | Vol Trend: ${volRatio.toFixed(2)}x\n`;
    report += `→ ${volumeNarrative}\n\n`;

    // SECTION 2: CONFLUENCE DEEP-DIVE
    report += `[🔬 CONFLUENCE OF EVIDENCE]\n`;
    
    // VOLUME X PRICE CONFLUENCE (The "So What")
    let confluenceLogic = "";
    if (priceChange < 0 && volRatio < 1.0) {
      confluenceLogic = "DIVERGENCE: Price is falling, but volume is drying up. This suggests the selling pressure is exhausting and a bounce is mathematically likely.";
    } else if (priceChange < 0 && volRatio > 1.5) {
      confluenceLogic = "CONFIRMED DOWNTREND: High volume accompanying the price drop confirms institutional exit. Do not buy the dip yet.";
    }
    report += `• SMART MONEY: ${confluenceLogic}\n`;

    // RSI & MACD SIGNIFICANCE (Industry Standard Explainer)
    report += `• MOMENTUM: RSI ${rsiVal.toFixed(1)} | MACD ${d.macdHist.toFixed(2)}\n`;
    report += `  → ${rsiVal < 30 ? "Institutional Buy Zone (Oversold Extremity)." : "Momentum searching for equilibrium."}\n`;

    // VOLATILITY & ADX
    report += `• VOLATILITY: ADX ${adxVal.toFixed(2)} (${adxVal > 30 ? "Strong Power Trend" : "Weak/Coiling Trend"})\n`;
    report += `  → High ADX with Low Volume (${volRatio}) indicates the trend is 'Automated' (Algorithmic) rather than human-driven.\n\n`;

    // SECTION 3: PRICE ARCHITECTURE
    report += `[🎯 PRICE ARCHITECTURE]\n`;
    report += `🛡️ Support: $${d.support.toFixed(2)} | 🚧 Resistance: $${d.resistance.toFixed(2)}\n`;
    report += `📊 Bollinger %B: ${(d.bolB * 100).toFixed(1)}% → ${d.bolB < 0.2 ? "Statistical Value Floor" : "Inside Distribution Range"}\n\n`;

    // SECTION 4: STRATEGIC VERDICT
    report += `[💡 STRATEGIC VERDICT]\n`;
    const isOversold = rsiVal < 25;
    report += `ACTION: ${isOversold ? "TACTICAL BOUNCE ENTRY (Wait for 5m Reversal)" : "REDUCE / HOLD"}\n`;
    report += `Narrative: Despite the BEAR state, the Low Volume (${volRatio}) and Extreme RSI (${rsiVal}) create a 'Vacuum' where any minor buying will cause a sharp snap-back to $${d.sma50.toFixed(2)}.\n\n`;

    // SECTION 5: FOOTER
    report += `[📝 ANALYTICS FOOTER]\n`;
    report += `• Structural Reason: ${d.price < d.sma200 ? "GRAVITY: Below 200-Day SMA." : "Trend Strength Failure."}\n`;
    report += `• ATH Distance: ${((d.athDiff < 1) ? d.athDiff * 100 : d.athDiff).toFixed(2)}%\n`;
    report += `◈━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━◈`;

    return report;
  }
}


/**
 * MASTER INSTITUTIONAL INTELLIGENCE ENGINE v10.0 (GOLDEN-aligned)
 * - Uses ONLY the fields returned by getTickerDataFromBaseline()
 * - Produces a professional, indicator-complete narrative with explicit "Why" and "Why Not"
 * - Output is HTML (for showAnalysisPopup right pane)
 * - Conditionally shows sections based on trigger source (REPORT sheet vs others)
 */
class MasterAnalysisEngine {
  static analyze(d, isFromReportSheet = false) {
    // ---------------------------
    // Helpers
    // ---------------------------
    const n = (v) => {
      if (typeof v === "number") return isFinite(v) ? v : null;
      if (v === null || v === undefined) return null;
      const s = String(v).replace(/[$,%\s,]/g, "").trim();
      if (s === "" || s === "—" || s === "-") return null;
      const x = Number(s);
      return isFinite(x) ? x : null;
    };

    const t = (v) => (v == null || String(v).trim() === "" ? "—" : String(v));
    const pct = (v, digits = 2) => {
      const x = n(v);
      if (x == null) return "—";
      // changePct/athDiff are already ratios (0.05 = 5%)
      return (x * 100).toFixed(digits) + "%";
    };
    const p2 = (v) => {
      const x = n(v);
      if (x == null) return "—";
      return "$" + x.toFixed(2);
    };
    const f2 = (v) => {
      const x = n(v);
      if (x == null) return "—";
      return x.toFixed(2);
    };
    const f3 = (v) => {
      const x = n(v);
      if (x == null) return "—";
      return x.toFixed(3);
    };

    // ---------------------------
    // Normalize key fields
    // ---------------------------
    const ticker     = t(d.ticker).toUpperCase();
    const signal     = t(d.signal);
    const fundamental= t(d.fundamental);  // FUNDAMENTAL bucket (VALUE/FAIR/EXPENSIVE/...)
    const decision   = t(d.decision);     // DECISION label (Trade Long / Hold / Take Profit / Reduce / Stop-Out / Avoid / ...)
    const price      = n(d.price) ?? 0;
    const chg        = n(d.changePct);
    const rvol       = n(d.volRatio);
    const ath        = n(d.isATH);
    const athDiff    = n(d.athDiff);
    const rr         = n(d.rrQuality);
    const trendState = t(d.trendState);

    const sma20      = n(d.sma20);
    const sma50      = n(d.sma50);
    const sma200     = n(d.sma200);

    const rsi        = n(d.rsi);
    const macd       = n(d.macdHist);
    const divg       = t(d.divergence);
    const adx        = n(d.adx);
    const stochK     = n(d.stochK);

    const sup        = n(d.support);
    const res        = n(d.resistance);
    const tgt        = n(d.target);
    const atr        = n(d.atr);
    const bb         = n(d.bolB);

    // ---------------------------
    // Derived diagnostics
    // ---------------------------
    const aboveSMA200 = (price > 0 && sma200 != null) ? price >= sma200 : null;
    const aboveSMA50  = (price > 0 && sma50  != null) ? price >= sma50  : null;
    const aboveSMA20  = (price > 0 && sma20  != null) ? price >= sma20  : null;

    const atrPct = (price > 0 && atr != null) ? (atr / price) : null;
    const stretchATR = (price > 0 && sma20 != null && atr != null && atr > 0)
      ? ((price - sma20) / atr)
      : null;

    const stochPct = (stochK != null) ? (stochK * 100) : null;
    const bbPct    = (bb != null) ? (bb * 100) : null;

    // ---------------------------
    // Classification (consistent with your sheet rules)
    // ---------------------------
    const regime =
      aboveSMA200 === null ? "UNKNOWN"
      : (aboveSMA200 ? "RISK-ON (Above SMA200)" : "RISK-OFF (Below SMA200)");

    const rsiBand =
      (rsi == null) ? "Unknown"
      : (rsi <= 30 ? "Oversold (tactical value)" : (rsi >= 70 ? "Overbought (pullback risk)" : "Neutral"));

    const macdBand =
      (macd == null) ? "Unknown"
      : (macd >= 0 ? "Positive impulse" : "Negative impulse");

    const adxBand =
      (adx == null) ? "Unknown"
      : (adx < 15 ? "Range/Chop" : (adx < 25 ? "Weak/Developing trend" : "Strong trend"));

    const stochBand =
      (stochK == null) ? "Unknown"
      : (stochK <= 0.2 ? "Oversold timing" : (stochK >= 0.8 ? "Overbought timing" : "Neutral timing"));

    const rvolBand =
      (rvol == null) ? "Unknown"
      : (rvol >= 1.5 ? "High participation (sponsorship)" : (rvol < 1.0 ? "Low participation (drift)" : "Normal participation"));

    const bbBand =
      (bb == null) ? "Unknown"
      : (bb <= 0.2 ? "Lower band zone (statistical low)" : (bb >= 0.8 ? "Upper band zone (statistical high)" : "Mid-band"));

    // ---------------------------
    // “Why” and “Why Not” (explicit, based on your SIGNAL hierarchy)
    // ---------------------------
    const why = [];
    const whyNot = [];

    // Primary invalidations / regime
    if (sup != null && price > 0 && price < sup) {
      why.push(`Price (${p2(price)}) is below Support (${p2(sup)}) → invalidation risk (Stop-Out condition).`);
    } else {
      whyNot.push(`Stop-Out not triggered: price is not below Support.`);
    }

    if (sma200 != null && price > 0 && price < sma200) {
      why.push(`Price (${p2(price)}) is below SMA200 (${p2(sma200)}) → Risk-Off regime, trend trades are structurally blocked.`);
    } else if (sma200 != null && price > 0) {
      why.push(`Price (${p2(price)}) is above SMA200 (${p2(sma200)}) → Risk-On regime, trend setups are structurally permitted.`);
    }

    // Volume / breakout sponsorship
    if (rvol != null) {
      if (rvol >= 1.5) {
        why.push(`RVOL ${f2(rvol)}x indicates strong participation (good for breakouts / continuation).`);
      } else if (rvol < 1.0) {
        why.push(`RVOL ${f2(rvol)}x indicates low participation (moves are less reliable; expect drift/chop).`);
        whyNot.push(`Breakout confirmation is weaker because RVOL < 1.5x.`);
      } else {
        why.push(`RVOL ${f2(rvol)}x is normal (no strong sponsorship signal).`);
      }
    }

    // Momentum
    if (rsi != null) {
      why.push(`RSI ${f2(rsi)} → ${rsiBand}.`);
      if (rsi <= 30) why.push(`RSI in oversold band supports tactical mean-reversion (only if structure holds above Support).`);
      if (rsi >= 70) why.push(`RSI in overbought band supports profit-taking/trim near Resistance/Target.`);
    }

    if (macd != null) {
      why.push(`MACD histogram ${f3(macd)} → ${macdBand}.`);
      if (macd < 0 && aboveSMA50 === false) {
        why.push(`MACD < 0 and price below SMA50 → momentum weakening (aligns with Reduce logic if PURCHASED).`);
      }
    }

    // Trend strength
    if (adx != null) {
      why.push(`ADX ${f2(adx)} → ${adxBand}.`);
      if (adx < 15) {
        whyNot.push(`Trend continuation signals are suppressed because ADX < 15 (range regime).`);
      }
    }

    // Timing oscillators
    if (stochK != null) {
      why.push(`Stoch %K ${(stochPct != null ? stochPct.toFixed(1) : "—")} % → ${stochBand}.`);
    }
    if (bb != null) {
      why.push(`Bollinger %B ${(bbPct != null ? bbPct.toFixed(1) : "—")} % → ${bbBand}.`);
    }

    // Levels / asymmetry
    if (res != null && price > 0) {
      const distToRes = (res > 0) ? ((res - price) / price) : null;
      if (distToRes != null) {
        if (distToRes <= 0.01) {
          why.push(`Price is within ~1% of Resistance (${p2(res)}) → supply zone; profit-taking risk increases.`);
          whyNot.push(`New long entries are less attractive when price is near Resistance unless RVOL is strong and breakout confirms.`);
        } else if (distToRes >= 0.05) {
          why.push(`Resistance (${p2(res)}) is meaningfully above price → room for upside if structure is supportive.`);
        }
      }
    }

    if (rr != null) {
      if (rr >= 3) why.push(`R:R ${f2(rr)}x is favorable (>=3).`);
      else if (rr >= 1.5) why.push(`R:R ${f2(rr)}x is acceptable but not elite (1.5–3).`);
      else {
        why.push(`R:R ${f2(rr)}x is weak (<1.5) → poor asymmetry for new exposure.`);
        whyNot.push(`Trade entries are de-prioritized because R:R is below threshold.`);
      }
    }

    // Overextension (Reduce Overextended)
    if (stretchATR != null) {
      if (stretchATR >= 2) {
        why.push(`Stretch vs SMA20 is ${stretchATR.toFixed(1)}x ATR (>=2x) → overextended; pullback risk elevated.`);
      }
    }

    // Fundamental risk flag
    const fundRisk =
      /ZOMBIE|PRICED FOR PERFECTION|EXPENSIVE/i.test(fundamental);

    // ---------------------------
    // Final verdict narrative (ties SIGNAL + FUNDAMENTAL + DECISION)
    // ---------------------------
    const actionTone = (() => {
      const up = String(decision).toUpperCase();
      if (/STOP|AVOID/.test(up)) return "bear";
      if (/TAKE PROFIT|REDUCE/.test(up)) return "trim";
      if (/TRADE LONG|ACCUMULATE|ADD IN DIP/.test(up)) return "bull";
      return "neutral";
    })();

    const verdictTitle =
      actionTone === "bull" ? "EXECUTION BIAS: LONG"
      : actionTone === "trim" ? "EXECUTION BIAS: TRIM / RISK-REDUCE"
      : actionTone === "bear" ? "EXECUTION BIAS: DEFENSIVE"
      : "EXECUTION BIAS: WAIT / MONITOR";

    const riskFlags = [];
    if (fundRisk && (/BREAKOUT|TREND CONTINUATION/i.test(signal))) {
      riskFlags.push("Momentum is constructive, but fundamentals are in a fragile bucket (valuation/earnings risk). Size down and take profits faster.");
    }
    if (aboveSMA200 === false && (/TREND CONTINUATION|BREAKOUT/i.test(signal))) {
      riskFlags.push("Risk-Off regime conflicts with directional trend trades. Prefer tactical setups only with strict invalidation.");
    }
    if (rvol != null && rvol < 1.0 && (/BREAKOUT/i.test(signal))) {
      riskFlags.push("Breakout narrative is under-sponsored (RVOL < 1.0). Treat as tentative until volume expands.");
    }

    // ---------------------------
    // Build Professional HTML
    // ---------------------------
    const chipClass = (tone) => {
      if (tone === "bull") return "chip chip-bull";
      if (tone === "trim") return "chip chip-trim";
      if (tone === "bear") return "chip chip-bear";
      return "chip chip-neutral";
    };

    const decisionChip = chipClass(actionTone);

    // If triggered from REPORT sheet, get D18 content and rows 7-8 for custom narrative
    let customNarrative = "";
    let marketRating = "";
    let consensusPrice = "";
    
    if (isFromReportSheet) {
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const reportSheet = ss.getSheetByName("REPORT");
        if (reportSheet) {
          // Get MARKET RATING from B7
          const b7Value = reportSheet.getRange("B7").getValue();
          if (b7Value && String(b7Value).trim() !== "" && String(b7Value).trim() !== "—") {
            marketRating = String(b7Value).trim();
          }
          
          // Get CONSENSUS PRICE from B8
          const b8Value = reportSheet.getRange("B8").getValue();
          if (b8Value && String(b8Value).trim() !== "" && String(b8Value).trim() !== "—") {
            consensusPrice = String(b8Value).trim();
          }
          
          // Get AI analysis from D18
          const d18Value = reportSheet.getRange("D18").getValue();
          if (d18Value && String(d18Value).trim() !== "") {
            customNarrative = String(d18Value).trim();
          }
        }
      } catch (err) {
        console.error("Error reading REPORT sheet data:", err);
      }
    }

    const html = `
      <div class="mi-wrap">
        <div class="mi-head">
          <div class="mi-head">
            <div class="mi-sub">
              <span class="mi-chip mi-chip-yellow">SIGNAL: ${t(signal)}</span>
              <span class="mi-chip mi-chip-yellow">FUNDAMENTAL: ${t(fundamental)}</span>
              <span class="mi-chip mi-chip-yellow">DECISION: ${t(decision)}</span>
            </div>
          </div>
          <div class="mi-sub">
            <span class="mi-ticker">${ticker}</span>
            <span class="mi-chip ${decisionChip}">${t(decision)}</span>
            <span class="mi-regime">${regime}</span>
          </div>
        </div>

        <div class="mi-grid">
          <div class="mi-card">
            <div class="mi-card-title">Snapshot</div>
            <div class="mi-kv">
              <div class="k">SIGNAL</div><div class="v">${t(signal)}</div>
              <div class="k">FUNDAMENTAL</div><div class="v">${t(fundamental)}</div>
              <div class="k">PRICE</div><div class="v">${p2(price)}</div>
              <div class="k">CHG%</div><div class="v">${chg == null ? "—" : pct(chg)}</div>
              <div class="k">RVOL</div><div class="v">${rvol == null ? "—" : (f2(rvol) + "x")} (${rvolBand})</div>
              <div class="k">TREND</div><div class="v">${t(trendState)}</div>
              <div class="k">R:R</div><div class="v">${rr == null ? "—" : (f2(rr) + "x")}</div>
            </div>
          </div>

          <div class="mi-card">
            <div class="mi-card-title">Trend & Structure</div>
            <div class="mi-kv">
              <div class="k">SMA20</div><div class="v">${sma20 == null ? "—" : p2(sma20)} ${aboveSMA20 === null ? "" : (aboveSMA20 ? " (price above)" : " (price below)")}</div>
              <div class="k">SMA50</div><div class="v">${sma50 == null ? "—" : p2(sma50)} ${aboveSMA50 === null ? "" : (aboveSMA50 ? " (price above)" : " (price below)")}</div>
              <div class="k">SMA200</div><div class="v">${sma200 == null ? "—" : p2(sma200)} ${aboveSMA200 === null ? "" : (aboveSMA200 ? " (risk-on)" : " (risk-off)")}</div>
              <div class="k">ADX</div><div class="v">${adx == null ? "—" : f2(adx)} (${adxBand})</div>
              <div class="k">Stretch</div><div class="v">${stretchATR == null ? "—" : (stretchATR.toFixed(1) + "x ATR vs SMA20")}</div>
              <div class="k">ATR</div><div class="v">${atr == null ? "—" : f2(atr)}${atrPct == null ? "" : (" (" + (atrPct*100).toFixed(2) + "% of price)")}</div>
            </div>
          </div>

          <div class="mi-card">
            <div class="mi-card-title">Momentum & Timing</div>
            <div class="mi-kv">
              <div class="k">RSI</div><div class="v">${rsi == null ? "—" : f2(rsi)} (${rsiBand})</div>
              <div class="k">MACD Hist</div><div class="v">${macd == null ? "—" : f3(macd)} (${macdBand})</div>
              <div class="k">Divergence</div><div class="v">${t(divg)}</div>
              <div class="k">Stoch %K</div><div class="v">${stochPct == null ? "—" : stochPct.toFixed(1) + "%"} (${stochBand})</div>
              <div class="k">Bollinger %B</div><div class="v">${bbPct == null ? "—" : bbPct.toFixed(1) + "%"} (${bbBand})</div>
            </div>
          </div>

          <div class="mi-card">
            <div class="mi-card-title">Levels & Planning</div>
            <div class="mi-kv">
              <div class="k">Support</div><div class="v">${sup == null ? "—" : p2(sup)}</div>
              <div class="k">Resistance</div><div class="v">${res == null ? "—" : p2(res)}</div>
              <div class="k">Target</div><div class="v">${tgt == null ? "—" : p2(tgt)}</div>
              <div class="k">ATH</div><div class="v">${ath == null ? "—" : p2(ath)}</div>
              <div class="k">ATH Diff</div><div class="v">${athDiff == null ? "—" : pct(athDiff)}</div>
            </div>
          </div>
        </div>

        ${isFromReportSheet && (marketRating || consensusPrice || customNarrative) ? `
        ${marketRating || consensusPrice ? `
        <div class="mi-section mi-analyst-section">
          <div class="mi-section-title">Analyst Consensus</div>
          <div class="mi-analyst-grid">
            ${marketRating ? `
            <div class="mi-analyst-item">
              <div class="mi-analyst-label">Market Rating</div>
              <div class="mi-analyst-value mi-rating-${marketRating.toLowerCase().replace(/\s+/g, '-')}">${marketRating}</div>
            </div>
            ` : ''}
            ${consensusPrice ? `
            <div class="mi-analyst-item">
              <div class="mi-analyst-label">12-Month Target</div>
              <div class="mi-analyst-value mi-price-target">${consensusPrice}</div>
            </div>
            ` : ''}
          </div>
        </div>
        ` : ''}
        ${customNarrative ? `
        <div class="mi-section">
          <div class="mi-section-title">Investment Analysis</div>
          <div class="mi-narrative">${customNarrative.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>').replace(/\*([^*]+)\*/g, '<em>$1</em>').replace(/\n/g, '<br>')}</div>
        </div>
        ` : ''}
        ` : `
        <div class="mi-section">
          <div class="mi-section-title">Why this decision</div>
          <ul class="mi-list">
            ${(why.length ? why : ["No supporting conditions could be computed (data missing)."]).map(x => `<li>${x}</li>`).join("")}
          </ul>
        </div>

        <div class="mi-section">
          <div class="mi-section-title">Why not the alternatives</div>
          <ul class="mi-list mi-list-muted">
            ${(whyNot.length ? whyNot : ["No explicit blockers detected beyond the current decision gates."]).map(x => `<li>${x}</li>`).join("")}
          </ul>
        </div>

        <div class="mi-section mi-verdict">
          <div class="mi-verdict-title">${verdictTitle}</div>
          <div class="mi-verdict-body">
            <div><b>Decision:</b> ${t(decision)}</div>
            <div><b>Signal:</b> ${t(signal)} | <b>Fundamental:</b> ${t(fundamental)}</div>
            <div><b>Execution framing:</b> ${(
              actionTone === "bull"
                ? "Prefer disciplined entries at/near support with RVOL confirmation for breakouts; respect invalidation below Support."
                : actionTone === "trim"
                  ? "Reduce risk into strength/overextension; avoid chasing new entries at resistance."
                  : actionTone === "bear"
                    ? "Stand aside or exit; wait for structure repair (reclaim SMA200 / rebuild base)."
                    : "Monitor; wait for higher-quality setup (RVOL + structure + asymmetry)."
            )}</div>
          </div>
          ${
            riskFlags.length
              ? `<div class="mi-risk">
                   <div class="mi-risk-title">Risk flags</div>
                   <ul class="mi-list">${riskFlags.map(x => `<li>${x}</li>`).join("")}</ul>
                 </div>`
              : ""
          }
        </div>
        `}
      </div>
    `;

    return html;
  }
}

/**
 * Shows professional 3-pane dialog:
 * - Left: compact indicator panel (buildIndicatorPanelHtml_)
 * - Middle: INPUT F data formatted professionally
 * - Bottom: DASH_REPORT data formatted
 * - Top: CALCULATIONS B-F data highlighted
 */
function showAnalysisPopup(ticker, reportHtml, d) {
  const leftPanel = buildIndicatorPanelHtml_(d);
  const middlePanel = buildInputFPanel_(ticker);
  const bottomPanel = buildDashReportPanel_(ticker);
  const topPanel = buildCalculationsTopPanel_(ticker);

  const html = `
    <div class="wrap">
      <div class="top-section">
        ${topPanel}
      </div>
      
      <div class="main-content">
        <div class="left">
          <div class="leftTitle">${String(ticker || "").toUpperCase()}</div>
          <div class="leftBody">
            ${leftPanel}
          </div>
          <button class="btnClose" onclick="google.script.host.close()">CLOSE</button>
        </div>

        <div class="right-container">
          <div class="middle">
            <div class="middleTitle">ANALYSIS DETAILS</div>
            <div class="middleBody">
              ${middlePanel}
            </div>
          </div>
          
          <div class="bottom-section">
            <div class="bottomTitle">DASHBOARD METRICS</div>
            <div class="bottomBody">
              ${bottomPanel}
            </div>
          </div>
        </div>
      </div>
    </div>

    <style>
      * {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Helvetica Neue', Arial, sans-serif;
      }
      
      .wrap{
        display:flex;
        flex-direction:column;
        gap:8px;
        padding:8px;
        background:#0b0f14;
        color:#e5e7eb;
        height:100vh;
        overflow:hidden;
      }

      /* TOP SECTION - CALCULATIONS B-F */
      .top-section{
        background:#0b1220;
        border:1px solid #30363d;
        border-radius:8px;
        padding:8px;
        min-height:60px;
      }

      /* MAIN CONTENT - LEFT + RIGHT CONTAINER */
      .main-content{
        display:flex;
        gap:8px;
        flex:1;
        overflow:hidden;
      }

      /* LEFT PANE - NARROWER */
      .left{
        width:22%;
        min-width:220px;
        background:#0b1220;
        border:1px solid #30363d;
        border-radius:8px;
        padding:8px;
        display:flex;
        flex-direction:column;
      }
      .leftTitle{
        font-weight:700;
        color:#93c5fd;
        letter-spacing:0.3px;
        font-size:13px;
        padding:3px 0 6px 0;
        border-bottom:1px solid #30363d;
        margin-bottom:6px;
      }
      .leftBody{
        flex:1;
        overflow:auto;
      }
      .btnClose{
        margin-top:8px;
        width:100%;
        padding:8px;
        border-radius:8px;
        cursor:pointer;
        font-weight:600;
        background:#111827;
        color:#93c5fd;
        border:1px solid #30363d;
        font-size:12px;
      }
      .btnClose:hover{ background:#1f2937; }

      /* RIGHT CONTAINER - MIDDLE + BOTTOM STACKED EQUALLY */
      .right-container{
        flex:1;
        display:flex;
        flex-direction:column;
        gap:8px;
        overflow:hidden;
      }

      /* MIDDLE PANE - INPUT F - EQUAL HEIGHT */
      .middle{
        flex:1;
        background:#0d1117;
        border:1px solid #30363d;
        border-radius:8px;
        display:flex;
        flex-direction:column;
        overflow:hidden;
        min-height:0;
      }
      .middleTitle{
        font-weight:600;
        color:#58a6ff;
        letter-spacing:0.5px;
        font-size:13px;
        padding:8px 10px;
        border-bottom:1px solid #30363d;
        background:#0b1220;
        border-top-left-radius:8px;
        border-top-right-radius:8px;
      }
      .middleBody{
        padding:14px;
        overflow:auto;
        flex:1;
      }

      /* BOTTOM SECTION - DASH_REPORT - EQUAL HEIGHT */
      .bottom-section{
        flex:1;
        background:#0b1220;
        border:1px solid #30363d;
        border-radius:8px;
        padding:8px;
        display:flex;
        flex-direction:column;
        overflow:hidden;
        min-height:0;
      }
      .bottomTitle{
        font-weight:600;
        color:#fbbf24;
        letter-spacing:0.5px;
        font-size:13px;
        margin-bottom:6px;
        padding-bottom:4px;
        border-bottom:1px solid #30363d;
      }
      .bottomBody{
        font-size:12px;
        line-height:1.6;
        overflow:auto;
        flex:1;
      }

      /* Formatting helpers */
      .data-row{
        display:flex;
        justify-content:space-between;
        padding:4px 6px;
        margin:3px 0;
        border-radius:4px;
        background:#0b1220;
        border:1px solid #30363d;
      }
      .data-label{
        font-weight:600;
        color:#9ca3af;
        font-size:11px;
      }
      .data-value{
        font-weight:500;
        color:#e5e7eb;
        font-size:11px;
      }
      .section-header{
        font-weight:600;
        color:#93c5fd;
        font-size:11px;
        letter-spacing:0.3px;
        margin:8px 0 4px 0;
        padding:3px 6px;
        background:#111827;
        border-radius:4px;
        border:1px solid #30363d;
      }
      .highlight-box{
        background:#1f2937;
        border:2px solid #fbbf24;
        border-radius:6px;
        padding:6px 8px;
        margin:3px 0;
      }
      .highlight-label{
        font-weight:600;
        color:#fbbf24;
        font-size:10px;
        letter-spacing:0.3px;
        text-transform:uppercase;
        margin-bottom:3px;
      }
      .highlight-value{
        font-weight:600;
        color:#e5e7eb;
        font-size:12px;
      }
      .text-content{
        font-size:12px;
        line-height:1.6;
        color:#e5e7eb;
        white-space:pre-wrap;
        word-wrap:break-word;
      }
      .text-content strong{
        font-weight:600;
        color:#fbbf24;
      }
      
      /* Professional analysis formatting */
      .analysis-heading{
        font-weight:600;
        color:#58a6ff;
        font-size:14px;
        letter-spacing:0.3px;
        margin:14px 0 8px 0;
        padding-bottom:4px;
        border-bottom:1px solid #30363d;
      }
      .analysis-paragraph{
        font-size:13px;
        line-height:1.7;
        color:#e5e7eb;
        margin-bottom:12px;
        text-align:left;
      }
      .analysis-paragraph strong{
        font-weight:600;
        color:#fbbf24;
      }
      .analysis-paragraph em{
        font-style:italic;
        color:#93c5fd;
      }
      .analysis-text{
        font-size:13px;
        line-height:1.7;
        color:#e5e7eb;
      }
      
      /* Dashboard metrics formatting */
      .dash-section{
        font-weight:600;
        color:#fbbf24;
        font-size:12px;
        letter-spacing:0.3px;
        margin:10px 0 6px 0;
        padding:4px 8px;
        background:#1f2937;
        border-radius:4px;
        border-left:3px solid #fbbf24;
      }
      .dash-row{
        display:flex;
        justify-content:space-between;
        align-items:flex-start;
        padding:4px 8px;
        margin:2px 0;
        border-bottom:1px solid #1f2937;
      }
      .dash-label{
        font-weight:600;
        color:#9ca3af;
        font-size:11px;
        flex:0 0 40%;
        padding-right:8px;
      }
      .dash-value{
        font-weight:500;
        color:#e5e7eb;
        font-size:11px;
        flex:1;
        text-align:left;
      }
      .dash-text{
        font-size:11px;
        line-height:1.6;
        color:#e5e7eb;
        margin:4px 0;
        padding:2px 8px;
      }
      .dash-category{
        font-weight:700;
        color:#60a5fa;
        font-size:11px;
        letter-spacing:0.3px;
        margin:8px 0 4px 0;
        padding:2px 8px;
        text-transform:uppercase;
      }
      .dash-bullets{
        margin:4px 0;
        padding-left:8px;
      }
      .dash-bullet{
        font-size:11px;
        line-height:1.6;
        color:#e5e7eb;
        margin:2px 0;
        padding:2px 0 2px 8px;
      }
      .dash-bullet strong{
        color:#fbbf24;
        font-weight:600;
      }
      .dash-bullet .dash-value{
        color:#34d399;
        font-weight:600;
      }


      /* REPORT (scoped) */
      .rightBody .mi-wrap{ }
      .rightBody .mi-head{
        border-bottom:1px solid #30363d;
        padding-bottom:10px;
        margin-bottom:12px;
      }
      .rightBody .mi-title{
        font-weight:900;
        font-size:16px;
        color:#e5e7eb;
        text-align:left;
        margin-bottom:6px;
      }
      .rightBody .mi-sub{
        display:flex;
        gap:10px;
        flex-wrap:wrap;
        align-items:center;
        font-size:12px;
        color:#9ca3af;
      }
      .rightBody .mi-ticker{
        font-weight:900;
        color:#fbbf24;
      }
      .rightBody .mi-regime{
        padding:2px 8px;
        border:1px solid #30363d;
        border-radius:999px;
        background:#0b1220;
        color:#cbd5e1;
        font-weight:800;
      }

      .rightBody .chip{
        padding:2px 8px;
        border-radius:999px;
        font-weight:900;
        border:1px solid #30363d;
      }
      .rightBody .chip-bull{ background:#0f3d2e; color:#a7f3d0; }
      .rightBody .chip-trim{ background:#3b2f11; color:#fde68a; }
      .rightBody .chip-bear{ background:#4a1414; color:#fecaca; }
      .rightBody .chip-neutral{ background:#1f2937; color:#e5e7eb; }

      .rightBody .mi-grid{
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:10px;
        margin:12px 0;
      }

      .rightBody .mi-card{
        border:1px solid #30363d;
        border-radius:12px;
        background:#0b1220;
        padding:10px;
      }
      .rightBody .mi-card-title{
        font-weight:900;
        color:#93c5fd;
        font-size:12px;
        letter-spacing:0.6px;
        margin-bottom:8px;
      }

      .rightBody .mi-kv{
        display:grid;
        grid-template-columns:140px 1fr;
        gap:6px 10px;
        font-size:12px;
      }
      .rightBody .mi-kv .k{
        color:#9ca3af;
        font-weight:800;
      }
      .rightBody .mi-kv .v{
        color:#e5e7eb;
        font-weight:700;
      }

      .rightBody .mi-section{
        border:1px solid #30363d;
        border-radius:12px;
        background:#0b1220;
        padding:10px;
        margin:10px 0;
      }
      .rightBody .mi-section-title{
        font-weight:900;
        color:#ff7b72;
        font-size:12px;
        letter-spacing:0.6px;
        margin-bottom:6px;
      }
      .rightBody .mi-list{
        margin:0;
        padding-left:18px;
        font-size:12px;
        line-height:1.45;
        color:#e5e7eb;
      }
      .rightBody .mi-list-muted{
        color:#cbd5e1;
      }
      .rightBody .mi-verdict{
        background:#0d1117;
      }
      .rightBody .mi-verdict-title{
        font-weight:900;
        color:#fbbf24;
        font-size:13px;
        margin-bottom:8px;
      }
      .rightBody .mi-verdict-body{
        font-size:12px;
        line-height:1.5;
      }
      .rightBody .mi-risk{
        margin-top:10px;
        border-top:1px dashed #30363d;
        padding-top:10px;
      }
      .rightBody .mi-risk-title{
        font-weight:900;
        color:#f87171;
        font-size:12px;
        margin-bottom:6px;
      }
      .rightBody .mi-narrative{
        font-size:12px;
        line-height:1.6;
        color:#e5e7eb;
        white-space:pre-wrap;
        word-wrap:break-word;
      }
      .rightBody .mi-narrative strong{
        font-weight:900;
        color:#fbbf24;
      }
      .rightBody .mi-narrative em{
        font-style:italic;
        color:#93c5fd;
      }
      .rightBody .mi-analyst-section{
        background:#0d1117;
        border:1px solid #30363d;
      }
      .rightBody .mi-analyst-grid{
        display:grid;
        grid-template-columns:1fr 1fr;
        gap:12px;
        margin-top:8px;
      }
      .rightBody .mi-analyst-item{
        background:#0b1220;
        border:1px solid #30363d;
        border-radius:8px;
        padding:10px;
        text-align:center;
      }
      .rightBody .mi-analyst-label{
        font-size:10px;
        font-weight:800;
        color:#9ca3af;
        text-transform:uppercase;
        letter-spacing:0.5px;
        margin-bottom:6px;
      }
      .rightBody .mi-analyst-value{
        font-size:14px;
        font-weight:900;
        letter-spacing:0.3px;
      }
      .rightBody .mi-rating-strong-buy{
        color:#10b981;
      }
      .rightBody .mi-rating-buy{
        color:#34d399;
      }
      .rightBody .mi-rating-hold{
        color:#fbbf24;
      }
      .rightBody .mi-rating-sell{
        color:#f87171;
      }
      .rightBody .mi-rating-strong-sell{
        color:#dc2626;
      }
      .rightBody .mi-price-target{
        color:#93c5fd;
      }
    </style>
  `;

  const out = HtmlService.createHtmlOutput(html)
    .setWidth(1200)
    .setHeight(900);

  SpreadsheetApp.getUi().showModalDialog(out, `Terminal Intelligence: ${String(ticker).toUpperCase()}`);
}


/**
 * Builds the TOP panel showing CALCULATIONS B-F data (highlighted)
 * CALCULATIONS structure:
 * B=MARKET RATING, C=DECISION, D=SIGNAL, E=PATTERNS, F=CONSENSUS PRICE, G=PRICE, H=CHANGE%
 */
function buildCalculationsTopPanel_(ticker) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calcSheet = ss.getSheetByName("CALCULATIONS");
  
  if (!calcSheet) {
    return '<div class="text-content">CALCULATIONS sheet not found</div>';
  }

  const data = calcSheet.getDataRange().getValues();
  let rowData = null;
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString().toUpperCase() === ticker.toString().toUpperCase()) {
      rowData = data[i];
      break;
    }
  }
  
  if (!rowData) {
    return '<div class="text-content">Ticker not found in CALCULATIONS</div>';
  }

  // CALCULATIONS columns: A=0(Ticker), B=1(MARKET RATING), C=2(DECISION), D=3(SIGNAL), E=4(PATTERNS), F=5(CONSENSUS PRICE)
  const marketRating = rowData[1] || "—";
  const decision = rowData[2] || "—";
  const signal = rowData[3] || "—";
  const patterns = rowData[4] || "—";
  const consensusPrice = rowData[5];
  
  // Format consensus price
  const consensusPriceDisplay = (typeof consensusPrice === 'number') ? '$' + consensusPrice.toFixed(2) : (consensusPrice || "—");

  return `
    <div style="display:grid; grid-template-columns:repeat(5, 1fr); gap:10px;">
      <div class="highlight-box">
        <div class="highlight-label">Market Rating</div>
        <div class="highlight-value">${marketRating}</div>
      </div>
      <div class="highlight-box">
        <div class="highlight-label">Decision</div>
        <div class="highlight-value">${decision}</div>
      </div>
      <div class="highlight-box">
        <div class="highlight-label">Signal</div>
        <div class="highlight-value">${signal}</div>
      </div>
      <div class="highlight-box">
        <div class="highlight-label">Patterns</div>
        <div class="highlight-value">${patterns}</div>
      </div>
      <div class="highlight-box">
        <div class="highlight-label">Consensus Price</div>
        <div class="highlight-value">${consensusPriceDisplay}</div>
      </div>
    </div>
  `;
}

/**
 * Builds the MIDDLE panel showing INPUT F data formatted professionally
 * Reads directly from INPUT sheet column F
 */
function buildInputFPanel_(ticker) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("INPUT");
  
  if (!inputSheet) {
    return '<div class="analysis-text">INPUT sheet not found</div>';
  }

  const data = inputSheet.getDataRange().getValues();
  let columnFData = "";
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString().toUpperCase() === ticker.toString().toUpperCase()) {
      // Column F is index 5
      columnFData = data[i][5] || "";
      break;
    }
  }
  
  if (!columnFData) {
    return '<div class="analysis-text">No analysis data available for this ticker</div>';
  }

  // Format the data professionally with proper structure
  let formattedContent = String(columnFData);
  
  // Split into paragraphs and format
  const paragraphs = formattedContent.split(/\n\n+/);
  let html = '';
  
  for (let para of paragraphs) {
    para = para.trim();
    if (!para) continue;
    
    // Check if it's a heading (starts with ## or **HEADING**)
    if (para.match(/^##\s+(.+)$/)) {
      const heading = para.replace(/^##\s+/, '');
      html += `<div class="analysis-heading">${heading}</div>`;
    } else if (para.match(/^\*\*([^*]+)\*\*:?\s*$/)) {
      const heading = para.replace(/^\*\*([^*]+)\*\*:?\s*$/, '$1');
      html += `<div class="analysis-heading">${heading}</div>`;
    } else {
      // Regular paragraph - format bold and italic
      para = para
        .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
        .replace(/\*([^*]+)\*/g, '<em>$1</em>')
        .replace(/\n/g, '<br>');
      html += `<div class="analysis-paragraph">${para}</div>`;
    }
  }

  return html || '<div class="analysis-text">No formatted content available</div>';
}

/**
 * Builds the BOTTOM panel showing DASH_REPORT data formatted professionally
 * Uses formulaEvaluator.js functions to generate institutional-grade narratives
 * Displays SIGNAL and DECISION narratives with professional formatting
 */
function buildDashReportPanel_(ticker) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const calcSheet = ss.getSheetByName("CALCULATIONS");
    const dashboardSheet = ss.getSheetByName("DASHBOARD");
    
    if (!calcSheet) {
      return '<div class="dash-text">CALCULATIONS sheet not found</div>';
    }
    
    if (!dashboardSheet) {
      return '<div class="dash-text">DASHBOARD sheet not found</div>';
    }
    
    // Get mode flag from DASHBOARD!H1 (TRUE = INVEST/long-term, FALSE = TRADE/short-term)
    const modeValue = dashboardSheet.getRange("H1").getValue();
    const isInvestMode = (modeValue === true || String(modeValue).toUpperCase() === "TRUE");
    
    // Find ticker row in CALCULATIONS sheet
    const data = calcSheet.getDataRange().getValues();
    let tickerRow = -1;
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toString().toUpperCase() === ticker.toString().toUpperCase()) {
        tickerRow = i;
        break;
      }
    }
    
    if (tickerRow === -1) {
      return '<div class="dash-text">Ticker not found in CALCULATIONS sheet</div>';
    }
    
    // Extract ticker data from CALCULATIONS row (0-based indexing)
    const rowData = data[tickerRow];
    const tickerData = {
      ticker: rowData[0],           // Column A
      marketRating: rowData[1],     // Column B
      decision: rowData[2],         // Column C
      signal: rowData[3],           // Column D
      patterns: rowData[4],         // Column E
      consensusPrice: rowData[5],   // Column F
      price: rowData[6],            // Column G
      changePct: rowData[7],        // Column H
      volTrend: rowData[8],         // Column I
      athTrue: rowData[9],          // Column J
      athDiff: rowData[10],         // Column K
      athZone: rowData[11],         // Column L
      fundamental: rowData[12],     // Column M
      trendState: rowData[13],      // Column N
      sma20: rowData[14],           // Column O
      sma50: rowData[15],           // Column P
      sma200: rowData[16],          // Column Q
      rsi: rowData[17],             // Column R
      macdHist: rowData[18],        // Column S
      divergence: rowData[19],      // Column T
      adx: rowData[20],             // Column U
      stochK: rowData[21],          // Column V
      volRegime: rowData[22],       // Column W
      bbpSignal: rowData[23],       // Column X
      atr: rowData[24],             // Column Y
      bollingerB: rowData[25],      // Column Z
      target: rowData[26],          // Column AA
      rrQuality: rowData[27],       // Column AB
      support: rowData[28],         // Column AC
      resistance: rowData[29],      // Column AD
      atrStop: rowData[30],         // Column AE
      atrTarget: rowData[31],       // Column AF
      positionSize: rowData[32],    // Column AG
      isPurchased: false            // Default to false (not purchased)
    };
    
    // Call evaluateSignalFormula() to get SIGNAL narrative
    let signalResult = null;
    let signalNarrative = null;
    
    try {
      // Check if function is available
      if (typeof evaluateSignalFormula !== 'function') {
        throw new Error("evaluateSignalFormula function not available");
      }
      
      signalResult = evaluateSignalFormula(tickerData, isInvestMode);
      signalNarrative = signalResult ? signalResult.narrative : null;
    } catch (error) {
      Logger.log(`Error calling evaluateSignalFormula: ${error.message}`);
      // Fall back to generic explanation
      signalNarrative = `🎯 WHY '${tickerData.signal}' TRIGGERED:\n\n⚠️ Using generic explanation (formula evaluation unavailable)\n\nSignal criteria: ${tickerData.signal} - see formula logic for details`;
    }
    
    // Call evaluateDecisionFormula() to get DECISION narrative
    let decisionResult = null;
    let decisionNarrative = null;
    
    try {
      // Check if function is available
      if (typeof evaluateDecisionFormula !== 'function') {
        throw new Error("evaluateDecisionFormula function not available");
      }
      
      const signalValue = signalResult ? signalResult.signal : tickerData.signal;
      decisionResult = evaluateDecisionFormula(tickerData, signalValue, isInvestMode);
      decisionNarrative = decisionResult ? decisionResult.narrative : null;
    } catch (error) {
      Logger.log(`Error calling evaluateDecisionFormula: ${error.message}`);
      // Fall back to generic explanation
      decisionNarrative = `🎯 HOW DECISION WAS DERIVED:\n\n⚠️ Using generic explanation (formula evaluation unavailable)\n\n1. SIGNAL: ${tickerData.signal} (technical setup)\n2. PATTERNS: ${tickerData.patterns === "—" ? "not detected" : tickerData.patterns + " detected"}\n3. DECISION: ${tickerData.decision}\n\n✅ Positive signal supports entry/add`;
    }
    
    // Format narratives into professional HTML
    let html = '';
    
    // SIGNAL Section
    html += '<div class="dash-section">📊 SIGNAL ANALYSIS</div>';
    html += formatNarrative_(signalNarrative || "No narrative available");
    
    // DECISION Section
    html += '<div class="dash-section">🎯 DECISION RATIONALE</div>';
    html += formatNarrative_(decisionNarrative || "No narrative available");
    
    return html;
    
  } catch (error) {
    Logger.log(`Error in buildDashReportPanel_: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    return `<div class="dash-text">Error generating dashboard report: ${error.message}</div>`;
  }
}

/**
 * Helper function to format narrative text into professional HTML
 * Converts markdown-style formatting to HTML with proper styling
 */
function formatNarrative_(narrative) {
  if (!narrative || narrative.trim() === "") {
    return '<div class="dash-text">No narrative available</div>';
  }
  
  let html = '';
  
  // Split narrative into sections (separated by double newlines)
  const sections = narrative.split(/\n\n+/);
  
  for (let section of sections) {
    section = section.trim();
    if (!section) continue;
    
    // Check if section is a category header (e.g., "PRICE ACTION:", "TREND STRUCTURE:")
    if (section.match(/^[A-Z\s]+:$/)) {
      const header = section.replace(/:$/, '').trim();
      html += `<div class="dash-category">${header}</div>`;
      continue;
    }
    
    // Check if section contains bullet points
    if (section.includes('•') || section.includes('-')) {
      const lines = section.split(/\n/);
      html += '<div class="dash-bullets">';
      
      for (let line of lines) {
        line = line.trim();
        if (!line) continue;
        
        // Remove bullet markers
        line = line.replace(/^[•\-]\s*/, '');
        
        // Format bold text (**text**)
        line = line.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
        
        // Format percentages and values with highlighting
        line = line.replace(/(\$[\d,]+\.?\d*)/g, '<span class="dash-value">$1</span>');
        line = line.replace(/([\d.]+%)/g, '<span class="dash-value">$1</span>');
        
        html += `<div class="dash-bullet">• ${line}</div>`;
      }
      
      html += '</div>';
    } else {
      // Regular paragraph
      let text = section;
      
      // Format bold text (**text**)
      text = text.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
      
      // Format percentages and values with highlighting
      text = text.replace(/(\$[\d,]+\.?\d*)/g, '<span class="dash-value">$1</span>');
      text = text.replace(/([\d.]+%)/g, '<span class="dash-value">$1</span>');
      
      // Replace newlines with <br>
      text = text.replace(/\n/g, '<br>');
      
      html += `<div class="dash-text">${text}</div>`;
    }
  }
  
  return html;
}


/* ---------- popup helpers ---------- */

/**
 * Builds the LEFT-side indicator panel HTML (2-column grid layout with dashboard color coding).
 * - Uses 2-column grid structure for compact display
 * - Dashboard color coding: green/red/grey backgrounds
 */
function buildIndicatorPanelHtml_(d) {
  const num = (v) => {
    const n = (typeof v === "number") ? v : parseFloat(String(v || "").replace(/[%,$,]/g, ""));
    return isFinite(n) ? n : null;
  };

  const fmt = {
    price: (v) => (num(v) == null ? "—" : `$${num(v).toFixed(2)}`),
    pct: (v) => (num(v) == null ? "—" : `${(num(v) * 100).toFixed(2)}%`),
    pctAlready: (v) => (num(v) == null ? "—" : `${num(v).toFixed(2)}%`),
    ratio: (v) => (num(v) == null ? "—" : `${num(v).toFixed(2)}x`),
    n2: (v) => (num(v) == null ? "—" : num(v).toFixed(2)),
    n3: (v) => (num(v) == null ? "—" : num(v).toFixed(3)),
    text: (v) => (v == null || String(v).trim() === "" ? "—" : String(v)),
    stochPct: (v) => {
      const n = num(v);
      if (n == null) return "—";
      return `${(n * 100).toFixed(1)}%`;
    },
    bollPct: (v) => {
      const n = num(v);
      if (n == null) return "—";
      return `${(n * 100).toFixed(1)}%`;
    }
  };

  // Dashboard color coding
  const BG_GREEN = "#0f3d2e";
  const BG_RED   = "#4a1414";
  const BG_GREY  = "#1f2937";
  const FG_LIGHT = "#e5e7eb";

  const cellStyle = (bg) => [
    `background:${bg}`,
    `color:${FG_LIGHT}`,
    `border:1px solid #30363d`,
    `border-radius:4px`,
    `padding:6px 8px`,
    `display:flex`,
    `flex-direction:column`,
    `gap:2px`
  ].join(";");

  const labelStyle = [
    `font-weight:600`,
    `letter-spacing:0.2px`,
    `font-size:11px`,
    `color:#9ca3af`,
    `white-space:nowrap`,
    `overflow:hidden`,
    `text-overflow:ellipsis`
  ].join(";");

  const valStyle = [
    `font-weight:600`,
    `font-size:12px`,
    `color:${FG_LIGHT}`,
    `white-space:nowrap`
  ].join(";");

  const sectionStyle = [
    `margin:10px 0 6px 0`,
    `padding:4px 8px`,
    `border-radius:6px`,
    `background:#111827`,
    `border:1px solid #30363d`,
    `color:#93c5fd`,
    `font-weight:600`,
    `font-size:10px`,
    `letter-spacing:0.4px`,
    `text-transform:uppercase`,
    `grid-column:1 / -1`
  ].join(";");

  // Color rules matching DASHBOARD sheet
  const colorPriceVsSMA200 = () => {
    const p = num(d.price), sma200 = num(d.sma200);
    if (p == null || sma200 == null) return BG_GREY;
    return p >= sma200 ? BG_GREEN : BG_RED;
  };

  const colorChangePct = () => {
    const c = num(d.changePct);
    if (c == null) return BG_GREY;
    if (c > 0) return BG_GREEN;
    if (c < 0) return BG_RED;
    return BG_GREY;
  };

  const colorVol = () => {
    const v = num(d.volRatio);
    if (v == null) return BG_GREY;
    if (v >= 1.5) return BG_GREEN;
    if (v < 0.85) return BG_GREY;
    return BG_GREY;
  };

  const colorAthDiff = () => {
    const x = num(d.athDiff);
    if (x == null) return BG_GREY;
    return x >= 0 ? BG_GREEN : BG_RED;
  };

  const colorAthZone = () => {
    const v = String(d.athZone || "").toUpperCase();
    if (v.includes("ATH") || v.includes("ZONE")) return BG_GREEN;
    if (v.includes("PULLBACK")) return BG_GREY;
    if (v.includes("CORRECTION")) return BG_RED;
    return BG_GREY;
  };

  const colorFundamental = () => {
    const v = String(d.fundamental || "").toUpperCase();
    if (v.includes("VALUE") || v.includes("DEEP VALUE")) return BG_GREEN;
    if (v.includes("FAIR")) return BG_GREY;
    if (v.includes("EXPENSIVE") || v.includes("ZOMBIE") || v.includes("PERFECTION")) return BG_RED;
    return BG_GREY;
  };

  const colorVolRegime = () => {
    const v = String(d.volRegime || "").toUpperCase();
    if (v.includes("LOW") || v.includes("SQUEEZE")) return BG_GREEN;
    if (v.includes("HIGH") || v.includes("BREAKOUT")) return BG_RED;
    return BG_GREY;
  };

  const colorRR = () => {
    const rr = num(d.rrQuality);
    if (rr == null) return BG_GREY;
    if (rr >= 3) return BG_GREEN;
    if (rr < 1.5) return BG_RED;
    return BG_GREY;
  };

  const colorTrendState = () => {
    const v = String(d.trendState || "").toUpperCase();
    if (v === "BULL") return BG_GREEN;
    if (v === "BEAR") return BG_RED;
    return BG_GREY;
  };

  const colorRSI = () => {
    const r = num(d.rsi);
    if (r == null) return BG_GREY;
    if (r <= 30) return BG_GREEN;
    if (r >= 70) return BG_RED;
    return BG_GREY;
  };

  const colorMACD = () => {
    const m = num(d.macdHist);
    if (m == null) return BG_GREY;
    return m >= 0 ? BG_GREEN : BG_RED;
  };

  const colorDiv = () => {
    const v = String(d.divergence || "").toUpperCase();
    if (v.includes("BULL")) return BG_GREEN;
    if (v.includes("BEAR")) return BG_RED;
    return BG_GREY;
  };

  const colorADX = () => {
    const a = num(d.adx);
    if (a == null) return BG_GREY;
    if (a >= 25) return BG_GREEN;
    if (a < 15) return BG_GREY;
    return BG_GREY;
  };

  const colorStoch = () => {
    const k = num(d.stochK);
    if (k == null) return BG_GREY;
    if (k <= 0.2) return BG_GREEN;
    if (k >= 0.8) return BG_RED;
    return BG_GREY;
  };

  const colorLevels = (key) => {
    const p = num(d.price);
    const v = num(d[key]);
    if (p == null || v == null || v <= 0) return BG_GREY;
    if (key === "support") return (p <= v * 1.01) ? BG_GREEN : BG_GREY;
    if (key === "resistance") return (p >= v * 0.99) ? BG_RED : BG_GREY;
    return BG_GREY;
  };

  const colorBoll = () => {
    const b = num(d.bolB);
    if (b == null) return BG_GREY;
    if (b <= 0.2) return BG_GREEN;
    if (b >= 0.8) return BG_RED;
    return BG_GREY;
  };

  // Cell builder
  const cell = (label, value, bg) => {
    return `
      <div style="${cellStyle(bg)}">
        <div style="${labelStyle}">${label}</div>
        <div style="${valStyle}">${value}</div>
      </div>
    `;
  };

  // 2-column grid container
  let html = `<div style="display:grid; grid-template-columns:1fr 1fr; gap:6px;">`;

  // PRICE / VOLUME SECTION (matches mobile report rows 9-14)
  html += `<div style="${sectionStyle}">PRICE / VOLUME</div>`;
  html += cell("PRICE", fmt.price(d.price), colorPriceVsSMA200());
  html += cell("CHG%", fmt.pct(d.changePct), colorChangePct());
  html += cell("VOL TREND", fmt.ratio(d.volRatio), colorVol());
  html += cell("P/E", "—", BG_GREY); // From GOOGLEFINANCE, not in d object
  html += cell("EPS", "—", BG_GREY); // From GOOGLEFINANCE, not in d object
  html += cell("RANGE %", "—", BG_GREY); // Calculated from historical data, not in d object

  // PERFORMANCE SECTION (matches mobile report rows 16-19)
  html += `<div style="${sectionStyle}">PERFORMANCE</div>`;
  html += cell("ATH TRUE", fmt.price(d.isATH), BG_GREY);
  html += cell("ATH DIFF %", fmt.pct(d.athDiff), colorAthDiff());
  html += cell("ATH ZONE", fmt.text(d.athZone), colorAthZone());
  html += cell("FUNDAMENTAL", fmt.text(d.fundamental), colorFundamental());

  // TREND SECTION (matches mobile report rows 21-24)
  html += `<div style="${sectionStyle}">TREND</div>`;
  html += cell("TREND STATE", fmt.text(d.trendState), colorTrendState());
  html += cell("SMA 20", fmt.price(d.sma20), BG_GREY);
  html += cell("SMA 50", fmt.price(d.sma50), BG_GREY);
  html += cell("SMA 200", fmt.price(d.sma200), BG_GREY);

  // MOMENTUM SECTION (matches mobile report rows 26-30)
  html += `<div style="${sectionStyle}">MOMENTUM</div>`;
  html += cell("RSI", fmt.n2(d.rsi), colorRSI());
  html += cell("MACD HIST", fmt.n3(d.macdHist), colorMACD());
  html += cell("DIVERGENCE", fmt.text(d.divergence), colorDiv());
  html += cell("ADX (14)", fmt.n2(d.adx), colorADX());
  html += cell("STOCH %K", fmt.stochPct(d.stochK), colorStoch());

  // VOLATILITY SECTION (matches mobile report rows 32-35)
  html += `<div style="${sectionStyle}">VOLATILITY</div>`;
  html += cell("VOL REGIME", fmt.text(d.volRegime), colorVolRegime());
  html += cell("BBP SIGNAL", fmt.text(d.bbpSignal), BG_GREY);
  html += cell("ATR (14)", fmt.n2(d.atr), BG_GREY);
  html += cell("BOLLINGER %B", fmt.bollPct(d.bolB), colorBoll());

  // TARGET SECTION (matches mobile report rows 37-43)
  html += `<div style="${sectionStyle}">TARGET</div>`;
  html += cell("TARGET (3:1)", fmt.price(d.target), BG_GREY);
  html += cell("R:R QUALITY", fmt.ratio(d.rrQuality), colorRR());
  html += cell("SUPPORT", fmt.price(d.support), colorLevels("support"));
  html += cell("RESISTANCE", fmt.price(d.resistance), colorLevels("resistance"));
  html += cell("ATR STOP", fmt.price(d.atrStop), BG_GREY);
  html += cell("ATR TARGET", fmt.price(d.atrTarget), BG_GREY);
  html += cell("POSITION SIZE", fmt.text(d.positionSize), BG_GREY);

  html += `</div>`;

  return html;
}


/**
 * UI Trigger Function
 */
function runMasterAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Detect which sheet triggered the function
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet ? activeSheet.getName() : "";
  const isFromReportSheet = sheetName === "REPORT";
  const isFromDashboard = sheetName === "DASHBOARD";
  
  let ticker = "";
  
  // Get ticker based on source sheet
  if (isFromDashboard) {
    // Get ticker from column A of the selected row
    const activeCell = ss.getActiveCell();
    if (activeCell) {
      const row = activeCell.getRow();
      ticker = activeSheet.getRange(row, 1).getValue(); // Column A
    }
  } else {
    // Get ticker from ACTIVE CELL or REPORT A1
    const activeCell = ss.getActiveCell();
    ticker = activeCell ? activeCell.getValue() : "";
    
    // Fallback: Check REPORT A1
    if (!ticker || typeof ticker !== 'string' || ticker === "") {
      const reportSheet = ss.getSheetByName("REPORT");
      if (reportSheet) {
        ticker = reportSheet.getRange("A1").getValue();
      }
    }
  }

  if (!ticker || typeof ticker !== 'string' || ticker === "") {
    ss.toast("⚠️ Please select a cell with a ticker symbol", "INPUT NEEDED");
    return;
  }

  // Fetch the data
  const d = getTickerDataFromBaseline(ticker);
  if (!d) return;

  // Generate and show the report
  const report = MasterAnalysisEngine.analyze(d, isFromReportSheet);
  showAnalysisPopup(ticker, report, d);
}

/**
 * Helper: Searches the Golden Baseline for a specific ticker and returns the data object.
 * Mapping follows the CURRENT CALCULATIONS column structure (A-AI)
 * @param {string} ticker - The stock symbol to search for.
 */
function getTickerDataFromBaseline(ticker) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const DATA_TAB_NAME = "CALCULATIONS"; 
  const dataSheet = ss.getSheetByName(DATA_TAB_NAME);

  if (!dataSheet) {
    ss.toast(`❌ Tab '${DATA_TAB_NAME}' not found!`, "ERROR", 5);
    return null;
  }

  const data = dataSheet.getDataRange().getValues();
  
  for (let i = 0; i < data.length; i++) {
    // Column A is Index 0
    if (data[i][0].toString().toUpperCase() === ticker.toString().toUpperCase()) {
      const rowData = data[i];
      
      // CORRECT MAPPING TABLE based on generateCalculations.js:
      // A=0(Ticker), B=1(MARKET RATING), C=2(DECISION), D=3(SIGNAL), E=4(PATTERNS), 
      // F=5(CONSENSUS PRICE), G=6(Price), H=7(Change %), I=8(Vol Trend), 
      // J=9(ATH TRUE), K=10(ATH Diff %), L=11(ATH ZONE), M=12(FUNDAMENTAL),
      // N=13(Trend State), O=14(SMA 20), P=15(SMA 50), Q=16(SMA 200),
      // R=17(RSI), S=18(MACD Hist), T=19(Divergence), U=20(ADX 14), V=21(Stoch %K 14),
      // W=22(VOL REGIME), X=23(BBP SIGNAL), Y=24(ATR 14), Z=25(Bollinger %B),
      // AA=26(Target 3:1), AB=27(R:R Quality), AC=28(Support), AD=29(Resistance),
      // AE=30(ATR STOP), AF=31(ATR TARGET), AG=32(POSITION SIZE), AH=33(LAST STATE), AI=34(ANALYSIS SUMMARY)
      
      return {
        ticker:      rowData[0],  // Col A: Ticker
        marketRating:rowData[1],  // Col B: MARKET RATING
        decision:    rowData[2],  // Col C: DECISION
        signal:      rowData[3],  // Col D: SIGNAL
        patterns:    rowData[4],  // Col E: PATTERNS
        consensusPrice: rowData[5], // Col F: CONSENSUS PRICE
        price:       rowData[6],  // Col G: Price
        changePct:   rowData[7],  // Col H: Change %
        volRatio:    rowData[8],  // Col I: Vol Trend (RVOL)
        isATH:       rowData[9],  // Col J: ATH (TRUE)
        athDiff:     rowData[10], // Col K: ATH Diff %
        athZone:     rowData[11], // Col L: ATH ZONE
        fundamental: rowData[12], // Col M: FUNDAMENTAL
        trendState:  rowData[13], // Col N: Trend State
        sma20:       rowData[14], // Col O: SMA 20
        sma50:       rowData[15], // Col P: SMA 50
        sma200:      rowData[16], // Col Q: SMA 200
        rsi:         rowData[17], // Col R: RSI
        macdHist:    rowData[18], // Col S: MACD Hist
        divergence:  rowData[19], // Col T: Divergence
        adx:         rowData[20], // Col U: ADX (14)
        stochK:      rowData[21], // Col V: Stoch %K (14)
        volRegime:   rowData[22], // Col W: Vol Regime
        bbpSignal:   rowData[23], // Col X: BBP Signal
        atr:         rowData[24], // Col Y: ATR (14)
        bolB:        rowData[25], // Col Z: Bollinger %B
        target:      rowData[26], // Col AA: Target (3:1)
        rrQuality:   rowData[27], // Col AB: R:R Quality
        support:     rowData[28], // Col AC: Support
        resistance:  rowData[29], // Col AD: Resistance
        atrStop:     rowData[30], // Col AE: ATR STOP
        atrTarget:   rowData[31], // Col AF: ATR TARGET
        positionSize:rowData[32], // Col AG: POSITION SIZE
        lastState:   rowData[33], // Col AH: Last State
        stopLoss:    rowData[28]  // Mapping Support as StopLoss
      };
    }
  }
  
  ss.toast(`Ticker '${ticker}' not found in baseline.`, "⚠️ SEARCH FAILED", 3);
  return null;
}