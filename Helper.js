

/**
 * Helper: reads sidebar values by labels (case-insensitive)
 * Scans A8:B200 (your sidebar region)
 */
function getSidebarValuesByLabels_(chartSheet, labels) {
  const want = new Set(labels.map(l => String(l).trim().toUpperCase()));
  const keys = chartSheet.getRange("A8:A200").getValues().flat().map(v => String(v || "").trim().toUpperCase());
  const vals = chartSheet.getRange("B8:B200").getValues().flat();

  const out = {};
  for (let i = 0; i < keys.length; i++) {
    if (want.has(keys[i])) {
      const original = labels.find(l => String(l).trim().toUpperCase() === keys[i]);
      out[original] = vals[i];
    }
  }
  labels.forEach(l => { if (out[l] === undefined) out[l] = 0; });
  return out;
}

function getSidebarLevels_(chartSheet) {
  const labelRange = chartSheet.getRange("A5:A120").getValues().flat();
  const valueRange = chartSheet.getRange("B5:B120").getValues().flat();

  const findValueAny = (labels) => {
    const want = new Set(labels.map(l => String(l).trim().toUpperCase()));
    const idx = labelRange.findIndex(v => want.has(String(v || "").trim().toUpperCase()));
    if (idx === -1) return 0;
    return Number(valueRange[idx]) || 0;
  };

  return {
    support: findValueAny(["SUPPORT", "SUPPORT FLOOR"]),
    resistance: findValueAny(["RESISTANCE", "RESISTANCE CEILING"])
  };
}

/**
* ------------------------------------------------------------------
* REFERENCE GUIDE (UPDATED: SELL states + aligned to DECISION/SIGNAL formulas)
* - Keeps your structure; only updates vocabulary tables and explanations.
* ------------------------------------------------------------------
*/
function generateReferenceSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = "REFERENCE_GUIDE";
  let sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear().clearFormats();

  const rows = [];

  // Title
  rows.push(["INSTITUTIONAL TERMINAL — REFERENCE GUIDE", "", "", ""]);
  rows.push(["Dashboard/Chart vocabulary, column definitions, and action playbook (aligned to current formulas).", "", "", ""]);
  rows.push(["", "", "", ""]);

  // NEW: Position tags policy
  rows.push(["0) POSITION TAGS (INPUT!C) — PRACTICAL USAGE (IMPORTANT)", "", "", ""]);
  rows.push(["RULE", "MEANING", "HOW IT AFFECTS THE ENGINE", "WHAT YOU SHOULD DO"]);
  rows.push([
    "PURCHASED (only behavioral tag)",
    "Indicates you currently hold this ticker (an open position).",
    "Enables sell-side + position-management actions in DECISION (Stop-Out / Take Profit / Reduce / Add in Dip).",
    "When you buy, add PURCHASED to INPUT!C for that ticker. When you fully exit, remove PURCHASED."
  ]);
  rows.push([
    "All other tags (e.g., M7, P0, etc.)",
    "Custom labels for your own grouping and filtering.",
    "NO IMPACT on buy/sell logic. They are ignored by the engine for decisions.",
    "Use them only to filter lists (INPUT filters / watchlists). Do not expect them to alter decisions."
  ]);
  rows.push(["", "", "", ""]);

  // Column definitions
  rows.push(["1) DASHBOARD COLUMN DEFINITIONS (TECHNICAL)", "", "", ""]);
  rows.push(["COLUMN", "WHAT IT IS", "HOW IT IS USED", "USER ACTION"]);
  const cols = [
    ["Ticker", "Symbol (key)", "Join key across DATA/CALCULATIONS/CHART", "Select for chart / review notes."],
    ["SIGNAL", "Technical setup label (rules engine)", "Describes setup type (breakout / trend / mean-rev / risk-off / stop-out)", "Use as setup classification; DECISION is what you act on."],
    ["FUNDAMENTAL", "EPS + P/E risk bucket", "Blocks trades in weak quality/extreme valuation regimes", "Prefer VALUE/FAIR; avoid ZOMBIE/PRICED FOR PERFECTION when momentum is fragile."],
    ["DECISION", "Action label (position-aware)", "Final instruction (trade/accumulate/avoid/stop/trim/profit/add-in-dip)", "Primary action field."],
    ["Price", "Live last price (GOOGLEFINANCE)", "Used for regime tests, distance-to-levels, ATR stretch", "Confirm price vs SMA200 & levels."],
    ["Change %", "Daily % change", "Tape context; not a signal alone", "Do not chase without a setup."],
    ["Vol Trend", "Relative volume proxy (RVOL)", "Conviction filter for breakouts", "Prefer >=1.5x for breakout validation."],
    ["ATH (TRUE)", "All-time high reference", "Context for overhead supply / price discovery", "Avoid chasing into ceilings without RVOL."],
    ["ATH Diff %", "Distance from ATH", "Pullback vs near-ATH classification", "Use with regime + levels."],
    ["R:R Quality", "Reward/Risk ratio proxy", "Trade quality gate", ">=3 excellent; 2–3 acceptable; <2 poor."],
    ["Trend Score", "★ count (Price above SMAs)", "Quick structure strength read", "3★ strongest; <2★ caution."],
    ["Trend State", "Bull/Bear via SMA200", "Defines risk-on vs risk-off playbook", "Below SMA200 = risk-off bias."],
    ["SMA 20", "Short-term mean", "Stretch anchor; mean reversion reference", "Avoid buying when >2x ATR above SMA20."],
    ["SMA 50", "Medium trend line", "Momentum/structure confirmation", "If lost with MACD<0, reduce risk."],
    ["SMA 200", "Long-term regime line", "Primary risk-on/risk-off filter", "Below: avoid trend-chasing."],
    ["RSI", "Momentum oscillator (0–100)", "Overbought/oversold + bias filter", "<30 oversold; >70 overbought; 50 bias."],
    ["MACD Hist", "Impulse (positive/negative)", "Momentum confirmation / deterioration", "Negative impulse with SMA50 loss = reduce."],
    ["Divergence", "Price vs RSI divergence heuristic", "Early reversal warning", "Bull div supports bounce; bear div warns."],
    ["ADX (14)", "Trend strength", "Chop vs trend filter", "<15 range; 15–25 weak; >=25 trend."],
    ["Stoch %K (14)", "Fast oscillator (0–1)", "Timing within regimes", "<0.2 oversold; >0.8 overbought."],
    ["Support", "20-day min low proxy", "Risk line / invalidation reference", "Break below = Stop-Out."],
    ["Resistance", "50-day max high proxy", "Ceiling / profit-taking reference", "Near resistance + overbought = Take Profit."],
    ["Target (3:1)", "Tactical take-profit projection", "Planning exits; not a forecast", "Use for planning only."],
    ["ATR (14)", "Volatility proxy", "Sizing/stops + stretch detection", "Higher ATR = wider stops / smaller size."],
    ["Bollinger %B", "Band position proxy", "Compression/range heuristic", "Low %B + low ADX = chop."],
    ["TECH NOTES", "Narrative (indicator values + rationale)", "Explains what is driving setup and timing", "Read before acting."],
    ["FUND NOTES", "Narrative (fund + signal + action + flags)", "Explains why decision is allowed/blocked", "Respect blockers (risk-off / fragile valuation)."]
  ];
  cols.forEach(r => rows.push(r));

  // SIGNAL vocabulary
  rows.push(["", "", "", ""]);
  rows.push(["2) SIGNAL — FULL VOCABULARY (WHAT IT MEANS + WHAT TO DO)", "", "", ""]);
  rows.push(["SIGNAL VALUE", "TECHNICAL DEFINITION", "WHEN IT TRIGGERS", "EXPECTED USER ACTION"]);
  const signal = [
    ["Stop-Out", "Price < Support", "Breakdown through support floor", "Exit / do not average down. Wait for base."],
    ["Risk-Off (Below SMA200)", "Price < SMA200", "Long-term risk-off regime", "Avoid chasing; only tactical with strict risk."],
    ["Volatility Squeeze (Coiling)", "ATR compressed vs recent lows", "Compression / coiling", "Wait for breakout confirmation (RVOL + levels)."],
    ["Range-Bound (Low ADX)", "ADX < 15", "No trend / chop regime", "Range tactics only; smaller size; tighter targets."],
    ["Breakout (High Volume)", "RVOL high + price near/above resistance", "Breakout attempt with sponsorship", "Only act when DECISION allows (Trade Long)."],
    ["Mean Reversion (Oversold)", "StochK<=0.20 above support", "Oversold timing in structure", "If PURCHASED: Add in Dip (when allowed). If not: Tactical long only if DECISION says Trade Long."],
    ["Trend Continuation", "Above SMA200 with momentum + trend", "Uptrend continuation regime", "Accumulate on pullbacks; avoid chasing stretch."]
  ];
  signal.forEach(r => rows.push(r));

  // FUNDAMENTAL vocabulary
  rows.push(["", "", "", ""]);
  rows.push(["3) FUNDAMENTAL — FULL VOCABULARY (FILTER + RISK)", "", "", ""]);
  rows.push(["FUNDAMENTAL VALUE", "WHAT IT MEANS (IN THIS MODEL)", "RISK PROFILE", "EXPECTED USER ACTION"]);
  const fund = [
    ["VALUE", "EPS positive with supportive P/E", "Lower valuation risk vs others", "Prefer for breakouts/trend setups when tech confirms."],
    ["FAIR", "Neutral valuation bucket (fallback)", "Neutral valuation risk", "Trade only when technical gates pass."],
    ["EXPENSIVE", "Valuation premium (lower margin for error)", "Multiple compression risk", "Be selective; prefer strongest technical setups."],
    ["PRICED FOR PERFECTION", "Very elevated P/E (high expectations)", "Fragile; sharp drawdown risk on misses", "Only best setups; size down; take profits faster."],
    ["ZOMBIE", "EPS <= 0 / weak earnings quality", "High blow-up risk", "Avoid longs; treat as high risk."]
  ];
  fund.forEach(r => rows.push(r));

  // DECISION vocabulary (position-aware)
  rows.push(["", "", "", ""]);
  rows.push(["4) DECISION — FULL VOCABULARY (WHAT TO DO)", "", "", ""]);
  rows.push(["DECISION VALUE", "WHY IT HAPPENS (ENGINE RULE)", "POSITION REQUIREMENT", "EXPECTED USER ACTION"]);
  const decision = [
    ["Stop-Out", "Price < Support (invalidation)", "Purchased or not", "Exit / stand aside."],
    ["Avoid", "Risk-off regime or blocked conditions", "Not required", "No trade; deprioritize."],
    ["Trade Long", "Breakout / setup allowed by gates", "Not purchased", "Enter with stop at Support; plan to Resistance/Target."],
    ["Accumulate", "Trend continuation in acceptable conditions", "Not purchased (or add later manually)", "Scale in on pullbacks; avoid chasing."],
    ["Hold", "No edge or gates not met", "Any", "Do nothing; monitor levels and signals."],
    ["Take Profit", "Target hit OR resistance + overbought", "PURCHASED only", "Trim/sell into strength; do not chase."],
    ["Reduce (Momentum Weak)", "MACD < 0 AND Price < SMA50", "PURCHASED only", "Reduce exposure; tighten risk."],
    ["Reduce (Overextended)", "Stretch >= 2x ATR above SMA20", "PURCHASED only", "Trim; wait for mean reversion."],
    ["Add in Dip", "Oversold timing above support in risk-on regime", "PURCHASED only", "Add small / staged; never add below Support."],
    ["LOADING", "Data not ready", "N/A", "Wait for refresh; do not act."]
  ];
  decision.forEach(r => rows.push(r));

  // Quick playbook
  rows.push(["", "", "", ""]);
  rows.push(["5) QUICK PLAYBOOK (HOW TO USE THE TERMINAL)", "", "", ""]);
  rows.push(["RULE", "WHY", "WHAT TO LOOK FOR", "WHAT TO AVOID"]);
  rows.push(["Position flagging", "Ensures sell logic only triggers for holdings", "INPUT!C contains PURCHASED", "Expecting M7/P0/etc. to change decisions (they will not)."]);
  rows.push(["Trend trades", "Best expectancy in strong regimes", "Above SMA200, ADX>=25, MACD>0, RVOL>=1.5", "Buying in Risk-Off or with ADX<15."]);
  rows.push(["Range trades", "Chop markets are mean-reverting", "ADX<15 and price near Support/Resistance", "Chasing mid-range; poor R:R."]);
  rows.push(["Profit-taking", "Avoid giving back gains", "Take Profit near Resistance/Target", "Adding new longs when stretched/overbought."]);
  rows.push(["Loss avoidance", "Stops define survival", "Stop-Out (Price<Support)", "Averaging down below Support."]);
  rows.push(["R:R gating", "Prevents low-quality trades", "R:R>=2 tactical; >=3 preferred", "R:R<2 unless exceptional setup."]);

  // Write
  sh.getRange(1, 1, rows.length, 4).setValues(rows);

  // Styling (professional)
  sh.setColumnWidth(1, 240);
  sh.setColumnWidth(2, 520);
  sh.setColumnWidth(3, 360);
  sh.setColumnWidth(4, 280);
  sh.setRowHeights(1, Math.min(rows.length, 900), 18);
  sh.setFrozenRows(3);

  // Title bars
  sh.getRange("A1:D1").merge()
    .setBackground("#0B5394").setFontColor("white")
    .setFontWeight("bold").setFontSize(13)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sh.getRange("A2:D2").merge()
    .setBackground("#073763").setFontColor("#FFFF00")
    .setFontWeight("bold").setFontSize(9)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // Section headers
  for (let r = 1; r <= rows.length; r++) {
    const v = String(sh.getRange(r, 1).getValue() || "");
    if (/^\d\)|^0\)/.test(v)) {
      sh.getRange(r, 1, 1, 4).merge()
        .setBackground("#212121").setFontColor("white")
        .setFontWeight("bold").setFontSize(10)
        .setHorizontalAlignment("left");
    }
  }

  // Table header rows
  for (let r = 1; r <= rows.length; r++) {
    const a = String(sh.getRange(r, 1).getValue() || "").trim();
    if (["RULE", "COLUMN", "SIGNAL VALUE", "FUNDAMENTAL VALUE", "DECISION VALUE"].includes(a)) {
      sh.getRange(r, 1, 1, 4)
        .setBackground("#F3F3F3")
        .setFontWeight("bold")
        .setFontColor("#111111")
        .setHorizontalAlignment("center");
    }
  }

  sh.getRange(1, 1, rows.length, 4).setWrap(true).setVerticalAlignment("top");
  sh.getRange(1, 1, rows.length, 4)
    .setBorder(true, true, true, true, true, true, "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID);

  const band = sh.getRange(4, 1, Math.max(1, rows.length - 3), 4).applyRowBanding();
  band.setHeaderRowColor("#FFFFFF");
  band.setFirstRowColor("#FFFFFF");
  band.setSecondRowColor("#FAFAFA");

  ss.toast("REFERENCE_GUIDE updated: PURCHASED is the only behavioral tag; others are filter-only.", "✅ DONE", 3);
}