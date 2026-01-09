

/**
* ------------------------------------------------------------------
* 7. AUTOMATED ALERT & MONITOR SYSTEM
* ------------------------------------------------------------------
*/

function startMarketMonitor() {
  stopMarketMonitor();

  ScriptApp.newTrigger('checkSignalsAndSendAlerts')
    .timeBased()
    .everyMinutes(30)
    .create();

  SpreadsheetApp.getUi().alert(
    '🔔 MONITOR ACTIVE',
    'Checking DECISION changes (CALCULATIONS!D) every 30 minutes.\n\n' +
    'You will be emailed only when a DECISION changes, including:\n' +
    '- Trade Long / Accumulate\n' +
    '- Take Profit / Reduce\n' +
    '- Stop-Out / Avoid\n',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
* ------------------------------------------------------------------
* STOP MONITOR (UPDATED TEXT: DECISION monitor)
* ------------------------------------------------------------------
*/
function stopMarketMonitor() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'checkSignalsAndSendAlerts') {
      ScriptApp.deleteTrigger(t);
    }
  });

  SpreadsheetApp.getUi().alert(
    '🔕 MONITOR STOPPED',
    'Automated DECISION checks disabled.\n\n' +
    'No further emails will be sent for DECISION changes (CALCULATIONS!D) until you start the monitor again.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


function checkSignalsAndSendAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calcSheet = ss.getSheetByName("CALCULATIONS");
  if (!calcSheet) return;

  const lastRow = calcSheet.getLastRow();
  if (lastRow < 3) return;

  // A..AB (28 cols)
  const range = calcSheet.getRange(3, 1, lastRow - 2, 28);
  const data = range.getValues();

  const alerts = [];

  data.forEach((r, i) => {
    const ticker = (r[0] || "").toString().trim();     // A
    const decision = (r[2] || "").toString().trim();   // C (DECISION) ✅
    // Note: LAST_STATE column removed - state tracking simplified

    if (!ticker || !decision || decision === "LOADING") return;
    if (decision === lastState) return;

    // Actionable states: includes SELL/trim/profit + buy/trade + stops/avoid
    const isActionable = /STOP|AVOID|TAKE PROFIT|REDUCE|TRADE LONG|ACCUMULATE/i.test(decision);

    if (isActionable) {
      alerts.push(
        `TICKER: ${ticker}\nNEW DECISION: ${decision}\nPREVIOUS: ${lastState || "Initial Scan"}`
      );
    }

    // Persist the new last notified decision into AB
    calcSheet.getRange(i + 3, 28).setValue(decision);
  });

  if (alerts.length === 0) return;

  const email = Session.getActiveUser().getEmail();
  const subject = `📈 Terminal Alert: ${alerts.length} Decision Change(s)`;
  const body =
    "Institutional Terminal detected DECISION changes (CALCULATIONS!D):\n\n" +
    alerts.join("\n\n") +
    "\n\nView Terminal:\n" + ss.getUrl();

  MailApp.sendEmail(email, subject, body);
}
