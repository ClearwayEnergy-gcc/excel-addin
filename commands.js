// commands.js — Ribbon button handlers
//
// This script runs inside the hidden function file iframe (commands.html).
// Each ribbon button mapped to ExecuteFunction calls one of these functions.
// Results are written to localStorage so the task pane can display them.
//
// To add a new ribbon button:
//   1. Define a function here following the same pattern (accepts `event`, calls event.completed())
//   2. Register it below with Office.actions.associate()
//   3. Add a <Control> block in manifest.xml pointing to the function name

Office.onReady(function () {
  // Register functions so Office can find them by name from the manifest.
  Office.actions.associate('readCell',  readCell);
  Office.actions.associate('writeData', writeData);
  Office.actions.associate('clearRange', clearRange);
});

// ── Logging helper ────────────────────────────────────────────────────────────
// Writes a log entry to localStorage. The task pane picks it up via the
// storage event or its polling fallback.

function writeLog(message, type) {
  var timestamp = new Date().toLocaleTimeString();

  var counter = parseInt(localStorage.getItem('addin_log_counter') || '0') + 1;
  localStorage.setItem('addin_log_counter', counter);

  var entry = {
    id:        counter,
    timestamp: timestamp,
    message:   message,
    type:      type || 'info'   // 'info' | 'success' | 'error'
  };

  var logs = [];
  try {
    logs = JSON.parse(localStorage.getItem('addin_logs') || '[]');
  } catch (e) { /* start fresh if corrupted */ }

  logs.push(entry);
  localStorage.setItem('addin_logs', JSON.stringify(logs));

  // Writing a separate key triggers the 'storage' event in other contexts
  // (task pane iframe) listening to the same origin.
  localStorage.setItem('addin_log_latest', JSON.stringify(entry));
}

// ── Button 1: Read Cell ───────────────────────────────────────────────────────
// Reads the currently selected cell and logs its address + value.

function readCell(event) {
  Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load(['address', 'values']);

    return context.sync().then(function () {
      var address = range.address;
      var value   = range.values[0][0];
      var display = (value === null || value === '') ? '(empty)' : value;
      writeLog('Read Cell [' + address + ']: ' + display, 'success');
    });
  })
  .catch(function (error) {
    writeLog('Read Cell error: ' + error.message, 'error');
  })
  .then(function () {
    // event.completed() must always be called to release the ribbon button.
    event.completed();
  });
}

// ── Button 2: Write Data ──────────────────────────────────────────────────────
// Writes a small sample table to A1:C3 on the active sheet.

function writeData(event) {
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    var dataRange   = sheet.getRange('A1:C3');
    dataRange.values = [
      ['Name',   'Value', 'Status'],
      ['Item A',  100,    'Active'],
      ['Item B',  200,    'Pending']
    ];

    // Bold the header row
    sheet.getRange('A1:C1').format.font.bold = true;

    return context.sync().then(function () {
      writeLog('Write Data: Sample table written to A1:C3', 'success');
    });
  })
  .catch(function (error) {
    writeLog('Write Data error: ' + error.message, 'error');
  })
  .then(function () {
    event.completed();
  });
}

// ── Button 3: Clear Range ─────────────────────────────────────────────────────
// Clears all content and formatting from A1:C3.

function clearRange(event) {
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange('A1:C3').clear();

    return context.sync().then(function () {
      writeLog('Clear Range: A1:C3 cleared', 'success');
    });
  })
  .catch(function (error) {
    writeLog('Clear Range error: ' + error.message, 'error');
  })
  .then(function () {
    event.completed();
  });
}
