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
  Office.actions.associate('goalSeek',  goalSeek);
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

// ── Button 3: Goal Seek ───────────────────────────────────────────────────────
// Requires three named ranges in the workbook:
//   CEG_Target — a formula cell that should reach 0 (e.g. model output - target)
//   CEG_Input  — the input cell whose value is overwritten each iteration
//   CEG_Guess  — a formula cell that computes the next candidate value for CEG_Input
//
// Each iteration:
//   1. Read CEG_Target — stop if |value| ≤ TOLERANCE
//   2. Read CEG_Guess
//   3. Write CEG_Guess → CEG_Input  (triggers Excel recalculation)
//   4. Repeat

function goalSeek(event) {
  var MAX_ITERATIONS = 1000;
  var TOLERANCE      = 1e-10;  // treat |CEG_Target| <= this as "equal to zero"

  Excel.run(function (context) {
    var names = context.workbook.names;

    // getItemOrNullObject avoids a hard error if a name is missing
    var targetItem = names.getItemOrNullObject('CEG_Target');
    var inputItem  = names.getItemOrNullObject('CEG_Input');
    var guessItem  = names.getItemOrNullObject('CEG_Guess');

    targetItem.load('isNullObject');
    inputItem.load('isNullObject');
    guessItem.load('isNullObject');

    return context.sync().then(function () {

      // ── Step 1: Verify all three named ranges exist ─────────────────────────
      var missing = [];
      if (targetItem.isNullObject) missing.push('CEG_Target');
      if (inputItem.isNullObject)  missing.push('CEG_Input');
      if (guessItem.isNullObject)  missing.push('CEG_Guess');

      if (missing.length > 0) {
        writeLog('Goal Seek: Missing named range(s): ' + missing.join(', '), 'error');
        return;
      }

      writeLog('Goal Seek: Named ranges found — CEG_Target, CEG_Input, CEG_Guess.', 'info');

      var targetRange = targetItem.getRange();
      var inputRange  = inputItem.getRange();
      var guessRange  = guessItem.getRange();

      // ── Step 2: Enter the iterative loop ────────────────────────────────────
      return goalSeekIterate(context, targetRange, inputRange, guessRange,
                             0, MAX_ITERATIONS, TOLERANCE);
    });
  })
  .catch(function (error) {
    writeLog('Goal Seek error: ' + error.message, 'error');
  })
  .then(function () {
    event.completed();
  });
}

// Recursive helper — one Promise chain per iteration so Excel can recalculate
// between each sync() call.
function goalSeekIterate(context, targetRange, inputRange, guessRange,
                         iteration, maxIter, tol) {

  if (iteration >= maxIter) {
    writeLog('Goal Seek: Did not converge after ' + maxIter + ' iterations.', 'error');
    return Promise.resolve();
  }

  // Load both values before the sync so we get them in one round-trip
  targetRange.load('values');
  guessRange.load('values');

  return context.sync().then(function () {
    var targetValue = targetRange.values[0][0];
    var guessValue  = guessRange.values[0][0];

    // Log progress on first iteration and every 50 after that
    if (iteration === 0 || iteration % 50 === 0) {
      writeLog('Goal Seek [iter ' + iteration + ']: CEG_Target = ' + targetValue, 'info');
    }

    // Convergence check — stop when |CEG_Target| is close enough to zero
    if (Math.abs(targetValue) <= tol) {
      writeLog(
        'Goal Seek: Converged in ' + iteration + ' iteration(s). ' +
        'CEG_Target = ' + targetValue,
        'success'
      );
      return;
    }

    // Copy CEG_Guess → CEG_Input to advance the iteration
    inputRange.values = [[guessValue]];

    // Sync triggers Excel recalculation; then recurse for the next iteration
    return context.sync().then(function () {
      return goalSeekIterate(context, targetRange, inputRange, guessRange,
                             iteration + 1, maxIter, tol);
    });
  });
}
