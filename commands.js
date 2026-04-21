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
  Office.actions.associate('readNamedRanges', readNamedRanges);
  Office.actions.associate('writeData',       writeData);
  Office.actions.associate('goalSeek',        goalSeek);
});

// ═══════════════════════════════════════════════════════════════════════════════
// SHARED STORAGE UTILITIES
// Functions below read/write to localStorage so that:
//   • All ribbon button handlers (this file) can access persisted data
//   • The task pane can read the same data (same GitHub Pages origin)
//   • Data survives task pane close/reopen and Excel session restarts
// ═══════════════════════════════════════════════════════════════════════════════

// ── Activity log ──────────────────────────────────────────────────────────────
// Writes a timestamped entry to localStorage. The task pane picks it up via
// the storage event (cross-iframe) or a polling fallback.

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

// ── Named-range list ──────────────────────────────────────────────────────────
// Returns the full list stored by readNamedRanges(), or [] if not yet populated.
// Each item: { name, type, formula, scope }
//   name    — the range name (e.g. "CEG_Target")
//   type    — "Range" | "String" | "Integer" | etc.
//   formula — the reference string (e.g. "=Sheet1!$A$1")
//   scope   — "Workbook" | "<SheetName>" for sheet-scoped names

function getNamedRangesList() {
  try {
    return JSON.parse(localStorage.getItem('addin_named_ranges') || '[]');
  } catch (e) {
    return [];
  }
}

// Returns the stored metadata object for a single named range, or null.
// Usage: var meta = getNamedRange('CEG_Target');
function getNamedRange(name) {
  var list = getNamedRangesList();
  for (var i = 0; i < list.length; i++) {
    if (list[i].name === name) { return list[i]; }
  }
  return null;
}

// ═══════════════════════════════════════════════════════════════════════════════
// BUTTON 1 — Read Named-Ranges
// ═══════════════════════════════════════════════════════════════════════════════
// Sweeps all named ranges from the open workbook (both workbook-scoped and
// sheet-scoped) and writes them to localStorage for use by other functions.
// The task pane Named Ranges panel auto-updates when this runs.

function readNamedRanges(event) {
  Excel.run(function (context) {
    // Queue: load workbook-level names + all worksheet names in one round-trip
    var wbNames = context.workbook.names;
    wbNames.load(['name', 'type', 'formula', 'visible']);

    var sheets = context.workbook.worksheets;
    sheets.load('items/name');

    return context.sync().then(function () {
      var rangeList = [];

      // Collect workbook-scoped named ranges
      wbNames.items.forEach(function (item) {
        rangeList.push({
          name:    item.name,
          type:    item.type,
          formula: item.formula,
          scope:   'Workbook'
        });
      });

      // Queue loading of each worksheet's named ranges
      var sheetNameLoaders = sheets.items.map(function (sheet) {
        var sn = sheet.names;
        sn.load(['name', 'type', 'formula', 'visible']);
        return { sheetName: sheet.name, names: sn };
      });

      // Second sync: resolve all sheet-level named ranges
      return context.sync().then(function () {
        sheetNameLoaders.forEach(function (loader) {
          loader.names.items.forEach(function (item) {
            rangeList.push({
              name:    item.name,
              type:    item.type,
              formula: item.formula,
              scope:   loader.sheetName
            });
          });
        });

        // ── Persist to localStorage ────────────────────────────────────────
        localStorage.setItem('addin_named_ranges', JSON.stringify(rangeList));
        localStorage.setItem('addin_named_ranges_ts', new Date().toLocaleTimeString());
        // Separate key write triggers the task pane storage listener
        localStorage.setItem('addin_named_ranges_updated', String(Date.now()));

        // ── Log results ────────────────────────────────────────────────────
        if (rangeList.length === 0) {
          writeLog('Read Named-Ranges: No named ranges found in this workbook.', 'info');
        } else {
          writeLog(
            'Read Named-Ranges: Loaded ' + rangeList.length + ' named range(s).',
            'success'
          );
          rangeList.forEach(function (r) {
            writeLog('  → ' + r.name + ' [' + r.scope + ']: ' + r.formula, 'info');
          });
        }
      });
    });
  })
  .catch(function (error) {
    writeLog('Read Named-Ranges error: ' + error.message, 'error');
  })
  .then(function () {
    event.completed();
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// BUTTON 2 — Write Data
// ═══════════════════════════════════════════════════════════════════════════════
// Writes a small sample table to A1:C3 on the active sheet.

function writeData(event) {
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    var dataRange = sheet.getRange('A1:C3');
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

// ═══════════════════════════════════════════════════════════════════════════════
// BUTTON 3 — Goal Seek
// ═══════════════════════════════════════════════════════════════════════════════
// Requires three named ranges in the workbook:
//   CEG_Target — formula cell that should reach 0 (e.g. model output - target)
//   CEG_Input  — input cell whose value is overwritten each iteration
//   CEG_Guess  — formula cell computing the next candidate value for CEG_Input
//
// Each iteration:
//   1. Read CEG_Target — stop if |value| ≤ TOLERANCE
//   2. Read CEG_Guess
//   3. Write CEG_Guess → CEG_Input  (triggers Excel recalculation)
//   4. Repeat

function goalSeek(event) {
  var MAX_ITERATIONS = 1000;
  var TOLERANCE      = 1e-10;

  Excel.run(function (context) {
    var names = context.workbook.names;

    var targetItem = names.getItemOrNullObject('CEG_Target');
    var inputItem  = names.getItemOrNullObject('CEG_Input');
    var guessItem  = names.getItemOrNullObject('CEG_Guess');

    targetItem.load('isNullObject');
    inputItem.load('isNullObject');
    guessItem.load('isNullObject');

    return context.sync().then(function () {

      // Verify all three named ranges exist
      var missing = [];
      if (targetItem.isNullObject) missing.push('CEG_Target');
      if (inputItem.isNullObject)  missing.push('CEG_Input');
      if (guessItem.isNullObject)  missing.push('CEG_Guess');

      if (missing.length > 0) {
        writeLog('Goal Seek: Missing named range(s): ' + missing.join(', '), 'error');
        return;
      }

      writeLog('Goal Seek: Found CEG_Target, CEG_Input, CEG_Guess. Starting...', 'info');

      var targetRange = targetItem.getRange();
      var inputRange  = inputItem.getRange();
      var guessRange  = guessItem.getRange();

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

// Recursive helper — one Promise chain per iteration so Excel recalculates
// between each sync() call.
function goalSeekIterate(context, targetRange, inputRange, guessRange,
                         iteration, maxIter, tol) {

  if (iteration >= maxIter) {
    writeLog('Goal Seek: Did not converge after ' + maxIter + ' iterations.', 'error');
    return Promise.resolve();
  }

  targetRange.load('values');
  guessRange.load('values');

  return context.sync().then(function () {
    var targetValue = targetRange.values[0][0];
    var guessValue  = guessRange.values[0][0];

    // Log progress on first iteration and every 50 after that
    if (iteration === 0 || iteration % 50 === 0) {
      writeLog('Goal Seek [iter ' + iteration + ']: CEG_Target = ' + targetValue, 'info');
    }

    // Convergence check
    if (Math.abs(targetValue) <= tol) {
      writeLog(
        'Goal Seek: Converged in ' + iteration + ' iteration(s). ' +
        'CEG_Target = ' + targetValue,
        'success'
      );
      return;
    }

    // Copy CEG_Guess → CEG_Input, then sync to trigger recalculation
    inputRange.values = [[guessValue]];

    return context.sync().then(function () {
      return goalSeekIterate(context, targetRange, inputRange, guessRange,
                             iteration + 1, maxIter, tol);
    });
  });
}
