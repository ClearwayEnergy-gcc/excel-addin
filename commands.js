// commands.js — Ribbon button handlers
//
// Runs inside the hidden function file iframe (commands.html).
// All ribbon buttons registered via Office.actions.associate() below.
//
// localStorage keys used (shared with taskpane.js, same GitHub Pages origin):
//   addin_logs              — activity log entries array
//   addin_log_latest        — latest single entry (triggers storage event)
//   addin_log_counter       — monotonic ID counter for log entries
//   addin_rename_map        — [{oldName, newName}] array set by readRenameList()
//   addin_rename_map_ts     — human timestamp of last readRenameList() call
//   addin_rename_map_updated — written to trigger task pane storage event

Office.onReady(function () {
  Office.actions.associate('readNamedRanges',    readNamedRanges);
  Office.actions.associate('readRenameList',     readRenameList);
  Office.actions.associate('applyNamedRangeRename', applyNamedRangeRename);
  Office.actions.associate('goalSeek',           goalSeek);
});

// ── Manage Named-Ranges sheet name (shared constant) ─────────────────────────
var MGMT_SHEET = 'Manage Named-Ranges';

// ═══════════════════════════════════════════════════════════════════════════════
// SHARED UTILITIES
// ═══════════════════════════════════════════════════════════════════════════════

function writeLog(message, type) {
  var timestamp = new Date().toLocaleTimeString();
  var counter   = parseInt(localStorage.getItem('addin_log_counter') || '0') + 1;
  localStorage.setItem('addin_log_counter', counter);

  var entry = { id: counter, timestamp: timestamp, message: message, type: type || 'info' };

  var logs = [];
  try { logs = JSON.parse(localStorage.getItem('addin_logs') || '[]'); } catch (e) {}
  logs.push(entry);
  localStorage.setItem('addin_logs', JSON.stringify(logs));
  localStorage.setItem('addin_log_latest', JSON.stringify(entry));
}

// Returns the stored rename map, or [] if none loaded yet.
function getRenameMap() {
  try { return JSON.parse(localStorage.getItem('addin_rename_map') || '[]'); } catch (e) { return []; }
}

// ═══════════════════════════════════════════════════════════════════════════════
// BUTTON 1 — Read Named-Ranges
// Creates (or refreshes) a worksheet called "Manage Named-Ranges".
//   Column A — all named-range names found in the workbook (read-only reference)
//   Column B — user fills in new names; leave blank to skip that range
// Sweeps both workbook-scoped and sheet-scoped named ranges.
// ═══════════════════════════════════════════════════════════════════════════════

function readNamedRanges(event) {
  Excel.run(function (context) {

    // ── Pass 1: load workbook names + worksheet list ─────────────────────────
    var wbNames = context.workbook.names;
    wbNames.load(['name', 'type', 'formula', 'visible']);

    var sheets = context.workbook.worksheets;
    sheets.load('items/name');

    return context.sync().then(function () {

      // Collect workbook-scoped names
      var allNames = [];
      wbNames.items.forEach(function (item) {
        allNames.push({ name: item.name, scope: 'Workbook' });
      });

      // Queue loading of each sheet's named ranges
      var sheetLoaders = sheets.items
        .filter(function (s) { return s.name !== MGMT_SHEET; }) // skip the mgmt sheet itself
        .map(function (sheet) {
          var sn = sheet.names;
          sn.load(['name', 'type', 'formula']);
          return { sheetName: sheet.name, names: sn };
        });

      // ── Pass 2: resolve sheet-scoped names ───────────────────────────────
      return context.sync().then(function () {
        sheetLoaders.forEach(function (loader) {
          loader.names.items.forEach(function (item) {
            allNames.push({ name: item.name, scope: loader.sheetName });
          });
        });

        // ── Get or create the management sheet ───────────────────────────
        var mgmtSheet = context.workbook.worksheets.getItemOrNullObject(MGMT_SHEET);
        mgmtSheet.load('isNullObject');

        return context.sync().then(function () {
          var sheet;
          if (mgmtSheet.isNullObject) {
            sheet = context.workbook.worksheets.add(MGMT_SHEET);
          } else {
            sheet = mgmtSheet;
            // Clear previous content (columns A & B, generous row range)
            sheet.getRange('A1:B2000').clear();
          }

          // ── Write headers ────────────────────────────────────────────────
          var headerRange = sheet.getRange('A1:B1');
          headerRange.values = [['Existing Name', 'New Name (leave blank to skip)']];
          headerRange.format.font.bold    = true;
          headerRange.format.fill.color   = '#D6E4F0';
          headerRange.format.font.color   = '#1A3A5C';

          // ── Write named-range names to Column A ──────────────────────────
          if (allNames.length > 0) {
            var dataRange = sheet.getRangeByIndexes(1, 0, allNames.length, 1);
            dataRange.values = allNames.map(function (n) { return [n.name]; });
          }

          // ── Column formatting ────────────────────────────────────────────
          sheet.getRange('A:A').format.columnWidth = 220;
          sheet.getRange('B:B').format.columnWidth = 220;

          // Freeze header row
          sheet.freezePanes.freezeRows(1);

          // Bring the sheet into view
          sheet.activate();

          return context.sync().then(function () {
            writeLog(
              'Read Named-Ranges: "' + MGMT_SHEET + '" sheet created/refreshed with ' +
              allNames.length + ' named range(s). Fill Column B with new names, then click "Read Rename List".',
              'success'
            );
          });
        });
      });
    });
  })
  .catch(function (error) {
    writeLog('Read Named-Ranges error: ' + error.message, 'error');
  })
  .then(function () { event.completed(); });
}

// ═══════════════════════════════════════════════════════════════════════════════
// BUTTON 2 — Read Rename List
// Reads Columns A and B from the "Manage Named-Ranges" sheet.
// Builds an array of {oldName, newName} pairs, skipping rows where Column B
// is empty, and stores it to localStorage for use by Apply Rename.
// ═══════════════════════════════════════════════════════════════════════════════

function readRenameList(event) {
  Excel.run(function (context) {

    var sheet = context.workbook.worksheets.getItemOrNullObject(MGMT_SHEET);
    sheet.load('isNullObject');

    return context.sync().then(function () {

      if (sheet.isNullObject) {
        writeLog(
          'Read Rename List: Sheet "' + MGMT_SHEET + '" not found. ' +
          'Run "Read Named-Ranges" first.',
          'error'
        );
        return;
      }

      // Read all used data from the sheet
      var usedRange = sheet.getUsedRange();
      usedRange.load(['values', 'rowCount']);

      return context.sync().then(function () {
        var values    = usedRange.values;
        var renameMap = [];

        // Row 0 is the header — start from row 1
        for (var i = 1; i < values.length; i++) {
          var oldName = String(values[i][0] || '').trim();
          var newName = String(values[i][1] || '').trim();

          if (!oldName) { continue; }           // blank old-name row — skip
          if (!newName) { continue; }           // no new name entered — skip
          if (oldName === newName) { continue; } // identity — nothing to do

          renameMap.push({ oldName: oldName, newName: newName });
        }

        // ── Persist to localStorage ──────────────────────────────────────
        var ts = new Date().toLocaleTimeString();
        localStorage.setItem('addin_rename_map',         JSON.stringify(renameMap));
        localStorage.setItem('addin_rename_map_ts',      ts);
        localStorage.setItem('addin_rename_map_updated', String(Date.now())); // triggers task pane

        // ── Log ──────────────────────────────────────────────────────────
        if (renameMap.length === 0) {
          writeLog(
            'Read Rename List: No rename pairs found. ' +
            'Make sure Column B has at least one new name entered.',
            'info'
          );
        } else {
          writeLog('Read Rename List: ' + renameMap.length + ' pair(s) loaded into memory.', 'success');
          renameMap.forEach(function (pair) {
            writeLog('  "' + pair.oldName + '"  →  "' + pair.newName + '"', 'info');
          });
          writeLog('Switch to the target workbook, then click "Apply Rename".', 'info');
        }
      });
    });
  })
  .catch(function (error) {
    writeLog('Read Rename List error: ' + error.message, 'error');
  })
  .then(function () { event.completed(); });
}

// ═══════════════════════════════════════════════════════════════════════════════
// BUTTON 3 — Apply Named-Range Rename
// Reads the rename map from localStorage and applies it to the currently active
// workbook. Works on any open workbook — intended to be run on the TARGET file
// after using the source file to build the rename map.
//
// Rename strategy (Office.js has no direct rename API):
//   1. Load formula references for all old names (1 sync)
//   2. Batch-add all new names with the same formula references (1 sync)
//   3. Batch-delete all old names (1 sync)
// ═══════════════════════════════════════════════════════════════════════════════

function applyNamedRangeRename(event) {
  var renameMap = getRenameMap();

  if (renameMap.length === 0) {
    writeLog(
      'Apply Rename: No rename map in memory. ' +
      'Run "Read Rename List" on the source workbook first.',
      'error'
    );
    event.completed();
    return;
  }

  writeLog('Apply Rename: Applying ' + renameMap.length + ' rename(s) to the active workbook…', 'info');

  Excel.run(function (context) {
    var names = context.workbook.names;

    // ── Step 1: Load existence + formula for every old name ──────────────────
    var loaders = renameMap.map(function (pair) {
      var item = names.getItemOrNullObject(pair.oldName);
      item.load(['isNullObject', 'formula']);
      return { pair: pair, item: item };
    });

    return context.sync().then(function () {

      // Separate found vs not-found
      var toRename = [];
      loaders.forEach(function (loader) {
        if (loader.item.isNullObject) {
          writeLog(
            '  Skipped "' + loader.pair.oldName + '" — not found in this workbook.',
            'info'
          );
        } else {
          toRename.push({
            oldName: loader.pair.oldName,
            newName: loader.pair.newName,
            formula: loader.item.formula,
            oldItem: loader.item
          });
        }
      });

      var skipped = renameMap.length - toRename.length;

      if (toRename.length === 0) {
        writeLog(
          'Apply Rename: None of the named ranges in the map exist in this workbook. ' +
          'Make sure you have the correct workbook active.',
          'error'
        );
        return;
      }

      // ── Step 2: Batch-add all new names ──────────────────────────────────
      toRename.forEach(function (entry) {
        names.add(entry.newName, entry.formula);
      });

      return context.sync().then(function () {

        // ── Step 3: Batch-delete all old names ───────────────────────────
        toRename.forEach(function (entry) {
          entry.oldItem.delete();
        });

        return context.sync().then(function () {
          // Log individual results
          toRename.forEach(function (entry) {
            writeLog('  ✓  "' + entry.oldName + '"  →  "' + entry.newName + '"', 'success');
          });

          writeLog(
            'Apply Rename: Complete — ' + toRename.length + ' renamed' +
            (skipped > 0 ? ', ' + skipped + ' not found in this workbook.' : '.'),
            'success'
          );
        });
      });
    });
  })
  .catch(function (error) {
    writeLog('Apply Rename error: ' + error.message, 'error');
  })
  .then(function () { event.completed(); });
}

// ═══════════════════════════════════════════════════════════════════════════════
// BUTTON 4 — Goal Seek
// Requires three named ranges: CEG_Target, CEG_Input, CEG_Guess.
// Iterates CEG_Guess → CEG_Input until |CEG_Target| ≤ TOLERANCE.
// ═══════════════════════════════════════════════════════════════════════════════

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
      var missing = [];
      if (targetItem.isNullObject) missing.push('CEG_Target');
      if (inputItem.isNullObject)  missing.push('CEG_Input');
      if (guessItem.isNullObject)  missing.push('CEG_Guess');

      if (missing.length > 0) {
        writeLog('Goal Seek: Missing named range(s): ' + missing.join(', '), 'error');
        return;
      }

      writeLog('Goal Seek: Found CEG_Target, CEG_Input, CEG_Guess. Starting…', 'info');

      return goalSeekIterate(
        context,
        targetItem.getRange(),
        inputItem.getRange(),
        guessItem.getRange(),
        0, MAX_ITERATIONS, TOLERANCE
      );
    });
  })
  .catch(function (error) {
    writeLog('Goal Seek error: ' + error.message, 'error');
  })
  .then(function () { event.completed(); });
}

function goalSeekIterate(context, targetRange, inputRange, guessRange, iteration, maxIter, tol) {
  if (iteration >= maxIter) {
    writeLog('Goal Seek: Did not converge after ' + maxIter + ' iterations.', 'error');
    return Promise.resolve();
  }

  targetRange.load('values');
  guessRange.load('values');

  return context.sync().then(function () {
    var targetValue = targetRange.values[0][0];
    var guessValue  = guessRange.values[0][0];

    if (iteration === 0 || iteration % 50 === 0) {
      writeLog('Goal Seek [iter ' + iteration + ']: CEG_Target = ' + targetValue, 'info');
    }

    if (Math.abs(targetValue) <= tol) {
      writeLog(
        'Goal Seek: Converged in ' + iteration + ' iteration(s). CEG_Target = ' + targetValue,
        'success'
      );
      return;
    }

    inputRange.values = [[guessValue]];
    return context.sync().then(function () {
      return goalSeekIterate(context, targetRange, inputRange, guessRange,
                             iteration + 1, maxIter, tol);
    });
  });
}
