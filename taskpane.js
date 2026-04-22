// taskpane.js — Excel Macro Add-in task pane
//
// All commands run directly from the task pane. The ribbon only opens the
// task pane; there are no ExecuteFunction ribbon buttons, so no separate
// function file (commands.html/js) is needed.

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById('clearLogBtn').addEventListener('click', clearLog);
    document.getElementById('goalSeekBtn').addEventListener('click', goalSeek);
    showEmptyState();
  }
});

// ═══════════════════════════════════════════════════════════════════════════════
// ACTIVITY LOG
// ═══════════════════════════════════════════════════════════════════════════════

function writeLog(message, type) {
  removeEmptyState();

  var list = document.getElementById('log-list');
  var li   = document.createElement('li');
  li.className = 'log-entry ' + (type || 'info');

  var indicator = { success: '✓', error: '✕', info: '●' }[type] || '●';
  var timestamp = new Date().toLocaleTimeString();

  li.innerHTML =
    '<span class="timestamp">' + escapeHtml(timestamp) + '</span>' +
    '<span class="indicator">' + indicator              + '</span>' +
    '<span class="message">'   + escapeHtml(message)    + '</span>';

  list.appendChild(li);
  li.scrollIntoView({ behavior: 'smooth', block: 'end' });
}

function clearLog() {
  document.getElementById('log-list').innerHTML = '';
  showEmptyState();
}

function showEmptyState() {
  if (!document.getElementById('empty-state')) {
    var p = document.createElement('p');
    p.id          = 'empty-state';
    p.className   = 'empty-state';
    p.textContent = 'No activity yet. Click a command button to get started.';
    document.getElementById('log-list').appendChild(p);
  }
}

function removeEmptyState() {
  var el = document.getElementById('empty-state');
  if (el) { el.parentNode.removeChild(el); }
}

function escapeHtml(str) {
  var div = document.createElement('div');
  div.appendChild(document.createTextNode(String(str)));
  return div.innerHTML;
}

// ═══════════════════════════════════════════════════════════════════════════════
// COMMAND — Goal Seek
// Requires three named ranges: CEG_Target, CEG_Input, CEG_Guess.
// Iterates CEG_Guess → CEG_Input until |CEG_Target| ≤ TOLERANCE.
// ═══════════════════════════════════════════════════════════════════════════════

function goalSeek() {
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
  });
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
