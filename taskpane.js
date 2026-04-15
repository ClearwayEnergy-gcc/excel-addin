// taskpane.js — Activity Log task pane
//
// Reads log entries written by commands.js via localStorage.
// New entries are picked up via the 'storage' event (Excel Online / WebView2)
// with a polling fallback for environments where cross-iframe storage events
// are not fired.

var lastDisplayedId = 0;

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById('clearLogBtn').addEventListener('click', clearLog);
    loadExistingLogs();
    listenForNewLogs();
  }
});

// ── Load logs that already exist in localStorage on task pane open ────────────
function loadExistingLogs() {
  var logs = getStoredLogs();
  if (logs.length === 0) {
    showEmptyState();
  } else {
    logs.forEach(function (entry) {
      appendLogEntry(entry);
      if (entry.id > lastDisplayedId) {
        lastDisplayedId = entry.id;
      }
    });
  }
}

// ── Listen for new log entries written by commands.js ─────────────────────────
function listenForNewLogs() {
  // Primary: storage event fires when a different context (commands iframe)
  // writes to localStorage at the same origin.
  window.addEventListener('storage', function (e) {
    if (e.key === 'addin_log_latest' && e.newValue) {
      try {
        var entry = JSON.parse(e.newValue);
        if (entry.id > lastDisplayedId) {
          removeEmptyState();
          appendLogEntry(entry);
          lastDisplayedId = entry.id;
        }
      } catch (err) { /* ignore malformed entries */ }
    }
  });

  // Fallback: poll every 500 ms — catches cases where the task pane and
  // function file share the same browsing context (no cross-context event).
  setInterval(function () {
    var logs = getStoredLogs();
    var newLogs = logs.filter(function (e) { return e.id > lastDisplayedId; });
    newLogs.forEach(function (entry) {
      removeEmptyState();
      appendLogEntry(entry);
      lastDisplayedId = entry.id;
    });
  }, 500);
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function getStoredLogs() {
  try {
    return JSON.parse(localStorage.getItem('addin_logs') || '[]');
  } catch (e) {
    return [];
  }
}

function appendLogEntry(entry) {
  var list = document.getElementById('log-list');
  var li = document.createElement('li');
  li.className = 'log-entry ' + (entry.type || 'info');

  var indicator = { success: '✓', error: '✕', info: '●' }[entry.type] || '●';

  li.innerHTML =
    '<span class="timestamp">' + escapeHtml(entry.timestamp) + '</span>' +
    '<span class="indicator">' + indicator + '</span>' +
    '<span class="message">' + escapeHtml(entry.message) + '</span>';

  list.appendChild(li);
  li.scrollIntoView({ behavior: 'smooth', block: 'end' });
}

function clearLog() {
  localStorage.removeItem('addin_logs');
  localStorage.removeItem('addin_log_latest');
  localStorage.removeItem('addin_log_counter');
  lastDisplayedId = 0;
  document.getElementById('log-list').innerHTML = '';
  showEmptyState();
}

function showEmptyState() {
  if (!document.getElementById('empty-state')) {
    var p = document.createElement('p');
    p.id = 'empty-state';
    p.className = 'empty-state';
    p.textContent = 'No activity yet. Use the ribbon buttons to get started.';
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
