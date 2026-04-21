// taskpane.js — Excel Macro Add-in task pane
//
// Manages two UI sections:
//   1. Named Ranges panel — reflects the list stored by readNamedRanges()
//   2. Activity Log       — reflects log entries written by all ribbon functions
//
// Both sections read from localStorage (same GitHub Pages origin as commands.js)
// and update in real time via the 'storage' event + a polling fallback.

var lastDisplayedId    = 0;   // highest log entry ID rendered so far
var nrPanelExpanded    = false;

// ── Initialise ────────────────────────────────────────────────────────────────
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    // Wire up static buttons
    document.getElementById('clearLogBtn').addEventListener('click', clearLog);
    document.getElementById('nr-header').addEventListener('click', toggleNrPanel);

    // Populate panels from any data already in localStorage
    loadNamedRangesPanel();
    loadExistingLogs();

    // Start listening for updates written by commands.js
    listenForStorageChanges();
  }
});

// ═══════════════════════════════════════════════════════════════════════════════
// NAMED RANGES PANEL
// ═══════════════════════════════════════════════════════════════════════════════

function loadNamedRangesPanel() {
  var list = getNamedRangesList();
  var ts   = localStorage.getItem('addin_named_ranges_ts') || '';
  renderNamedRanges(list, ts);
}

function renderNamedRanges(list, ts) {
  var badge   = document.getElementById('nr-badge');
  var tsEl    = document.getElementById('nr-ts');
  var ul      = document.getElementById('nr-list');
  var emptyEl = document.getElementById('nr-empty');

  // Update badge
  badge.textContent = list.length;
  badge.className   = 'badge' + (list.length === 0 ? ' empty' : '');

  // Update timestamp
  tsEl.textContent = ts ? 'updated ' + ts : '';

  // Rebuild list
  ul.innerHTML = '';
  if (list.length === 0) {
    emptyEl.style.display = '';
  } else {
    emptyEl.style.display = 'none';
    list.forEach(function (r) {
      var li = document.createElement('li');
      li.className = 'nr-item';
      li.innerHTML =
        '<span class="nr-name">'    + escapeHtml(r.name)    + '</span>' +
        '<span class="nr-scope">'   + escapeHtml(r.scope)   + '</span>' +
        '<span class="nr-formula">' + escapeHtml(r.formula) + '</span>';
      ul.appendChild(li);
    });

    // Auto-expand the panel when new ranges arrive
    if (!nrPanelExpanded) { setNrPanelOpen(true); }
  }
}

function toggleNrPanel() {
  setNrPanelOpen(!nrPanelExpanded);
}

function setNrPanelOpen(open) {
  nrPanelExpanded = open;
  var body    = document.getElementById('nr-body');
  var chevron = document.getElementById('nr-chevron');
  var header  = document.getElementById('nr-header');

  if (open) {
    body.classList.remove('nr-collapsed');
    chevron.classList.add('open');
    header.setAttribute('aria-expanded', 'true');
  } else {
    body.classList.add('nr-collapsed');
    chevron.classList.remove('open');
    header.setAttribute('aria-expanded', 'false');
  }
}

function getNamedRangesList() {
  try {
    return JSON.parse(localStorage.getItem('addin_named_ranges') || '[]');
  } catch (e) {
    return [];
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// ACTIVITY LOG
// ═══════════════════════════════════════════════════════════════════════════════

function loadExistingLogs() {
  var logs = getStoredLogs();
  if (logs.length === 0) {
    showEmptyState();
  } else {
    logs.forEach(function (entry) {
      appendLogEntry(entry);
      if (entry.id > lastDisplayedId) { lastDisplayedId = entry.id; }
    });
  }
}

// ── Real-time listener ────────────────────────────────────────────────────────
function listenForStorageChanges() {
  // Primary path: storage event fires when a DIFFERENT context (commands iframe)
  // writes to localStorage at the same origin.
  window.addEventListener('storage', function (e) {

    // New log entry
    if (e.key === 'addin_log_latest' && e.newValue) {
      try {
        var entry = JSON.parse(e.newValue);
        if (entry.id > lastDisplayedId) {
          removeEmptyState();
          appendLogEntry(entry);
          lastDisplayedId = entry.id;
        }
      } catch (err) { /* ignore malformed */ }
    }

    // Named ranges updated
    if (e.key === 'addin_named_ranges_updated') {
      loadNamedRangesPanel();
    }
  });

  // Fallback: poll every 500 ms — handles environments where the task pane and
  // function file share the same browsing context (storage event not fired).
  setInterval(function () {
    // Check for new log entries
    var logs    = getStoredLogs();
    var newLogs = logs.filter(function (e) { return e.id > lastDisplayedId; });
    newLogs.forEach(function (entry) {
      removeEmptyState();
      appendLogEntry(entry);
      lastDisplayedId = entry.id;
    });

    // Check for named ranges updates (compare stored timestamp to what we show)
    var storedTs = localStorage.getItem('addin_named_ranges_ts') || '';
    var shownTs  = document.getElementById('nr-ts').textContent;
    if (storedTs && shownTs !== 'updated ' + storedTs) {
      loadNamedRangesPanel();
    }
  }, 500);
}

// ── Log helpers ───────────────────────────────────────────────────────────────
function getStoredLogs() {
  try {
    return JSON.parse(localStorage.getItem('addin_logs') || '[]');
  } catch (e) {
    return [];
  }
}

function appendLogEntry(entry) {
  var list = document.getElementById('log-list');
  var li   = document.createElement('li');
  li.className = 'log-entry ' + (entry.type || 'info');

  var indicator = { success: '✓', error: '✕', info: '●' }[entry.type] || '●';

  li.innerHTML =
    '<span class="timestamp">' + escapeHtml(entry.timestamp) + '</span>' +
    '<span class="indicator">' + indicator + '</span>' +
    '<span class="message">'   + escapeHtml(entry.message)   + '</span>';

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
    p.id          = 'empty-state';
    p.className   = 'empty-state';
    p.textContent = 'No activity yet. Use the ribbon buttons to get started.';
    document.getElementById('log-list').appendChild(p);
  }
}

function removeEmptyState() {
  var el = document.getElementById('empty-state');
  if (el) { el.parentNode.removeChild(el); }
}

// ── Shared utility ────────────────────────────────────────────────────────────
function escapeHtml(str) {
  var div = document.createElement('div');
  div.appendChild(document.createTextNode(String(str)));
  return div.innerHTML;
}
