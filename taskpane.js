// taskpane.js — Excel Macro Add-in task pane
//
// Two UI sections:
//   1. Rename Map panel  — reflects addin_rename_map written by readRenameList()
//   2. Activity Log      — reflects log entries written by all ribbon functions
//
// Both sections read from localStorage (same GitHub Pages origin as commands.js)
// and update via the 'storage' event + a 500 ms polling fallback.

var lastDisplayedId = 0;
var nrPanelExpanded = false;

// ── Initialise ────────────────────────────────────────────────────────────────
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById('clearLogBtn').addEventListener('click', clearLog);
    document.getElementById('nr-header').addEventListener('click', toggleNrPanel);

    loadRenamePanelFromStorage();
    loadExistingLogs();
    listenForStorageChanges();
  }
});

// ═══════════════════════════════════════════════════════════════════════════════
// RENAME MAP PANEL
// ═══════════════════════════════════════════════════════════════════════════════

function loadRenamePanelFromStorage() {
  var map = getRenameMap();
  var ts  = localStorage.getItem('addin_rename_map_ts') || '';
  renderRenameMap(map, ts);
}

function getRenameMap() {
  try { return JSON.parse(localStorage.getItem('addin_rename_map') || '[]'); } catch (e) { return []; }
}

function renderRenameMap(map, ts) {
  var badge   = document.getElementById('nr-badge');
  var tsEl    = document.getElementById('nr-ts');
  var ul      = document.getElementById('nr-list');
  var emptyEl = document.getElementById('nr-empty');

  badge.textContent = map.length;
  badge.className   = 'badge' + (map.length === 0 ? ' empty' : '');
  tsEl.textContent  = ts ? 'loaded ' + ts : '';

  ul.innerHTML = '';

  if (map.length === 0) {
    emptyEl.style.display = '';
  } else {
    emptyEl.style.display = 'none';
    map.forEach(function (pair) {
      var li = document.createElement('li');
      li.className = 'nr-item';
      li.innerHTML =
        '<span class="nr-old">'    + escapeHtml(pair.oldName) + '</span>' +
        '<span class="nr-arrow">'  + '→'                      + '</span>' +
        '<span class="nr-new">'    + escapeHtml(pair.newName) + '</span>';
      ul.appendChild(li);
    });

    // Auto-expand when fresh data arrives
    if (!nrPanelExpanded) { setNrPanelOpen(true); }
  }
}

function toggleNrPanel() { setNrPanelOpen(!nrPanelExpanded); }

function setNrPanelOpen(open) {
  nrPanelExpanded = open;
  document.getElementById('nr-body').classList.toggle('nr-collapsed', !open);
  document.getElementById('nr-chevron').classList.toggle('open', open);
  document.getElementById('nr-header').setAttribute('aria-expanded', String(open));
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

function listenForStorageChanges() {
  // Primary: storage event (fires when a different iframe writes to localStorage)
  window.addEventListener('storage', function (e) {
    if (e.key === 'addin_log_latest' && e.newValue) {
      try {
        var entry = JSON.parse(e.newValue);
        if (entry.id > lastDisplayedId) {
          removeEmptyState();
          appendLogEntry(entry);
          lastDisplayedId = entry.id;
        }
      } catch (err) { /* ignore */ }
    }
    if (e.key === 'addin_rename_map_updated') {
      loadRenamePanelFromStorage();
    }
  });

  // Fallback: poll every 500 ms (same-context environments skip the storage event)
  var lastPanelTs = '';
  setInterval(function () {
    // New log entries
    var logs    = getStoredLogs();
    var newLogs = logs.filter(function (e) { return e.id > lastDisplayedId; });
    newLogs.forEach(function (entry) {
      removeEmptyState();
      appendLogEntry(entry);
      lastDisplayedId = entry.id;
    });

    // Rename map updates
    var storedTs = localStorage.getItem('addin_rename_map_ts') || '';
    if (storedTs && storedTs !== lastPanelTs) {
      lastPanelTs = storedTs;
      loadRenamePanelFromStorage();
    }
  }, 500);
}

// ── Log helpers ───────────────────────────────────────────────────────────────
function getStoredLogs() {
  try { return JSON.parse(localStorage.getItem('addin_logs') || '[]'); } catch (e) { return []; }
}

function appendLogEntry(entry) {
  var list = document.getElementById('log-list');
  var li   = document.createElement('li');
  li.className = 'log-entry ' + (entry.type || 'info');

  var indicator = { success: '✓', error: '✕', info: '●' }[entry.type] || '●';

  li.innerHTML =
    '<span class="timestamp">' + escapeHtml(entry.timestamp) + '</span>' +
    '<span class="indicator">' + indicator                   + '</span>' +
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

function escapeHtml(str) {
  var div = document.createElement('div');
  div.appendChild(document.createTextNode(String(str)));
  return div.innerHTML;
}
