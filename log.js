// log.js — Clearway Project Financial Model Add-in
// Activity log utilities, shared across modules.

export function writeLog(message, type) {
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

export function clearLog() {
  document.getElementById('log-list').innerHTML = '';
  showEmptyState();
}

export function showEmptyState() {
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
