// skill-package/log.js — Clearway Project Financial Model Add-in
//
// Skill-context implementation of the log interface.
//
// writeLog() is synchronous — it buffers entries in memory so commands.js
// can call it identically to the add-in DOM version with no code changes.
//
// flushLog() is async — it writes all buffered entries to the "Claude Log"
// worksheet (creating it if it doesn't exist, appending if it does) and
// returns the entries array for inclusion in the skill result.

var _buffer = [];

// ── Public interface (mirrors add-in log.js) ─────────────────────────────────

export function writeLog(message, level) {
  _buffer.push({
    timestamp: new Date().toISOString(),
    level:     level || 'info',
    message:   message
  });
}

// Not called by commands.js; provided for interface completeness.
export function clearLog() {
  _buffer = [];
}

// Not called by commands.js; no-op in skill context.
export function showEmptyState() {}

// ── Skill-only export ─────────────────────────────────────────────────────────

// Called by skills.js after each command completes.
// Drains the buffer, writes all entries to the "Claude Log" worksheet,
// and returns the entries so the skill wrapper can pass them back to Claude.
export function flushLog() {
  var entries = _buffer.slice();
  _buffer = [];

  if (entries.length === 0) return Promise.resolve(entries);

  var SHEET_NAME = 'Claude Log';

  return Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAME);
    sheet.load('isNullObject');

    return context.sync().then(function () {
      var rows = entries.map(function (e) {
        return [e.timestamp, e.level, e.message];
      });

      if (sheet.isNullObject) {
        // First run — create sheet with a header row, then write entries
        sheet = context.workbook.worksheets.add(SHEET_NAME);
        sheet.getRangeByIndexes(0, 0, 1, 3).values = [['Timestamp', 'Level', 'Message']];
        sheet.getRangeByIndexes(1, 0, rows.length, 3).values = rows;
        return context.sync().then(function () { return entries; });
      }

      // Sheet exists — find the first empty row and append
      var usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load(['isNullObject', 'rowCount']);
      return context.sync().then(function () {
        var nextRow = usedRange.isNullObject ? 0 : usedRange.rowCount;
        sheet.getRangeByIndexes(nextRow, 0, rows.length, 3).values = rows;
        return context.sync().then(function () { return entries; });
      });
    });
  });
}
