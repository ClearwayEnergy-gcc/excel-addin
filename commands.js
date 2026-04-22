// commands.js — Clearway Project Financial Model Add-in
//
// Command functions triggered by task-pane buttons.
// Relies on writeLog() defined in taskpane.js (shared global scope).

// ═══════════════════════════════════════════════════════════════════════════════
// COMMAND — Check for Clearway Project Financial Model
// Looks for the named range "CEG_ModelTemplateVersion" in the workbook.
// ═══════════════════════════════════════════════════════════════════════════════

function checkModel() {
  Excel.run(function (context) {
    var item = context.workbook.names.getItemOrNullObject('CEG_ModelTemplateVersion');
    item.load('isNullObject,value');

    return context.sync().then(function () {
      if (item.isNullObject) {
        writeLog('This workbook does not appear to be a Clearway Project Financial Model (CEG_ModelTemplateVersion not found).', 'error');
      } else {
        writeLog('Clearway Project Financial Model detected. CEG_ModelTemplateVersion = ' + item.value, 'success');
      }
    });
  })
  .catch(function (error) {
    writeLog('Check Model error: ' + error.message, 'error');
  });
}
