// skill-package/skills.js — Clearway Project Financial Model Add-in
//
// Thin wrapper layer for the Claude on Excel skill context.
//
// Each exported function calls the matching command from commands.js, then
// flushes the log buffer to the "Claude Log" worksheet and returns a
// structured result that Claude can incorporate into its reply:
//
//   { success: true,  log: [ { timestamp, level, message }, ... ] }
//   { success: false, log: [ ... ] }           // errors captured in log
//   { success: false, error: '...', log: [...] } // unexpected rejection
//
// commands.js and log.js are identical-interface counterparts to the add-in
// versions. No changes to commands.js are needed between environments.

import * as commands from './commands.js';
import { flushLog }  from './log.js';

// ── Wrapper factory ───────────────────────────────────────────────────────────

function _wrap(fn) {
  return function () {
    return fn()
      .then(function ()         { return flushLog(); })
      .then(function (entries)  {
        return {
          success: !entries.some(function (e) { return e.level === 'error'; }),
          log:     entries
        };
      })
      .catch(function (err) {
        // Unexpected rejection (commands.js catches internally, but just in case)
        return flushLog().then(function (entries) {
          return { success: false, error: err.message, log: entries };
        });
      });
  };
}

// ── Exported skills ───────────────────────────────────────────────────────────

export var checkModel                 = _wrap(commands.checkModel);
export var solveTEIUpfrontInvestment  = _wrap(commands.solveTEIUpfrontInvestment);
export var findTEPshipFlipDate        = _wrap(commands.findTEPshipFlipDate);
export var solveTermDebt              = _wrap(commands.solveTermDebt);
export var iterateTermDebt            = _wrap(commands.iterateTermDebt);
export var solveCWENUpfrontInvestment = _wrap(commands.solveCWENUpfrontInvestment);
export var solveCE2UpfrontInvestment  = _wrap(commands.solveCE2UpfrontInvestment);
export var solveCapexCFCircularity    = _wrap(commands.solveCapexCFCircularity);
export var solveCapitalStack          = _wrap(commands.solveCapitalStack);
export var runScenarios               = _wrap(commands.runScenarios);
export var pasteMetrics               = _wrap(commands.pasteMetrics);
