// buttons.js — Clearway Project Financial Model Add-in
//
// Task-pane button handlers for the add-in.
//
// Each export wraps the corresponding commands.js function with a .catch that
// logs the error to writeLog. This keeps commands.js free of error-handling
// boilerplate so it can be shared unchanged with the Claude on Excel skill
// package (skill-package/skills.js provides its own equivalent wrapper).

import * as commands from './commands.js';
import { writeLog }  from './log.js';

// ── Wrapper factory ───────────────────────────────────────────────────────────

function _buttonWrap(fn, name) {
  return function () {
    return fn().catch(function (error) {
      writeLog(name + ' error: ' + error.message, 'error');
    });
  };
}

// ── Button handlers ───────────────────────────────────────────────────────────

export var checkModel                 = _buttonWrap(commands.checkModel,                 'Check Model');
export var solveTEIUpfrontInvestment  = _buttonWrap(commands.solveTEIUpfrontInvestment,  'Solve TEI Upfront Investment');
export var findTEPshipFlipDate        = _buttonWrap(commands.findTEPshipFlipDate,         'Find TE Partnership Flip Date');
export var solveTermDebt              = _buttonWrap(commands.solveTermDebt,               'Solve Term Debt');
export var iterateTermDebt            = _buttonWrap(commands.iterateTermDebt,             'Iterate Term Debt');
export var solveCWENUpfrontInvestment = _buttonWrap(commands.solveCWENUpfrontInvestment,  'Solve CWEN Investment');
export var solveCE2UpfrontInvestment  = _buttonWrap(commands.solveCE2UpfrontInvestment,   'Solve Third-Party CE Investment');
export var solveCapexCFCircularity    = _buttonWrap(commands.solveCapexCFCircularity,     'Solve CapEx CF Circularity');
export var solveCapitalStack          = _buttonWrap(commands.solveCapitalStack,           'Solve Capital Stack');
export var runScenarios               = _buttonWrap(commands.runScenarios,                'Run Scenarios');
export var pasteMetrics               = _buttonWrap(commands.pasteMetrics,                'Paste Metrics');
