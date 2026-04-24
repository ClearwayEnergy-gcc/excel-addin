// buttons.js — Clearway Project Financial Model Add-in
//
// Task-pane button handlers for the add-in.
//
// Each export wraps the corresponding commands.js function with:
//   1. Visual lock — disables all command buttons and highlights the active one
//      so users cannot trigger a second command while one is running.
//   2. Error catch — logs any unhandled rejection to writeLog.
//   3. Visual unlock — restores all buttons once the command completes (or fails).
//
// This keeps commands.js free of UI and error-handling boilerplate so it can
// be shared unchanged with the Claude on Excel skill package.

import * as commands from './commands.js';
import { writeLog }  from './log.js';

// ── Button ID registry ────────────────────────────────────────────────────────
// All command button IDs. clearLogBtn is intentionally excluded — clearing the
// log while a command runs is harmless, so it stays enabled.

var _CMD_BUTTON_IDS = [
  'checkModelBtn',
  'solveTEUpfrontBtn',
  'flipDateBtn',
  'termDebtSolveBtn',
  'iterateTermDebtBtn',
  'solveCWENUpfrontInvestmentBtn',
  'solveCE2UpfrontInvestmentBtn',
  'solveCapexCFCircularityBtn',
  'solveCapitalStackBtn',
  'runScenariosBtn',
  'pasteMetricsBtn'
];

// ── UI lock / unlock helpers ──────────────────────────────────────────────────

function _setRunning(activeId) {
  _CMD_BUTTON_IDS.forEach(function (id) {
    var btn = document.getElementById(id);
    if (!btn) return;
    btn.disabled = true;
    btn.classList.add(id === activeId ? 'cmd-btn-running' : 'cmd-btn-waiting');
  });
}

function _clearRunning() {
  _CMD_BUTTON_IDS.forEach(function (id) {
    var btn = document.getElementById(id);
    if (!btn) return;
    btn.disabled = false;
    btn.classList.remove('cmd-btn-running', 'cmd-btn-waiting');
  });
}

// ── Wrapper factory ───────────────────────────────────────────────────────────

function _buttonWrap(fn, name, buttonId) {
  return function () {
    _setRunning(buttonId);
    return fn()
      .catch(function (error) {
        writeLog(name + ' error: ' + error.message, 'error');
      })
      .then(function () {
        _clearRunning();
      });
  };
}

// ── Button handlers ───────────────────────────────────────────────────────────

export var checkModel                 = _buttonWrap(commands.checkModel,                 'Check Model',                      'checkModelBtn');
export var solveTEIUpfrontInvestment  = _buttonWrap(commands.solveTEIUpfrontInvestment,  'Solve TEI Upfront Investment',     'solveTEUpfrontBtn');
export var findTEPshipFlipDate        = _buttonWrap(commands.findTEPshipFlipDate,         'Find TE Partnership Flip Date',    'flipDateBtn');
export var solveTermDebt              = _buttonWrap(commands.solveTermDebt,               'Solve Term Debt',                  'termDebtSolveBtn');
export var iterateTermDebt            = _buttonWrap(commands.iterateTermDebt,             'Iterate Term Debt',                'iterateTermDebtBtn');
export var solveCWENUpfrontInvestment = _buttonWrap(commands.solveCWENUpfrontInvestment,  'Solve CWEN Investment',            'solveCWENUpfrontInvestmentBtn');
export var solveCE2UpfrontInvestment  = _buttonWrap(commands.solveCE2UpfrontInvestment,   'Solve Third-Party CE Investment',  'solveCE2UpfrontInvestmentBtn');
export var solveCapexCFCircularity    = _buttonWrap(commands.solveCapexCFCircularity,     'Solve CapEx CF Circularity',       'solveCapexCFCircularityBtn');
export var solveCapitalStack          = _buttonWrap(commands.solveCapitalStack,           'Solve Capital Stack',              'solveCapitalStackBtn');
export var runScenarios               = _buttonWrap(commands.runScenarios,                'Run Scenarios',                    'runScenariosBtn');
export var pasteMetrics               = _buttonWrap(commands.pasteMetrics,                'Paste Metrics',                    'pasteMetricsBtn');
