// taskpane.js — Clearway Project Financial Model Add-in
//
// Entry point. Wires up Office.onReady and button event listeners.
// All commands run directly from the task pane — the ribbon button only
// opens the task pane; there are no ExecuteFunction ribbon buttons.

import { clearLog, showEmptyState } from './log.js';
import { checkModel, solveTEIUpfrontInvestment, findTEPshipFlipDate, solveTermDebt, iterateTermDebt, solveCWEN, solveCE2 } from './commands.js';

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById('clearLogBtn').addEventListener('click', clearLog);
    document.getElementById('checkModelBtn').addEventListener('click', checkModel);
    document.getElementById('solveTEUpfrontBtn').addEventListener('click', solveTEIUpfrontInvestment);
    document.getElementById('flipDateBtn').addEventListener('click', findTEPshipFlipDate);
    document.getElementById('termDebtSolveBtn').addEventListener('click', solveTermDebt);
    document.getElementById('iterateTermDebtBtn').addEventListener('click', iterateTermDebt);
    document.getElementById('solveCWENBtn').addEventListener('click', solveCWEN);
    document.getElementById('solveCE2Btn').addEventListener('click', solveCE2);
    showEmptyState();
  }
});
