// taskpane.js — Clearway Project Financial Model Add-in
//
// Entry point. Wires up Office.onReady and button event listeners.
// All commands run directly from the task pane — the ribbon button only
// opens the task pane; there are no ExecuteFunction ribbon buttons.

import { clearLog, showEmptyState } from './log.js';
import { checkModel, solveTEIUpfrontInvestment, findTEPshipFlipDate, solveTermDebt, iterateTermDebt, solveCWENUpfrontInvestment, solveCE2UpfrontInvestment, solveCapExCircularity } from './commands.js';

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById('clearLogBtn').addEventListener('click', clearLog);
    document.getElementById('checkModelBtn').addEventListener('click', checkModel);
    document.getElementById('solveTEUpfrontBtn').addEventListener('click', solveTEIUpfrontInvestment);
    document.getElementById('flipDateBtn').addEventListener('click', findTEPshipFlipDate);
    document.getElementById('termDebtSolveBtn').addEventListener('click', solveTermDebt);
    document.getElementById('iterateTermDebtBtn').addEventListener('click', iterateTermDebt);
    document.getElementById('solveCWENUpfrontInvestmentBtn').addEventListener('click', solveCWENUpfrontInvestment);
    document.getElementById('solveCE2UpfrontInvestmentBtn').addEventListener('click', solveCE2UpfrontInvestment);
    document.getElementById('solveCapExCircularityBtn').addEventListener('click', solveCapExCircularity);
    showEmptyState();
  }
});
