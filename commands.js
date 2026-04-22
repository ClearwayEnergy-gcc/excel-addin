// commands.js — Clearway Project Financial Model Add-in
//
// Exported functions are the public command API (called from taskpane.js).
// Private helpers (_prefixed) are module-scoped and unreachable from outside.

import { writeLog } from './log.js';

// ═══════════════════════════════════════════════════════════════════════════════
// COMMAND — Check for Clearway Project Financial Model
// Looks for the named range "CEG_ModelTemplateVersion" in the workbook.
// ═══════════════════════════════════════════════════════════════════════════════

export function checkModel() {
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

// ═══════════════════════════════════════════════════════════════════════════════
// COMMAND — Solve Tax Equity Upfront
// Port of TaxEquity.SolveTEUpfront().
//
// Sets calculation to automatic, initialises scenario/baseline ranges, sets
// CEG_FlipDate to CEG_TargetFlip, then iterates CEG_TEUpfront_HC ← CEG_TEUpfront_Live
// until CEG_TEUpfront_Diff = 0.
// ═══════════════════════════════════════════════════════════════════════════════

export function solveTEIUpfrontInvestment() {
  var MAX_ITER = 50;

  Excel.run(function (context) {
    var wb = context.workbook;
    wb.application.calculationMode = Excel.CalculationMode.automatic;

    var nFinancingScenario   = wb.names.getItemOrNullObject('CEG_FinancingScenario');
    var nFlipProrate         = wb.names.getItemOrNullObject('CEG_FlipProrate');
    var nFlipProrateRuledOut = wb.names.getItemOrNullObject('CEG_FlipProrateRuledOutActual');
    var nPrefPaste           = wb.names.getItemOrNullObject('CEG_PrefBaselinePaste');
    var nPrefCopy            = wb.names.getItemOrNullObject('CEG_PrefBaselineCopy');
    var nPaygoHC             = wb.names.getItemOrNullObject('CEG_PAYGOBaseline.HC');
    var nPaygoLive           = wb.names.getItemOrNullObject('CEG_PAYGOBaseline.Live');
    var nFlipDate            = wb.names.getItemOrNullObject('CEG_FlipDate');
    var nTargetFlip          = wb.names.getItemOrNullObject('CEG_TargetFlip');
    var nTEDiff              = wb.names.getItemOrNullObject('CEG_TEUpfront_Diff');
    var nTEHC                = wb.names.getItemOrNullObject('CEG_TEUpfront_HC');
    var nTELive              = wb.names.getItemOrNullObject('CEG_TEUpfront_Live');

    nFinancingScenario.load('isNullObject');
    nFlipProrate.load('isNullObject');
    nFlipProrateRuledOut.load('isNullObject');
    nPrefPaste.load('isNullObject');
    nPrefCopy.load('isNullObject');
    nPaygoHC.load('isNullObject');
    nPaygoLive.load('isNullObject');
    nFlipDate.load('isNullObject');
    nTargetFlip.load('isNullObject');
    nTEDiff.load('isNullObject');
    nTEHC.load('isNullObject');
    nTELive.load('isNullObject');

    return context.sync().then(function () {
      var missing = [];
      if (nFinancingScenario.isNullObject)   missing.push('CEG_FinancingScenario');
      if (nFlipProrate.isNullObject)         missing.push('CEG_FlipProrate');
      if (nFlipProrateRuledOut.isNullObject) missing.push('CEG_FlipProrateRuledOutActual');
      if (nPrefPaste.isNullObject)           missing.push('CEG_PrefBaselinePaste');
      if (nPrefCopy.isNullObject)            missing.push('CEG_PrefBaselineCopy');
      if (nPaygoHC.isNullObject)             missing.push('CEG_PAYGOBaseline.HC');
      if (nPaygoLive.isNullObject)           missing.push('CEG_PAYGOBaseline.Live');
      if (nFlipDate.isNullObject)            missing.push('CEG_FlipDate');
      if (nTargetFlip.isNullObject)          missing.push('CEG_TargetFlip');
      if (nTEDiff.isNullObject)              missing.push('CEG_TEUpfront_Diff');
      if (nTEHC.isNullObject)               missing.push('CEG_TEUpfront_HC');
      if (nTELive.isNullObject)             missing.push('CEG_TEUpfront_Live');

      if (missing.length > 0) {
        writeLog('Solve TE Upfront: Missing named range(s): ' + missing.join(', '), 'error');
        return;
      }

      writeLog('Solve TE Upfront: Starting…', 'info');

      // Step 1: CEG_FinancingScenario = 1
      nFinancingScenario.getRange().values = [[1]];

      // Step 2: Clear prorate hardcodes
      nFlipProrate.getRange().clear(Excel.ClearApplyTo.contents);
      nFlipProrateRuledOut.getRange().clear(Excel.ClearApplyTo.contents);

      // Step 3: Paste baseline values
      var rPrefCopy  = nPrefCopy.getRange();
      var rPaygoLive = nPaygoLive.getRange();
      rPrefCopy.load('values');
      rPaygoLive.load('values');

      return context.sync().then(function () {
        nPrefPaste.getRange().values = rPrefCopy.values;
        nPaygoHC.getRange().values   = rPaygoLive.values;

        // Step 4: CEG_FlipDate = CEG_TargetFlip
        var rTargetFlip = nTargetFlip.getRange();
        rTargetFlip.load('values');

        return context.sync().then(function () {
          nFlipDate.getRange().values = rTargetFlip.values;

          // Step 5: Iterate until CEG_TEUpfront_Diff = 0
          return context.sync().then(function () {
            return _solveTEUpfrontLoop(
              context,
              nTEDiff.getRange(),
              nTEHC.getRange(),
              nTELive.getRange(),
              0, MAX_ITER
            );
          });
        });
      });
    });
  })
  .catch(function (error) {
    writeLog('Solve TE Upfront error: ' + error.message, 'error');
  });
}

// Loop helper: copies CEG_TEUpfront_Live → CEG_TEUpfront_HC each iteration
// until CEG_TEUpfront_Diff = 0 (exact, matching VBA behaviour).
function _solveTEUpfrontLoop(context, rDiff, rHC, rLive, iter, maxIter) {
  if (iter >= maxIter) {
    writeLog('Solve TE Upfront: Did not converge after ' + maxIter + ' iterations.', 'error');
    return Promise.resolve();
  }

  rDiff.load('values');

  return context.sync().then(function () {
    var diff = rDiff.values[0][0];

    if (iter === 0 || iter % 5 === 0) {
      writeLog('Solve TE Upfront [iter ' + iter + ']: CEG_TEUpfront_Diff = ' + diff, 'info');
    }

    if (diff === 0) {
      writeLog('Solve TE Upfront: Converged in ' + iter + ' iteration(s).', 'success');
      return;
    }

    rLive.load('values');
    return context.sync().then(function () {
      rHC.values = rLive.values;
      return context.sync().then(function () {
        return _solveTEUpfrontLoop(context, rDiff, rHC, rLive, iter + 1, maxIter);
      });
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// COMMAND — Flip Date Solve
// Port of TaxEquity.FlipDate().
//
// If CEG_TEStructure = "Fixed Flip": sets CEG_FlipDate = CEG_TargetFlip and exits.
// Otherwise: clears prorate hardcodes, seeds the flip date from CEG_FlipInitialGuess,
// then walks CEG_FlipDate forward or backward via CEG_FlipGuess until
// CEG_FlipIRR_Live crosses the CEG_FlipIRR target. If CEG_FlipProrateOn is true,
// subsequently solves for CEG_FlipProrate.
// ═══════════════════════════════════════════════════════════════════════════════

export function findTEPshipFlipDate() {
  var MAX_ITER = 50;

  Excel.run(function (context) {
    var wb = context.workbook;
    wb.application.calculationMode = Excel.CalculationMode.automatic;

    var nTEStructure             = wb.names.getItemOrNullObject('CEG_TEStructure');
    var nFlipDate                = wb.names.getItemOrNullObject('CEG_FlipDate');
    var nTargetFlip              = wb.names.getItemOrNullObject('CEG_TargetFlip');
    var nFlipProrate             = wb.names.getItemOrNullObject('CEG_FlipProrate');
    var nFlipProrateRuledOut     = wb.names.getItemOrNullObject('CEG_FlipProrateRuledOutActual');
    var nLiquidation             = wb.names.getItemOrNullObject('CEG_Liquidation');
    var nFlipInitialGuess        = wb.names.getItemOrNullObject('CEG_FlipInitialGuess');
    var nFlipIRRLive             = wb.names.getItemOrNullObject('CEG_FlipIRR_Live');
    var nFlipIRRTarget           = wb.names.getItemOrNullObject('CEG_FlipIRR');
    var nFlipGuess               = wb.names.getItemOrNullObject('CEG_FlipGuess');
    var nFlipProrateOn           = wb.names.getItemOrNullObject('CEG_FlipProrateOn');
    var nFlipProrateGuess        = wb.names.getItemOrNullObject('CEG_FlipProrateGuess');
    var nFlipProrateRuledOutCalc = wb.names.getItemOrNullObject('CEG_FlipProrateRuledOutCalc');

    nTEStructure.load('isNullObject');
    nFlipDate.load('isNullObject');
    nTargetFlip.load('isNullObject');
    nFlipProrate.load('isNullObject');
    nFlipProrateRuledOut.load('isNullObject');
    nLiquidation.load('isNullObject');
    nFlipInitialGuess.load('isNullObject');
    nFlipIRRLive.load('isNullObject');
    nFlipIRRTarget.load('isNullObject');
    nFlipGuess.load('isNullObject');
    nFlipProrateOn.load('isNullObject');
    nFlipProrateGuess.load('isNullObject');
    nFlipProrateRuledOutCalc.load('isNullObject');

    return context.sync().then(function () {
      var missing = [];
      if (nTEStructure.isNullObject)             missing.push('CEG_TEStructure');
      if (nFlipDate.isNullObject)                missing.push('CEG_FlipDate');
      if (nTargetFlip.isNullObject)              missing.push('CEG_TargetFlip');
      if (nFlipProrate.isNullObject)             missing.push('CEG_FlipProrate');
      if (nFlipProrateRuledOut.isNullObject)     missing.push('CEG_FlipProrateRuledOutActual');
      if (nLiquidation.isNullObject)             missing.push('CEG_Liquidation');
      if (nFlipInitialGuess.isNullObject)        missing.push('CEG_FlipInitialGuess');
      if (nFlipIRRLive.isNullObject)             missing.push('CEG_FlipIRR_Live');
      if (nFlipIRRTarget.isNullObject)           missing.push('CEG_FlipIRR');
      if (nFlipGuess.isNullObject)               missing.push('CEG_FlipGuess');
      if (nFlipProrateOn.isNullObject)           missing.push('CEG_FlipProrateOn');
      if (nFlipProrateGuess.isNullObject)        missing.push('CEG_FlipProrateGuess');
      if (nFlipProrateRuledOutCalc.isNullObject) missing.push('CEG_FlipProrateRuledOutCalc');

      if (missing.length > 0) {
        writeLog('Flip Date: Missing named range(s): ' + missing.join(', '), 'error');
        return;
      }

      // Read CEG_TEStructure to check for Fixed Flip
      var rTEStructure = nTEStructure.getRange();
      rTEStructure.load('values');

      return context.sync().then(function () {
        var structure = rTEStructure.values[0][0];

        // ── Fixed Flip: just set the date and exit ────────────────────────────
        if (structure === 'Fixed Flip') {
          var rTargetFlip = nTargetFlip.getRange();
          rTargetFlip.load('values');
          return context.sync().then(function () {
            nFlipDate.getRange().values = rTargetFlip.values;
            return context.sync().then(function () {
              writeLog('Flip Date: Fixed Flip — flip date set to CEG_TargetFlip. No solve needed.', 'info');
            });
          });
        }

        // ── Variable Flip: solve for the flip date ────────────────────────────
        writeLog('Flip Date: Starting solve…', 'info');

        // Clear previous prorate hardcodes
        nFlipProrate.getRange().clear(Excel.ClearApplyTo.contents);
        nFlipProrateRuledOut.getRange().clear(Excel.ClearApplyTo.contents);

        // Seed: set flip date to end-of-life, then to initial guess (two syncs,
        // matching VBA's two sequential assignments to force recalc between them)
        var rLiquidation      = nLiquidation.getRange();
        var rFlipInitialGuess = nFlipInitialGuess.getRange();
        var rFlipDate         = nFlipDate.getRange();
        rLiquidation.load('values');

        return context.sync().then(function () {
          rFlipDate.values = rLiquidation.values;

          return context.sync().then(function () {
            rFlipInitialGuess.load('values');

            return context.sync().then(function () {
              rFlipDate.values = rFlipInitialGuess.values;
              writeLog('Flip Date: Seeded to initial guess.', 'info');

              return context.sync().then(function () {
                // Read live IRR and target IRR to determine walk direction
                var rFlipIRRLive   = nFlipIRRLive.getRange();
                var rFlipIRRTarget = nFlipIRRTarget.getRange();
                var rFlipGuess     = nFlipGuess.getRange();
                rFlipIRRLive.load('values');
                rFlipIRRTarget.load('values');

                return context.sync().then(function () {
                  var irrLive   = rFlipIRRLive.values[0][0];
                  var irrTarget = rFlipIRRTarget.values[0][0];

                  var walkPromise;
                  if (irrLive > irrTarget) {
                    writeLog('Flip Date: IRR exceeded at initial guess — walking back.', 'info');
                    walkPromise = _flipDateWalkBack(
                      context, rFlipDate, rFlipGuess, rFlipIRRLive, irrTarget, 0, MAX_ITER
                    );
                  } else {
                    writeLog('Flip Date: IRR not yet met at initial guess — walking forward.', 'info');
                    walkPromise = _flipDateWalkForward(
                      context, rFlipDate, rFlipGuess, rFlipIRRLive, irrTarget, 0, MAX_ITER
                    );
                  }

                  // After the walk, solve proration if enabled
                  return walkPromise.then(function () {
                    var rFlipProrateOn = nFlipProrateOn.getRange();
                    rFlipProrateOn.load('values');

                    return context.sync().then(function () {
                      if (!rFlipProrateOn.values[0][0]) {
                        return;
                      }

                      writeLog('Flip Date: Solving pro-rated last pre-flip distribution…', 'info');
                      return _flipDateProrateLoop(
                        context,
                        nFlipProrate.getRange(),
                        nFlipProrateGuess.getRange(),
                        nFlipProrateRuledOut.getRange(),
                        nFlipProrateRuledOutCalc.getRange(),
                        0, MAX_ITER
                      );
                    });
                  });
                });
              });
            });
          });
        });
      });
    });
  })
  .catch(function (error) {
    writeLog('Flip Date error: ' + error.message, 'error');
  });
}

// Walk-back loop helper: advances CEG_FlipDate ← CEG_FlipGuess until
// CEG_FlipIRR_Live <= irrTarget OR FlipGuess = FlipDate (safety stop).
// After the loop, performs one final CEG_FlipDate ← CEG_FlipGuess assignment
// (matching VBA behaviour: relies on FlipGuess formula to land on the exact date).
function _flipDateWalkBack(context, rFlipDate, rFlipGuess, rFlipIRRLive, irrTarget, iter, maxIter) {
  if (iter >= maxIter) {
    writeLog('Flip Date: Walk-back did not converge after ' + maxIter + ' iterations.', 'error');
    return Promise.resolve();
  }

  rFlipIRRLive.load('values');
  rFlipGuess.load('values');
  rFlipDate.load('values');

  return context.sync().then(function () {
    var irrLive = rFlipIRRLive.values[0][0];
    var guess   = rFlipGuess.values[0][0];
    var current = rFlipDate.values[0][0];

    if (irrLive <= irrTarget || _valuesEqual(guess, current)) {
      writeLog('Flip Date: Walk-back complete — setting final flip date.', 'info');
      rFlipDate.values = [[guess]];
      return context.sync().then(function () {
        writeLog('Flip Date: Flip date found.', 'success');
      });
    }

    rFlipDate.values = [[guess]];
    return context.sync().then(function () {
      return _flipDateWalkBack(context, rFlipDate, rFlipGuess, rFlipIRRLive, irrTarget, iter + 1, maxIter);
    });
  });
}

// Walk-forward loop helper: advances CEG_FlipDate ← CEG_FlipGuess until
// CEG_FlipIRR_Live >= irrTarget OR FlipGuess = FlipDate (safety stop).
function _flipDateWalkForward(context, rFlipDate, rFlipGuess, rFlipIRRLive, irrTarget, iter, maxIter) {
  if (iter >= maxIter) {
    writeLog('Flip Date: Walk-forward did not converge after ' + maxIter + ' iterations.', 'error');
    return Promise.resolve();
  }

  rFlipIRRLive.load('values');
  rFlipGuess.load('values');
  rFlipDate.load('values');

  return context.sync().then(function () {
    var irrLive = rFlipIRRLive.values[0][0];
    var guess   = rFlipGuess.values[0][0];
    var current = rFlipDate.values[0][0];

    if (irrLive >= irrTarget || _valuesEqual(guess, current)) {
      writeLog('Flip Date: Flip date found.', 'success');
      return;
    }

    rFlipDate.values = [[guess]];
    return context.sync().then(function () {
      return _flipDateWalkForward(context, rFlipDate, rFlipGuess, rFlipIRRLive, irrTarget, iter + 1, maxIter);
    });
  });
}

// Prorate loop helper: iterates CEG_FlipProrate ← CEG_FlipProrateGuess and
// CEG_FlipProrateRuledOutActual ← CEG_FlipProrateRuledOutCalc until
// CEG_FlipProrateGuess = CEG_FlipProrate.
function _flipDateProrateLoop(context, rProrate, rProrateGuess, rProrateRuledOutActual, rProrateRuledOutCalc, iter, maxIter) {
  if (iter >= maxIter) {
    writeLog('Flip Date: Prorate solve did not converge after ' + maxIter + ' iterations.', 'error');
    return Promise.resolve();
  }

  rProrateGuess.load('values');
  rProrate.load('values');

  return context.sync().then(function () {
    var guess   = rProrateGuess.values[0][0];
    var current = rProrate.values[0][0];

    if (_valuesEqual(guess, current)) {
      writeLog('Flip Date: Pro-rated last distribution found.', 'success');
      return;
    }

    rProrate.values = [[guess]];

    return context.sync().then(function () {
      rProrateRuledOutCalc.load('values');
      return context.sync().then(function () {
        rProrateRuledOutActual.values = rProrateRuledOutCalc.values;
        return context.sync().then(function () {
          return _flipDateProrateLoop(
            context, rProrate, rProrateGuess,
            rProrateRuledOutActual, rProrateRuledOutCalc,
            iter + 1, maxIter
          );
        });
      });
    });
  });
}

// Compares two cell values that may be date serial numbers, Date objects, or strings.
function _valuesEqual(a, b) {
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  return a === b;
}

// ═══════════════════════════════════════════════════════════════════════════════
// COMMAND — Term Debt Solve
// Port of TermDebt.TermDebtSolve().
//
// If CEG_TDActive is false: clears CEG_PrincipalHC and exits.
// Otherwise: clears the sweep mini-perm cap, pastes CFADS for each active
// scenario (optionally solving the flip date first), then calls iterateTermDebt
// to size the loan.
//
// Named-range notes (VBA used worksheet-object qualifiers):
//   CEG_TDActive, CEG_ProjectDebt, CEG_SweepActive — accessed via FinancingInputs in VBA
//   CEG_Scenario2/3/4Active, CEG_FinancingScenario  — accessed via ScenarioManager in VBA
// ═══════════════════════════════════════════════════════════════════════════════

export function solveTermDebt() {
  var flags = null; // shared across the promise chain

  // ── Step 1: validate ranges, check TDActive, read all scenario flags ───────
  return Excel.run(function (context) {
    var wb = context.workbook;
    wb.application.calculationMode = Excel.CalculationMode.automatic;

    var nTDActive         = wb.names.getItemOrNullObject('CEG_TDActive');
    var nPrincipalHC      = wb.names.getItemOrNullObject('CEG_PrincipalHC');
    var nSweepMiniPermCap = wb.names.getItemOrNullObject('CEG_SweepMiniPermCap');
    var nProjectDebt      = wb.names.getItemOrNullObject('CEG_ProjectDebt');
    var nSweepActive      = wb.names.getItemOrNullObject('CEG_SweepActive');
    var nScenario2Active  = wb.names.getItemOrNullObject('CEG_Scenario2Active');
    var nScenario3Active  = wb.names.getItemOrNullObject('CEG_Scenario3Active');
    var nScenario4Active  = wb.names.getItemOrNullObject('CEG_Scenario4Active');
    var nFinancingScenario = wb.names.getItemOrNullObject('CEG_FinancingScenario');
    var nP50CFADS         = wb.names.getItemOrNullObject('CEG_P50CFADS');
    var nP50CFADSCopy     = wb.names.getItemOrNullObject('CEG_P50CFADSCopy');
    var nP99CFADS         = wb.names.getItemOrNullObject('CEG_P99CFADS');
    var nP99CFADSCopy     = wb.names.getItemOrNullObject('CEG_P99CFADSCopy');
    var nSweepCFADS       = wb.names.getItemOrNullObject('CEG_SweepCFADS');
    var nSweepCFADSCopy   = wb.names.getItemOrNullObject('CEG_SweepCFADSCopy');
    var nPrincipalDiff    = wb.names.getItemOrNullObject('CEG_PrincipalDiff');
    var nPrincipalLive    = wb.names.getItemOrNullObject('CEG_PrincipalLive');
    var nSweepDiff        = wb.names.getItemOrNullObject('CEG_SweepMiniPermCap_Diff');
    var nSweepGuess       = wb.names.getItemOrNullObject('CEG_SweepMiniPermCap_Guess');

    [nTDActive, nPrincipalHC, nSweepMiniPermCap, nProjectDebt, nSweepActive,
     nScenario2Active, nScenario3Active, nScenario4Active, nFinancingScenario,
     nP50CFADS, nP50CFADSCopy, nP99CFADS, nP99CFADSCopy,
     nSweepCFADS, nSweepCFADSCopy, nPrincipalDiff, nPrincipalLive,
     nSweepDiff, nSweepGuess].forEach(function (n) { n.load('isNullObject'); });

    return context.sync().then(function () {
      var missing = [];
      if (nTDActive.isNullObject)          missing.push('CEG_TDActive');
      if (nPrincipalHC.isNullObject)       missing.push('CEG_PrincipalHC');
      if (nSweepMiniPermCap.isNullObject)  missing.push('CEG_SweepMiniPermCap');
      if (nProjectDebt.isNullObject)       missing.push('CEG_ProjectDebt');
      if (nSweepActive.isNullObject)       missing.push('CEG_SweepActive');
      if (nScenario2Active.isNullObject)   missing.push('CEG_Scenario2Active');
      if (nScenario3Active.isNullObject)   missing.push('CEG_Scenario3Active');
      if (nScenario4Active.isNullObject)   missing.push('CEG_Scenario4Active');
      if (nFinancingScenario.isNullObject) missing.push('CEG_FinancingScenario');
      if (nP50CFADS.isNullObject)          missing.push('CEG_P50CFADS');
      if (nP50CFADSCopy.isNullObject)      missing.push('CEG_P50CFADSCopy');
      if (nP99CFADS.isNullObject)          missing.push('CEG_P99CFADS');
      if (nP99CFADSCopy.isNullObject)      missing.push('CEG_P99CFADSCopy');
      if (nSweepCFADS.isNullObject)        missing.push('CEG_SweepCFADS');
      if (nSweepCFADSCopy.isNullObject)    missing.push('CEG_SweepCFADSCopy');
      if (nPrincipalDiff.isNullObject)     missing.push('CEG_PrincipalDiff');
      if (nPrincipalLive.isNullObject)     missing.push('CEG_PrincipalLive');
      if (nSweepDiff.isNullObject)         missing.push('CEG_SweepMiniPermCap_Diff');
      if (nSweepGuess.isNullObject)        missing.push('CEG_SweepMiniPermCap_Guess');

      if (missing.length > 0) {
        writeLog('Term Debt Solve: Missing named range(s): ' + missing.join(', '), 'error');
        return null;
      }

      // Check TDActive
      var rTDActive = nTDActive.getRange();
      rTDActive.load('values');

      return context.sync().then(function () {
        if (!rTDActive.values[0][0]) {
          nPrincipalHC.getRange().clear(Excel.ClearApplyTo.contents);
          return context.sync().then(function () {
            writeLog('Term Debt Solve: Term loan inactive (CEG_TDActive = false).', 'info');
            return null;
          });
        }

        // Clear sweep mini-perm cap
        nSweepMiniPermCap.getRange().clear(Excel.ClearApplyTo.contents);

        // Read all scenario flags
        var rProjectDebt     = nProjectDebt.getRange();
        var rSweepActive     = nSweepActive.getRange();
        var rScenario2Active = nScenario2Active.getRange();
        var rScenario3Active = nScenario3Active.getRange();
        var rScenario4Active = nScenario4Active.getRange();
        rProjectDebt.load('values');
        rSweepActive.load('values');
        rScenario2Active.load('values');
        rScenario3Active.load('values');
        rScenario4Active.load('values');

        return context.sync().then(function () {
          var projectDebt     = rProjectDebt.values[0][0];
          var sweepActive     = rSweepActive.values[0][0];
          var scenario2Active = rScenario2Active.values[0][0];
          var scenario3Active = rScenario3Active.values[0][0];
          var scenario4Active = rScenario4Active.values[0][0];

          writeLog('Term Debt Solve: Solving Term Debt — ' + projectDebt, 'info');

          // If sweep inactive, clear sweep mini-perm cap again (matches VBA)
          if (!sweepActive) {
            nSweepMiniPermCap.getRange().clear(Excel.ClearApplyTo.contents);
          }

          return context.sync().then(function () {
            return {
              projectDebt:     projectDebt,
              sweepActive:     sweepActive,
              scenario2Active: scenario2Active,
              scenario3Active: scenario3Active,
              scenario4Active: scenario4Active
            };
          });
        });
      });
    });
  })

  // ── Step 2: Scenario 2 — P50 CFADS ────────────────────────────────────────
  .then(function (f) {
    if (!f) return;
    flags = f;
    if (!flags.scenario2Active) return;

    writeLog('Term Debt Solve: Pasting P50 CFADS (scenario 2)…', 'info');
    return Excel.run(function (context) {
      context.workbook.names.getItem('CEG_FinancingScenario').getRange().values = [[2]];
      return context.sync();
    })
    .then(function () {
      if (flags.projectDebt !== 'Project') return findTEPshipFlipDate();
    })
    .then(function () {
      return Excel.run(function (context) {
        var rCopy = context.workbook.names.getItem('CEG_P50CFADSCopy').getRange();
        rCopy.load('values');
        return context.sync().then(function () {
          context.workbook.names.getItem('CEG_P50CFADS').getRange().values = rCopy.values;
          return context.sync();
        });
      });
    });
  })

  // ── Step 3: Scenario 3 — P99 CFADS ────────────────────────────────────────
  .then(function () {
    if (!flags) return;
    if (!flags.scenario3Active) return;

    writeLog('Term Debt Solve: Pasting P99 CFADS (scenario 3)…', 'info');
    return Excel.run(function (context) {
      context.workbook.names.getItem('CEG_FinancingScenario').getRange().values = [[3]];
      return context.sync();
    })
    .then(function () {
      if (flags.projectDebt !== 'Project') return findTEPshipFlipDate();
    })
    .then(function () {
      return Excel.run(function (context) {
        var rCopy = context.workbook.names.getItem('CEG_P99CFADSCopy').getRange();
        rCopy.load('values');
        return context.sync().then(function () {
          context.workbook.names.getItem('CEG_P99CFADS').getRange().values = rCopy.values;
          return context.sync();
        });
      });
    });
  })

  // ── Step 4: Scenario 4 — Sweep CFADS (only if sweep active) ───────────────
  .then(function () {
    if (!flags) return;
    if (!flags.scenario4Active || !flags.sweepActive) return;

    writeLog('Term Debt Solve: Pasting Sweep CFADS (scenario 4)…', 'info');
    return Excel.run(function (context) {
      context.workbook.names.getItem('CEG_FinancingScenario').getRange().values = [[4]];
      return context.sync();
    })
    .then(function () {
      if (flags.projectDebt !== 'Project') return findTEPshipFlipDate();
    })
    .then(function () {
      return Excel.run(function (context) {
        var rCopy = context.workbook.names.getItem('CEG_SweepCFADSCopy').getRange();
        rCopy.load('values');
        return context.sync().then(function () {
          context.workbook.names.getItem('CEG_SweepCFADS').getRange().values = rCopy.values;
          return context.sync();
        });
      });
    });
  })

  // ── Step 5: Size the debt ──────────────────────────────────────────────────
  .then(function () {
    if (!flags) return;
    writeLog('Term Debt Solve: Sizing term debt…', 'info');
    return iterateTermDebt();
  })

  .catch(function (error) {
    writeLog('Term Debt Solve error: ' + error.message, 'error');
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// COMMAND — Iterate Term Debt
// Port of TermDebt.IterateTermDebt().
//
// Reads CEG_SweepActive and dispatches to the appropriate sizing routine:
// sweep-cap + debt sizing if active, plain debt sizing otherwise.
// ═══════════════════════════════════════════════════════════════════════════════

export function iterateTermDebt() {
  return Excel.run(function (context) {
    var nSweepActive = context.workbook.names.getItemOrNullObject('CEG_SweepActive');
    nSweepActive.load('isNullObject');

    return context.sync().then(function () {
      if (nSweepActive.isNullObject) {
        writeLog('Iterate Term Debt: Missing CEG_SweepActive.', 'error');
        return null;
      }
      var rSweepActive = nSweepActive.getRange();
      rSweepActive.load('values');
      return context.sync().then(function () {
        return rSweepActive.values[0][0];
      });
    });
  })
  .then(function (sweepActive) {
    if (sweepActive === null) return;
    if (sweepActive) {
      return _sweepCapAndDebtSizing();
    } else {
      return _debtSizing();
    }
  })
  .catch(function (error) {
    writeLog('Iterate Term Debt error: ' + error.message, 'error');
  });
}

// Debt sizing helper: clears CEG_PrincipalHC (if bClear), then iterates
// CEG_PrincipalHC ← CEG_PrincipalLive until CEG_PrincipalDiff = 0.
// Port of TermDebt.DebtSizing().
function _debtSizing(bClear) {
  if (bClear === undefined) bClear = true;
  var MAX_ITER = 50;

  return Excel.run(function (context) {
    var wb             = context.workbook;
    var nPrincipalHC   = wb.names.getItemOrNullObject('CEG_PrincipalHC');
    var nPrincipalDiff = wb.names.getItemOrNullObject('CEG_PrincipalDiff');
    var nPrincipalLive = wb.names.getItemOrNullObject('CEG_PrincipalLive');

    nPrincipalHC.load('isNullObject');
    nPrincipalDiff.load('isNullObject');
    nPrincipalLive.load('isNullObject');

    return context.sync().then(function () {
      var missing = [];
      if (nPrincipalHC.isNullObject)   missing.push('CEG_PrincipalHC');
      if (nPrincipalDiff.isNullObject)  missing.push('CEG_PrincipalDiff');
      if (nPrincipalLive.isNullObject)  missing.push('CEG_PrincipalLive');
      if (missing.length > 0) {
        writeLog('Debt Sizing: Missing named range(s): ' + missing.join(', '), 'error');
        return;
      }

      var rHC   = nPrincipalHC.getRange();
      var rDiff = nPrincipalDiff.getRange();
      var rLive = nPrincipalLive.getRange();

      if (!bClear) {
        return _debtSizingLoop(context, rDiff, rHC, rLive, 0, MAX_ITER);
      }
      rHC.clear(Excel.ClearApplyTo.contents);
      return context.sync().then(function () {
        return _debtSizingLoop(context, rDiff, rHC, rLive, 0, MAX_ITER);
      });
    });
  })
  .catch(function (error) {
    writeLog('Debt Sizing error: ' + error.message, 'error');
  });
}

// Loop helper: copies CEG_PrincipalLive → CEG_PrincipalHC each iteration
// until CEG_PrincipalDiff = 0.
function _debtSizingLoop(context, rDiff, rHC, rLive, iter, maxIter) {
  if (iter >= maxIter) {
    writeLog('Debt Sizing: Did not converge after ' + maxIter + ' iterations.', 'error');
    return Promise.resolve();
  }

  rDiff.load('values');
  return context.sync().then(function () {
    var diff = rDiff.values[0][0];

    if (iter === 0 || iter % 5 === 0) {
      writeLog('Debt Sizing [iter ' + iter + ']: CEG_PrincipalDiff = ' + diff, 'info');
    }

    if (diff === 0) {
      writeLog('Debt Sizing: Converged in ' + iter + ' iteration(s).', 'success');
      return;
    }

    rLive.load('values');
    return context.sync().then(function () {
      rHC.values = rLive.values;
      return context.sync().then(function () {
        return _debtSizingLoop(context, rDiff, rHC, rLive, iter + 1, maxIter);
      });
    });
  });
}

// Sweep cap + debt sizing helper: clears CEG_PrincipalHC and CEG_SweepMiniPermCap
// (if bClear), then iterates both until CEG_PrincipalDiff = 0 AND
// CEG_SweepMiniPermCap_Diff = 0. Port of TermDebt.SweepCapAndDebtSizing().
function _sweepCapAndDebtSizing(bClear) {
  if (bClear === undefined) bClear = true;
  var MAX_ITER = 50;

  return Excel.run(function (context) {
    var wb            = context.workbook;
    var nPrincipalHC  = wb.names.getItemOrNullObject('CEG_PrincipalHC');
    var nPrincipalDiff = wb.names.getItemOrNullObject('CEG_PrincipalDiff');
    var nPrincipalLive = wb.names.getItemOrNullObject('CEG_PrincipalLive');
    var nSweepHC      = wb.names.getItemOrNullObject('CEG_SweepMiniPermCap');
    var nSweepDiff    = wb.names.getItemOrNullObject('CEG_SweepMiniPermCap_Diff');
    var nSweepGuess   = wb.names.getItemOrNullObject('CEG_SweepMiniPermCap_Guess');

    [nPrincipalHC, nPrincipalDiff, nPrincipalLive,
     nSweepHC, nSweepDiff, nSweepGuess].forEach(function (n) { n.load('isNullObject'); });

    return context.sync().then(function () {
      var missing = [];
      if (nPrincipalHC.isNullObject)   missing.push('CEG_PrincipalHC');
      if (nPrincipalDiff.isNullObject)  missing.push('CEG_PrincipalDiff');
      if (nPrincipalLive.isNullObject)  missing.push('CEG_PrincipalLive');
      if (nSweepHC.isNullObject)        missing.push('CEG_SweepMiniPermCap');
      if (nSweepDiff.isNullObject)      missing.push('CEG_SweepMiniPermCap_Diff');
      if (nSweepGuess.isNullObject)     missing.push('CEG_SweepMiniPermCap_Guess');
      if (missing.length > 0) {
        writeLog('Sweep Cap & Debt Sizing: Missing named range(s): ' + missing.join(', '), 'error');
        return;
      }

      var rPrincipalHC   = nPrincipalHC.getRange();
      var rPrincipalDiff = nPrincipalDiff.getRange();
      var rPrincipalLive = nPrincipalLive.getRange();
      var rSweepHC       = nSweepHC.getRange();
      var rSweepDiff     = nSweepDiff.getRange();
      var rSweepGuess    = nSweepGuess.getRange();

      if (!bClear) {
        return _sweepCapLoop(context,
          rPrincipalDiff, rPrincipalHC, rPrincipalLive,
          rSweepDiff, rSweepHC, rSweepGuess, 0, MAX_ITER);
      }
      rPrincipalHC.clear(Excel.ClearApplyTo.contents);
      rSweepHC.clear(Excel.ClearApplyTo.contents);
      return context.sync().then(function () {
        return _sweepCapLoop(context,
          rPrincipalDiff, rPrincipalHC, rPrincipalLive,
          rSweepDiff, rSweepHC, rSweepGuess, 0, MAX_ITER);
      });
    });
  })
  .catch(function (error) {
    writeLog('Sweep Cap & Debt Sizing error: ' + error.message, 'error');
  });
}

// Loop helper: iterates CEG_PrincipalHC ← CEG_PrincipalLive and
// CEG_SweepMiniPermCap ← CEG_SweepMiniPermCap_Guess each iteration until
// BOTH CEG_PrincipalDiff = 0 AND CEG_SweepMiniPermCap_Diff = 0.
function _sweepCapLoop(context, rPrincipalDiff, rPrincipalHC, rPrincipalLive, rSweepDiff, rSweepHC, rSweepGuess, iter, maxIter) {
  if (iter >= maxIter) {
    writeLog('Sweep Cap & Debt Sizing: Did not converge after ' + maxIter + ' iterations.', 'error');
    return Promise.resolve();
  }

  rPrincipalDiff.load('values');
  rSweepDiff.load('values');

  return context.sync().then(function () {
    var principalDiff = rPrincipalDiff.values[0][0];
    var sweepDiff     = rSweepDiff.values[0][0];

    if (iter === 0 || iter % 5 === 0) {
      writeLog('Sweep Cap & Debt Sizing [iter ' + iter + ']: PrincipalDiff = ' + principalDiff + ', SweepDiff = ' + sweepDiff, 'info');
    }

    // Both must be 0 to converge (matches VBA: "... = 0 And ... = 0")
    if (principalDiff === 0 && sweepDiff === 0) {
      writeLog('Sweep Cap & Debt Sizing: Converged in ' + iter + ' iteration(s).', 'success');
      return;
    }

    rPrincipalLive.load('values');
    rSweepGuess.load('values');
    return context.sync().then(function () {
      rPrincipalHC.values = rPrincipalLive.values;
      rSweepHC.values     = rSweepGuess.values;
      return context.sync().then(function () {
        return _sweepCapLoop(context,
          rPrincipalDiff, rPrincipalHC, rPrincipalLive,
          rSweepDiff, rSweepHC, rSweepGuess,
          iter + 1, maxIter);
      });
    });
  });
}
