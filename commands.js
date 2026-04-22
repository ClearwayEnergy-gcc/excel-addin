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
  var MAX_ITER = 1000;

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

    if (iter === 0 || iter % 50 === 0) {
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
  var MAX_ITER = 10000;

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
