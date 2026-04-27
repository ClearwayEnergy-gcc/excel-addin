// commands.js — Clearway Project Financial Model Add-in
//
// Pure-work command functions shared between the Office.js add-in and the
// Claude on Excel skill package.
//
// Each exported function calls _checkModel() first, which silently populates
// the workbook variant config on the first call and is a no-op thereafter.
// The exported checkModel() delegates to _checkModel() and additionally logs
// "already run" when called a second time explicitly.
//
// Errors are NOT caught here — callers handle them:
//   buttons.js  (add-in task-pane buttons)
//   skills.js   (Claude on Excel skill package)
//
// Private helpers (_prefixed) are module-scoped and unreachable from outside.

import { writeLog }            from './log.js';
import { getConfig, setConfig } from './config.js';

// ═══════════════════════════════════════════════════════════════════════════════
// COMMAND — Check for Clearway Project Financial Model
// Looks for the named range "CEG_ModelTemplateVersion" in the workbook and
// populates the variant config. Idempotent: logs "already run" and returns
// immediately if config is already populated.
// ═══════════════════════════════════════════════════════════════════════════════

export function checkModel() {
  if (getConfig().populated) {
    writeLog('Check Model: already run — skipping.', 'info');
    return Promise.resolve();
  }
  return _checkModel();
}

// Private — contains the actual workbook detection work.
// Called by the exported checkModel() and at the start of every other command.
// Silent no-op if config is already populated; otherwise reads named ranges,
// calls setConfig(), and logs findings.
function _checkModel() {
  if (getConfig().populated) return Promise.resolve();

  var t0 = Date.now();
  return Excel.run(function (context) {
    var item = context.workbook.names.getItemOrNullObject('CEG_ModelTemplateVersion');
    item.load('isNullObject,value');

    // ── Add reads for variant-detection named ranges here as they are ────────
    // ── identified. Load them alongside CEG_ModelTemplateVersion, then ───────
    // ── include their values in the setConfig({ ... }) call below. ───────────

    return context.sync().then(function () {
      if (item.isNullObject) {
        writeLog('This workbook does not appear to be a Clearway Project Financial Model (CEG_ModelTemplateVersion not found).', 'error');
        setConfig({ modelVersion: null });
      } else {
        writeLog('Clearway Project Financial Model detected. CEG_ModelTemplateVersion = ' + item.value, 'success');
        setConfig({
          modelVersion: item.value
          // ── Add variant flags here as they are identified, e.g.: ──────────
          // someFlag: nSomeRange.isNullObject ? false : nSomeRange.getRange().values[0][0],
        });
      }
    });
  })
  .then(function () {
    writeLog('Check Model: completed in ' + ((Date.now() - t0) / 1000).toFixed(2) + 's.', 'info');
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
  var t0 = Date.now();
  var MAX_ITER = 50;
  return _checkModel().then(function () {
    return Excel.run(function (context) {
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
          writeLog('Solve TEI Upfront Investment: Missing named range(s): ' + missing.join(', '), 'error');
          return;
        }

        writeLog('Solve TEI Upfront Investment: Starting…', 'info');

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
              return _solveTEIUpfrontInvestmentLoop(
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
    });
  })
  .then(function () {
    writeLog('Solve TEI Upfront Investment: completed in ' + ((Date.now() - t0) / 1000).toFixed(2) + 's.', 'info');
  });
}

// Loop helper: copies CEG_TEUpfront_Live → CEG_TEUpfront_HC each iteration
// until CEG_TEUpfront_Diff = 0 (exact, matching VBA behaviour).
function _solveTEIUpfrontInvestmentLoop(context, rDiff, rHC, rLive, iter, maxIter) {
  if (iter >= maxIter) {
    writeLog('Solve TEI Upfront Investment: Did not converge after ' + maxIter + ' iterations.', 'error');
    return Promise.resolve();
  }

  rDiff.load('values');

  return context.sync().then(function () {
    var diff = rDiff.values[0][0];

    if (iter === 0 || iter % 5 === 0) {
      writeLog('Solve TEI Upfront Investment [iter ' + iter + ']: CEG_TEUpfront_Diff = ' + diff, 'info');
    }

    if (diff === 0) {
      writeLog('Solve TEI Upfront Investment: Converged in ' + iter + ' iteration(s).', 'success');
      return;
    }

    rLive.load('values');
    return context.sync().then(function () {
      rHC.values = rLive.values;
      return context.sync().then(function () {
        return _solveTEIUpfrontInvestmentLoop(context, rDiff, rHC, rLive, iter + 1, maxIter);
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
//
// Called directly by solveTermDebt, solveCWENUpfrontInvestment, and
// solveCE2UpfrontInvestment so that errors propagate to those callers.
// ═══════════════════════════════════════════════════════════════════════════════

export function findTEPshipFlipDate() {
  var t0 = Date.now();
  var MAX_ITER = 50;

  return _checkModel().then(function () {
    return Excel.run(function (context) {
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

          // ── Fixed Flip: just set the date and exit ──────────────────────────
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

          // ── Variable Flip: solve for the flip date ──────────────────────────
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
    });
  })
  .then(function () {
    writeLog('Find TE Partnership Flip Date: completed in ' + ((Date.now() - t0) / 1000).toFixed(2) + 's.', 'info');
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
// ═══════════════════════════════════════════════════════════════════════════════

export function solveTermDebt() {
  var t0 = Date.now();
  var flags = null; // shared across the promise chain

  return _checkModel().then(function () {

    // ── Step 1: validate ranges, check TDActive, read all scenario flags ─────
    return Excel.run(function (context) {
      var wb = context.workbook;
      wb.application.calculationMode = Excel.CalculationMode.automatic;

      var nTDActive          = wb.names.getItemOrNullObject('CEG_TDActive');
      var nPrincipalHC       = wb.names.getItemOrNullObject('CEG_PrincipalHC');
      var nSweepMiniPermCap  = wb.names.getItemOrNullObject('CEG_SweepMiniPermCap');
      var nProjectDebt       = wb.names.getItemOrNullObject('CEG_ProjectDebt');
      var nSweepActive       = wb.names.getItemOrNullObject('CEG_SweepActive');
      var nScenario2Active   = wb.names.getItemOrNullObject('CEG_Scenario2Active');
      var nScenario3Active   = wb.names.getItemOrNullObject('CEG_Scenario3Active');
      var nScenario4Active   = wb.names.getItemOrNullObject('CEG_Scenario4Active');
      var nFinancingScenario = wb.names.getItemOrNullObject('CEG_FinancingScenario');
      var nP50CFADS          = wb.names.getItemOrNullObject('CEG_P50CFADS');
      var nP50CFADSCopy      = wb.names.getItemOrNullObject('CEG_P50CFADSCopy');
      var nP99CFADS          = wb.names.getItemOrNullObject('CEG_P99CFADS');
      var nP99CFADSCopy      = wb.names.getItemOrNullObject('CEG_P99CFADSCopy');
      var nSweepCFADS        = wb.names.getItemOrNullObject('CEG_SweepCFADS');
      var nSweepCFADSCopy    = wb.names.getItemOrNullObject('CEG_SweepCFADSCopy');
      var nPrincipalDiff     = wb.names.getItemOrNullObject('CEG_PrincipalDiff');
      var nPrincipalLive     = wb.names.getItemOrNullObject('CEG_PrincipalLive');
      var nSweepDiff         = wb.names.getItemOrNullObject('CEG_SweepMiniPermCap_Diff');
      var nSweepGuess        = wb.names.getItemOrNullObject('CEG_SweepMiniPermCap_Guess');

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

    // ── Step 2: Scenario 2 — P50 CFADS ──────────────────────────────────────
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

    // ── Step 3: Scenario 3 — P99 CFADS ──────────────────────────────────────
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

    // ── Step 4: Scenario 4 — Sweep CFADS (only if sweep active) ─────────────
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

    // ── Step 5: Size the debt ────────────────────────────────────────────────
    .then(function () {
      if (!flags) return;
      writeLog('Term Debt Solve: Sizing term debt…', 'info');
      return iterateTermDebt();
    });

  })
  .then(function () {
    writeLog('Solve Term Debt: completed in ' + ((Date.now() - t0) / 1000).toFixed(2) + 's.', 'info');
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
  var t0 = Date.now();
  return _checkModel().then(function () {
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
    });
  })
  .then(function () {
    writeLog('Iterate Term Debt: completed in ' + ((Date.now() - t0) / 1000).toFixed(2) + 's.', 'info');
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
    var wb             = context.workbook;
    var nPrincipalHC   = wb.names.getItemOrNullObject('CEG_PrincipalHC');
    var nPrincipalDiff = wb.names.getItemOrNullObject('CEG_PrincipalDiff');
    var nPrincipalLive = wb.names.getItemOrNullObject('CEG_PrincipalLive');
    var nSweepHC       = wb.names.getItemOrNullObject('CEG_SweepMiniPermCap');
    var nSweepDiff     = wb.names.getItemOrNullObject('CEG_SweepMiniPermCap_Diff');
    var nSweepGuess    = wb.names.getItemOrNullObject('CEG_SweepMiniPermCap_Guess');

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

// ═══════════════════════════════════════════════════════════════════════════════
// COMMAND — Solve CWEN Investment
// Port of CashEquity.SolveCWEN().
//
// Sets CEG_FinancingScenario = 5, solves the flip date, then iterates:
//   CEG_MerchantNPVPaste  ← CEG_MerchantNPVCopy
//   CEG_MerchantCAFDPaste ← CEG_MerchantCAFDCopy
//   CEG_CWENPP_HC         ← CEG_CWENPP_Live
// until CEG_CWENPP_Diff = 0.
// ═══════════════════════════════════════════════════════════════════════════════

export function solveCWENUpfrontInvestment() {
  var t0 = Date.now();
  var MAX_ITER = 50;
  var ok = false;

  return _checkModel().then(function () {

    // Step 1: validate ranges, set calculation and scenario
    return Excel.run(function (context) {
      var wb = context.workbook;
      wb.application.calculationMode = Excel.CalculationMode.automatic;

      var nFinancingScenario = wb.names.getItemOrNullObject('CEG_FinancingScenario');
      var nNPVPaste          = wb.names.getItemOrNullObject('CEG_MerchantNPVPaste');
      var nNPVCopy           = wb.names.getItemOrNullObject('CEG_MerchantNPVCopy');
      var nCAFDPaste         = wb.names.getItemOrNullObject('CEG_MerchantCAFDPaste');
      var nCAFDCopy          = wb.names.getItemOrNullObject('CEG_MerchantCAFDCopy');
      var nCWENPPHC          = wb.names.getItemOrNullObject('CEG_CWENPP_HC');
      var nCWENPPLive        = wb.names.getItemOrNullObject('CEG_CWENPP_Live');
      var nCWENPPDiff        = wb.names.getItemOrNullObject('CEG_CWENPP_Diff');

      [nFinancingScenario, nNPVPaste, nNPVCopy, nCAFDPaste, nCAFDCopy,
       nCWENPPHC, nCWENPPLive, nCWENPPDiff].forEach(function (n) { n.load('isNullObject'); });

      return context.sync().then(function () {
        var missing = [];
        if (nFinancingScenario.isNullObject) missing.push('CEG_FinancingScenario');
        if (nNPVPaste.isNullObject)          missing.push('CEG_MerchantNPVPaste');
        if (nNPVCopy.isNullObject)           missing.push('CEG_MerchantNPVCopy');
        if (nCAFDPaste.isNullObject)         missing.push('CEG_MerchantCAFDPaste');
        if (nCAFDCopy.isNullObject)          missing.push('CEG_MerchantCAFDCopy');
        if (nCWENPPHC.isNullObject)          missing.push('CEG_CWENPP_HC');
        if (nCWENPPLive.isNullObject)        missing.push('CEG_CWENPP_Live');
        if (nCWENPPDiff.isNullObject)        missing.push('CEG_CWENPP_Diff');

        if (missing.length > 0) {
          writeLog('Solve CWEN: Missing named range(s): ' + missing.join(', '), 'error');
          return;
        }

        writeLog('Solve CWEN: Solving CWEN Investment…', 'info');
        nFinancingScenario.getRange().values = [[5]];
        return context.sync().then(function () { ok = true; });
      });
    })

    // Step 2: solve flip date
    .then(function () {
      if (!ok) return;
      return findTEPshipFlipDate();
    })

    // Step 3: convergence loop
    .then(function () {
      if (!ok) return;
      return Excel.run(function (context) {
        var wb       = context.workbook;
        var rDiff     = wb.names.getItem('CEG_CWENPP_Diff').getRange();
        var rNPVPaste = wb.names.getItem('CEG_MerchantNPVPaste').getRange();
        var rNPVCopy  = wb.names.getItem('CEG_MerchantNPVCopy').getRange();
        var rCAFDPaste = wb.names.getItem('CEG_MerchantCAFDPaste').getRange();
        var rCAFDCopy  = wb.names.getItem('CEG_MerchantCAFDCopy').getRange();
        var rHC       = wb.names.getItem('CEG_CWENPP_HC').getRange();
        var rLive     = wb.names.getItem('CEG_CWENPP_Live').getRange();

        return _solveCWENUpfrontInvestmentLoop(context,
          rDiff, rNPVPaste, rNPVCopy, rCAFDPaste, rCAFDCopy, rHC, rLive, 0, MAX_ITER);
      });
    });

  })
  .then(function () {
    writeLog('Solve CWEN Investment: completed in ' + ((Date.now() - t0) / 1000).toFixed(2) + 's.', 'info');
  });
}

// Loop helper: each iteration copies NPV, CAFD, and CWENPP Live → HC/Paste
// until CEG_CWENPP_Diff = 0.
function _solveCWENUpfrontInvestmentLoop(context, rDiff, rNPVPaste, rNPVCopy, rCAFDPaste, rCAFDCopy, rHC, rLive, iter, maxIter) {
  if (iter >= maxIter) {
    writeLog('Solve CWEN: Did not converge after ' + maxIter + ' iterations.', 'error');
    return Promise.resolve();
  }

  rDiff.load('values');
  return context.sync().then(function () {
    var diff = rDiff.values[0][0];

    if (iter === 0 || iter % 5 === 0) {
      writeLog('Solve CWEN [iter ' + iter + ']: CEG_CWENPP_Diff = ' + diff, 'info');
    }

    if (diff === 0) {
      writeLog('Solve CWEN: Converged in ' + iter + ' iteration(s).', 'success');
      return;
    }

    rNPVCopy.load('values');
    rCAFDCopy.load('values');
    rLive.load('values');
    return context.sync().then(function () {
      rNPVPaste.values  = rNPVCopy.values;
      rCAFDPaste.values = rCAFDCopy.values;
      rHC.values        = rLive.values;
      return context.sync().then(function () {
        return _solveCWENUpfrontInvestmentLoop(context,
          rDiff, rNPVPaste, rNPVCopy, rCAFDPaste, rCAFDCopy, rHC, rLive,
          iter + 1, maxIter);
      });
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// COMMAND — Solve Third-Party CE Investment
// Port of CashEquity.SolveCE2().
//
// Sets CEG_FinancingScenario = 6, solves the flip date, then iterates
// CEG_CE2PP_HC ← CEG_CE2PP_Live until CEG_CE2PP_Diff = 0.
// ═══════════════════════════════════════════════════════════════════════════════

export function solveCE2UpfrontInvestment() {
  var t0 = Date.now();
  var MAX_ITER = 50;
  var ok = false;

  return _checkModel().then(function () {

    // Step 1: validate ranges, set calculation and scenario
    return Excel.run(function (context) {
      var wb = context.workbook;
      wb.application.calculationMode = Excel.CalculationMode.automatic;

      var nFinancingScenario = wb.names.getItemOrNullObject('CEG_FinancingScenario');
      var nCE2PPHC           = wb.names.getItemOrNullObject('CEG_CE2PP_HC');
      var nCE2PPLive         = wb.names.getItemOrNullObject('CEG_CE2PP_Live');
      var nCE2PPDiff         = wb.names.getItemOrNullObject('CEG_CE2PP_Diff');

      [nFinancingScenario, nCE2PPHC, nCE2PPLive, nCE2PPDiff]
        .forEach(function (n) { n.load('isNullObject'); });

      return context.sync().then(function () {
        var missing = [];
        if (nFinancingScenario.isNullObject) missing.push('CEG_FinancingScenario');
        if (nCE2PPHC.isNullObject)           missing.push('CEG_CE2PP_HC');
        if (nCE2PPLive.isNullObject)         missing.push('CEG_CE2PP_Live');
        if (nCE2PPDiff.isNullObject)         missing.push('CEG_CE2PP_Diff');

        if (missing.length > 0) {
          writeLog('Solve CE2: Missing named range(s): ' + missing.join(', '), 'error');
          return;
        }

        writeLog('Solve CE2: Solving Third-Party CE Investment…', 'info');
        nFinancingScenario.getRange().values = [[6]];
        return context.sync().then(function () { ok = true; });
      });
    })

    // Step 2: solve flip date
    .then(function () {
      if (!ok) return;
      return findTEPshipFlipDate();
    })

    // Step 3: convergence loop
    .then(function () {
      if (!ok) return;
      return Excel.run(function (context) {
        var wb    = context.workbook;
        var rDiff = wb.names.getItem('CEG_CE2PP_Diff').getRange();
        var rHC   = wb.names.getItem('CEG_CE2PP_HC').getRange();
        var rLive = wb.names.getItem('CEG_CE2PP_Live').getRange();

        return _solveCE2UpfrontInvestmentLoop(context, rDiff, rHC, rLive, 0, MAX_ITER);
      });
    });

  })
  .then(function () {
    writeLog('Solve Third-Party CE Investment: completed in ' + ((Date.now() - t0) / 1000).toFixed(2) + 's.', 'info');
  });
}

// Loop helper: copies CEG_CE2PP_Live → CEG_CE2PP_HC each iteration
// until CEG_CE2PP_Diff = 0.
function _solveCE2UpfrontInvestmentLoop(context, rDiff, rHC, rLive, iter, maxIter) {
  if (iter >= maxIter) {
    writeLog('Solve CE2: Did not converge after ' + maxIter + ' iterations.', 'error');
    return Promise.resolve();
  }

  rDiff.load('values');
  return context.sync().then(function () {
    var diff = rDiff.values[0][0];

    if (iter === 0 || iter % 5 === 0) {
      writeLog('Solve CE2 [iter ' + iter + ']: CEG_CE2PP_Diff = ' + diff, 'info');
    }

    if (diff === 0) {
      writeLog('Solve CE2: Converged in ' + iter + ' iteration(s).', 'success');
      return;
    }

    rLive.load('values');
    return context.sync().then(function () {
      rHC.values = rLive.values;
      return context.sync().then(function () {
        return _solveCE2UpfrontInvestmentLoop(context, rDiff, rHC, rLive, iter + 1, maxIter);
      });
    });
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

// ═══════════════════════════════════════════════════════════════════════════════
// COMMAND — Solve CapEx CF Circularity
// Port of ProjectCosts.CapExCircularitySolve().
//
// Switches to manual calculation, then iterates:
//   CEG_CF_TakeOut_Paste              ← CEG_CF_TakeOut_Copy
//   CEG_TCInsuranceHCValues           ← CEG_TCInsuranceLiveValues
//   CEG_CapEx.ConstructionCosts.Paste ← CEG_CapEx.ConstructionCosts.Copy
// followed by a full manual recalculate each iteration, until both
// CEG_Capex.ConstructionCostDelta and CEG_CF_TakeOut_Diff are <=
// CEG_General.CircularityDeltaPrecision. Restores automatic calculation on exit.
// ═══════════════════════════════════════════════════════════════════════════════

export function solveCapexCFCircularity() {
  var t0 = Date.now();
  var MAX_ITER = 50;
  return _checkModel().then(function () {
    return Excel.run(function (context) {
      var wb = context.workbook;

      var nCostDelta         = wb.names.getItemOrNullObject('CEG_Capex.ConstructionCostDelta');
      var nTakeOutDiff       = wb.names.getItemOrNullObject('CEG_CF_TakeOut_Diff');
      var nTakeOutPaste      = wb.names.getItemOrNullObject('CEG_CF_TakeOut_Paste');
      var nTakeOutCopy       = wb.names.getItemOrNullObject('CEG_CF_TakeOut_Copy');
      var nTCInsuranceHC     = wb.names.getItemOrNullObject('CEG_TCInsuranceHCValues');
      var nTCInsuranceLive   = wb.names.getItemOrNullObject('CEG_TCInsuranceLiveValues');
      var nConstructionPaste = wb.names.getItemOrNullObject('CEG_CapEx.ConstructionCosts.Paste');
      var nConstructionCopy  = wb.names.getItemOrNullObject('CEG_CapEx.ConstructionCosts.Copy');

      [nCostDelta, nTakeOutDiff, nTakeOutPaste, nTakeOutCopy,
       nTCInsuranceHC, nTCInsuranceLive, nConstructionPaste, nConstructionCopy]
        .forEach(function (n) { n.load('isNullObject'); });

      return context.sync().then(function () {
        var missing = [];
        if (nCostDelta.isNullObject)        missing.push('CEG_Capex.ConstructionCostDelta');
        if (nTakeOutDiff.isNullObject)       missing.push('CEG_CF_TakeOut_Diff');
        if (nTakeOutPaste.isNullObject)      missing.push('CEG_CF_TakeOut_Paste');
        if (nTakeOutCopy.isNullObject)       missing.push('CEG_CF_TakeOut_Copy');
        if (nTCInsuranceHC.isNullObject)     missing.push('CEG_TCInsuranceHCValues');
        if (nTCInsuranceLive.isNullObject)   missing.push('CEG_TCInsuranceLiveValues');
        if (nConstructionPaste.isNullObject) missing.push('CEG_CapEx.ConstructionCosts.Paste');
        if (nConstructionCopy.isNullObject)  missing.push('CEG_CapEx.ConstructionCosts.Copy');

        if (missing.length > 0) {
          writeLog('Solve CapEx CF Circularity: Missing named range(s): ' + missing.join(', '), 'error');
          return;
        }

        writeLog('Solve CapEx CF Circularity: Solving Construction Financing/CapEx…', 'info');

        return _solveCapexCFCircularityLoop(
          context,
          nCostDelta.getRange(),
          nTakeOutDiff.getRange(),
          nTakeOutPaste.getRange(),
          nTakeOutCopy.getRange(),
          nTCInsuranceHC.getRange(),
          nTCInsuranceLive.getRange(),
          nConstructionPaste.getRange(),
          nConstructionCopy.getRange(),
          0, MAX_ITER
        );
      });
    });
  })
  .then(function () {
    writeLog('Solve CapEx CF Circularity: completed in ' + ((Date.now() - t0) / 1000).toFixed(2) + 's.', 'info');
  });
}

// Loop helper: each iteration pastes TakeOut, TCInsurance, and ConstructionCosts,
// then triggers a full manual recalculate, until both delta values === 0.
function _solveCapexCFCircularityLoop(context, rCostDelta, rTakeOutDiff,
    rTakeOutPaste, rTakeOutCopy, rTCInsuranceHC, rTCInsuranceLive,
    rConstructionPaste, rConstructionCopy, iter, maxIter) {

  if (iter >= maxIter) {
    writeLog('Solve CapEx CF Circularity: Did not converge after ' + maxIter + ' iterations.', 'error');
    return Promise.resolve();
  }

  rCostDelta.load('values');
  rTakeOutDiff.load('values');

  return context.sync().then(function () {
    var costDelta   = rCostDelta.values[0][0];
    var takeOutDiff = rTakeOutDiff.values[0][0];

    if (iter === 0 || iter % 5 === 0) {
      writeLog('Solve CapEx CF Circularity [iter ' + iter + ']: ConstructionCostDelta = ' + costDelta + ', TakeOutDiff = ' + takeOutDiff, 'info');
    }

    if (costDelta === 0 && takeOutDiff === 0) {
      writeLog('Solve CapEx CF Circularity: Converged in ' + iter + ' iteration(s).', 'success');
      return;
    }

    rTakeOutCopy.load('values');
    rTCInsuranceLive.load('values');
    rConstructionCopy.load('values');

    return context.sync().then(function () {
      rTakeOutPaste.values      = rTakeOutCopy.values;
      rTCInsuranceHC.values     = rTCInsuranceLive.values;
      rConstructionPaste.values = rConstructionCopy.values;

      return context.sync().then(function () {
        return _solveCapexCFCircularityLoop(
          context,
          rCostDelta, rTakeOutDiff,
          rTakeOutPaste, rTakeOutCopy,
          rTCInsuranceHC, rTCInsuranceLive,
          rConstructionPaste, rConstructionCopy,
          iter + 1, maxIter
        );
      });
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// COMMAND — Paste Metrics
// Port of ScenarioManager.PasteMetrics().
//
// Pastes CEG_KeyMetrics into the column at the offset specified by
// CEG_MetricsPasteColumn relative to the CEG_KeyMetrics range itself.
// ═══════════════════════════════════════════════════════════════════════════════

export function pasteMetrics() {
  var t0 = Date.now();
  return _checkModel().then(function () {
    return Excel.run(function (context) {
      var wb = context.workbook;
      wb.application.calculationMode = Excel.CalculationMode.automatic;

      var nKeyMetrics  = wb.names.getItemOrNullObject('CEG_KeyMetrics');
      var nPasteColumn = wb.names.getItemOrNullObject('CEG_MetricsPasteColumn');

      nKeyMetrics.load('isNullObject');
      nPasteColumn.load('isNullObject');

      return context.sync().then(function () {
        var missing = [];
        if (nKeyMetrics.isNullObject)  missing.push('CEG_KeyMetrics');
        if (nPasteColumn.isNullObject) missing.push('CEG_MetricsPasteColumn');
        if (missing.length > 0) {
          writeLog('Paste Metrics: Missing named range(s): ' + missing.join(', '), 'error');
          return;
        }

        var rPasteColumn = nPasteColumn.getRange();
        var rKeyMetrics  = nKeyMetrics.getRange();
        rPasteColumn.load('values');
        rKeyMetrics.load('values');

        return context.sync().then(function () {
          var colOffset = rPasteColumn.values[0][0];
          var rDest     = rKeyMetrics.getOffsetRange(0, colOffset);
          rDest.values  = rKeyMetrics.values;

          return context.sync().then(function () {
            writeLog('Paste Metrics: Key metrics pasted at column offset ' + colOffset + '.', 'success');
          });
        });
      });
    });
  })
  .then(function () {
    writeLog('Paste Metrics: completed in ' + ((Date.now() - t0) / 1000).toFixed(2) + 's.', 'info');
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// FMV Income Approach loop helper
// Used by solveCapitalStack when CEG_Scenario1Active = true and
// CEG_FMVOverride = "".
// Iterates CEG_IncomeApproachHC ← CEG_IncomeApproachLive until
// CEG_IncomeApproachDiff = 0.
// ═══════════════════════════════════════════════════════════════════════════════

function _fmvIncomeApproachLoop(context, rDiff, rHC, rLive, iter, maxIter) {
  if (iter >= maxIter) {
    writeLog('FMV Income Approach: Did not converge after ' + maxIter + ' iterations.', 'error');
    return Promise.resolve();
  }

  rDiff.load('values');
  return context.sync().then(function () {
    var diff = rDiff.values[0][0];

    if (iter === 0 || iter % 5 === 0) {
      writeLog('FMV Income Approach [iter ' + iter + ']: CEG_IncomeApproachDiff = ' + diff, 'info');
    }

    if (diff === 0) {
      writeLog('FMV Income Approach: Converged in ' + iter + ' iteration(s).', 'success');
      return;
    }

    rLive.load('values');
    return context.sync().then(function () {
      rHC.values = rLive.values;
      return context.sync().then(function () {
        return _fmvIncomeApproachLoop(context, rDiff, rHC, rLive, iter + 1, maxIter);
      });
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// COMMAND — Solve Capital Stack
// Port of ScenarioManager.SolveCapitalStack().
//
// Orchestrates the full capital stack solve in order:
//   1. CapEx CF circularity with approximate take-out amounts
//   2. Project-level term debt (if CEG_ProjectDebt = "Project")
//   3. Tax equity (if CEG_Scenario1Active):
//        — If CEG_FMVOverride = "": iterate FMV income approach
//        — Solve TEI upfront investment
//   4. Backleverage term debt (if CEG_ProjectDebt = "BL")
//   5. Third-party CE investment (if CEG_Scenario6Active)
//   6. CWEN investment (if CEG_Scenario5Active)
//   7. CapEx CF circularity with resolved take-out amounts
//   8. Paste metrics
// ═══════════════════════════════════════════════════════════════════════════════

export function solveCapitalStack() {
  var t0 = Date.now();
  var MAX_ITER = 50;
  var projectDebt = null;

  return _checkModel().then(function () {

    writeLog('Solve Capital Stack: Starting…', 'info');

    // ── Step 1: Set ignore flag, solve CapEx CF circularity ───────────────────
    writeLog('Solve Capital Stack: Solving CF with approximate take-out amounts…', 'info');
    return Excel.run(function (context) {
      var nIgnore = context.workbook.names.getItemOrNullObject('CEG_CF_IgnoreTakeoutAmounts');
      nIgnore.load('isNullObject');
      return context.sync().then(function () {
        if (nIgnore.isNullObject) {
          writeLog('Solve Capital Stack: Missing CEG_CF_IgnoreTakeoutAmounts.', 'error');
          return false;
        }
        nIgnore.getRange().values = [[true]];
        return context.sync().then(function () { return true; });
      });
    })
    .then(function (ok) {
      if (!ok) return;
      return solveCapexCFCircularity();
    })

    // ── Step 2: Read CEG_ProjectDebt ────────────────────────────────────────
    .then(function () {
      return Excel.run(function (context) {
        var nProjectDebt = context.workbook.names.getItemOrNullObject('CEG_ProjectDebt');
        nProjectDebt.load('isNullObject');
        return context.sync().then(function () {
          if (nProjectDebt.isNullObject) {
            writeLog('Solve Capital Stack: Missing CEG_ProjectDebt.', 'error');
            return null;
          }
          var r = nProjectDebt.getRange();
          r.load('values');
          return context.sync().then(function () {
            projectDebt = r.values[0][0];
            return projectDebt;
          });
        });
      });
    })

    // ── Step 3: Project-level term debt ─────────────────────────────────────
    .then(function (pd) {
      if (pd === 'Project') {
        writeLog('Solve Capital Stack: Solving project-level term debt…', 'info');
        return solveTermDebt();
      }
    })

    // ── Step 4: Tax equity — check scenario flag ─────────────────────────────
    .then(function () {
      return Excel.run(function (context) {
        var nScenario1Active = context.workbook.names.getItemOrNullObject('CEG_Scenario1Active');
        nScenario1Active.load('isNullObject');
        return context.sync().then(function () {
          if (nScenario1Active.isNullObject) return null;
          var r = nScenario1Active.getRange();
          r.load('values');
          return context.sync().then(function () { return r.values[0][0]; });
        });
      });
    })
    .then(function (scenario1Active) {
      if (!scenario1Active) return;

      // Set financing scenario 1, then read FMV override flag
      return Excel.run(function (context) {
        context.workbook.names.getItem('CEG_FinancingScenario').getRange().values = [[1]];

        var nFMVOverride = context.workbook.names.getItemOrNullObject('CEG_FMVOverride');
        nFMVOverride.load('isNullObject');

        return context.sync().then(function () {
          if (nFMVOverride.isNullObject) return undefined;
          var r = nFMVOverride.getRange();
          r.load('values');
          return context.sync().then(function () { return r.values[0][0]; });
        });
      })
      .then(function (fmvOverride) {
        // Run FMV income approach solve only when override field is blank
        if (fmvOverride !== '') return;

        writeLog('Solve Capital Stack: Solving FMV income approach…', 'info');
        return Excel.run(function (context) {
          var wb    = context.workbook;
          var nDiff = wb.names.getItemOrNullObject('CEG_IncomeApproachDiff');
          var nHC   = wb.names.getItemOrNullObject('CEG_IncomeApproachHC');
          var nLive = wb.names.getItemOrNullObject('CEG_IncomeApproachLive');
          [nDiff, nHC, nLive].forEach(function (n) { n.load('isNullObject'); });

          return context.sync().then(function () {
            var missing = [];
            if (nDiff.isNullObject) missing.push('CEG_IncomeApproachDiff');
            if (nHC.isNullObject)   missing.push('CEG_IncomeApproachHC');
            if (nLive.isNullObject) missing.push('CEG_IncomeApproachLive');
            if (missing.length > 0) {
              writeLog('Solve Capital Stack: Missing FMV range(s): ' + missing.join(', '), 'error');
              return;
            }
            return _fmvIncomeApproachLoop(context,
              nDiff.getRange(), nHC.getRange(), nLive.getRange(), 0, MAX_ITER);
          });
        });
      })
      .then(function () {
        writeLog('Solve Capital Stack: Solving tax equity upfront investment…', 'info');
        return solveTEIUpfrontInvestment();
      });
    })

    // ── Step 5: Backleverage term debt ───────────────────────────────────────
    .then(function () {
      if (projectDebt === 'BL') {
        writeLog('Solve Capital Stack: Solving backleverage term debt…', 'info');
        return solveTermDebt();
      }
    })

    // ── Step 6: Third-party CE investment ────────────────────────────────────
    .then(function () {
      return Excel.run(function (context) {
        var nScenario6Active = context.workbook.names.getItemOrNullObject('CEG_Scenario6Active');
        nScenario6Active.load('isNullObject');
        return context.sync().then(function () {
          if (nScenario6Active.isNullObject) return false;
          var r = nScenario6Active.getRange();
          r.load('values');
          return context.sync().then(function () { return r.values[0][0]; });
        });
      });
    })
    .then(function (scenario6Active) {
      if (scenario6Active) {
        writeLog('Solve Capital Stack: Solving third-party CE investment…', 'info');
        return solveCE2UpfrontInvestment();
      }
    })

    // ── Step 7: CWEN investment ──────────────────────────────────────────────
    .then(function () {
      return Excel.run(function (context) {
        var nScenario5Active = context.workbook.names.getItemOrNullObject('CEG_Scenario5Active');
        nScenario5Active.load('isNullObject');
        return context.sync().then(function () {
          if (nScenario5Active.isNullObject) return false;
          var r = nScenario5Active.getRange();
          r.load('values');
          return context.sync().then(function () { return r.values[0][0]; });
        });
      });
    })
    .then(function (scenario5Active) {
      if (scenario5Active) {
        writeLog('Solve Capital Stack: Solving CWEN investment…', 'info');
        return solveCWENUpfrontInvestment();
      }
    })

    // ── Step 8: Solve CF with resolved take-out amounts ──────────────────────
    .then(function () {
      writeLog('Solve Capital Stack: Solving CF with resolved take-out amounts…', 'info');
      return Excel.run(function (context) {
        context.workbook.names.getItem('CEG_CF_IgnoreTakeoutAmounts').getRange().values = [[false]];
        return context.sync();
      });
    })
    .then(function () {
      return solveCapexCFCircularity();
    })

    // ── Step 9: Paste metrics ────────────────────────────────────────────────
    .then(function () {
      writeLog('Solve Capital Stack: Pasting metrics…', 'info');
      return pasteMetrics();
    });

  })
  .then(function () {
    writeLog('Solve Capital Stack: completed in ' + ((Date.now() - t0) / 1000).toFixed(2) + 's.', 'info');
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// COMMAND — Run Scenarios
// Port of ScenarioManager.RunScenarios().
//
// Iterates CEG_InputsScenario from 1 to CEG_TotalScenarios, calling
// solveCapitalStack for each scenario in sequence.
// ═══════════════════════════════════════════════════════════════════════════════

function _runScenariosLoop(totalScenarios, currentScenario) {
  if (currentScenario > totalScenarios) {
    writeLog('Run Scenarios: All ' + totalScenarios + ' scenario(s) complete.', 'success');
    return Promise.resolve();
  }

  writeLog('Run Scenarios: Starting scenario ' + currentScenario + ' of ' + totalScenarios + '…', 'info');

  return Excel.run(function (context) {
    context.workbook.names.getItem('CEG_InputsScenario').getRange().values = [[currentScenario]];
    return context.sync();
  })
  .then(function () {
    return solveCapitalStack();
  })
  .then(function () {
    return _runScenariosLoop(totalScenarios, currentScenario + 1);
  });
}

export function runScenarios() {
  var t0 = Date.now();

  return _checkModel().then(function () {
    return Excel.run(function (context) {
      var nTotalScenarios = context.workbook.names.getItemOrNullObject('CEG_TotalScenarios');
      var nInputsScenario = context.workbook.names.getItemOrNullObject('CEG_InputsScenario');
      [nTotalScenarios, nInputsScenario].forEach(function (n) { n.load('isNullObject'); });

      return context.sync().then(function () {
        var missing = [];
        if (nTotalScenarios.isNullObject) missing.push('CEG_TotalScenarios');
        if (nInputsScenario.isNullObject) missing.push('CEG_InputsScenario');
        if (missing.length > 0) {
          writeLog('Run Scenarios: Missing named range(s): ' + missing.join(', '), 'error');
          return null;
        }

        var r = nTotalScenarios.getRange();
        r.load('values');
        return context.sync().then(function () {
          return r.values[0][0];
        });
      });
    })
    .then(function (totalScenarios) {
      if (totalScenarios === null) return;
      writeLog('Run Scenarios: Running ' + totalScenarios + ' scenario(s)…', 'info');
      return _runScenariosLoop(totalScenarios, 1);
    });
  })
  .then(function () {
    writeLog('Run Scenarios: completed in ' + ((Date.now() - t0) / 1000).toFixed(2) + 's.', 'info');
  });
}
