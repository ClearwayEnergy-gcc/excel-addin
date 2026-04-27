// config.js — Clearway Project Financial Model Add-in
//
// Workbook-level configuration detected by checkModel().
// Shared unchanged between the Office.js add-in and the Claude skill package
// (no environment-specific code here).
//
// ── HOW TO ADD A NEW VARIANT ────────────────────────────────────────────────
//
//   When a new workbook divergence is discovered after deployment:
//
//   1. Add a flag below with a safe default (base-template behavior = no flag).
//      Document which named range drives it and what values to expect.
//
//   2. In commands.js → _checkModel(): read the named range and include the
//      flag in the setConfig({ ... }) call.
//
//   3. In the affected command(s) in commands.js: read the flag with
//      getConfig().yourFlag and add the branching logic.
//
//   4. Commit — the pre-commit hook syncs this file and commands.js to
//      skill-package/ automatically.
//
// ── CONFIG SCHEMA ───────────────────────────────────────────────────────────

var _defaults = {
  populated:    false,  // true once _checkModel() has run successfully
  modelVersion: null,   // string value of CEG_ModelTemplateVersion, or null

  // ── Variant flags ─────────────────────────────────────────────────────────
  // None yet. Add here as workbook divergences are identified post-deployment.
  // Example shape (do not uncomment — for reference only):
  //
  //   hasTaxEquity:    true,   // false if CEG_TEStructure = "None"
  //   projectDebtType: 'Project',  // 'Project' | 'BL' | 'None'
};

var _config = Object.assign({}, _defaults);

// ── Public interface ─────────────────────────────────────────────────────────

// Read the current config. Commands call this to branch on variant flags.
export function getConfig() {
  return _config;
}

// Write variant flags detected by _checkModel(). Always marks populated: true.
export function setConfig(values) {
  Object.assign(_config, values, { populated: true });
}

// Reset to defaults — useful if the user switches workbooks mid-session.
export function resetConfig() {
  _config = Object.assign({}, _defaults);
}
