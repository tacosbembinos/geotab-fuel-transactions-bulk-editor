/*
 * Fuel Transactions Bulk Editor — units.js
 *
 * Single source of truth for unit conversion + display formatting.
 *
 * Canonical (Geotab API) storage is always Liters / km / UTC ISO 8601.
 * These helpers convert between canonical and the user-selected DISPLAY
 * unit; they never mutate stored values. main.js wraps each helper so
 * existing call sites keep the same signature and continue to read the
 * display unit from `ui.volUnit` / `ui.odoUnit`.
 *
 * Exposed via window.FTBE_Units (same pattern as window.CSVUtil in csv.js)
 * so the bootstrap order in fuelBulkEditor.html stays predictable:
 *   1. units.js   — pure helpers, no dependencies
 *   2. csv.js     — pure helpers, no dependencies
 *   3. main.js    — addin entry, consumes both
 */
(function (root) {
  'use strict';

  // Conversion factors — exact NIST values where they exist.
  // 1 US gallon = 3.785411784 L (exact, by NIST 2008 redefinition).
  // 1 mile     = 1.609344 km    (exact, by international yard/pound agreement).
  var L_PER_GAL_US = 3.785411784;
  var KM_PER_MI    = 1.609344;

  function fromDisplayVolume(v, volUnit) {
    return volUnit === 'gal' ? v * L_PER_GAL_US : v;
  }
  function toDisplayVolume(v, volUnit) {
    return volUnit === 'gal' ? v / L_PER_GAL_US : v;
  }
  function fromDisplayOdo(v, odoUnit) {
    return odoUnit === 'mi' ? v * KM_PER_MI : v;
  }
  function toDisplayOdo(v, odoUnit) {
    return odoUnit === 'mi' ? v / KM_PER_MI : v;
  }

  function fmtNum(v, digits) {
    return v == null || isNaN(v)
      ? ''
      : Number(v).toFixed(digits != null ? digits : 2);
  }

  function fmtDateTime(iso) {
    if (!iso) return '';
    try { return new Date(iso).toLocaleString(); }
    catch (_) { return String(iso); }
  }

  // ISO timestamp → value suitable for <input type="datetime-local">.
  // Browser-local wall clock; seconds dropped because datetime-local
  // inputs don't accept them by default.
  function isoToLocalInput(iso) {
    if (!iso) return '';
    var d = new Date(iso);
    var pad = function (n) { return String(n).padStart(2, '0'); };
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) +
           'T' + pad(d.getHours()) + ':' + pad(d.getMinutes());
  }

  // datetime-local string → canonical UTC ISO. Browser-local interpretation
  // is intentional today; PR-3 will replace this with an explicit TZ choice.
  function localInputToIso(s) {
    if (!s) return null;
    var d = new Date(s);
    return isNaN(d) ? null : d.toISOString();
  }

  root.FTBE_Units = {
    L_PER_GAL_US: L_PER_GAL_US,
    KM_PER_MI:    KM_PER_MI,
    fromDisplayVolume: fromDisplayVolume,
    toDisplayVolume:   toDisplayVolume,
    fromDisplayOdo:    fromDisplayOdo,
    toDisplayOdo:      toDisplayOdo,
    fmtNum:            fmtNum,
    fmtDateTime:       fmtDateTime,
    isoToLocalInput:   isoToLocalInput,
    localInputToIso:   localInputToIso
  };
}(typeof window !== 'undefined' ? window : this));
