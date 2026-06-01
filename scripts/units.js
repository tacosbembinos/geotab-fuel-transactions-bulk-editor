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

  // datetime-local string → canonical UTC ISO. Browser-local interpretation.
  function localInputToIso(s) {
    if (!s) return null;
    var d = new Date(s);
    return isNaN(d) ? null : d.toISOString();
  }

  // Detect a string that already declares its offset (trailing Z, +HH:MM,
  // -HH:MM, or +HHMM). Such strings are unambiguous and should bypass the
  // tzMode interpretation below.
  var EXPLICIT_OFFSET_RE = /(Z|[+\-]\d{2}:?\d{2})\s*$/;
  function hasExplicitOffset(s) {
    return typeof s === 'string' && EXPLICIT_OFFSET_RE.test(s.trim());
  }

  // Given a naive wall-clock (y/mo/d/h/mi/s) interpreted in an IANA zone,
  // return the corresponding UTC epoch ms. Trick: assume UTC first, ask
  // Intl what that instant LOOKS LIKE in the target zone, measure the drift,
  // and correct. Works for any historical or future offset Intl knows about.
  function zonedWallClockToUtcMs(y, mo, d, h, mi, s, zone) {
    var candidate = Date.UTC(y, mo - 1, d, h, mi, s || 0);
    try {
      var parts = new Intl.DateTimeFormat('en-US', {
        timeZone: zone, hour12: false,
        year: 'numeric', month: '2-digit', day: '2-digit',
        hour: '2-digit', minute: '2-digit', second: '2-digit'
      }).formatToParts(new Date(candidate));
      var o = {};
      for (var i = 0; i < parts.length; i++) {
        if (parts[i].type !== 'literal') o[parts[i].type] = parts[i].value;
      }
      // Intl quirk: hour can come back as "24" at midnight in some locales.
      var hh = +o.hour === 24 ? 0 : +o.hour;
      var asZoneMs = Date.UTC(+o.year, +o.month - 1, +o.day, hh, +o.minute, +o.second);
      return candidate - (asZoneMs - candidate);
    } catch (_) {
      return candidate;   // fall back to UTC interpretation if zone unknown
    }
  }

  // Parse the date-portion of a CSV datetime string into [y, mo, d, h, mi, s].
  // Accepts the formats MyGeotab's bulk-import tool and common fuel-card
  // exports emit, including "YYYY-MM-DDTHH:MM:SS", "YYYY-MM-DD HH:MM",
  // "M/D/YYYY H:MM[:SS] [AM|PM]", and bare "YYYY-MM-DD".
  function parseNaiveDateTime(s) {
    if (typeof s !== 'string') return null;
    var str = s.trim().replace(/(Z|[+\-]\d{2}:?\d{2})\s*$/, '');
    var m = str.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$/);
    if (m) return [+m[1], +m[2], +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)];
    m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM|am|pm)?)?$/);
    if (m) {
      var hr = +(m[4] || 0);
      var ampm = (m[7] || '').toUpperCase();
      if (ampm === 'PM' && hr < 12) hr += 12;
      if (ampm === 'AM' && hr === 12) hr = 0;
      return [+m[3], +m[1], +m[2], hr, +(m[5] || 0), +(m[6] || 0)];
    }
    return null;
  }

  // CSV cell → canonical UTC ISO 8601 string. opts:
  //   tzMode: 'browser' (default — interpret wall-clock as browser-local)
  //         | 'utc'                — interpret wall-clock as UTC
  //         | 'iana'               — interpret in opts.tzIana
  //   tzIana: IANA zone name when tzMode === 'iana'
  // Strings that already carry Z/+HH:MM are returned as-is regardless of mode.
  function csvDateToIso(s, opts) {
    if (s == null || s === '') return null;
    var str = String(s).trim();
    if (!str) return null;
    if (hasExplicitOffset(str)) {
      var d = new Date(str);
      return isNaN(d) ? null : d.toISOString();
    }
    var mode = (opts && opts.tzMode) || 'browser';
    var parts = parseNaiveDateTime(str);
    if (!parts) {
      var fallback = new Date(str);    // last-ditch: trust the engine
      return isNaN(fallback) ? null : fallback.toISOString();
    }
    var y = parts[0], mo = parts[1], dy = parts[2], h = parts[3], mi = parts[4], se = parts[5];
    var ms;
    if (mode === 'utc') {
      ms = Date.UTC(y, mo - 1, dy, h, mi, se);
    } else if (mode === 'iana' && opts && opts.tzIana) {
      ms = zonedWallClockToUtcMs(y, mo, dy, h, mi, se, opts.tzIana);
    } else {
      // browser-local: build a Date in local time, then take its UTC ISO.
      var local = new Date(y, mo - 1, dy, h, mi, se);
      ms = local.getTime();
    }
    if (isNaN(ms)) return null;
    return new Date(ms).toISOString();
  }

  // Detect a string that already names its unit (e.g. "12.3 gal", "150 L").
  // Returns the bare numeric string with the suffix stripped, or null if the
  // input has no recognisable unit hint.
  var UNIT_HINT_RE = /\s*(L|l|liter|liters|litre|litres|gal|gallons?|US gal|km|kilometer|kilometers|kilometre|kilometres|mi|miles?)\b/;
  function stripUnitHint(s) {
    if (typeof s !== 'string') return null;
    var t = s.trim();
    if (!UNIT_HINT_RE.test(t)) return null;
    return t.replace(UNIT_HINT_RE, '').trim();
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
    localInputToIso:   localInputToIso,
    csvDateToIso:      csvDateToIso,
    hasExplicitOffset: hasExplicitOffset,
    stripUnitHint:     stripUnitHint,
    parseNaiveDateTime: parseNaiveDateTime,
    zonedWallClockToUtcMs: zonedWallClockToUtcMs
  };
}(typeof window !== 'undefined' ? window : this));
