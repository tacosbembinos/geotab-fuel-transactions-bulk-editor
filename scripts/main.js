/* global geotab */
/*
 * Fuel Transactions Bulk Editor — main.js
 *
 * Lessons applied from BUILD_GUIDE.md / BUILDING_A_GENERIC_ADDIN.md:
 *  - shim window.geotab before assignment
 *  - callback() invoked unconditionally in initialize
 *  - blur() aborts in-flight, clears timers; unload() full teardown
 *  - all bulk writes via api.multiCall, chunked at MULTICALL_CHUNK = 100,
 *    SEQUENTIAL chunks (to respect the 100/min/user rate cap)
 *  - escapeHtml() before any innerHTML; textContent preferred otherwise
 *  - whitelist enums (productType, currencyCode) before composing entities
 *  - optimistic-concurrency: keep `version` from Get; on Set failure re-Get and resurface
 *  - modal a11y: toggle [hidden] + [inert] + aria-hidden; restore focus on close
 *  - standalone bootstrap shim at bottom for file:// preview
 */
(function () {
  'use strict';
  if (typeof window !== 'undefined' && typeof window.geotab === 'undefined') {
    window.geotab = { addin: {} };
  }
})();

geotab.addin.fuelBulkEditor = function () {
  'use strict';

  // ── Constants ────────────────────────────────────────────────────────────
  const MULTICALL_CHUNK = 200;            // BUILD_GUIDE §6
  const RESULTS_LIMIT   = 50000;          // matches official fuel-tracker-edit
  const L_PER_GAL_US    = 3.785411784;
  const KM_PER_MI       = 1.609344;

  // Geotab per-user API quota. The HAR capture shows this tenant is provisioned
  // at "Maximum admitted 200 per 1m." Sub-calls inside an ExecuteMultiCall
  // count individually, so a naïve 56-chunk save of ~5,600 Set sub-calls trips
  // the bucket on the 3rd chunk. We throttle proactively on a sliding window,
  // auto-calibrate from the server's own error message (some tenants run at
  // 100/min, others higher), and back off reactively when OverLimitException
  // still slips through (e.g. other MyGeotab tabs sharing the same quota).
  // Kept as `let` (not frozen) so parseQuotaFromMessage() can adjust live.
  const RATE_LIMIT = {
    callsPerWindow:    200,
    windowMs:          60 * 1000,
    safetyMargin:      0.9,         // effective cap ≈ 180/min
    maxRetries:        5,
    defaultCooldownMs: 60 * 1000    // when err carries no usable hint
  };

  const PRODUCT_TYPES = Object.freeze([
    'Unknown', 'Regular', 'Midgrade', 'Premium', 'Super',
    'Diesel', 'DieselExhaustFluid', 'E85', 'CNG', 'LPG',
    'Electric', 'Hydrogen', 'NonFuel'
  ]);
  const CURRENCIES = Object.freeze([
    'USD', 'CAD', 'GBP', 'EUR', 'AUD', 'MXN', 'JPY', 'PHP', 'INR', 'BRL'
  ]);
  // Editable on FuelTransaction Set (others are server-controlled or pass-through).
  const EDITABLE_FIELDS = Object.freeze([
    'dateTime', 'volume', 'cost', 'currencyCode',
    'odometer', 'productType', 'comments'
  ]);

  // ── State ────────────────────────────────────────────────────────────────
  let api = null;
  let state = null;
  const ui = {
    initialized: false,
    inflight: new Set(),       // AbortController-like { abort() } shims
    opGen: 0,                  // bumped on every blur/unload — callers capture it and bail on stale
    rows: [],                  // current loaded FuelTransactions (canonical/raw shape)
    edited: new Map(),         // id -> patch object (pending UI edits)
    selected: new Set(),       // selected row ids
    unmatchedPreview: [],      // CSV rows with no FT match — shown grey at top of table; ⇒ Force Import
    duplicateTargets: new Map(), // tx.id -> count of CSV rows that matched it (>1 = warning)
    sortKey: 'dateTime',
    sortDir: 'desc',
    deviceById: new Map(),     // deviceId -> name
    driverById: new Map(),     // driverId -> name
    volUnit: 'L',
    odoUnit: 'km',
    lastFocusEl: null          // for modal focus restoration
  };

  // Cancellation contract:
  //   - blur()/unload() bump ui.opGen and abort every handle in ui.inflight.
  //   - apiCall rejects its Promise with CANCELLED so chained .then()s don't
  //     orphan and status updates can run their .catch path.
  //   - apiMultiCall short-circuits between chunks: any chunks not yet issued
  //     are padded with { __cancelled: true } so callers can count partials.
  //   - Top-level handlers capture `const myGen = ui.opGen` at entry and
  //     short-circuit if it no longer matches before touching shared state.
  //     That prevents stale loads from clobbering a fresher one and stops
  //     post-completion side effects (e.g. auto-reload after Save) from
  //     firing on an add-in the user has navigated away from.
  const CANCELLED = Object.freeze({ __cancelled: true });
  function isCancelled(e) { return !!(e && e.__cancelled); }
  function isStale(gen) { return gen !== ui.opGen; }

  // ── DOM helpers ──────────────────────────────────────────────────────────
  const $ = (id) => document.getElementById(id);
  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
  }
  function setStatus(msg, kind) {
    const el = $('ftbe-status');
    if (!el) return;
    el.textContent = msg || '';
    el.classList.remove('is-error', 'is-success');
    if (kind === 'error')   el.classList.add('is-error');
    if (kind === 'success') el.classList.add('is-success');
  }

  // ── API plumbing ─────────────────────────────────────────────────────────
  // Sliding-window subcall ledger. Each entry { t, n } records how many
  // sub-calls were emitted at time t. We prune entries older than windowMs.
  const rateState = { events: [] };
  function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

  // Single-flight gate: only ONE api.call / api.multiCall is in flight at a
  // time. The next request waits for the prior to settle (success OR error)
  // before being issued. This is the hard guarantee the throttle relies on —
  // without it, two callers can both pass awaitCapacity() in the same tick,
  // both call recordSubcalls(), and double-emit before the ledger reflects
  // either one. With it, the rate-window math always sees a consistent state.
  let apiGate = Promise.resolve();
  function withApiGate(fn) {
    const run = apiGate.then(fn, fn);
    apiGate = run.then(() => {}, () => {});
    return run;
  }
  function pruneRateEvents(now) {
    const cutoff = now - RATE_LIMIT.windowMs;
    while (rateState.events.length && rateState.events[0].t < cutoff) rateState.events.shift();
  }
  function subcallsInWindow(now) {
    pruneRateEvents(now);
    let sum = 0;
    for (const e of rateState.events) sum += e.n;
    return sum;
  }
  function recordSubcalls(n) { rateState.events.push({ t: Date.now(), n }); }

  async function awaitCapacity(n, abortHandle) {
    const cap = Math.floor(RATE_LIMIT.callsPerWindow * RATE_LIMIT.safetyMargin);
    // If a single chunk is larger than the cap (shouldn't happen — chunk≤100,
    // cap=180), drain the whole window first.
    while (!(abortHandle && abortHandle.aborted)) {
      const now = Date.now();
      const used = subcallsInWindow(now);
      if (used + n <= cap) return;
      // Walk events oldest-first until enough will have aged out to fit.
      const need = used + n - cap;
      let freed = 0, waitUntil = now;
      for (const e of rateState.events) {
        freed += e.n;
        waitUntil = e.t + RATE_LIMIT.windowMs;
        if (freed >= need) break;
      }
      const waitMs = Math.max(250, waitUntil - now + 50);
      setStatus('Rate-limit throttle: waiting ' + Math.ceil(waitMs / 1000) + 's before next batch…');
      await sleep(waitMs);
    }
  }

  // Walks the entire error envelope. The MyGeotab SDK has delivered OverLimit
  // in three shapes the wild: top-level `{name: 'OverLimitException', ...}`,
  // wrapped `{error: {errors: [{name: 'OverLimitException'}]}, requestIndex: 0}`,
  // and the per-subcall inline form returned via multiCall's success callback.
  // A shallow check on err.name misses the latter two; a recursive dig catches
  // all of them. Bounded by depth + a cycle guard.
  const OVER_LIMIT_RX = /OverLimitException|quota exceeded/i;
  function isOverLimitError(err) {
    if (err == null) return false;
    const seen = new Set();
    function dig(node, depth) {
      if (node == null || depth > 6) return false;
      if (typeof node === 'string') return OVER_LIMIT_RX.test(node);
      if (typeof node !== 'object') return false;
      if (seen.has(node)) return false;
      seen.add(node);
      if (Array.isArray(node)) {
        for (const item of node) if (dig(item, depth + 1)) return true;
        return false;
      }
      for (const k of ['name', 'message', 'type']) {
        if (node[k] != null && OVER_LIMIT_RX.test(String(node[k]))) return true;
      }
      for (const k of ['error', 'data', 'cause']) {
        if (k in node && dig(node[k], depth + 1)) return true;
      }
      if (Array.isArray(node.errors)) {
        for (const e of node.errors) if (dig(e, depth + 1)) return true;
      }
      return false;
    }
    return dig(err, 0);
  }

  // Inspect a multiCall success result for inline JSON-RPC error envelopes.
  // `api.multiCall` may invoke the SUCCESS callback with an array where each
  // element is either a real result OR `{error: {...}, requestIndex: N}` — the
  // HAR shows this shape with OverLimitException. Without this check we'd
  // treat a fully-failed batch as success and immediately fire the next one.
  function inlineMultiCallErrors(results) {
    if (!Array.isArray(results)) return null;
    const errs = [];
    for (let i = 0; i < results.length; i++) {
      const r = results[i];
      if (r && typeof r === 'object' && r.error) {
        errs.push({ index: i, err: r.error });
      }
    }
    if (!errs.length) return null;
    const overLimit = errs.some((e) => isOverLimitError(e.err));
    return { errs, overLimit, firstErr: errs[0].err };
  }

  // Parse "Maximum admitted N per Wm" / "...per Ws" out of the JSON-RPC error
  // message and adjust RATE_LIMIT live. Geotab provisions quotas per tenant;
  // a tenant on 100/min would otherwise see this code over-throttle.
  function parseQuotaFromMessage(err) {
    const msg = (err && err.message) || '';
    const m = msg.match(/Maximum admitted\s+(\d+)\s+per\s+(\d+)\s*([smh])/i);
    if (!m) return;
    const n = parseInt(m[1], 10);
    const w = parseInt(m[2], 10);
    const unit = m[3].toLowerCase();
    const ms = (unit === 'h' ? 3600 : unit === 'm' ? 60 : 1) * 1000 * w;
    if (n > 0 && ms > 0) {
      RATE_LIMIT.callsPerWindow = n;
      RATE_LIMIT.windowMs = ms;
      RATE_LIMIT.defaultCooldownMs = ms;
    }
  }

  function parseRetryAfterMs(err) {
    // api.multiCall's error callback receives the parsed JSON-RPC error and
    // does NOT expose the HTTP Retry-After header. Best-effort: honour
    // err.retryAfter if a custom transport forwarded it; otherwise fall back
    // to the full window (the server resets the bucket on a fixed cadence).
    const ra = err && (err.retryAfter || (err.data && err.data.retryAfter));
    if (typeof ra === 'number' && isFinite(ra) && ra > 0) return Math.max(1000, ra * 1000);
    return RATE_LIMIT.defaultCooldownMs;
  }

  // Wraps api.call in a Promise + tracks a cancel handle in ui.inflight.
  // Counts as 1 sub-call against the rate window. Throttles proactively,
  // retries on OverLimitException with exponential backoff, and REJECTS with
  // CANCELLED on blur/unload (never silently orphans the callback chain).
  function apiCall(method, params, opts) {
    opts = opts || {};
    if (!api || typeof api.call !== 'function') {
      return Promise.reject(new Error('API unavailable (standalone preview)'));
    }
    const handle = { aborted: false, abort() { this.aborted = true; } };
    ui.inflight.add(handle);

    const invoke = () => new Promise((resolve) => {
      api.call(method, params,
        (res) => resolve({ ok: true, res }),
        (err) => resolve({ ok: false, err })
      );
    });

    return (async () => {
      let attempt = 0;
      try {
        while (true) {
          if (handle.aborted || (opts.gen != null && isStale(opts.gen))) throw CANCELLED;
          await awaitCapacity(1, handle);
          if (handle.aborted) throw CANCELLED;
          // Single-flight: the call below cannot start until any prior api
          // request has settled. recordSubcalls is moved inside the gate so
          // the ledger is updated atomically with the actual emission.
          const out = await withApiGate(async () => {
            recordSubcalls(1);
            return invoke();
          });
          if (handle.aborted || (opts.gen != null && isStale(opts.gen))) throw CANCELLED;
          if (out.ok) return out.res;
          if (isOverLimitError(out.err) && attempt < RATE_LIMIT.maxRetries) {
            parseQuotaFromMessage(out.err);
            const waitMs = Math.floor(parseRetryAfterMs(out.err) * Math.pow(1.25, attempt));
            setStatus('Rate limit hit. Cooling down ' + Math.ceil(waitMs / 1000) + 's and retrying…', 'error');
            await sleep(waitMs);
            // Do NOT clear rateState.events — they prune naturally once older
            // than windowMs, and clearing lets the next attempt skip throttling
            // when other tabs are still consuming the same per-tenant bucket.
            attempt++;
            continue;
          }
          throw out.err;
        }
      } finally {
        ui.inflight.delete(handle);
      }
    })();
  }

  // multiCall in SEQUENTIAL chunks of MULTICALL_CHUNK.
  //   - Validates each sub-call (server aborts a whole batch on one bad shape).
  //   - Throttles against RATE_LIMIT proactively (sliding window of sub-calls).
  //   - Retries the current chunk on OverLimitException — atomic per-batch
  //     (HAR shows requestIndex:0, no sub-calls committed), so re-send is safe.
  //   - Real cancellation between chunks: remaining sub-calls are padded with
  //     { __cancelled: true } and resolve immediately. We can't recall chunks
  //     already in flight on the server, but we stop issuing new ones — which
  //     is what matters for blur during a bulk Save.
  //   - Progress: when chunks > 1, opts.label drives a "N / total" status.
  // opts: { label?: string, gen?: number }
  function apiMultiCall(calls, opts) {
    opts = opts || {};
    if (!api || typeof api.multiCall !== 'function') {
      return Promise.reject(new Error('multiCall unavailable'));
    }
    const valid = calls.filter((c) => Array.isArray(c) && typeof c[0] === 'string' && c[1] && typeof c[1] === 'object');
    if (valid.length !== calls.length) {
      return Promise.reject(new Error('multiCall: malformed sub-call detected; refusing to send.'));
    }
    const chunks = [];
    for (let i = 0; i < valid.length; i += MULTICALL_CHUNK) {
      chunks.push(valid.slice(i, i + MULTICALL_CHUNK));
    }
    const handle = { aborted: false, abort() { this.aborted = true; } };
    ui.inflight.add(handle);

    const sendChunk = (chunk) => new Promise((resolve) => {
      api.multiCall(chunk,
        (results) => resolve({ ok: true, results: results || [] }),
        (err)     => resolve({ ok: false, err })
      );
    });

    const total = valid.length;
    const label = opts.label || null;
    const checkAborted = () =>
      handle.aborted || (opts.gen != null && isStale(opts.gen));

    return (async () => {
      const all = [];
      let done = 0;
      try {
        for (let idx = 0; idx < chunks.length; idx++) {
          const chunk = chunks[idx];
          if (checkAborted()) {
            const rest = chunks.slice(idx).reduce((n, c) => n + c.length, 0);
            for (let k = 0; k < rest; k++) all.push(CANCELLED);
            return all;
          }
          let attempt = 0;
          while (true) {
            await awaitCapacity(chunk.length, handle);
            if (checkAborted()) {
              const rest = chunks.slice(idx).reduce((n, c) => n + c.length, 0);
              for (let k = 0; k < rest; k++) all.push(CANCELLED);
              return all;
            }
            if (label && chunks.length > 1) {
              setStatus(label + ' — ' + done + ' / ' + total +
                        ' (chunk ' + (idx + 1) + ' of ' + chunks.length + ')…');
            }
            // Single-flight + atomic ledger update: the next multiCall is
            // not dispatched until the previous one fully settles. This is
            // the load-bearing guarantee the task asks for — "only send the
            // next multiCall when the last one is completed."
            const out = await withApiGate(async () => {
              recordSubcalls(chunk.length);
              return sendChunk(chunk);
            });
            // multiCall can deliver per-subcall errors via either callback.
            // Treat both paths uniformly.
            const inline = out.ok ? inlineMultiCallErrors(out.results) : null;
            const overLimit =
              (!out.ok && isOverLimitError(out.err)) ||
              (inline && inline.overLimit);
            if (overLimit && attempt < RATE_LIMIT.maxRetries) {
              const errForMsg = out.ok ? inline.firstErr : out.err;
              parseQuotaFromMessage(errForMsg);
              const waitMs = Math.floor(parseRetryAfterMs(errForMsg) * Math.pow(1.25, attempt));
              setStatus('Rate limit hit on chunk ' + (idx + 1) + ' / ' + chunks.length +
                        '. Cooling down ' + Math.ceil(waitMs / 1000) + 's and retrying…', 'error');
              await sleep(waitMs);
              // Keep rateState.events intact — the cooldown is >= windowMs so
              // they prune naturally on the next awaitCapacity() pass, and
              // leaving them shields against per-tenant buckets shared with
              // other tabs that may still be draining the quota.
              attempt++;
              continue;
            }
            if (out.ok) {
              if (inline) {
                // Some subcalls have non-rate-limit errors. Surface each one
                // as __error in its original position so the caller's
                // results.forEach((res, idx) => ...) sees them.
                for (let i = 0; i < out.results.length; i++) {
                  const r = out.results[i];
                  if (r && typeof r === 'object' && r.error) {
                    all.push({ __error: r.error });
                  } else {
                    all.push(r);
                  }
                }
              } else {
                all.push.apply(all, out.results);
              }
              done += chunk.length;
              break;
            }
            // Transport-level failure (and not OverLimit, or retries exhausted).
            for (let k = 0; k < chunk.length; k++) all.push({ __error: out.err });
            done += chunk.length;
            break;
          }
        }
        return all;
      } finally {
        ui.inflight.delete(handle);
      }
    })();
  }

  // ── Unit + format helpers ────────────────────────────────────────────────
  function fromDisplayVolume(v) { return ui.volUnit === 'gal' ? v * L_PER_GAL_US : v; }
  function toDisplayVolume(v)   { return ui.volUnit === 'gal' ? v / L_PER_GAL_US : v; }
  function fromDisplayOdo(v)    { return ui.odoUnit === 'mi'  ? v * KM_PER_MI    : v; }
  function toDisplayOdo(v)      { return ui.odoUnit === 'mi'  ? v / KM_PER_MI    : v; }
  function fmtNum(v, digits)    { return v == null || isNaN(v) ? '' : Number(v).toFixed(digits != null ? digits : 2); }
  function fmtDateTime(iso) {
    if (!iso) return '';
    try { return new Date(iso).toLocaleString(); } catch (_) { return String(iso); }
  }
  function isoToLocalInput(iso) {
    if (!iso) return '';
    const d = new Date(iso);
    const pad = (n) => String(n).padStart(2, '0');
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) +
           'T' + pad(d.getHours()) + ':' + pad(d.getMinutes());
  }
  function localInputToIso(s) {
    if (!s) return null;
    const d = new Date(s);
    return isNaN(d) ? null : d.toISOString();
  }

  // ── Reference data loaders ───────────────────────────────────────────────
  function loadReferenceData() {
    const calls = [
      ['Get', { typeName: 'Device', resultsLimit: 10000 }],
      ['Get', { typeName: 'Driver', resultsLimit: 10000 }]
    ];
    const myGen = ui.opGen;
    return apiMultiCall(calls, { gen: myGen }).then((results) => {
      if (isStale(myGen)) return;
      const devices = (results && results[0]) || [];
      const drivers = (results && results[1]) || [];
      // Skip cancellation placeholders.
      if (devices && devices.__cancelled) return;
      ui.deviceById.clear();
      ui.driverById.clear();
      (Array.isArray(devices) ? devices : []).forEach((d) => {
        if (d && d.id) ui.deviceById.set(d.id, d.name || d.id);
      });
      (Array.isArray(drivers) ? drivers : []).forEach((d) => {
        if (d && d.id) ui.driverById.set(d.id, [d.firstName, d.lastName].filter(Boolean).join(' ') || d.name || d.id);
      });
      populateSelect($('ftbe-device'), ui.deviceById, 'All devices');
      populateSelect($('ftbe-driver'), ui.driverById, 'All drivers');
    }).catch((err) => {
      if (isCancelled(err)) return;            // expected on blur/unload — not a failure
      console.warn('[fuelBulkEditor] reference data load failed', err);
    });
  }
  function populateSelect(sel, map, allLabel) {
    if (!sel) return;
    while (sel.options.length > 1) sel.remove(1);
    sel.options[0].text = allLabel;
    const items = Array.from(map.entries()).sort((a, b) => String(a[1]).localeCompare(String(b[1])));
    items.forEach(([id, name]) => {
      const opt = document.createElement('option');
      opt.value = id;
      opt.textContent = name;
      sel.appendChild(opt);
    });
  }

  // ── Fuel transactions: load ──────────────────────────────────────────────
  function loadTransactions() {
    const fromIso = localInputToIso($('ftbe-from').value);
    const toIso   = localInputToIso($('ftbe-to').value);
    if (!fromIso || !toIso) {
      setStatus('Pick a valid From and To date.', 'error');
      return;
    }
    const search = { fromDate: fromIso, toDate: toIso };
    const devId = $('ftbe-device').value;
    const drvId = $('ftbe-driver').value;
    if (devId) search.deviceSearch = { id: devId };
    if (drvId) search.driverSearch = { id: drvId };

    setStatus('Loading transactions…');
    ui.rows = [];
    ui.edited.clear();
    ui.selected.clear();
    ui.unmatchedPreview = [];
    ui.duplicateTargets.clear();
    renderTable();

    const myGen = ui.opGen;
    apiCall('Get', { typeName: 'FuelTransaction', search, resultsLimit: RESULTS_LIMIT })
      .then((rows) => {
        if (isStale(myGen)) return;          // user moved on; drop result
        ui.rows = Array.isArray(rows) ? rows : [];
        applySortAndRender();
        setStatus(ui.rows.length + ' transactions loaded.', 'success');
      })
      .catch((err) => {
        if (isStale(myGen) || isCancelled(err)) return;
        setStatus('Load failed: ' + (err && err.message ? err.message : err), 'error');
      });
  }

  // ── Filter / sort / render ───────────────────────────────────────────────
  // HAR-confirmed: r.driver may be the bare string "UnknownDriverId" OR an
  // object { id }. Normalise both cases so downstream code can rely on it.
  function driverIdOf(r) {
    if (!r || r.driver == null) return null;
    if (typeof r.driver === 'string') return r.driver;
    return r.driver.id || null;
  }
  function deviceIdOf(r) {
    if (!r || r.device == null) return null;
    if (typeof r.device === 'string') return r.device;
    return r.device.id || null;
  }

  function deriveDisplayRows() {
    const q = ($('ftbe-search').value || '').trim().toLowerCase();
    const arr = ui.rows.map((r) => {
      const devId = deviceIdOf(r);
      const drvId = driverIdOf(r);
      return {
        raw: r,
        id: r.id,
        version: r.version,
        dateTime: r.dateTime,
        deviceId: devId,
        deviceName: (devId && ui.deviceById.get(devId)) || r.description || '',
        driverId: drvId,
        driverName: r.driverName || (drvId && ui.driverById.get(drvId)) || '',
        productType: r.productType,
        volume: typeof r.volume === 'number' ? r.volume : null,
        cost: typeof r.cost === 'number' ? r.cost : null,
        currencyCode: r.currencyCode || '',
        odometer: typeof r.odometer === 'number' ? r.odometer : null,
        siteName: r.siteName || '',
        cardNumber: r.cardNumber || '',
        comments: r.comments || ''
      };
    });
    const filtered = q ? arr.filter((d) => {
      return [d.driverName, d.siteName, d.comments, d.cardNumber, d.deviceName]
        .some((v) => String(v).toLowerCase().indexOf(q) !== -1);
    }) : arr;
    const dir = ui.sortDir === 'asc' ? 1 : -1;
    const k = ui.sortKey;
    filtered.sort((a, b) => {
      const av = a[k], bv = b[k];
      if (av == null && bv == null) return 0;
      if (av == null) return 1;
      if (bv == null) return -1;
      if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * dir;
      return String(av).localeCompare(String(bv)) * dir;
    });
    return filtered;
  }

  // ── Change-visualization helpers ─────────────────────────────────────────
  // The visual vocabulary the table uses to communicate edits:
  //   amber cell + left accent  = value differs from server (pending edit)
  //   ↑ / ↓                     = numeric direction of change
  //   ✎                         = non-numeric change (text/enum/datetime)
  //   • (dot)                   = field is the SAME as server (no change staged
  //                               for this cell, even if other cells on the row changed)
  //   title= tooltip            = exact before → after values in raw units
  //   grey row + dashed border  = unmatched CSV row that needs Force Import
  //   amber row stripe          = "duplicate target" warning (≥2 CSV rows hit same FT)
  function isCellChanged(oldVal, newVal) {
    if (newVal == null) return false;
    if (oldVal == null) return true;
    if (typeof oldVal === 'number' && typeof newVal === 'number') {
      return Math.abs(oldVal - newVal) > 1e-9;
    }
    return String(oldVal) !== String(newVal);
  }
  function diffMarker(oldVal, newVal, isNumeric) {
    if (isNumeric && oldVal != null && newVal != null) {
      const a = Number(oldVal), b = Number(newVal);
      if (isFinite(a) && isFinite(b)) {
        if (b > a) return '↑';
        if (b < a) return '↓';
      }
    }
    return '✎';
  }
  // Build a single cell. If `oldRaw !== newRaw`, wrap in .cell-edit with a
  // marker + tooltip; otherwise return the plain escaped display value.
  function diffCell(oldRaw, newRaw, displayVal, opts) {
    opts = opts || {};
    const displayOld = opts.displayOld != null ? opts.displayOld : oldRaw;
    const changed = isCellChanged(oldRaw, newRaw);
    const safeDisplay = escapeHtml(displayVal == null ? '' : String(displayVal));
    if (!changed) return safeDisplay;
    const marker = diffMarker(oldRaw, newRaw, !!opts.numeric);
    const title  = 'was: ' + (oldRaw == null || oldRaw === '' ? '(empty)' : displayOld) +
                   '  →  now: ' + (newRaw == null || newRaw === '' ? '(empty)' : (displayVal == null ? newRaw : displayVal));
    return '<span class="cell-edit" title="' + escapeHtml(title) + '">' +
             safeDisplay +
             ' <span class="cell-edit__marker" aria-hidden="true">' + marker + '</span>' +
           '</span>';
  }

  // Render the small "unmatched CSV preview" rows pinned to the top of the
  // tbody so the user can see what would NOT be staged before they save.
  // These rows are read-only — they have no Geotab `id` yet (that's the whole
  // point) — so they don't participate in selection or bulk operations.
  function unmatchedPreviewRowsHtml() {
    if (!ui.unmatchedPreview.length) return '';
    return ui.unmatchedPreview.map((u, idx) => {
      const csv = u.csvRow || u;
      const reason = u.reason || '';
      const key = csv.vehicleIdentificationNumber || csv.serialNumber || csv.licencePlate || '(no key)';
      const dt = csv.dateTime ? fmtDateTime(csv.dateTime) : '';
      const vol = csv.volume != null && csv.volume !== '' ? csv.volume : '';
      const cost = csv.cost != null && csv.cost !== '' ? csv.cost : '';
      const odo = csv.odometer != null && csv.odometer !== '' ? csv.odometer : '';
      const tipParts = ['No existing FuelTransaction matched this CSV row.'];
      if (reason) tipParts.push(reason);
      tipParts.push('To add it as a new record, re-run import with “Force Import…” or use Geotab’s native Fuel Transactions import.');
      const tip = tipParts.join('\n');
      return '<tr class="is-unmatched-preview" data-pending-idx="' + idx + '" title="' + escapeHtml(tip) + '">' +
        '<td class="addin-col-check"><span class="cell-pending-badge" aria-label="Unmatched CSV row">+</span></td>' +
        '<td>' + escapeHtml(dt) + '</td>' +
        '<td colspan="2"><span class="cell-pending-label">UNMATCHED · ' + escapeHtml(key) + '</span></td>' +
        '<td>' + escapeHtml(csv.productType || '') + '</td>' +
        '<td class="addin-col-num">' + escapeHtml(vol) + '</td>' +
        '<td class="addin-col-num">' + escapeHtml(cost) + '</td>' +
        '<td>' + escapeHtml(csv.currencyCode || '') + '</td>' +
        '<td class="addin-col-num">' + escapeHtml(odo) + '</td>' +
        '<td>' + escapeHtml(csv.siteName || '') + '</td>' +
        '<td>' + escapeHtml(csv.cardNumber || '') + '</td>' +
        '<td><i>' + escapeHtml(reason) + '</i></td>' +
        '<td class="addin-col-actions">' +
          '<button type="button" class="addin-row-btn" data-action="dismiss-pending" data-pending-idx="' + idx + '">Dismiss</button>' +
        '</td>' +
      '</tr>';
    }).join('');
  }

  function renderTable() {
    const tbody = $('ftbe-tbody');
    if (!tbody) return;
    const rows = deriveDisplayRows();
    const unmatchedHtml = unmatchedPreviewRowsHtml();
    if (!rows.length && !unmatchedHtml) {
      tbody.innerHTML = '<tr><td colspan="13" class="addin-empty">No transactions match.</td></tr>';
      updateBulkButtons();
      updateSortIndicators();
      updateLegendStrip();
      return;
    }
    const html = rows.map((d) => {
      const isSel = ui.selected.has(d.id);
      const isEdited = ui.edited.has(d.id);
      const isDupTarget = (ui.duplicateTargets.get(d.id) || 0) > 1;
      const patch = ui.edited.get(d.id) || {};

      // Old (server) vs new (post-patch) raw values, by field.
      const newVol = patch.volume       != null ? patch.volume       : d.volume;
      const newCost = patch.cost        != null ? patch.cost         : d.cost;
      const newOdo  = patch.odometer    != null ? patch.odometer     : d.odometer;
      const newCur  = patch.currencyCode!= null ? patch.currencyCode : d.currencyCode;
      const newProd = patch.productType != null ? patch.productType  : d.productType;
      const newCmt  = patch.comments    != null ? patch.comments     : d.comments;
      const newDt   = patch.dateTime    != null ? patch.dateTime     : d.dateTime;

      const dtCell   = diffCell(d.dateTime, patch.dateTime != null ? newDt : null,
                                fmtDateTime(newDt),
                                { displayOld: fmtDateTime(d.dateTime) });
      const prodCell = diffCell(d.productType, patch.productType != null ? newProd : null, newProd || '');
      const volCell  = diffCell(d.volume, patch.volume != null ? newVol : null,
                                newVol != null ? fmtNum(toDisplayVolume(newVol), 3) : '',
                                { numeric: true,
                                  displayOld: d.volume != null ? fmtNum(toDisplayVolume(d.volume), 3) + ' ' + ui.volUnit : '(empty)' });
      const costCell = diffCell(d.cost, patch.cost != null ? newCost : null,
                                newCost != null ? fmtNum(newCost, 2) : '',
                                { numeric: true });
      const curCell  = diffCell(d.currencyCode, patch.currencyCode != null ? newCur : null, newCur || '');
      const odoCell  = diffCell(d.odometer, patch.odometer != null ? newOdo : null,
                                newOdo != null ? fmtNum(toDisplayOdo(newOdo), 0) : '',
                                { numeric: true,
                                  displayOld: d.odometer != null ? fmtNum(toDisplayOdo(d.odometer), 0) + ' ' + ui.odoUnit : '(empty)' });
      const cmtCell  = diffCell(d.comments, patch.comments != null ? newCmt : null, newCmt || '');

      const rowClasses = ['ftbe-row'];
      if (isSel) rowClasses.push('is-selected');
      if (isEdited) rowClasses.push('is-edited');
      if (isDupTarget) rowClasses.push('is-dup-target');

      const dupTip = isDupTarget
        ? ' title="Warning: ' + ui.duplicateTargets.get(d.id) + ' CSV rows matched this same FuelTransaction. Only the last set of edits will be kept."'
        : '';

      return '<tr data-id="' + escapeHtml(d.id) + '" class="' + rowClasses.join(' ') + '"' + dupTip + '>' +
        '<td class="addin-col-check"><input type="checkbox" class="ftbe-row-check" ' +
            (isSel ? 'checked' : '') + ' aria-label="Select row"></td>' +
        '<td>' + dtCell + '</td>' +
        '<td>' + escapeHtml(d.deviceName) + '</td>' +
        '<td>' + escapeHtml(d.driverName) + '</td>' +
        '<td>' + prodCell + '</td>' +
        '<td class="addin-col-num">' + volCell + '</td>' +
        '<td class="addin-col-num">' + costCell + '</td>' +
        '<td>' + curCell + '</td>' +
        '<td class="addin-col-num">' + odoCell + '</td>' +
        '<td>' + escapeHtml(d.siteName) + '</td>' +
        '<td>' + escapeHtml(d.cardNumber) + '</td>' +
        '<td>' + cmtCell + '</td>' +
        '<td class="addin-col-actions">' +
          '<button type="button" class="addin-row-btn" data-action="edit">Edit</button>' +
          (isEdited ? ' <button type="button" class="addin-row-btn" data-action="revert">Revert</button>' : '') +
          ' <button type="button" class="addin-row-btn addin-row-btn--danger" data-action="delete">Delete</button>' +
        '</td>' +
      '</tr>';
    }).join('');
    tbody.innerHTML = unmatchedHtml + html;
    updateBulkButtons();
    updateSortIndicators();
    updateLegendStrip();
    const dirtyCount = ui.edited.size;
    if (dirtyCount > 0) {
      setStatus(dirtyCount + ' pending edit(s). Click "Save edits" to commit.');
      ensureSaveEditsButton();
    } else {
      removeSaveEditsButton();
    }
  }

  // ── Legend strip ────────────────────────────────────────────────────────
  // Shown above the table whenever the table has anything non-trivial to
  // explain (pending edits, unmatched preview rows, duplicate-target warnings).
  // Anchored to the table wrapper so it scrolls with the toolbar, not the body.
  function updateLegendStrip() {
    const wrap = document.querySelector('.addin-table-wrap');
    if (!wrap) return;
    let strip = $('ftbe-legend');
    const dirty = ui.edited.size;
    const dupes = Array.from(ui.duplicateTargets.values()).filter((n) => n > 1).length;
    const unmatched = ui.unmatchedPreview.length;
    const need = dirty || dupes || unmatched;
    if (!need) { if (strip) strip.remove(); return; }
    if (!strip) {
      strip = document.createElement('div');
      strip.id = 'ftbe-legend';
      strip.className = 'addin-legend';
      wrap.parentNode.insertBefore(strip, wrap);
    }
    const parts = [];
    if (dirty) {
      parts.push(
        '<span class="addin-legend__group">' +
          '<span class="addin-legend__count">' + dirty + ' pending edit' + (dirty === 1 ? '' : 's') + '</span>' +
          '<span class="addin-legend__sample cell-edit">value <span class="cell-edit__marker">↑</span></span>' +
          '<span class="addin-legend__hint">amber cell = changed; <b>↑</b>/<b>↓</b> direction; <b>✎</b> text/enum. Hover a cell to see the previous value.</span>' +
        '</span>'
      );
    }
    if (dupes) {
      parts.push(
        '<span class="addin-legend__group addin-legend__group--warn">' +
          '<span class="addin-legend__count">' + dupes + ' duplicate target' + (dupes === 1 ? '' : 's') + '</span>' +
          '<span class="addin-legend__hint">Multiple CSV rows matched the same FuelTransaction. Only the last patch will survive.</span>' +
        '</span>'
      );
    }
    if (unmatched) {
      parts.push(
        '<span class="addin-legend__group addin-legend__group--ghost">' +
          '<span class="addin-legend__count">' + unmatched + ' unmatched CSV row' + (unmatched === 1 ? '' : 's') + '</span>' +
          '<span class="addin-legend__hint">No existing FuelTransaction matched. Use <b>Force Import…</b> to add them as new records.</span>' +
          '<button type="button" class="addin-row-btn" id="ftbe-legend-clear-unmatched">Clear preview</button>' +
        '</span>'
      );
    }
    strip.innerHTML = parts.join('');
    const clearBtn = $('ftbe-legend-clear-unmatched');
    if (clearBtn) clearBtn.addEventListener('click', () => {
      ui.unmatchedPreview = [];
      renderTable();
    });
  }

  function updateSortIndicators() {
    const ths = document.querySelectorAll('#ftbe-table thead th[data-sort]');
    ths.forEach((th) => {
      th.classList.remove('is-sort-asc', 'is-sort-desc');
      if (th.getAttribute('data-sort') === ui.sortKey) {
        th.classList.add(ui.sortDir === 'asc' ? 'is-sort-asc' : 'is-sort-desc');
      }
    });
  }

  function applySortAndRender() { renderTable(); }

  function updateBulkButtons() {
    const has = ui.selected.size > 0;
    $('ftbe-bulk-edit').disabled = !has;
    $('ftbe-bulk-delete').disabled = !has;
    const all = $('ftbe-check-all');
    if (all) {
      const total = deriveDisplayRows().length;
      all.checked = total > 0 && ui.selected.size === total;
      all.indeterminate = ui.selected.size > 0 && ui.selected.size < total;
    }
  }

  function ensureSaveEditsButton() {
    if ($('ftbe-save-edits')) return;
    const btn = document.createElement('button');
    btn.id = 'ftbe-save-edits';
    btn.className = 'btn btn--primary';
    btn.type = 'button';
    btn.textContent = 'Save edits';
    btn.addEventListener('click', saveAllEdits);
    const row = document.querySelector('.addin-toolbar__row--secondary');
    if (row) row.appendChild(btn);
  }
  function removeSaveEditsButton() {
    const btn = $('ftbe-save-edits');
    if (btn) btn.remove();
  }

  // ── Validation ───────────────────────────────────────────────────────────
  function sanitizePatch(patch) {
    // Whitelist enums + coerce numerics. Defence-in-depth (BUILD_GUIDE §7 / Phase 7).
    const out = {};
    for (const k of Object.keys(patch)) {
      if (EDITABLE_FIELDS.indexOf(k) === -1) continue;
      const v = patch[k];
      if (v === '' || v == null) continue;
      if (k === 'productType') {
        if (PRODUCT_TYPES.indexOf(v) === -1) continue;
        out.productType = v;
      } else if (k === 'currencyCode') {
        if (!/^[A-Z]{3}$/.test(v)) continue;
        out.currencyCode = v;
      } else if (k === 'volume' || k === 'cost' || k === 'odometer') {
        const n = Number(v);
        if (!isFinite(n) || n < 0) continue;
        out[k] = n;
      } else if (k === 'dateTime') {
        const d = new Date(v);
        if (isNaN(d)) continue;
        out.dateTime = d.toISOString();
      } else if (k === 'comments') {
        out.comments = String(v).slice(0, 1024);
      }
    }
    return out;
  }

  // ── Single-row edit modal ────────────────────────────────────────────────
  function openEditModal(id) {
    const row = ui.rows.find((r) => r.id === id);
    if (!row) return;
    const patch = ui.edited.get(id) || {};
    const cur = (k) => patch[k] != null ? patch[k] : row[k];
    const body = $('ftbe-modal-body');
    body.innerHTML =
      '<div class="addin-form-grid">' +
        '<label class="addin-field"><span>Date / time</span>' +
          '<input type="datetime-local" id="ftbe-edit-dateTime" value="' + escapeHtml(isoToLocalInput(cur('dateTime'))) + '"></label>' +
        '<label class="addin-field"><span>Product</span>' +
          '<select id="ftbe-edit-productType">' +
            PRODUCT_TYPES.map((p) => '<option value="' + escapeHtml(p) + '"' +
              (cur('productType') === p ? ' selected' : '') + '>' + escapeHtml(p) + '</option>').join('') +
          '</select></label>' +
        '<label class="addin-field"><span>Volume (' + escapeHtml(ui.volUnit) + ')</span>' +
          '<input type="number" step="0.001" min="0" id="ftbe-edit-volume" value="' +
          escapeHtml(cur('volume') != null ? fmtNum(toDisplayVolume(cur('volume')), 3) : '') + '"></label>' +
        '<label class="addin-field"><span>Cost</span>' +
          '<input type="number" step="0.01" min="0" id="ftbe-edit-cost" value="' +
          escapeHtml(cur('cost') != null ? fmtNum(cur('cost'), 2) : '') + '"></label>' +
        '<label class="addin-field"><span>Currency</span>' +
          '<select id="ftbe-edit-currencyCode">' +
            CURRENCIES.map((c) => '<option value="' + escapeHtml(c) + '"' +
              (cur('currencyCode') === c ? ' selected' : '') + '>' + escapeHtml(c) + '</option>').join('') +
          '</select></label>' +
        '<label class="addin-field"><span>Odometer (' + escapeHtml(ui.odoUnit) + ')</span>' +
          '<input type="number" step="1" min="0" id="ftbe-edit-odometer" value="' +
          escapeHtml(cur('odometer') != null ? fmtNum(toDisplayOdo(cur('odometer')), 0) : '') + '"></label>' +
        '<label class="addin-field full"><span>Comments</span>' +
          '<textarea id="ftbe-edit-comments" rows="3" maxlength="1024">' +
          escapeHtml(cur('comments') || '') + '</textarea></label>' +
      '</div>';
    $('ftbe-modal-title').textContent = 'Edit transaction';
    showModal(() => {
      const raw = {
        dateTime:     localInputToIso($('ftbe-edit-dateTime').value),
        productType:  $('ftbe-edit-productType').value,
        volume:       $('ftbe-edit-volume').value !== '' ? fromDisplayVolume(Number($('ftbe-edit-volume').value)) : '',
        cost:         $('ftbe-edit-cost').value !== '' ? Number($('ftbe-edit-cost').value) : '',
        currencyCode: $('ftbe-edit-currencyCode').value,
        odometer:     $('ftbe-edit-odometer').value !== '' ? fromDisplayOdo(Number($('ftbe-edit-odometer').value)) : '',
        comments:     $('ftbe-edit-comments').value
      };
      const clean = sanitizePatch(raw);
      const existing = ui.edited.get(id) || {};
      ui.edited.set(id, Object.assign({}, existing, clean));
      renderTable();
    });
  }

  // ── Bulk-edit modal ──────────────────────────────────────────────────────
  function openBulkEditModal() {
    if (!ui.selected.size) return;
    const body = $('ftbe-modal-body');
    body.innerHTML =
      '<p>Apply changes to <b>' + ui.selected.size + '</b> selected transaction(s). ' +
        'Leave a field blank to leave it unchanged.</p>' +
      '<div class="addin-form-grid">' +
        '<label class="addin-field"><span>Product</span>' +
          '<select id="ftbe-bulk-productType"><option value="">— unchanged —</option>' +
            PRODUCT_TYPES.map((p) => '<option value="' + escapeHtml(p) + '">' + escapeHtml(p) + '</option>').join('') +
          '</select></label>' +
        '<label class="addin-field"><span>Currency</span>' +
          '<select id="ftbe-bulk-currencyCode"><option value="">— unchanged —</option>' +
            CURRENCIES.map((c) => '<option value="' + escapeHtml(c) + '">' + escapeHtml(c) + '</option>').join('') +
          '</select></label>' +
        '<label class="addin-field"><span>Set volume (' + escapeHtml(ui.volUnit) + ')</span>' +
          '<input type="number" step="0.001" min="0" id="ftbe-bulk-volume"></label>' +
        '<label class="addin-field"><span>Set cost</span>' +
          '<input type="number" step="0.01" min="0" id="ftbe-bulk-cost"></label>' +
        '<label class="addin-field"><span>Set odometer (' + escapeHtml(ui.odoUnit) + ')</span>' +
          '<input type="number" step="1" min="0" id="ftbe-bulk-odometer"></label>' +
        '<label class="addin-field full"><span>Append to comments</span>' +
          '<input type="text" id="ftbe-bulk-comments-append" maxlength="512"></label>' +
      '</div>';
    $('ftbe-modal-title').textContent = 'Bulk edit ' + ui.selected.size + ' transaction(s)';
    showModal(() => {
      const raw = {
        productType:  $('ftbe-bulk-productType').value,
        currencyCode: $('ftbe-bulk-currencyCode').value,
        volume:       $('ftbe-bulk-volume').value !== '' ? fromDisplayVolume(Number($('ftbe-bulk-volume').value)) : '',
        cost:         $('ftbe-bulk-cost').value !== '' ? Number($('ftbe-bulk-cost').value) : '',
        odometer:     $('ftbe-bulk-odometer').value !== '' ? fromDisplayOdo(Number($('ftbe-bulk-odometer').value)) : ''
      };
      const append = ($('ftbe-bulk-comments-append').value || '').trim();
      const clean = sanitizePatch(raw);
      if (!Object.keys(clean).length && !append) return;
      ui.selected.forEach((id) => {
        const existing = ui.edited.get(id) || {};
        const row = ui.rows.find((r) => r.id === id) || {};
        const merged = Object.assign({}, existing, clean);
        if (append) {
          const base = existing.comments != null ? existing.comments : (row.comments || '');
          merged.comments = (base ? base + ' ' : '') + append;
          merged.comments = String(merged.comments).slice(0, 1024);
        }
        ui.edited.set(id, merged);
      });
      renderTable();
    });
  }

  // ── Save pending edits (multiCall, sequential, optimistic concurrency) ───
  function saveAllEdits() {
    if (!ui.edited.size) return;
    const entries = Array.from(ui.edited.entries());
    const calls = [];
    const callIdToTxId = [];
    for (const [id, patch] of entries) {
      const row = ui.rows.find((r) => r.id === id);
      if (!row) continue;
      const entity = Object.assign({}, row, sanitizePatch(patch));
      // Ensure required identity + concurrency fields:
      entity.id = row.id;
      entity.version = row.version;
      // Driver shape: Geotab returns the bare string "UnknownDriverId" when
      // no driver is assigned, and the object { id } when one is. Pass
      // through unchanged; default to "UnknownDriverId" if missing entirely.
      // (Do NOT wrap the unknown-driver string in an object — that's wrong.)
      if (entity.driver == null) entity.driver = 'UnknownDriverId';
      if (typeof entity.device === 'string') entity.device = { id: entity.device };
      // sourceData must be a string (preserve as-is or empty string).
      if (entity.sourceData != null && typeof entity.sourceData !== 'string') {
        try { entity.sourceData = JSON.stringify(entity.sourceData); } catch (_) { entity.sourceData = ''; }
      }
      if (!entity.id) continue;
      calls.push(['Set', { typeName: 'FuelTransaction', entity }]);
      callIdToTxId.push(id);
    }
    if (!calls.length) return;
    setStatus('Saving ' + calls.length + ' edit(s)…');
    const myGen = ui.opGen;
    apiMultiCall(calls, { label: 'Saving edits', gen: myGen }).then((results) => {
      if (isStale(myGen)) return;
      let ok = 0, fail = 0, cancelled = 0;
      const failures = [];
      results.forEach((res, idx) => {
        if (res && res.__cancelled) { cancelled++; }
        else if (res && res.__error) { fail++; failures.push({ id: callIdToTxId[idx], err: res.__error }); }
        else { ok++; ui.edited.delete(callIdToTxId[idx]); }
      });
      if (cancelled) {
        setStatus(ok + ' saved, ' + cancelled + ' not sent (cancelled).', 'error');
      } else if (fail) {
        setStatus(ok + ' saved, ' + fail + ' failed. Reloading to refresh versions…', 'error');
        console.warn('[fuelBulkEditor] save failures', failures);
      } else {
        setStatus(ok + ' edit(s) saved.', 'success');
      }
      // Re-Get so stale `version` tokens refresh (optimistic-concurrency hygiene).
      // Gate on gen — if blur happened during save, don't kick off a fresh load.
      if (!isStale(myGen)) loadTransactions();
    }).catch((err) => {
      if (isStale(myGen) || isCancelled(err)) return;
      setStatus('Save failed: ' + (err && err.message ? err.message : err), 'error');
    });
  }

  // ── Delete ───────────────────────────────────────────────────────────────
  function deleteRows(ids) {
    if (!ids.length) return;
    if (!confirm('Delete ' + ids.length + ' transaction(s)? This cannot be undone.')) return;
    const calls = ids.map((id) => ['Remove', { typeName: 'FuelTransaction', entity: { id } }]);
    setStatus('Deleting ' + calls.length + '…');
    const myGen = ui.opGen;
    apiMultiCall(calls, { label: 'Deleting', gen: myGen }).then((results) => {
      if (isStale(myGen)) return;
      const cancelled = results.filter((r) => r && r.__cancelled).length;
      const fail = results.filter((r) => r && r.__error).length;
      const ok = results.length - fail - cancelled;
      ui.selected.clear();
      ids.forEach((id) => ui.edited.delete(id));
      const msg = ok + ' deleted' +
        (fail ? (', ' + fail + ' failed') : '') +
        (cancelled ? (', ' + cancelled + ' not sent (cancelled)') : '');
      setStatus(msg, (fail || cancelled) ? 'error' : 'success');
      if (!isStale(myGen)) loadTransactions();
    }).catch((err) => {
      if (isStale(myGen) || isCancelled(err)) return;
      setStatus('Delete failed: ' + (err && err.message ? err.message : err), 'error');
    });
  }

  // ── CSV export ───────────────────────────────────────────────────────────
  // Mirrors the native Geotab "Fuel Transactions Import Template" exactly:
  // identical 16 headers, identical order, identical formatting. Round-trip
  // record identity is via VIN/Serial/Plate + dateTime on re-import — the
  // native template carries no `id` column either, and our matcher already
  // handles that path.
  const NATIVE_CSV_COLUMNS = Object.freeze([
    { key: 'Date & Time',                   header: 'Date & Time' },
    { key: 'Vehicle Identification Number', header: 'Vehicle Identification Number' },
    { key: 'Serial Number',                 header: 'Serial Number' },
    { key: 'License Plate',                 header: 'License Plate' },
    { key: 'Vehicle Description',           header: 'Vehicle Description' },
    { key: 'Cardholder',                    header: 'Cardholder' },
    { key: 'Card Number',                   header: 'Card Number' },
    { key: 'Volume (L)',                    header: 'Volume (L)' },
    { key: 'Cost',                          header: 'Cost' },
    { key: 'Currency Code',                 header: 'Currency Code' },
    { key: 'Product Type',                  header: 'Product Type' },
    { key: 'Transaction Odometer',          header: 'Transaction Odometer' },
    { key: 'Location Coordinates',          header: 'Location Coordinates' },
    { key: 'Location Address',              header: 'Location Address' },
    { key: 'Site Name',                     header: 'Site Name' },
    { key: 'Comments',                      header: 'Comments' }
  ]);

  // Local ISO 8601 without "Z" — matches the template example
  // "2026-02-01T14:30:00". Geotab interprets this column as tenant-local time
  // on import; emitting the same form keeps the round-trip lossless.
  function fmtNativeDateTime(iso) {
    if (!iso) return '';
    const d = new Date(iso);
    if (isNaN(d)) return '';
    const p = (n) => String(n).padStart(2, '0');
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) +
           'T' + p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds());
  }
  function fmtCoords(loc) {
    if (!loc || typeof loc !== 'object') return '';
    if (typeof loc.y !== 'number' || typeof loc.x !== 'number') return '';
    return loc.y + ',' + loc.x;     // "lat,lon" per template
  }
  function fmtPlainNumber(n) {
    return (n == null || n === '' || isNaN(n)) ? '' : String(Number(n));
  }

  function exportCsv() {
    if (!ui.rows.length) { setStatus('Nothing to export.', 'error'); return; }
    const data = (ui.selected.size ? ui.rows.filter((r) => ui.selected.has(r.id)) : ui.rows)
      .map((r) => ({
        'Date & Time':                   fmtNativeDateTime(r.dateTime),
        'Vehicle Identification Number': r.vehicleIdentificationNumber || '',
        'Serial Number':                 r.serialNumber || '',
        'License Plate':                 r.licencePlate || '',
        'Vehicle Description':           r.description || '',
        'Cardholder':                    r.driverName || '',
        'Card Number':                   r.cardNumber || '',
        'Volume (L)':                    fmtPlainNumber(r.volume),
        'Cost':                          fmtPlainNumber(r.cost),
        'Currency Code':                 r.currencyCode || '',
        'Product Type':                  r.productType || '',
        'Transaction Odometer':          fmtPlainNumber(r.odometer),
        'Location Coordinates':          fmtCoords(r.location),
        'Location Address':              '',                      // not stored on FuelTransaction entity
        'Site Name':                     r.siteName || '',
        'Comments':                      r.comments || ''
      }));
    const text = window.CSVUtil.serialize(data, NATIVE_CSV_COLUMNS);
    const blob = new Blob([text], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    a.download = 'fuel-transactions-' + stamp + '.csv';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 0);
    setStatus('Exported ' + data.length + ' row(s) (native Geotab template).', 'success');
  }

  // ── CSV import (bulk-edit upload) ────────────────────────────────────────
  function onImportFileChosen(file, mode) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const text = String(reader.result || '');
        const parsed = window.CSVUtil.parse(text);
        applyImport(parsed, mode);
      } catch (err) {
        setStatus('CSV parse failed: ' + (err && err.message ? err.message : err), 'error');
      }
    };
    reader.onerror = () => setStatus('Could not read CSV file.', 'error');
    reader.readAsText(file);
  }

  // External-format header aliases (case-insensitive). Maps third-party
  // fuel-card CSV columns to FuelTransaction field names.
  const EXTERNAL_HEADER_ALIASES = Object.freeze({
    'date & time':                    'dateTime',
    'datetime':                       'dateTime',
    'vehicle identification number':  'vehicleIdentificationNumber',
    'vin':                            'vehicleIdentificationNumber',
    'serial number':                  'serialNumber',
    'license plate':                  'licencePlate',
    'licence plate':                  'licencePlate',
    'vehicle description':            'description',
    'cardholder':                     'driverName',
    'card number':                    'cardNumber',
    'volume (l)':                     'volume',
    'volume':                         'volume',
    'cost':                           'cost',
    'currency code':                  'currencyCode',
    'product type':                   'productType',
    'transaction odometer':           'odometer',
    'odometer':                       'odometer',
    'location coordinates':           'locationCoordinates',
    'location address':               'siteAddress',
    'site name':                      'siteName',
    'comments':                       'comments'
  });

  // Lightweight productType synonym map. Anything else falls through to
  // sanitizePatch's whitelist (and gets rejected if not a known enum).
  const PRODUCT_TYPE_SYNONYMS = Object.freeze({
    'gasoline':         'Regular',
    'gas':              'Regular',
    'unleaded':         'Regular',
    'regular unleaded': 'Regular',
    'premium gasoline': 'Premium',
    'mid-grade':        'Midgrade',
    'midgrade':         'Midgrade',
    'diesel fuel':      'Diesel',
    'def':              'DieselExhaustFluid'
  });

  function stripCommas(s) { return typeof s === 'string' ? s.replace(/,/g, '') : s; }

  // Normalises a row from an external CSV into FuelTransaction-shaped fields.
  function normalizeExternalRow(headers, row) {
    const out = {};
    for (const h of headers) {
      const key = EXTERNAL_HEADER_ALIASES[h.toLowerCase()];
      if (!key) continue;
      let v = row[h];
      if (v == null || v === '') continue;
      if (key === 'volume' || key === 'cost' || key === 'odometer') v = stripCommas(v);
      if (key === 'productType') {
        const syn = PRODUCT_TYPE_SYNONYMS[String(v).toLowerCase()];
        if (syn) v = syn;
      }
      out[key] = v;
    }
    return out;
  }

  // The native Geotab Fuel Transactions Import Template uses VIN/Serial/Plate
  // as record-identity keys (no `id` column). We mirror that layout exactly,
  // so all imports flow through the VIN/Serial/Plate matcher.
  function isNativeTemplate(headers) {
    const lc = headers.map((h) => h.toLowerCase());
    return [
      'vehicle identification number', 'vin',
      'serial number', 'license plate', 'licence plate'
    ].some((h) => lc.indexOf(h) !== -1);
  }

  // Match CSV rows against loaded FuelTransactions by VIN → Serial → Plate,
  // narrowed by dateTime delta (default ±5 min). Returns {matched, unmatched, ambiguous}.
  function matchExternalRows(externalRows, fuelTx, opts) {
    const toleranceMs = (opts && opts.toleranceMinutes ? opts.toleranceMinutes : 5) * 60 * 1000;

    // Build indices once for O(1) candidate narrowing.
    const byVin    = new Map();
    const bySerial = new Map();
    const byPlate  = new Map();
    const pushTo = (map, key, tx) => {
      if (!key) return;
      const k = String(key).toUpperCase().trim();
      if (!k) return;
      if (!map.has(k)) map.set(k, []);
      map.get(k).push(tx);
    };
    fuelTx.forEach((tx) => {
      pushTo(byVin,    tx.vehicleIdentificationNumber, tx);
      pushTo(bySerial, tx.serialNumber,                tx);
      pushTo(byPlate,  tx.licencePlate,                tx);
    });

    const matched = [];     // [{ csvRow, tx, reason }]
    const unmatched = [];   // [{ csvRow, reason }]
    const ambiguous = [];   // [{ csvRow, candidates }]

    externalRows.forEach((csvRow) => {
      const csvIsoSrc = csvRow.dateTime;
      const csvIso = csvIsoSrc ? new Date(csvIsoSrc) : null;
      if (csvIso && isNaN(csvIso)) { unmatched.push({ csvRow, reason: 'bad dateTime' }); return; }

      // Returns one of:
      //   { tx, reason, deltaMs, candidateCount, keyLabel }   — confident pick
      //   { __amb: [...candidates], keyLabel }                — no timestamp, >1 candidate
      //   { __outOfWindow: { nearestMs, nearestTx, keyLabel } } — same VIN/Serial/Plate
      //     exists but none within tolerance (useful for the unmatched summary)
      //   null                                                — no candidates at all
      const tryKey = (map, raw, label) => {
        if (!raw) return null;
        const k = String(raw).toUpperCase().trim();
        const candidates = map.get(k) || [];
        if (!candidates.length) return null;
        if (!csvIso) {
          return candidates.length === 1
            ? { tx: candidates[0], reason: label + ' (no dateTime; unique)', deltaMs: null, candidateCount: 1, keyLabel: label }
            : { __amb: candidates, keyLabel: label };
        }
        const withDeltas = candidates
          .filter((tx) => tx.dateTime)
          .map((tx) => ({ tx, dms: Math.abs(new Date(tx.dateTime) - csvIso) }))
          .sort((a, b) => a.dms - b.dms);
        if (!withDeltas.length) return null;
        const inWindow = withDeltas.filter((x) => x.dms <= toleranceMs);
        if (!inWindow.length) {
          return { __outOfWindow: { nearestMs: withDeltas[0].dms, nearestTx: withDeltas[0].tx, keyLabel: label } };
        }
        if (inWindow.length === 1) {
          return { tx: inWindow[0].tx, reason: label, deltaMs: inWindow[0].dms, candidateCount: 1, keyLabel: label };
        }
        return {
          tx: inWindow[0].tx,
          reason: label + ' (closest of ' + inWindow.length + ')',
          deltaMs: inWindow[0].dms,
          candidateCount: inWindow.length,
          keyLabel: label
        };
      };

      const tries = [
        tryKey(byVin,    csvRow.vehicleIdentificationNumber, 'VIN'),
        tryKey(bySerial, csvRow.serialNumber,                'Serial'),
        tryKey(byPlate,  csvRow.licencePlate,                'Plate')
      ];
      const hit = tries.find((t) => t && t.tx);
      if (hit) {
        matched.push({
          csvRow, tx: hit.tx, reason: hit.reason,
          deltaMs: hit.deltaMs, candidateCount: hit.candidateCount, keyLabel: hit.keyLabel
        });
        return;
      }
      const amb = tries.find((t) => t && t.__amb);
      if (amb) { ambiguous.push({ csvRow, candidates: amb.__amb, keyLabel: amb.keyLabel }); return; }
      const oow = tries.find((t) => t && t.__outOfWindow);
      const reason = oow
        ? 'nearest ' + oow.__outOfWindow.keyLabel + ' match was ' + fmtDeltaMs(oow.__outOfWindow.nearestMs) + ' away (outside ±' + (toleranceMs / 60000) + ' min)'
        : 'no VIN/Serial/Plate found in loaded window';
      unmatched.push({
        csvRow, reason,
        nearestMissMs: oow ? oow.__outOfWindow.nearestMs : null,
        nearestMissKey: oow ? oow.__outOfWindow.keyLabel : null
      });
    });
    return { matched, unmatched, ambiguous };
  }

  // Humanise a millisecond delta for the match summary ("42s", "3m 12s", "2h 5m", "4d").
  function fmtDeltaMs(ms) {
    if (ms == null || !isFinite(ms)) return '?';
    const s = Math.round(ms / 1000);
    if (s < 60) return s + 's';
    const m = Math.round(s / 60);
    if (m < 60) return (s % 60) ? Math.floor(s / 60) + 'm ' + (s % 60) + 's' : m + 'm';
    const h = Math.floor(m / 60);
    if (h < 24) return (m % 60) ? h + 'h ' + (m % 60) + 'm' : h + 'h';
    const d = Math.floor(h / 24);
    return (h % 24) ? d + 'd ' + (h % 24) + 'h' : d + 'd';
  }

  function applyImport(parsed, mode) {
    if (!parsed || !parsed.headers || !parsed.headers.length) {
      setStatus('CSV appears empty.', 'error'); return;
    }
    if (!isNativeTemplate(parsed.headers)) {
      setStatus('CSV does not match the Geotab Fuel Transactions Import Template. Need at least one of: Vehicle Identification Number, Serial Number, License Plate.', 'error');
      return;
    }
    if (mode === 'force-add') return applyForceImport(parsed);
    return applyExternalImport(parsed, mode || 'stage-edits');
  }

  // Force Import — Add every CSV row as a NEW FuelTransaction. Bypasses the
  // VIN/Serial/Plate matcher.
  //
  // Entity fields populated here are limited to what the Geotab FuelTransaction
  // entity actually documents (geotab.com/sdk → FuelTransaction). We do NOT
  // mirror Drive-specific quirks observed in HAR captures — e.g. synthesising
  // `description: "eds_account-<serial>"` (a Drive internal device marker, not
  // a vehicle description) or stamping `sourceData` with a stringified clone
  // of the entity itself (sourceData per the spec is the raw upstream payload
  // from a fuel-card provider, not a self-referential blob). Both would inject
  // garbage into the audit trail. Only set a field if the CSV supplies a real
  // value for it.
  function buildAddEntity(csvRow) {
    const entity = {};
    if (csvRow.dateTime) {
      const d = new Date(csvRow.dateTime);
      if (!isNaN(d)) entity.dateTime = d.toISOString();
    }
    const vin    = csvRow.vehicleIdentificationNumber ? String(csvRow.vehicleIdentificationNumber).trim() : '';
    const serial = csvRow.serialNumber ? String(csvRow.serialNumber).trim() : '';
    const plate  = csvRow.licencePlate ? String(csvRow.licencePlate).trim() : '';
    if (vin)    entity.vehicleIdentificationNumber = vin;
    if (serial) entity.serialNumber = serial;
    if (plate)  entity.licencePlate = plate;

    if (csvRow.description) entity.description = String(csvRow.description);
    if (csvRow.driverName)  entity.driverName  = String(csvRow.driverName);
    if (csvRow.cardNumber)  entity.cardNumber  = String(csvRow.cardNumber);

    const volNum  = Number(stripCommas(csvRow.volume));
    const costNum = Number(stripCommas(csvRow.cost));
    const odoNum  = Number(stripCommas(csvRow.odometer));
    if (isFinite(volNum)  && volNum  >= 0) entity.volume   = volNum;
    if (isFinite(costNum) && costNum >= 0) entity.cost     = costNum;
    if (isFinite(odoNum)  && odoNum  >= 0) entity.odometer = odoNum;

    if (csvRow.currencyCode && /^[A-Z]{3}$/.test(String(csvRow.currencyCode).trim())) {
      entity.currencyCode = String(csvRow.currencyCode).trim();
    }
    if (csvRow.productType) {
      const syn = PRODUCT_TYPE_SYNONYMS[String(csvRow.productType).toLowerCase()];
      const pt  = syn || csvRow.productType;
      if (PRODUCT_TYPES.indexOf(pt) !== -1) entity.productType = pt;
    }

    if (csvRow.locationCoordinates) {
      const parts = String(csvRow.locationCoordinates).split(',').map((s) => Number(String(s).trim()));
      if (parts.length === 2 && isFinite(parts[0]) && isFinite(parts[1])) {
        entity.location = { y: parts[0], x: parts[1] };   // y=lat, x=lon
      }
    }
    if (csvRow.siteName) entity.siteName = String(csvRow.siteName);
    if (csvRow.comments) entity.comments = String(csvRow.comments).slice(0, 1024);

    // `provider` is a free-form attribution string. Stamp our add-in so the
    // audit trail correctly identifies the source.
    entity.provider = 'FuelTransactionsBulkEditor';
    return entity;
  }

  function applyForceImport(parsed) {
    const externals = parsed.rows
      .map((row) => normalizeExternalRow(parsed.headers, row))
      .filter((r) => r.vehicleIdentificationNumber || r.serialNumber || r.licencePlate);
    if (!externals.length) {
      setStatus('No usable rows (need VIN, Serial, or Plate on every row).', 'error');
      return;
    }
    const rejected = [];
    const adds = [];
    externals.forEach((csvRow, idx) => {
      const entity = buildAddEntity(csvRow);
      if (!entity.dateTime) { rejected.push({ idx, reason: 'missing or unparseable Date & Time' }); return; }
      if (!entity.vehicleIdentificationNumber && !entity.serialNumber && !entity.licencePlate) {
        rejected.push({ idx, reason: 'missing VIN / Serial / Plate' }); return;
      }
      adds.push({ csvRow, entity });
    });
    if (!adds.length) {
      setStatus('Force Import: every row was rejected (need Date & Time + VIN/Serial/Plate).', 'error');
      return;
    }
    const msg = 'Force Import will ADD ' + adds.length + ' NEW FuelTransaction(s) to your database.\n\n' +
                (rejected.length ? rejected.length + ' row(s) will be skipped (missing required fields).\n\n' : '') +
                'This bypasses the VIN/Serial/Plate matcher and cannot be undone except by Delete. Continue?';
    if (!confirm(msg)) { setStatus('Force Import cancelled.'); return; }

    const calls = adds.map(({ entity }) => ['Add', { typeName: 'FuelTransaction', entity }]);
    setStatus('Force-adding ' + calls.length + ' transaction(s)…');
    const myGen = ui.opGen;
    apiMultiCall(calls, { label: 'Force-adding', gen: myGen }).then((results) => {
      if (isStale(myGen)) return;
      const failures = [];
      let ok = 0, cancelled = 0;
      results.forEach((res, idx) => {
        if (res && res.__cancelled) cancelled++;
        else if (res && res.__error) failures.push({ csvRow: adds[idx].csvRow, err: res.__error });
        else ok++;
      });
      const fail = failures.length;
      if (cancelled) {
        setStatus(ok + ' added, ' + cancelled + ' not sent (cancelled), ' + fail + ' failed.', 'error');
      } else if (fail) {
        setStatus(ok + ' added, ' + fail + ' failed. Reloading…', 'error');
        console.warn('[fuelBulkEditor] force-import failures', failures);
      } else {
        setStatus(ok + ' transaction(s) added.', 'success');
      }
      showForceImportSummary({ added: ok, failed: fail, skipped: rejected.length, cancelled, failures });
      if (!isStale(myGen)) loadTransactions();
    }).catch((err) => {
      if (isStale(myGen) || isCancelled(err)) return;
      setStatus('Force Import failed: ' + (err && err.message ? err.message : err), 'error');
    });
  }

  function showForceImportSummary(s) {
    $('ftbe-modal-title').textContent = 'Force Import summary';
    const body = $('ftbe-modal-body');
    const failHtml = s.failures.slice(0, 25).map((f) => {
      const id = f.csvRow.vehicleIdentificationNumber || f.csvRow.serialNumber || f.csvRow.licencePlate || '(no key)';
      const msg = (f.err && f.err.message) ? f.err.message : String(f.err || 'unknown');
      return '<li class="addin-import-error">' + escapeHtml(id) + ' — ' + escapeHtml(msg) + '</li>';
    }).join('');
    body.innerHTML =
      '<p class="addin-import-summary">' +
        '<b>' + s.added + '</b> added, ' +
        '<b>' + s.failed + '</b> failed, ' +
        '<b>' + s.skipped + '</b> skipped (missing required fields)' +
        (s.cancelled ? ', <b>' + s.cancelled + '</b> not sent (cancelled)' : '') +
        '.</p>' +
      (failHtml
        ? '<h3 style="font-size:13px;margin:10px 0 4px">Failures (server-rejected)</h3>' +
          '<ul class="addin-import-summary">' + failHtml + '</ul>'
        : '');
    showModal(null, true);
  }

  // External-format import: match by VIN/Serial/Plate + dateTime, then stage
  // edits against the Geotab-assigned `id` + `version`. Server actions still
  // happen via chunked sequential multiCall.
  // mode:
  //   'stage-edits' — stage CSV values as pending edits on matched rows (default)
  //   'match-only'  — just select matched rows; do not touch ui.edited
  function applyExternalImport(parsed, mode) {
    mode = mode || 'stage-edits';
    const externals = parsed.rows
      .map((row) => normalizeExternalRow(parsed.headers, row))
      .filter((r) => r.vehicleIdentificationNumber || r.serialNumber || r.licencePlate);
    if (!externals.length) {
      setStatus('No usable rows (need VIN, Serial, or Plate).', 'error'); return;
    }
    // Derive a date window (±1 day padding) from CSV dateTime values, if any.
    const validDates = externals.map((r) => r.dateTime ? new Date(r.dateTime) : null).filter((d) => d && !isNaN(d));
    let needFetch = false;
    let fromIso, toIso;
    if (validDates.length) {
      const minMs = Math.min.apply(null, validDates.map((d) => d.getTime())) - 24 * 3600 * 1000;
      const maxMs = Math.max.apply(null, validDates.map((d) => d.getTime())) + 24 * 3600 * 1000;
      fromIso = new Date(minMs).toISOString();
      toIso   = new Date(maxMs).toISOString();
      needFetch = true;
    }

    // HAR-confirmed: FuelTransactionSearch accepts a top-level
    // vehicleIdentificationNumber filter server-side (undocumented but used
    // by the Drive App). Prefer per-VIN multiCall when the CSV has VINs —
    // far less data than a wide date-window pull.
    const uniqueVins = Array.from(new Set(
      externals.map((r) => (r.vehicleIdentificationNumber || '').trim()).filter(Boolean)
    ));
    const canUseServerVinSearch = uniqueVins.length > 0 && uniqueVins.length <= 100 && needFetch;

    const finishMatch = () => {
      const result = matchExternalRows(externals, ui.rows);
      ui.selected.clear();
      ui.duplicateTargets.clear();
      // Track how many CSV rows landed on the same tx — flagged in the table
      // (amber row stripe) AND surfaced as a dedicated warning section in the
      // summary modal, because "last write wins" is silently destructive.
      result.matched.forEach(({ tx }) => {
        ui.duplicateTargets.set(tx.id, (ui.duplicateTargets.get(tx.id) || 0) + 1);
      });
      // Compute per-row "fields that will actually change" for the summary
      // badges. We attach to the match record so showExternalMatchSummary
      // doesn't re-walk the data.
      result.matched.forEach((m) => {
        const csvRow = m.csvRow, tx = m.tx;
        const candidate = sanitizePatch({
          dateTime:     csvRow.dateTime,
          productType:  csvRow.productType,
          volume:       csvRow.volume,
          cost:         csvRow.cost,
          currencyCode: csvRow.currencyCode,
          odometer:     csvRow.odometer,
          comments:     csvRow.comments
        });
        const changes = [];
        for (const k of Object.keys(candidate)) {
          if (!isCellChanged(tx[k], candidate[k])) continue;
          changes.push({ field: k, oldVal: tx[k], newVal: candidate[k] });
        }
        m.changes = changes;
      });
      result.matched.forEach(({ csvRow, tx, changes }) => {
        if (mode === 'stage-edits' && changes && changes.length) {
          const patch = {};
          changes.forEach((c) => { patch[c.field] = c.newVal; });
          const existing = ui.edited.get(tx.id) || {};
          ui.edited.set(tx.id, Object.assign({}, existing, patch));
        }
        ui.selected.add(tx.id);   // both modes select; match-only stops here
      });
      // Stash unmatched rows so the table can preview them in grey above the
      // real data — answers "what is the CSV trying to add that I'm missing?"
      // without forcing the user back to the modal.
      ui.unmatchedPreview = result.unmatched.slice();
      renderTable();
      showExternalMatchSummary(result, mode);
    };

    const myGen = ui.opGen;

    if (canUseServerVinSearch) {
      setStatus('Fetching FuelTransactions by VIN (' + uniqueVins.length + ')…');
      const calls = uniqueVins.map((vin) => ['Get', {
        typeName: 'FuelTransaction',
        search: { fromDate: fromIso, toDate: toIso, vehicleIdentificationNumber: vin },
        resultsLimit: 10000
      }]);
      apiMultiCall(calls, { label: 'Fetching by VIN', gen: myGen }).then((results) => {
        if (isStale(myGen)) return;
        const cancelled = results.filter((r) => r && r.__cancelled).length;
        if (cancelled === results.length) {
          setStatus('CSV match cancelled before any data returned.'); return;
        }
        const merged = [];
        const seen = new Set();
        results.forEach((bucket) => {
          if (!bucket || bucket.__error || bucket.__cancelled) return;
          (bucket || []).forEach((tx) => {
            if (tx && tx.id && !seen.has(tx.id)) { seen.add(tx.id); merged.push(tx); }
          });
        });
        ui.rows = merged;
        if (fromIso) $('ftbe-from').value = isoToLocalInput(fromIso);
        if (toIso)   $('ftbe-to').value   = isoToLocalInput(toIso);
        if (cancelled) setStatus(cancelled + ' VIN bucket(s) were not fetched (cancelled). Match may be partial.', 'error');
        finishMatch();
      }).catch((err) => {
        if (isStale(myGen) || isCancelled(err)) return;
        setStatus('Per-VIN fetch failed: ' + (err && err.message ? err.message : err), 'error');
      });
    } else if (needFetch) {
      setStatus('Fetching FuelTransactions over CSV date window…');
      apiCall('Get', {
        typeName: 'FuelTransaction',
        search: { fromDate: fromIso, toDate: toIso },
        resultsLimit: RESULTS_LIMIT
      }).then((rows) => {
        if (isStale(myGen)) return;
        ui.rows = Array.isArray(rows) ? rows : [];
        if (fromIso) $('ftbe-from').value = isoToLocalInput(fromIso);
        if (toIso)   $('ftbe-to').value   = isoToLocalInput(toIso);
        finishMatch();
      }).catch((err) => {
        if (isStale(myGen) || isCancelled(err)) return;
        setStatus('Fetch failed: ' + (err && err.message ? err.message : err), 'error');
      });
    } else {
      // No timestamps in CSV — match against whatever's already loaded.
      if (!ui.rows.length) {
        setStatus('CSV has no dateTime values; load a date range first so the script has records to match against.', 'error');
        return;
      }
      finishMatch();
    }
  }

  // Compact, human-first match summary.
  //
  // The previous version dumped opaque Geotab IDs and labels like
  // "VIN (closest of 7)" with no field-level diff and no duplicate-target
  // warning. This version is organised as:
  //   1. Headline counts (legend, not just numbers)
  //   2. Duplicate-target warnings — most destructive footgun, surfaced first
  //   3. Matches — each row shows the VIN + CSV timestamp the user can
  //      recognise, a match-quality chip, and a per-field change list
  //   4. Unmatched — grouped block with a hint pointing to Force Import
  //   5. Ambiguous — needs CSV dateTime to disambiguate
  function showExternalMatchSummary(result, mode) {
    $('ftbe-modal-title').textContent = (mode === 'match-only') ? 'CSV lookup summary' : 'CSV match summary';
    const body = $('ftbe-modal-body');
    // Widen the modal panel for this view — the row format wants more room.
    const panel = document.querySelector('#ftbe-modal .addin-modal__panel');
    if (panel) panel.classList.add('addin-modal__panel--wide');

    // ── Group matches by target tx.id to detect duplicate targets ──────────
    const targets = new Map();
    result.matched.forEach((m) => {
      if (!targets.has(m.tx.id)) targets.set(m.tx.id, []);
      targets.get(m.tx.id).push(m);
    });
    const dupGroups = Array.from(targets.values()).filter((g) => g.length > 1);

    // ── Field abbreviations + value formatters for the change badge ────────
    const FIELD_LABELS = {
      dateTime: 'date', productType: 'product', volume: 'vol', cost: 'cost',
      currencyCode: 'cur', odometer: 'odo', comments: 'cmt'
    };
    function fmtFieldValue(field, v) {
      if (v == null || v === '') return '(empty)';
      if (field === 'volume')   return fmtNum(toDisplayVolume(Number(v)), 3) + ' ' + ui.volUnit;
      if (field === 'odometer') return fmtNum(toDisplayOdo(Number(v)), 0)    + ' ' + ui.odoUnit;
      if (field === 'cost')     return fmtNum(Number(v), 2);
      if (field === 'dateTime') return fmtDateTime(v);
      if (field === 'comments') return String(v).length > 40 ? String(v).slice(0, 40) + '…' : String(v);
      return String(v);
    }
    function changesBadge(changes) {
      if (!changes || !changes.length) return '<span class="match-badge match-badge--noop" title="CSV values match the existing record — nothing to stage">no change</span>';
      const tipLines = changes.map((c) =>
        FIELD_LABELS[c.field] + ': ' + fmtFieldValue(c.field, c.oldVal) + '  →  ' + fmtFieldValue(c.field, c.newVal)
      );
      const fieldList = changes.map((c) => FIELD_LABELS[c.field]).join(', ');
      return '<span class="match-badge match-badge--change" title="' + escapeHtml(tipLines.join('\n')) + '">' +
                changes.length + ' change' + (changes.length === 1 ? '' : 's') + ': ' + escapeHtml(fieldList) +
             '</span>';
    }
    function csvKey(csv) {
      return csv.vehicleIdentificationNumber || csv.serialNumber || csv.licencePlate || '(no key)';
    }
    function csvTime(csv) {
      return csv.dateTime ? fmtDateTime(csv.dateTime) : '(no time)';
    }
    function matchChip(m) {
      const label = m.keyLabel || 'match';
      const delta = (m.deltaMs != null) ? ' · Δ ' + fmtDeltaMs(m.deltaMs) : '';
      const fanout = (m.candidateCount && m.candidateCount > 1)
        ? ' · 1 of ' + m.candidateCount + ' in window'
        : '';
      const tip = label + ' key matched' + delta + fanout;
      return '<span class="match-chip" title="' + escapeHtml(tip) + '">' +
        escapeHtml(label) + escapeHtml(delta) + escapeHtml(fanout) + '</span>';
    }

    // ── Headline ───────────────────────────────────────────────────────────
    const totalChanges = result.matched.reduce((n, m) => n + ((m.changes && m.changes.length) || 0), 0);
    const headerText = (mode === 'match-only')
      ? 'Matched rows are <b>selected</b>. No edits were staged and nothing was sent to Geotab.'
      : 'Matched rows are now <b>selected and staged with edits</b>. Click <b>Save edits</b> ' +
        'to commit via the Geotab-assigned <code>id</code> + <code>version</code>.';

    // ── Section: duplicate-target warnings ────────────────────────────────
    const MAX_DUP_GROUPS = 15;
    const dupHtml = dupGroups.slice(0, MAX_DUP_GROUPS).map((group) => {
      const tx = group[0].tx;
      const head = escapeHtml(group.length + ' CSV rows → same FuelTransaction (VIN ' +
                              (tx.vehicleIdentificationNumber || tx.serialNumber || tx.licencePlate || '?') + ')');
      const items = group.map((m) =>
        '<li>' + escapeHtml(csvTime(m.csvRow)) + ' · ' + escapeHtml(csvKey(m.csvRow)) +
          ' — ' + (m.changes && m.changes.length
            ? escapeHtml(m.changes.length + ' change(s)')
            : '<i>no change</i>') +
        '</li>'
      ).join('');
      return '<div class="match-dupgroup">' +
        '<div class="match-dupgroup__head">' + head + ' — <b>only the last set of edits will survive.</b></div>' +
        '<ul class="match-dupgroup__list">' + items + '</ul>' +
      '</div>';
    }).join('');
    const dupOverflow = dupGroups.length > MAX_DUP_GROUPS
      ? '<p class="addin-import-summary"><i>… and ' + (dupGroups.length - MAX_DUP_GROUPS) + ' more duplicate group(s) not shown.</i></p>'
      : '';

    // ── Section: matches ─────────────────────────────────────────────────
    const MAX_ROWS = 50;
    const matchedRows = result.matched.slice(0, MAX_ROWS).map((m) =>
      '<li class="match-row">' +
        '<span class="match-row__key">' +
          '<span class="match-row__vin">' + escapeHtml(csvKey(m.csvRow)) + '</span>' +
          '<span class="match-row__time">' + escapeHtml(csvTime(m.csvRow)) + '</span>' +
        '</span>' +
        '<span class="match-row__meta">' +
          matchChip(m) +
          (mode === 'stage-edits' ? ' ' + changesBadge(m.changes) : '') +
        '</span>' +
      '</li>'
    ).join('');
    const matchedOverflow = result.matched.length > MAX_ROWS
      ? '<p class="addin-import-summary"><i>… and ' + (result.matched.length - MAX_ROWS) + ' more match(es) not shown — they are selected in the table.</i></p>'
      : '';

    // ── Section: unmatched ───────────────────────────────────────────────
    const unmatchedRows = result.unmatched.slice(0, MAX_ROWS).map((u) =>
      '<li class="match-row match-row--unmatched">' +
        '<span class="match-row__key">' +
          '<span class="match-row__vin">' + escapeHtml(csvKey(u.csvRow)) + '</span>' +
          '<span class="match-row__time">' + escapeHtml(csvTime(u.csvRow)) + '</span>' +
        '</span>' +
        '<span class="match-row__meta"><i>' + escapeHtml(u.reason) + '</i></span>' +
      '</li>'
    ).join('');
    const unmatchedOverflow = result.unmatched.length > MAX_ROWS
      ? '<p class="addin-import-summary"><i>… and ' + (result.unmatched.length - MAX_ROWS) + ' more unmatched row(s) not shown — preview them in the table.</i></p>'
      : '';

    // ── Section: ambiguous ───────────────────────────────────────────────
    const ambHtml = result.ambiguous.slice(0, 20).map((a) =>
      '<li>' + escapeHtml(csvKey(a.csvRow)) +
        ' — ' + a.candidates.length + ' candidates (need CSV dateTime to disambiguate)</li>'
    ).join('');

    body.innerHTML =
      // Headline
      '<div class="match-headline">' +
        '<div class="match-counts">' +
          '<span class="match-count match-count--ok">'   + result.matched.length   + ' matched</span>' +
          '<span class="match-count match-count--warn">' + dupGroups.length        + ' duplicate target' + (dupGroups.length === 1 ? '' : 's') + '</span>' +
          '<span class="match-count match-count--miss">' + result.unmatched.length + ' unmatched</span>' +
          '<span class="match-count match-count--amb">'  + result.ambiguous.length + ' ambiguous</span>' +
          (mode === 'stage-edits'
            ? '<span class="match-count match-count--diff">' + totalChanges + ' field change' + (totalChanges === 1 ? '' : 's') + ' staged</span>'
            : '') +
        '</div>' +
        '<p>' + headerText + '</p>' +
      '</div>' +
      // Duplicate-target warnings
      (dupGroups.length
        ? '<section class="match-section match-section--warn">' +
            '<h3>⚠ Duplicate targets — ' + dupGroups.length + '</h3>' +
            '<p class="match-section__hint">When multiple CSV rows match the same FuelTransaction, ' +
              'only the <b>last</b> patch survives. Re-check your CSV for duplicate VIN+timestamp rows.</p>' +
            dupHtml + dupOverflow +
          '</section>'
        : '') +
      // Matches
      (matchedRows
        ? '<section class="match-section">' +
            '<h3>Matches — ' + result.matched.length + '</h3>' +
            '<ul class="match-list">' + matchedRows + '</ul>' +
            matchedOverflow +
          '</section>'
        : '') +
      // Unmatched
      (unmatchedRows
        ? '<section class="match-section match-section--ghost">' +
            '<h3>Unmatched — ' + result.unmatched.length + '</h3>' +
            '<p class="match-section__hint">No existing FuelTransaction matched these CSV rows within the ±5&nbsp;min tolerance. ' +
              'To add them as <b>new</b> records, re-run with <b>Force Import…</b> ' +
              '(or use Geotab’s native Fuel Transactions Import). ' +
              'These rows are previewed in grey at the top of the table.</p>' +
            '<ul class="match-list">' + unmatchedRows + '</ul>' +
            unmatchedOverflow +
          '</section>'
        : '') +
      // Ambiguous
      (ambHtml
        ? '<section class="match-section">' +
            '<h3>Ambiguous — ' + result.ambiguous.length + '</h3>' +
            '<ul class="match-list">' + ambHtml + '</ul>' +
          '</section>'
        : '');
    showModal(null, true);
  }

  // ── Modal a11y ───────────────────────────────────────────────────────────
  function showModal(onSave, infoOnly) {
    const modal = $('ftbe-modal');
    if (!modal) return;
    ui.lastFocusEl = document.activeElement;
    modal.hidden = false;
    modal.removeAttribute('inert');
    modal.setAttribute('aria-hidden', 'false');
    const saveBtn = $('ftbe-modal-save');
    if (infoOnly) {
      saveBtn.style.display = 'none';
    } else {
      saveBtn.style.display = '';
      saveBtn.onclick = () => {
        try { if (onSave) onSave(); } finally { hideModal(); }
      };
    }
    // Focus first focusable inside the panel for keyboard users.
    const firstField = modal.querySelector('.addin-modal__body input, .addin-modal__body select, .addin-modal__body textarea, .addin-modal__body button');
    if (firstField) setTimeout(() => firstField.focus(), 0);
  }
  function hideModal() {
    const modal = $('ftbe-modal');
    if (!modal) return;
    modal.hidden = true;
    modal.setAttribute('inert', '');
    modal.setAttribute('aria-hidden', 'true');
    // Reset wide-modal class so the next modal (edit/bulk/force) opens at
    // standard width. Only the CSV match summary opts into the wide layout.
    const panel = modal.querySelector('.addin-modal__panel');
    if (panel) panel.classList.remove('addin-modal__panel--wide');
    const saveBtn = $('ftbe-modal-save');
    if (saveBtn) saveBtn.onclick = null;
    if (ui.lastFocusEl && typeof ui.lastFocusEl.focus === 'function') {
      try { ui.lastFocusEl.focus(); } catch (_) {}
    }
    ui.lastFocusEl = null;
  }

  // ── Default date range (last 30 days) ────────────────────────────────────
  function setDefaultDateRange() {
    const now = new Date();
    const from = new Date(now.getTime() - 30 * 24 * 3600 * 1000);
    const pad = (n) => String(n).padStart(2, '0');
    const toLocal = (d) =>
      d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) +
      'T' + pad(d.getHours()) + ':' + pad(d.getMinutes());
    if (!$('ftbe-from').value) $('ftbe-from').value = toLocal(from);
    if (!$('ftbe-to').value)   $('ftbe-to').value   = toLocal(now);
  }

  // ── Wiring ───────────────────────────────────────────────────────────────
  function bindControls() {
    $('ftbe-load').addEventListener('click', loadTransactions);
    $('ftbe-export-csv').addEventListener('click', exportCsv);
    $('ftbe-import-csv').addEventListener('click', () => $('ftbe-file').click());
    $('ftbe-file').addEventListener('change', (e) => {
      const f = e.target.files && e.target.files[0];
      onImportFileChosen(f, 'stage-edits');
      e.target.value = '';
    });
    $('ftbe-match-csv').addEventListener('click', () => $('ftbe-file-match').click());
    $('ftbe-file-match').addEventListener('change', (e) => {
      const f = e.target.files && e.target.files[0];
      onImportFileChosen(f, 'match-only');
      e.target.value = '';
    });
    $('ftbe-force-import').addEventListener('click', () => $('ftbe-file-force').click());
    $('ftbe-file-force').addEventListener('change', (e) => {
      const f = e.target.files && e.target.files[0];
      onImportFileChosen(f, 'force-add');
      e.target.value = '';
    });
    $('ftbe-search').addEventListener('input', () => renderTable());
    $('ftbe-vol-unit').addEventListener('change', (e) => { ui.volUnit = e.target.value; renderTable(); });
    $('ftbe-odo-unit').addEventListener('change', (e) => { ui.odoUnit = e.target.value; renderTable(); });

    $('ftbe-bulk-edit').addEventListener('click', openBulkEditModal);
    $('ftbe-bulk-delete').addEventListener('click', () => deleteRows(Array.from(ui.selected)));

    // Sort headers
    document.querySelectorAll('#ftbe-table thead th[data-sort]').forEach((th) => {
      th.addEventListener('click', () => {
        const k = th.getAttribute('data-sort');
        if (ui.sortKey === k) ui.sortDir = ui.sortDir === 'asc' ? 'desc' : 'asc';
        else { ui.sortKey = k; ui.sortDir = 'asc'; }
        applySortAndRender();
      });
    });

    // Select all
    $('ftbe-check-all').addEventListener('change', (e) => {
      const rows = deriveDisplayRows();
      if (e.target.checked) rows.forEach((r) => ui.selected.add(r.id));
      else ui.selected.clear();
      renderTable();
    });

    // Delegated tbody handler — survives re-renders.
    $('ftbe-tbody').addEventListener('click', (e) => {
      // Unmatched-CSV preview rows: only support "Dismiss" (they have no id).
      const pendingTr = e.target.closest('tr.is-unmatched-preview');
      if (pendingTr && e.target.matches('button[data-action="dismiss-pending"]')) {
        const idx = Number(e.target.getAttribute('data-pending-idx'));
        if (isFinite(idx)) {
          ui.unmatchedPreview.splice(idx, 1);
          renderTable();
        }
        return;
      }
      if (pendingTr) return;   // ignore other clicks inside ghost rows
      const tr = e.target.closest('tr[data-id]');
      if (!tr) return;
      const id = tr.getAttribute('data-id');
      if (e.target.classList && e.target.classList.contains('ftbe-row-check')) {
        if (e.target.checked) ui.selected.add(id); else ui.selected.delete(id);
        updateBulkButtons();
        tr.classList.toggle('is-selected', e.target.checked);
        return;
      }
      if (e.target.matches('button[data-action]')) {
        const action = e.target.getAttribute('data-action');
        if (action === 'edit')   openEditModal(id);
        if (action === 'revert') { ui.edited.delete(id); renderTable(); }
        if (action === 'delete') deleteRows([id]);
      }
    });

    // Modal close handlers
    document.querySelectorAll('#ftbe-modal [data-modal-close]').forEach((el) => {
      el.addEventListener('click', hideModal);
    });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && !$('ftbe-modal').hidden) hideModal();
    });
  }

  // ── Public lifecycle ─────────────────────────────────────────────────────
  return {
    initialize: function (freshApi, freshState, callback) {
      api = freshApi;
      state = freshState;
      try {
        if (!ui.initialized) {
          bindControls();
          setDefaultDateRange();
          ui.initialized = true;
        }
        // Reference data: best-effort; failure should not block load button.
        try { loadReferenceData(); } catch (e) { console.warn(e); }
      } catch (err) {
        console.error('[fuelBulkEditor] initialize failed', err);
      }
      if (typeof callback === 'function') callback();   // MANDATORY
    },

    focus: function (freshApi, freshState) {
      api = freshApi;
      state = freshState;
      // Sticky-header table is the only size-aware element; re-render in
      // case the iframe was display:none during blur.
      requestAnimationFrame(() => renderTable());
      // If a prior operation was cancelled by blur, the status line may still
      // read "Loading transactions…" or "Saving N edits…". Reset it to a
      // truthful state so the user isn't misled into thinking work is still
      // in flight after they return.
      const statusEl = $('ftbe-status');
      if (statusEl && /…$/.test(statusEl.textContent || '')) {
        setStatus('Resumed. Previous operation was cancelled when the tab lost focus — re-run if needed.');
      }
    },

    blur: function () {
      // Bump the generation token FIRST so any callback that lands after this
      // line sees a stale gen and short-circuits. Then abort each handle so
      // apiCall rejects with CANCELLED and apiMultiCall stops issuing further
      // chunks. We cannot recall a chunk already in flight on the server —
      // its writes will still land — but no NEW chunks will be sent.
      ui.opGen++;
      ui.inflight.forEach((c) => { try { c.abort(); } catch (_) {} });
      ui.inflight.clear();
    },

    unload: function () {
      this.blur();
      ui.rows = [];
      ui.edited.clear();
      ui.selected.clear();
      ui.deviceById.clear();
      ui.driverById.clear();
      const tbody = $('ftbe-tbody');
      if (tbody) tbody.innerHTML = '';
      api = null;
      state = null;
      ui.initialized = false;
    }
  };
};

// ── Standalone bootstrap (file:// preview) ─────────────────────────────────
(function () {
  if (typeof window === 'undefined' || typeof document === 'undefined') return;
  if (window.geotab && window.geotab.api) return;   // running inside MyGeotab
  document.addEventListener('DOMContentLoaded', () => {
    if (window.__fuelBulkEditorBootstrapped) return;
    window.__fuelBulkEditorBootstrapped = true;
    const lifecycle = window.geotab.addin.fuelBulkEditor();
    const stubApi = { getSession: (cb) => cb(null) };
    lifecycle.initialize(stubApi, {}, () => lifecycle.focus(stubApi, {}));
  });
})();
