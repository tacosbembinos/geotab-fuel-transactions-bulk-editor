/*
 * scripts/csv.js — RFC 4180 CSV serialise/parse, self-contained.
 * Exposes window.CSVUtil = { serialize(rows, columns), parse(text) }.
 */
(function () {
  'use strict';

  function quote(v) {
    if (v == null) return '';
    const s = String(v);
    if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  // rows: array of objects. columns: [{ key, header }] in output order.
  function serialize(rows, columns) {
    const header = columns.map((c) => quote(c.header || c.key)).join(',');
    const lines  = rows.map((r) => columns.map((c) => quote(r[c.key])).join(','));
    return header + '\r\n' + lines.join('\r\n') + (lines.length ? '\r\n' : '');
  }

  // Streaming-ish parse: returns { headers, rows: Array<object> }.
  // Tolerant of \r\n / \n line endings, BOM, blank trailing line.
  function parse(text) {
    if (!text) return { headers: [], rows: [] };
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);

    const cells = []; // flat: tokens with row breaks marked
    let cur = '';
    let inQuotes = false;
    let i = 0;
    const len = text.length;
    const records = [[]];
    function pushCell() { records[records.length - 1].push(cur); cur = ''; }
    function pushRow() { records.push([]); }

    while (i < len) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (text[i + 1] === '"') { cur += '"'; i += 2; continue; }
          inQuotes = false; i++; continue;
        }
        cur += ch; i++; continue;
      }
      if (ch === '"')         { inQuotes = true; i++; continue; }
      if (ch === ',')         { pushCell(); i++; continue; }
      if (ch === '\r')        { pushCell(); if (text[i + 1] === '\n') i++; pushRow(); i++; continue; }
      if (ch === '\n')        { pushCell(); pushRow(); i++; continue; }
      cur += ch; i++;
    }
    pushCell();
    // Drop trailing empty row if present
    if (records.length && records[records.length - 1].length === 1 &&
        records[records.length - 1][0] === '') records.pop();

    if (!records.length) return { headers: [], rows: [] };
    const headers = records[0].map((h) => h.trim());
    const rows = records.slice(1).map((cells) => {
      const obj = {};
      for (let k = 0; k < headers.length; k++) obj[headers[k]] = cells[k] != null ? cells[k] : '';
      return obj;
    });
    return { headers, rows };
  }

  window.CSVUtil = { serialize, parse };
})();
