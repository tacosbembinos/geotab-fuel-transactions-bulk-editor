# Fuel Transactions Bulk Editor

A MyGeotab add-in to view, search, export (CSV), and bulk-edit `FuelTransaction`
records via the Geotab API. Bulk edits are committed as chunked sequential
`multiCall` operations (≤100 sub-calls per round-trip) using each record's
Geotab-assigned `id` and concurrency `version`.

## Install in MyGeotab

1. Administration → **System** → **System Settings** → **Add-Ins** → **New Add-In**.
2. Paste this JSON into **Configuration**:

   ```json
   { "url": "https://raw.githubusercontent.com/tacosbembinos/geotab-fuel-transactions-bulk-editor/main/manifest.json" }
   ```

3. Save. The menu item appears under **Engine & Maintenance**.

## Hosting layout

| File | Hosted at | Role |
|---|---|---|
| `manifest.json` | `raw.githubusercontent.com` | Loader. Geotab follows its `url` field. |
| `config.json` | `raw.githubusercontent.com` | Real manifest (name, version, items[]). |
| `fuelBulkEditor.html` + `scripts/*.js` + `styles/main.css` + `images/icon.svg` | `cdn.jsdelivr.net/gh/.../@main/` | Rendered assets. jsDelivr serves correct `Content-Type` (raw.githubusercontent.com serves HTML/SVG as `text/plain`, which browsers refuse to render). |

`@main` follows the default branch's HEAD. Cut a release tag (e.g. `@v1.0.1`)
and update `config.json` to pin customers to it before tagging anything as GA.

## Local preview

Open `fuelBulkEditor.html` directly in a browser. The standalone bootstrap shim
at the bottom of `scripts/main.js` runs `initialize` with a stub `api` — UI
controls render, but Geotab API calls reject with `API unavailable`.

## CSV import formats

**Round-trip (`id`-based).** Export from this add-in; edit values in your editor
of choice; re-import. The script matches by `id` and re-sends with the original
`version` for optimistic-concurrency.

**External (third-party fuel-card feed).** No `id` column required. Provide any
of `Vehicle Identification Number` / `Serial Number` / `License Plate`. The
script prefers a server-side `vehicleIdentificationNumber` search (HAR-verified
the Drive App uses this filter) when the CSV has ≤100 unique VINs; otherwise it
falls back to a date-window pull. Matches are narrowed by ±5 min `dateTime`
tolerance, then staged as pending edits keyed by the Geotab-assigned `id`.
