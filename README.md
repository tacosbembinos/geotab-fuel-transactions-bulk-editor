# Fuel Transactions Bulk Editor

A MyGeotab add-in to view, search, export (CSV), and bulk-edit `FuelTransaction`
records via the Geotab API. Bulk edits are committed as chunked sequential
`multiCall` operations (≤100 sub-calls per round-trip) using each record's
Geotab-assigned `id` and concurrency `version`.

## Install in MyGeotab

1. Administration → **System** → **System Settings** → **Add-Ins** → **New Add-In**.
2. Paste this contents of the config.JSON into **Configuration**:
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

## CSV import / export — native Geotab template

Both export and import use the **native Geotab "Fuel Transactions Import
Template"** layout — identical 16 column headers in the same order. This is
the same file shape Geotab's built-in importer reads, and the same shape most
fuel-card providers (e.g. WEX, Comdata, Fuelman) emit:

```
Date & Time, Vehicle Identification Number, Serial Number, License Plate,
Vehicle Description, Cardholder, Card Number, Volume (L), Cost, Currency Code,
Product Type, Transaction Odometer, Location Coordinates, Location Address,
Site Name, Comments
```

Record identity is by **VIN → Serial Number → License Plate** (in that
precedence), narrowed by `±5 min` `Date & Time` tolerance. Matched rows are
staged as pending edits keyed by the Geotab-assigned `id` + `version`; the
user reviews, then commits via **Save edits** (chunked sequential `multiCall`
`Set`) or removes via **Delete selected** (`multiCall` `Remove`). Unmatched
rows are reported in a summary modal and not silently turned into `Add`
operations.

When the CSV has ≤100 unique VINs, the script issues per-VIN
`vehicleIdentificationNumber` searches (HAR-verified that the Drive App uses
this server-side filter); otherwise it falls back to a single date-window pull
covering min/max `Date & Time` ±1 day.
