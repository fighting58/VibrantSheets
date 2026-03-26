# Progress Update (2026-03-26)

## Scope
This update captures the latest fixes around Excel round-trip reliability, print behavior, and format preservation.

## Completed
- Fixed malformed currency number format in XLSX export that caused Excel repair (`xl/styles.xml`).
- Preserved merged ranges more safely during import/export flow.
- Preserved strikethrough during Excel save/load.
- Preserved border styles in save/load and improved thin-border rendering behavior on merged boundaries.
- Improved print workflow:
  - print-area oriented flow
  - safer preview generation
  - reduced editor-only artifact leakage into print output
- Added/expanded print controls in ribbon and print modal behavior.

## Currency Preservation (New)
- Added currency-aware format persistence for `KRW` and `USD`.
- Import now detects currency from Excel format code (`$` vs `₩`).
- Render/export now uses stored currency, not forced KRW.

## Validation Checklist
1. Import Excel file with USD/KRW mixed currency cells.
2. Save as XLSX from VibrantSheets.
3. Open saved copy in Excel:
- no repair dialog
- merged ranges remain intact
- currency unit matches source (USD remains USD, KRW remains KRW)
- border/background/strike styles remain intact

## Known Follow-ups
- Expand currency support beyond KRW/USD if required (EUR/JPY etc.).
- Add explicit UI control for currency unit selection per cell/range.
- Normalize legacy fallback exporter path to use the same format builder API everywhere.
