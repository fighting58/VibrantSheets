# Troubleshooting Addendum (2026-03-26)

## A. Excel "Repaired" Message on Open
**Symptom**
- Excel warns that the file was repaired and references `/xl/styles.xml`.

**Root Cause**
- Invalid currency `numFmt` string emitted in export (`"₩#,##0...` malformed quote boundary).

**Fix**
- Use valid format code generation (`"₩"#,##0...`) and apply consistently across export paths.

## B. Currency Unit Changed After Round-Trip
**Symptom**
- USD cells became KRW after import -> export.

**Root Cause**
- Internal format model did not preserve currency code, only `type/decimals`.

**Fix**
- Persist `currency` in internal format (`KRW`/`USD`).
- Detect from incoming `numFmt` and write back same currency symbol on export.

## C. Print Preview Included Editor Artifacts / Blank Extra Page
**Symptom**
- Selection tint or non-printable grid artifacts appeared; sometimes an empty second page was generated.

**Root Cause**
- Print composition included editor-only layers and over-broad effective range.

**Fix**
- Restrict print content to actual printable range and exclude editor overlays/helpers.
