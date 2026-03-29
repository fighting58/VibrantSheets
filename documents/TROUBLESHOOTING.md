# Troubleshooting Guide

This document lists critical bugs and technical debt encountered during the development of VibrantSheets, along with their root causes and permanent resolutions.

## 1. Grid Interaction & Selection Blocked

**Symptom:**  
The fill handle, multi-cell drag-selection, and range overlays were completely unresponsive ("frozen").

**Root Cause:**  
A JavaScript `TypeError` occurred within `setupEventListeners()` because certain DOM elements (like `find-case`, `find-exact` checkboxes) were removed from the ribbon UI in recent redesigns, but the JS code still tried to attach listeners to them. This crash prevented the subsequent global mouse and keyboard listeners from being registered.

**Permanent Resolution:**  
- Added robust null-checks for all ribbon UI elements in `setupEventListeners()`.
- Removed redundant/duplicate overlay creation calls that were causing DOM littering and z-index confusion.

---

## 8. Merged Cells Lose Merge on Paste

**Symptom:**  
Copying a selection that includes merged cells and pasting it would flatten the target area (merged cells are not preserved).

**Root Cause:**  
Paste logic prioritized the system clipboard when available. Even when copying inside the app, the system clipboard text was used, so the internal merge metadata was not applied. Additionally, minor newline differences caused internal clipboard comparisons to fail.

**Permanent Resolution:**  
- Stored merged ranges as relative offsets alongside internal clipboard data.
- Normalized clipboard text (line endings and trailing newlines) before comparing with the internal TSV.
- When the internal clipboard is used, reapply the stored merged ranges after pasting cell values.

---

## 9. Merged Cells Lost After Excel Save/Load

**Symptom:**  
Saving to `.xlsx` and re-opening caused merged ranges to appear broken or missing.

**Root Cause:**  
On import, merged ranges could extend beyond the computed grid size (based on data only). The grid was not expanded to include the full merge bounds, so merges were not applied correctly.

**Permanent Resolution:**  
- When importing `.xlsx`, expand `rows/cols` to at least the maximum merged range size before rendering.

## 2. Fill Handle Interaction Hidden/Blocked

**Symptom:**  
Users could see the fill handle but clicking/dragging it would accidentally start a new cell selection instead of a fill operation.

**Root Cause:**  
- **Z-Index Conflict:** The `.selection-overlay` (z-index 21) was placed above the `.fill-handle` (z-index 20). Even with `pointer-events: none`, the browser interaction occasionally failed or was confusing on certain coordinate points.
- **Missing Listener:** In some refactorings, the `mousedown` handler for the fill handle was not attached to the newly created DOM element.

**Permanent Resolution:**  
- Raised `.fill-handle` to `z-index: 30` (highest among grid overlays).
- Consolidated overlay creation into `init()` and ensured each element has its specific listeners attached immediately.

---

## 3. UI Contrast/Visibility (Light Grid Mode)

**Symptom:**  
After switching the grid background to white for readability, the selection range (`range-overlay`) became nearly invisible.

**Root Cause:**  
The alpha-channel opacity used for the selection background (`rgba(..., 0.05)`) was optimized for dark themes and washed out completely on solid white.

**Permanent Resolution:**  
- Updated CSS variables to use higher opacity (`0.15` - `0.2`) when in light-grid mode.
- Adjusted border widths and colors for overlays to provide high contrast against `#ffffff`.

---

## 4. "Overwrite on Type" (Input Lost)

**Symptom:**  
Typing on a selected cell would clear the content but fail to input the actual character, or the first character would be lost after moving to a new cell using `Enter`.

**Root Cause:**  
- **Caret Loss:** Setting `cell.innerText = ''` inside a `keydown` handler causes many browsers (Chrome/WebKit) to lose the current selection (Caret). Subsequent typing fails because there is no focused insertion point.
- **Race Condition:** Moving focus during the `Enter` key event sometimes created a race condition where the browser hadn't fully established the focus before the next key event.

**Permanent Resolution:**  
- **Forced Selection:** Immediately after clearing `innerText` in `prepareEnterMode`, the selection is manually re-established using `document.createRange()` and `window.getSelection()`.
- **Focus Stabilization:** Simplified the movement logic to rely on the natural browser `focus` event rather than forcing redundant handler calls.

---

## 247. IME (Korean) Composition Issues

**Symptom:**  
Starting a Korean character composition on an existing value would sometimes result in the first component being combined with existing text or cancelled.

**Root Cause:**  
The `compositionstart` event was initially tried without "Overwrite" logic, leading to mixed content. Furthermore, the "transparent caret" used to support IME in Ready Mode occasionally interfered with the visible focus.

**Permanent Resolution:**  
- Integrated `prepareEnterMode(cell, true)` into `handleCompositionStart` to ensure a clean slate for CJK input.
- Standardized `contentEditable = true` on all focused cells while maintaining spreadsheet-like behavior.

---

## 5. Thin Border Truncation by Grid Lines (Top / Left Border Missing)

**Symptom:**  
When a user applied a 1px solid thin border to a cell, the top and left borders were not visibly rendered or were truncated by the adjacent cells' default transparent grid lines (e.g., `var(--grid-line)`).

**Root Cause:**  
The browser's native `border-collapse: collapse` table property features a specific layout algorithm to resolve adjacent border conflicts. Specifically in Webkit-based browsers (Chrome/Edge), the bottom/right borders of an adjacent top/left cell tend to override the target cell's top/left borders. For instance, the transparent bottom border of cell A1 overrides the solid top border of cell A2.

**Permanent Resolution:**  
- Developed a "Bidirectional Mirror Rendering" system to counteract Webkit's rendering quirks.
- When applying borders in the grid rendering phase, the `renderBorders(cell)` logic collects and renders not only the cell's own border state but also the adjacent cells' boundaries.
- If a border is specified on a boundary, the CSS `!important` rule is applied symmetrically to both the target cell and its neighboring cell (e.g. applying the targeted border to both A2's top and A1's bottom).
- Re-rendering both edges identically forces the browser's resolving algorithm to use the user's selected style seamlessly, ensuring a 100% thick, visible border representation.

---

## 6. Multi-Row/Column Deletion Data Retention Bug

**Symptom:**  
When a user selected multiple column/row headers and executed a deletion (Delete Row/Col), the lower data shifted upwards to fill the gap visually, but the underlying data, formulas, and styles in the originally deleted area were not fully wiped from the internal state, causing visual and state corruption.

**Root Cause:**  
The internal data-shifting method `shiftData` safely handled moving coordinates (downwards/rightwards shift via `coord >= threshold ? coord + delta : coord`). However, during a destructive operation (negative delta), the code failed to explicitly wipe the memory for the cells within the deletion range (between `threshold + delta` and `threshold - 1`). Consequently, the deleted data remained untouched and was simply copied over as a new state iteration.

**Permanent Resolution:**  
- Upgraded the inner coordinate calculation helper `shiftCoord`.
- When detecting a destructive shifting operation (`d < 0`), it now actively identifies and evaluates the deletion coordinate range (`coord >= t + d && coord < t`), explicitly returning `-1` for invalid boundaries.
- The state replacement loops now skip updating indices that yield `-1` and rigorously delete all associated inner properties (data, formulas, styles, borders), ensuring the chunk vanishes perfectly.

---

## 7. IME (Korean) First Character Loss ("r" Issue)

**Symptom:**  
Immediately after pressing Enter to move the cell focus, if a user typed a Korean character like "가", the first consonant would detach (e.g. "r") or the entire set would remain uncombined english characters. This was a race condition arising from the gap between the window's composition startup and the application switching to edit mode.

**Root Cause:**  
1. The old mechanism manually cleared out the cell value via `activeCell.innerText = ''` inside the `keydown` event to transition to the `isEditing` mode.
2. During this manual clearing phase, the browser recognized that its ongoing "Composition Context" for the typed character was destroyed, failing to assemble the combined character and keeping the raw english input (e.g. "r").
3. By the time the IME engine restarted in the input field, the composition flow was broken.

**Permanent Resolution:**  
- **Select-All Strategy:** Within `handleCellFocus`, whenever a cell receives focus, the system immediately leverages `window.getSelection()` to establish a 'Select-All' state across the cell's contents.
- **Prevent Manual Clearing:** Explicit assignments like `innerText = ''` in `handleKeyDown` and `handleCompositionStart` were completely removed to respect existing selected text states.
- **Native Overwrite:** When typing begins on a fully selected cell, the browser utilizes its native "overwrite" interaction to replace the selection block. This preserves the IME composition seamlessly, ensuring the first CJK characters combine correctly.

## 10. Merged Cell Border Style Not Updating (Dot/Dash Overridden by Thin Solid)

**Symptom:**  
After applying thin solid borders to a merged range, later applying dot/dash borders to adjacent cells did not update the merged cell boundary. The boundary kept showing a solid thin line.

**Root Cause:**  
The thin-border mirror rendering (added to avoid border-collapse loss) persisted on merged-cell boundaries. When a new border was applied to an adjacent cell, the mirrored thin border was not synchronized to the merged anchor, so the older style won visually.

**Permanent Resolution:**  
- When applying a border to a cell adjacent to a merged range, also apply the same border to the merged range’s anchor on the corresponding boundary.
- This keeps the last-applied border style (dot/dash/solid) consistent on merged boundaries while preserving the thin-line mirror fix.


---

## 11. Browser Print Preview Ignored Horizontal Page Splits

**Symptom:**
- Page-break preview showed multiple horizontal pages, but the browser print preview still collapsed output into a single page.

**Root Cause:**
- The app was sending the live spreadsheet table directly to window.print().
- Chrome's print engine can shrink a wide HTML table to fit one page, even when the app's own page-break calculation expects multiple pages.

**Permanent Resolution:**
- Build dedicated print-only page DOM nodes before print.
- Clone the spreadsheet table once per computed print page and offset each clone so each page viewport shows only its own slice.
- Print the generated page containers instead of relying on the live grid layout to paginate itself.

---

## 12. First Printed Page Was Blank After Switching To Custom Print Pages

**Symptom:**
- After introducing custom print pages, the print preview showed the correct total page count, but page 1 was blank.

**Root Cause:**
- The normal app layout was still participating in print flow ahead of the generated print-page container, so the browser reserved the first printable page for non-print content.

**Permanent Resolution:**
- During print mode, hide all direct body children except the dedicated .vs-print-pages container.
- Keep only the generated print pages visible so page 1 starts with real printable content.

---

## 13. XLSX Import Failed With `anchors` / `Target` Errors

**Symptom:**
- Importing certain `.xlsx` files fails with errors like:
  - `Cannot read properties of undefined (reading 'anchors')`
  - `Cannot read properties of undefined (reading 'Target')`

**Root Cause:**
- ExcelJS choked on malformed/unsupported drawing anchors in `xl/drawings/*.xml`.
- Negative offsets or drawing references left in worksheet XML caused `anchors` or relationship targets to be undefined during reconcile.

**Permanent Resolution:**
- Preprocess drawing XML:
  - Clamp `<colOff>/<rowOff>` negative values to `0`.
- If loading still fails, strip drawing/vml parts and their worksheet references:
  - Remove `xl/drawings/*`
  - Remove drawing/vml relationships from `xl/worksheets/_rels/*.rels`
  - Remove `<drawing/>` and `<legacyDrawing/>` tags from worksheet XML
- Show an explicit user-facing warning when images are excluded.

---

## 14. Excel "Repaired" Message on Open

**Symptom:**  
Excel warns that the file was repaired and references `/xl/styles.xml`.

**Root Cause:**  
Invalid currency `numFmt` string emitted in export (`"₩#,##0...` malformed quote boundary).

**Permanent Resolution:**  
- Use valid format code generation (`"₩"#,##0...`) and apply consistently across export paths.

---

## 15. Currency Unit Changed After Round-Trip

**Symptom:**  
USD cells became KRW after import -> export.

**Root Cause:**  
Internal format model did not preserve currency code, only `type/decimals`.

**Permanent Resolution:**  
- Persist `currency` in internal format (`KRW`/`USD`).
- Detect from incoming `numFmt` and write back same currency symbol on export.

---

## 16. Print Preview Included Editor Artifacts / Blank Extra Page

**Symptom:**  
Selection tint or non-printable grid artifacts appeared; sometimes an empty second page was generated.

**Root Cause:**  
Print composition included editor-only layers and over-broad effective range.

**Permanent Resolution:**  
- Restrict print content to actual printable range and exclude editor overlays/helpers.
