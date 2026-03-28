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
????뉗? ?ㅼ꽑 ?뚮몢由?1px Solid)瑜??곸슜???? ?꾩껜 ?뚮몢由щ? ?곸슜?섎뜑?쇰룄 ?꾩そ(Top)怨??쇱そ(Left) ?뚮몢由??좎씠 ?몄젒 ???湲곕낯 洹몃━???낆? ??`var(--grid-line)`)???섑빐 ?ъ슜?먭? 吏?뺥븳 ?뚮몢由щ줈 洹몃━吏 ?딄퀬 媛?ㅼ????꾩긽??愿李곕릺?덉뒿?덈떎.

**Root Cause:**  
??釉뚮씪?곗???湲곕낯 ?쒓났 湲곕뒫??`border-collapse: collapse` ?뚯씠釉??띿꽦? ?몄젒????????뚮몢由?異⑸룎 ??怨좎쑀???곗꽑?쒖쐞 ?뚮뜑留??뚭퀬由ъ쬁???ъ슜?⑸땲??
?뱁엳 Webkit 湲곕컲 釉뚮씪?곗?(Chrome/Edge)???쇱そ ????곗륫(`border-right`)怨??꾩そ ????섎떒(`border-bottom`) ?뚮몢由щ? 洹?????ㅻⅨ履?? 諛??꾨옒履??)??醫뚯륫(`border-left`), 理쒖긽??`border-top`) ?뚮몢由щ낫??????쾶 洹몃젮 ?곗꽑 ?곸슜(??뼱?곌린)??踰꾨━??援ъ“瑜?吏?숇땲?? 利? A2 ????꾩そ??1px 寃????뚮몢由щ? ?낇???諛붾줈 ?꾩뿉 留욌떯? A1 ? ?꾨옯遺遺꾩뿉 議댁옱?섎뒗 1px ?щ챸 洹몃━???좊텇???섑빐 寃? ?뚮몢由ш? 媛?ㅼ쭊 寃껋엯?덈떎.

**Permanent Resolution:**  
- Webkit??怨좎쭏???뱀꽦????씠?⑺븯??**?묐갑???숆린???뚮뜑留?Mirror Rendering) ?쒖뒪??*??媛쒕컻???꾩엯?덉뒿?덈떎.
- ?뚮몢由щ? ?곸슜?????붾㈃??洹몃━??`renderBorders(cell)` ?④퀎?먯꽌, ?꾩옱 ?됯? 以묒씤 ?대떦 ????뚮몢由??뺣낫肉??꾨땲??**留욌떯???덈뒗 ?몄젒??????뚮몢由??곗씠?곌퉴吏 ?묐갑?μ쑝濡??숈떆 ?됯? 諛??섏쭛**?⑸땲??
- 留뚯빟 ?꾩옱 ??대굹 留욌떯? 諛섎???? 以??대뒓 ??履쎌뿉?쇰룄 ?뚮몢由ш? ?ㅼ젙?섏뼱 ?덈떎硫? **?묒そ ???留덉＜蹂대뒗 蹂 紐⑤몢???대떦 ?뚮몢由?CSS(`!important` 泥섎━)瑜?諛쒕씪踰꾨┰?덈떎.** (?? A2 ?꾩そ??洹몃젮吏???A1 ?꾨옒履쎌뿉???숈떆???숈씪???뚮몢由щ? 遺??
- 寃곌낵?곸쑝濡?釉뚮씪?곗? ?뚮뜑留??붿쭊????以??대뒓 諛⑺뼢???뚮몢由щ? ?곗꽑?섎뱺 愿怨꾩뾾?? ?곸슜?섎젮???ъ슜???뚮몢由ш? 100% ?숈씪???곹깭濡?洹몃젮???꾨꼍???몄텧?섎룄濡??뚮뜑留??뚯씠?꾨씪???고쉶瑜??ъ꽦?덉뒿?덈떎.

---

## 6. Multi-Row/Column Deletion Data Retention Bug

**Symptom:**  
?ъ슜?먭? ?????ㅻ뜑瑜??쒕옒洹명븯???щ윭 ?됱씠???댁쓣 ?숈떆???좏깮??????젣(Delete Row/Col)瑜??ㅽ뻾?대룄 ?곷떒???곗씠?곌? ?꾨옒濡?諛????뼱?⑥?湲곕쭔 ??肉? ?ㅼ쭏?곸씤 ?대떦 ?곸뿭???곗씠?곗? ?쒖떇???꾩쟾??吏?뚯?吏 ?딄퀬 ?붾㈃ 諛??대? ?곹깭(state)???붾쪟?섎뒗 ?꾩긽??諛쒖깮?덉뒿?덈떎.

**Root Cause:**  
?대? ?붿쭊??`shiftData` 硫붿꽌?쒕뒗 ?쎌엯 ??Down/Right Shift)?먮뒗 醫뚰몴瑜?諛?대궡????븷(`coord >= threshold ? coord + delta : coord`)?????섑뻾?섏?留? **??젣 ???뚯닔 ?명?) ?뱀젙 ??젣 ???援ш컙(`threshold + delta` 遺??`threshold - 1` 源뚯?) ?덉뿉 ?덈뒗 湲곗〈 ?곗씠?곕? 硫붾え由ъ뿉??紐낆떆?곸쑝濡??먭린?섎뒗 濡쒖쭅???꾨씫**?섏뼱 ?덉뿀?듬땲?? ?대줈 ?명빐 ?대룞 議곌굔???대떦?섏? ?딅뒗 ??젣 援ш컙 ?곗씠?곕뱾??洹몃?濡??좉퇋 ?곹깭(new state) 媛앹껜濡?蹂듭궗-?닿??섍퀬 留먯븯?듬땲??

**Permanent Resolution:**  
- `shiftData` ?댁쓽 醫뚰몴 怨꾩궛 ?ы띁 ?⑥닔 `shiftCoord`瑜??낃렇?덉씠?쒗뻽?듬땲??
- ?대룞 ?명?媛 ?뚯닔(`d < 0`)?????뚭린?섏뼱????醫뚰몴 踰붿쐞(`coord >= t + d && coord < t`)瑜?媛먯??섎㈃ **?좏슚?섏? ?딆? 醫뚰몴??`-1`??諛섑솚**?섎룄濡?諛⑹뼱 肄붾뱶瑜??묒꽦?덉뒿?덈떎.
- ?섏쐞 移섑솚 猷⑦봽?먯꽌 ??醫뚰몴媛 0蹂대떎 ????`nRow > 0 && nColNum > 0`)留???媛앹껜 留듭뿉 ?깅줉?섎?濡? ??젣 ???援ш컙 ?댁쓽 紐⑤뱺 ?곗씠??媛? ?쒖떇, ?섏떇, ?뚮몢由?媛 源붾걫?섍쾶 利앸컻?섏뿬 ?щ컮瑜???젣 ?숈옉???꾩닔?⑸땲??

---

## 7. IME (Korean) First Character Loss ("r?? Issue)

**Symptom:**  
?뷀꽣 ?ㅻ? ?뚮윭 ?ㅼ쓬 ?濡??대룞??吏곹썑, "媛"瑜??낅젰?섎㈃ ???"r??泥섎읆 ?먯쓬??遺꾨━?섍굅??泥??嫄댁씠 ?곷Ц?쇰줈 諛뺥엳???꾩긽??諛쒖깮?덉뒿?덈떎. ?대뒗 IME 湲고솕(Composition) ?쒖옉 ?쒖젏??釉뚮씪?곗????몄쭛 紐⑤뱶 ?꾪솚 ?띾룄? ?뉕컝??諛쒖깮?섎뒗 ?꾪삎?곸씤 ?ъ빱???덉씠??而⑤뵒?섏엯?덈떎.

**Root Cause:**  
1.  湲곗〈 濡쒖쭅? `keydown` ?대깽?몄뿉??吏곸젒 `activeCell.innerText = ''`濡????鍮꾩슦怨?`isEditing` ?곹깭濡??꾪솚?덉뒿?덈떎.
2.  ???섎룞 鍮꾩슦湲?怨쇱젙?먯꽌 釉뚮씪?곗????꾩옱???낅젰 而⑦뀓?ㅽ듃(Composition Context)媛 ?뚭눼?섏뿀?ㅺ퀬 ?먮떒?섏뿬, ?쒓? ?낅젰 ?붿쭊???딆뼱踰꾨━怨?泥?湲?먮? ?곷Ц 洹몃?濡??쎌엯?섍쾶 ?⑸땲??
3.  洹??댄썑???낅젰遺???ㅼ떆 IME媛 ?묐룞?섎㈃???쒓? 議고빀???쒖옉?섎?濡?"r" + "??媛 ?섎뒗 寃곌낵媛 ?섑??⑸땲??

**Permanent Resolution:**  
- **?ъ빱?????꾩껜 ?좏깮 (Select All Strategy)**: `handleCellFocus`?먯꽌 ????ъ빱?ㅻ? 諛쏅뒗 利됱떆 `window.getSelection()`???듯빐 ???紐⑤뱺 ?댁슜???좏깮(Select All)?섎룄濡??덉씠?꾩썐 ?붿쭊???섏젙?덉뒿?덈떎.
- **?섎룞 鍮꾩슦湲?諛⑹?**: `handleKeyDown` 諛?`handleCompositionStart`?먯꽌 吏곸젒 `innerText = ''`瑜?紐낆떆?곸쑝濡??몄텧?섎뒗 ??? 湲곗〈 ?댁슜???좏깮???곹깭濡??〓땲??
- **?ㅼ씠?곕툕 ??뼱?곌린 ?좊룄**: ???곹깭?먯꽌 ?ㅻ낫???낅젰???쒖옉?섎㈃, 釉뚮씪?곗???'?좏깮???곸뿭???낅젰媛믪쑝濡??泥??섎뒗 ?ㅼ씠?곕툕 ?숈옉???섑뻾?⑸땲?? ??怨쇱젙?먯꽌 IME ?붿쭊???먮쫫??源⑥?吏 ?딄퀬 ?먯뿰?ㅻ읇寃??꾩껜 ?댁슜???쒓?濡???뼱?뚯썙吏寃??섏뼱, 泥?湲?먮????꾨꼍??議고빀??蹂댁옣?⑸땲??

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
