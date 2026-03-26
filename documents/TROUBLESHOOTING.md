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
셀에 얇은 실선 테두리(1px Solid)를 적용할 때, 전체 테두리를 적용하더라도 위쪽(Top)과 왼쪽(Left) 테두리 선이 인접 셀의 기본 그리드 옅은 선(`var(--grid-line)`)에 의해 사용자가 지정한 테두리로 그리지 않고 가려지는 현상이 관찰되었습니다.

**Root Cause:**  
웹 브라우저의 기본 제공 기능인 `border-collapse: collapse` 테이블 속성은 인접한 두 셀의 테두리 충돌 시 고유한 우선순위 렌더링 알고리즘을 사용합니다.
특히 Webkit 기반 브라우저(Chrome/Edge)는 왼쪽 셀의 우측(`border-right`)과 위쪽 셀의 하단(`border-bottom`) 테두리를 그 대상(오른쪽 셀 및 아래쪽 셀)의 좌측(`border-left`), 최상단(`border-top`) 테두리보다 더 늦게 그려 우선 적용(덮어쓰기)해 버리는 구조를 지닙니다. 즉, A2 셀의 위쪽에 1px 검은색 테두리를 입혀도 바로 위에 맞닿은 A1 셀 아랫부분에 존재하는 1px 투명 그리드 선분에 의해 검은 테두리가 가려진 것입니다.

**Permanent Resolution:**  
- Webkit의 고질적 특성을 역이용하는 **양방향 동기화 렌더링(Mirror Rendering) 시스템**을 개발해 도입했습니다.
- 테두리를 적용한 뒤 화면에 그리는 `renderBorders(cell)` 단계에서, 현재 평가 중인 해당 셀의 테두리 정보뿐 아니라 **맞닿아 있는 인접된 셀의 테두리 데이터까지 양방향으로 동시 평가 및 수집**합니다.
- 만약 현재 셀이나 맞닿은 반대편 셀 중 어느 한 쪽에라도 테두리가 설정되어 있다면, **양쪽 셀의 마주보는 변 모두에 해당 테두리 CSS(`!important` 처리)를 발라버립니다.** (예: A2 위쪽이 그려질 땐 A1 아래쪽에도 동시에 동일한 테두리를 부여)
- 결과적으로 브라우저 렌더링 엔진이 둘 중 어느 방향의 테두리를 우선하든 관계없이, 적용하려는 사용자 테두리가 100% 동일한 상태로 그려져 완벽히 노출되도록 렌더링 파이프라인 우회를 달성했습니다.

---

## 6. Multi-Row/Column Deletion Data Retention Bug

**Symptom:**  
사용자가 행/열 헤더를 드래그하여 여러 행이나 열을 동시에 선택한 뒤 삭제(Delete Row/Col)를 실행해도 상단의 데이터가 아래로 밀려 덮어써지기만 할 뿐, 실질적인 해당 영역의 데이터와 서식이 완전히 지워지지 않고 화면 및 내부 상태(state)에 잔류하는 현상이 발생했습니다.

**Root Cause:**  
내부 엔진의 `shiftData` 메서드는 삽입 시(Down/Right Shift)에는 좌표를 밀어내는 역할(`coord >= threshold ? coord + delta : coord`)을 잘 수행하지만, **삭제 시(음수 델타) 특정 삭제 대상 구간(`threshold + delta` 부터 `threshold - 1` 까지) 안에 있는 기존 데이터를 메모리에서 명시적으로 폐기하는 로직이 누락**되어 있었습니다. 이로 인해 이동 조건에 해당하지 않는 삭제 구간 데이터들이 그대로 신규 상태(new state) 객체로 복사-이관되고 말았습니다.

**Permanent Resolution:**  
- `shiftData` 내의 좌표 계산 헬퍼 함수 `shiftCoord`를 업그레이드했습니다.
- 이동 델타가 음수(`d < 0`)일 때 파기되어야 할 좌표 범위(`coord >= t + d && coord < t`)를 감지하면 **유효하지 않은 좌표인 `-1`을 반환**하도록 방어 코드를 작성했습니다.
- 하위 치환 루프에서 새 좌표가 0보다 클 때(`nRow > 0 && nColNum > 0`)만 새 객체 맵에 등록하므로, 삭제 대상 구간 내의 모든 데이터(값, 서식, 수식, 테두리)가 깔끔하게 증발하여 올바른 삭제 동작을 완수합니다.

---

## 7. IME (Korean) First Character Loss ("rㅏ" Issue)

**Symptom:**  
엔터 키를 눌러 다음 셀로 이동한 직후, "가"를 입력하면 셀에 "rㅏ"처럼 자음이 분리되거나 첫 타건이 영문으로 박히는 현상이 발생했습니다. 이는 IME 기화(Composition) 시작 시점이 브라우저의 편집 모드 전환 속도와 엇갈려 발생하는 전형적인 포커스 레이스 컨디션입니다.

**Root Cause:**  
1.  기존 로직은 `keydown` 이벤트에서 직접 `activeCell.innerText = ''`로 셀을 비우고 `isEditing` 상태로 전환했습니다.
2.  이 수동 비우기 과정에서 브라우저는 현재의 입력 컨텍스트(Composition Context)가 파괴되었다고 판단하여, 한글 입력 엔진을 끊어버리고 첫 글자를 영문 그대로 삽입하게 됩니다.
3.  그 이후의 입력부터 다시 IME가 작동하면서 한글 조합이 시작되므로 "r" + "ㅏ"가 되는 결과가 나타납니다.

**Permanent Resolution:**  
- **포커스 시 전체 선택 (Select All Strategy)**: `handleCellFocus`에서 셀이 포커스를 받는 즉시 `window.getSelection()`을 통해 셀의 모든 내용을 선택(Select All)하도록 레이아웃 엔진을 수정했습니다.
- **수동 비우기 방지**: `handleKeyDown` 및 `handleCompositionStart`에서 직접 `innerText = ''`를 명시적으로 호출하는 대신, 기존 내용을 선택된 상태로 둡니다.
- **네이티브 덮어쓰기 유도**: 이 상태에서 키보드 입력이 시작되면, 브라우저는 '선택된 영역을 입력값으로 대체'하는 네이티브 동작을 수행합니다. 이 과정에서 IME 엔진의 흐름이 깨지지 않고 자연스럽게 전체 내용이 한글로 덮어씌워지게 되어, 첫 글자부터 완벽한 조합이 보장됩니다.
