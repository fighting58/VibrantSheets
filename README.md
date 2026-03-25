# VibrantSheets 📊

VibrantSheets is a modern, high-performance, and visually stunning web-based spreadsheet application built with Vanilla JavaScript and CSS. It features a premium Glassmorphism design with a dark mode aesthetic, focusing on a smooth and intuitive user experience.

![VibrantSheets Preview](https://github.com/fighting58/VibrantSheets/raw/master/preview.png) *(Note: Add a real preview image later)*

## ✨ Key Features (Implemented)

### 🎨 Premium UI/UX
- **Modern Design**: Glassmorphism-based ribbon menu and formula bar.
- **Dark Mode**: Sleek, eye-friendly dark theme optimized for productivity.
- **Micro-interactions**: Subtle animations, resize guides, and real-time status updates.

### ⌨️ Advanced Grid Interaction (Excel-Style)
- **Ready/Edit/Enter Modes**: 
    - **Ready**: Standard navigation (Arrow keys).
    - **Edit**: Refine content (`F2` or Double-click) with internal cursor.
    - **Enter**: Immediate overwriting mode when typing while selected.
- **Infinite Scroll**: High-performance lazy loading via dynamic row appending.
- **Precision Selection**: Dedicated 3px thick selection overlays for clear visual feedback.
- **Fill Handle**: Bidirectional fill handle for copying and pattern extension.
- **Smart Clipboard**: Custom TSV parser handles quoted cells and multi-line data from Excel/Google Sheets perfectly.
- **Escape Support**: `Esc` key cancels ongoing edits and reverts to original content.

### 📋 File & Data Management (Save & Save As)
- **Persistent FileHandle**: Maintains file linkage to allow one-click **"Save" (Overwrite)** after permissions.
- **Save As...**: Export to multiple formats including **.vsht**, **.xlsx**, and **.csv**.
- **.vsht Format**: High-fidelity custom JSON format preserving data, column widths, and row heights.
- **Excel (.xlsx) Integration**: Native binary import and export powered by **SheetJS**.
- **Deluxe CSV Support**: Smart parser for multi-line cells and UTF-8 BOM encoding for perfect Excel compatibility.

### 📏 Grid Customization
- **Independent Resizing**: Drag the edges of row/column headers to customize grid dimensions without affecting neighboring cells.
- **Status Indicator**: Real-time "Edited" (Yellow) and "Saved" (Green) badge to track document state.

---

## 🚀 Roadmap (Upcoming Implementation)

- [x] **Phase 5: Styling System & IME Fix** ✅
    - Toolbar for Bold, Italic, Underline, Strikethrough, and Colors.
    - Advanced 'Always-Editable' logic for perfect Korean input.
- [ ] **Phase 6: Table Operations** ⬅️ **Next Step**
    - Insert/Delete Rows and Columns.
    - Dynamic resizing and coordinate re-calculation.
- [ ] **Phase 7: Advanced Data Formatting**
    - Support for Currency, Percentage, Date, and Decimal control.
- [ ] **Phase 8: Formula Engine Core**
    - Implementation of a robust formula parser.
    - A1/B2 style cell reference evaluation.
- [ ] **Phase 9: Built-in Functions**
    - `SUM`, `AVG`, `COUNT`, `MIN`, `MAX` and Range (`B2:C10`) support.
- [ ] **Phase 10: Persistence & Recovery**
    - `localStorage` auto-save to prevent data loss.
    - Recent files management and session recovery.

## 🛠️ Tech Stack
- **Language**: Vanilla JavaScript (ES6+)
- **Styling**: Vanilla CSS (Premium Glassmorphism Design System)
- **Libraries**: [SheetJS (xlsx.full.min.js)](https://sheetjs.com/) for binary Excel support.
- **API**: Modern **File System Access API** for native OS-level file operations.
- **Optimization**: Dynamic DOM Row Management (Infinite Scroll).

---

## 📖 How to Run
Simply open `index.html` in any modern web browser. 
*Note: For the full 'Save As' experience, use a Chromium-based browser (Chrome, Edge).*

---
**Developed with ❤️ by [fighting58](https://github.com/fighting58)**
