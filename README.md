# VibrantSheets 📊

VibrantSheets is a modern, high-performance, and visually stunning web-based spreadsheet application built with Vanilla JavaScript and CSS. It features a premium Glassmorphism design with a dark mode aesthetic, focusing on a smooth and intuitive user experience.

![VibrantSheets Preview](https://github.com/fighting58/VibrantSheets/raw/master/preview.png) *(Note: Add a real preview image later)*

## ✨ Key Features (Implemented)

### 🎨 Premium UI/UX
- **Modern Design**: Glassmorphism-based ribbon menu and formula bar.
- **Dark Mode**: Sleek, eye-friendly dark theme optimized for productivity.
- **Micro-interactions**: Subtle animations, resize guides, and real-time status updates.

### ⌨️ Advanced Grid Interaction
- **Infinite Scroll**: Dynamic row appending to ensure high performance with large datasets.
- **Navigation**: Full keyboard support (Enter, Tab, Arrow keys, Shift+Tab/Enter).
- **Multi-line Editing**: Support for multiple lines within a single cell via `Alt + Enter`.
- **Range Selection**: Efficiently select multiple cells using mouse drag, `Shift + Click`, or `Shift + Arrow` keys.
- **Fill Handle**: Bidirectional fill handle to copy data or extend patterns horizontally and vertically.
- **Smart Clipboard**: Advanced TSV parser handles quoted cells and multi-line data from Excel/Google Sheets perfectly.

### 📋 File & Data Management
- **.vsht Format**: Custom JSON format that preserves not only data but also **column widths and row heights**.
- **Excel (.xlsx) Support**: Native binary `.xlsx` import powered by SheetJS.
- **Smart Clipboard**: Full support for Copy/Paste with TSV format for seamless compatibility with Microsoft Excel/Google Sheets.
- **CSV Support**: 
    - **Advanced Import**: Smart parser that handles multi-line cells and auto-detects delimiters.
    - **Native Save**: Uses `File System Access API` for a standard OS "Save As" experience.
    - **Excel Compatibility**: UTF-8 BOM encoding for perfect Korean character support in Excel.

### 📏 Grid Customization
- **Independent Resizing**: Drag the edges of row/column headers to customize grid dimensions without affecting neighboring cells.
- **Status Indicator**: Real-time "Edited" (Yellow) and "Saved" (Green) badge to track document state.

---

## 🚀 Roadmap (Upcoming Features)

- [ ] **Cell Styling**: Toolbar for Bold, Italic, Strikethrough, Text Color, and Background Color.
- [ ] **Formatting**: Support for Currency, Percentage, and Date formats.
- [ ] **Formula Engine**: Implementation of a formula parser for basic arithmetic and functions like `SUM`, `AVG`, `COUNT`, `MIN`, `MAX`.
- [ ] **Cell References**: Dynamic coordination between cells (e.g., `=A1+B1`).
- [ ] **Local Storage**: Automatic session persistence to prevent data loss.
- [ ] **Row/Col Operations**: Insert/Delete rows and columns dynamically.

## 🛠️ Tech Stack
- **Language**: Vanilla JavaScript (ES6+)
- **Styling**: Vanilla CSS (Custom Properties, Flexbox, Grid)
- **Architecture**: Object-Oriented JS (Class-based)

---

## 📖 How to Run
Simply open `index.html` in any modern web browser. 
*Note: For the full 'Save As' experience, use a Chromium-based browser (Chrome, Edge).*

---
**Developed with ❤️ by [fighting58](https://github.com/fighting58)**
