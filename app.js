class VibrantSheets {
    constructor() {
        this.rows = 50;
        this.cols = 26; // A to Z
        this.selectedCell = null;
        this.data = {}; // Store cell data here: { 'A1': 'value' }
        this.isDirty = false;
        
        // Range selection state
        this.selectionRange = null; // { startCol, startRow, endCol, endRow }
        this.isSelecting = false;
        this.selectionAnchor = null; // Cell ID where selection started

        // Fill handle state
        this.isFilling = false;
        this.fillStartCell = null;
        this.lastFillTargetCell = null;

        // Clipboard state
        this.clipboardData = null; // 2D array of copied values
        this.isCut = false;
        this.cutRange = null;

        // Resize state
        this.colWidths = new Array(this.cols).fill(100); // Default 100px per column
        this.rowHeights = new Array(this.rows + 1).fill(25); // Default 25px per row (1-indexed)
        this.isResizingCol = false;
        this.isResizingRow = false;
        this.resizeIndex = -1;
        this.resizeStartPos = 0;
        this.resizeStartSize = 0;

        this.init();
    }

    init() {
        this.container = document.getElementById('grid-container');
        this.formulaInput = document.getElementById('formula-input');
        this.cellAddress = document.getElementById('selected-cell-id');
        
        this.renderGrid();
        this.setupEventListeners();
    }

    // ─── Utility ───────────────────────────────────────────
    colToNumber(col) {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
            num = num * 26 + (col.charCodeAt(i) - 64);
        }
        return num;
    }

    numberToCol(num) {
        let col = '';
        while (num > 0) {
            let rem = (num - 1) % 26;
            col = String.fromCharCode(65 + rem) + col;
            num = Math.floor((num - 1) / 26);
        }
        return col;
    }

    parseCellId(id) {
        const col = id.match(/[A-Z]+/)[0];
        const row = parseInt(id.match(/\d+/)[0]);
        return { col, row, colNum: this.colToNumber(col) };
    }

    getCellEl(col, row) {
        const id = typeof col === 'number' ? `${this.numberToCol(col)}${row}` : `${col}${row}`;
        return document.querySelector(`[data-id="${id}"]`);
    }

    // ─── Grid Rendering ────────────────────────────────────
    renderGrid() {
        const table = document.createElement('table');
        table.className = 'spreadsheet-table';
        this.table = table;

        // Colgroup for dynamic column widths
        const colgroup = document.createElement('colgroup');
        const rowHeaderCol = document.createElement('col');
        rowHeaderCol.style.width = '40px';
        colgroup.appendChild(rowHeaderCol);
        for (let j = 0; j < this.cols; j++) {
            const col = document.createElement('col');
            col.style.width = `${this.colWidths[j]}px`;
            colgroup.appendChild(col);
        }
        table.appendChild(colgroup);
        this.colgroup = colgroup;
        
        // Header Row (A, B, C...)
        const headerRow = document.createElement('tr');
        const emptyHeader = document.createElement('th');
        emptyHeader.className = 'cell header row-header corner-header';
        headerRow.appendChild(emptyHeader);
        
        for (let j = 0; j < this.cols; j++) {
            const th = document.createElement('th');
            th.className = 'cell header col-header';
            th.innerText = String.fromCharCode(65 + j);
            th.dataset.colIndex = j;
            headerRow.appendChild(th);
        }
        table.appendChild(headerRow);
        
        // Data Rows
        this.tbody = document.createElement('tbody');
        table.appendChild(this.tbody);
        this.createRowElements(1, this.rows);
        
        this.container.appendChild(table);
    }

    createRowElements(startRow, endRow) {
        for (let i = startRow; i <= endRow; i++) {
            if (!this.rowHeights[i]) this.rowHeights[i] = 25;
            const tr = document.createElement('tr');
            tr.style.height = `${this.rowHeights[i]}px`;
            
            const rowHeader = document.createElement('td');
            rowHeader.className = 'cell header row-header';
            rowHeader.innerText = i;
            rowHeader.dataset.rowIndex = i;
            tr.appendChild(rowHeader);
            
            for (let j = 0; j < this.cols; j++) {
                const td = document.createElement('td');
                td.className = 'cell';
                td.contentEditable = true;
                const cellId = `${this.numberToCol(j + 1)}${i}`;
                td.dataset.id = cellId;
                
                if (this.data[cellId]) {
                    td.innerText = this.data[cellId];
                }
                
                td.addEventListener('focus', () => this.handleCellFocus(td));
                td.addEventListener('input', () => this.handleCellInput(td));
                td.addEventListener('blur', () => this.handleCellBlur(td));
                td.addEventListener('keydown', (e) => this.handleKeyDown(e));
                td.addEventListener('mousedown', (e) => this.handleCellMouseDown(td, e));
                
                tr.appendChild(td);
            }
            this.tbody.appendChild(tr);
        }
    }

    // ─── Event Listeners ───────────────────────────────────
    setupEventListeners() {
        // Toolbar buttons
        document.getElementById('btn-bold').addEventListener('click', () => {
            document.execCommand('bold', false, null);
        });
        
        document.getElementById('btn-italic').addEventListener('click', () => {
            document.execCommand('italic', false, null);
        });

        // CSV Open/Save buttons
        document.getElementById('btn-open').addEventListener('click', () => this.openFileDialog());
        document.getElementById('btn-save').addEventListener('click', () => this.exportFile());

        // Hidden file input for CSV import
        this.fileInput = document.getElementById('csv-file-input');
        this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));

        // Infinite Scroll logic to minimize browser DOM burden
        this.container.addEventListener('scroll', () => {
            // Add rows when scroll reaches near the bottom of the container
            if (this.container.scrollTop + this.container.clientHeight >= this.container.scrollHeight - 150) {
                const rowsToAdd = 30;
                this.createRowElements(this.rows + 1, this.rows + rowsToAdd);
                this.rows += rowsToAdd;
            }
        });

        // Formula bar input
        this.formulaInput.addEventListener('input', (e) => {
            if (this.selectedCell) {
                this.selectedCell.innerText = e.target.value;
                this.data[this.selectedCell.dataset.id] = e.target.value;
            }
        });

        // Mouse move/up for range selection, fill handle & resize (global)
        document.addEventListener('mousemove', (e) => {
            if (this.isResizingCol || this.isResizingRow) {
                this.handleResizeMove(e);
                return;
            }
            if (this.isSelecting) {
                const target = document.elementFromPoint(e.clientX, e.clientY);
                const cell = target ? target.closest('.cell') : null;
                if (cell && !cell.classList.contains('header') && cell.dataset.id) {
                    this.extendRangeSelection(cell);
                }
            }
            if (this.isFilling) {
                this.handleFillMove(e);
            }
        });

        document.addEventListener('mouseup', (e) => {
            if (this.isResizingCol || this.isResizingRow) {
                this.handleResizeEnd(e);
                return;
            }
            if (this.isSelecting) {
                this.endRangeSelection();
            }
            if (this.isFilling) {
                this.handleFillEnd(e);
            }
        });
        // Global keyboard shortcuts (clipboard, delete)
        document.addEventListener('keydown', (e) => {
            // Ignore if typing in formula bar
            if (e.target.id === 'formula-input') return;

            if (e.ctrlKey || e.metaKey) {
                switch (e.key.toLowerCase()) {
                    case 'c':
                        e.preventDefault();
                        this.copySelection();
                        break;
                    case 'x':
                        e.preventDefault();
                        this.cutSelection();
                        break;
                    case 'v':
                        e.preventDefault();
                        this.pasteAtSelection();
                        break;
                    case 'a':
                        e.preventDefault();
                        this.selectAll();
                        break;
                    case 's':
                        e.preventDefault();
                        this.exportFile();
                        break;
                    case 'o':
                        e.preventDefault();
                        this.openFileDialog();
                        break;
                }
            }

            if (e.key === 'Delete') {
                if (this.selectionRange) {
                    e.preventDefault();
                    this.deleteSelection();
                }
            }
        });

        // Fill Handle
        this.createFillHandle();
        this.createSelectionOverlay();
        this.createRangeOverlay();

        // Resize Handlers
        this.setupResizeHandlers();
    }

    // ─── Resize Handlers ─────────────────────────────────
    setupResizeHandlers() {
        const EDGE_ZONE = 5; // px from edge to trigger resize cursor

        // Create resize guide line
        this.resizeGuide = document.createElement('div');
        this.resizeGuide.className = 'resize-guide';
        this.resizeGuide.style.display = 'none';
        this.container.appendChild(this.resizeGuide);

        // Detect column header edge on mousemove
        this.container.addEventListener('mousemove', (e) => {
            if (this.isResizingCol || this.isResizingRow) return;

            const target = e.target;

            // Column header resize detection
            if (target.classList.contains('col-header')) {
                const rect = target.getBoundingClientRect();
                if (e.clientX >= rect.right - EDGE_ZONE) {
                    target.style.cursor = 'col-resize';
                    return;
                } else if (e.clientX <= rect.left + EDGE_ZONE && target.dataset.colIndex !== '0') {
                    target.style.cursor = 'col-resize';
                    return;
                } else {
                    target.style.cursor = '';
                }
            }

            // Row header resize detection
            if (target.classList.contains('row-header') && target.dataset.rowIndex) {
                const rect = target.getBoundingClientRect();
                if (e.clientY >= rect.bottom - EDGE_ZONE) {
                    target.style.cursor = 'row-resize';
                    return;
                } else {
                    target.style.cursor = '';
                }
            }
        });

        // Mousedown on header edge starts resize
        this.container.addEventListener('mousedown', (e) => {
            const target = e.target;

            // Column resize
            if (target.classList.contains('col-header')) {
                const rect = target.getBoundingClientRect();
                const colIdx = parseInt(target.dataset.colIndex);

                if (e.clientX >= rect.right - EDGE_ZONE) {
                    // Resize this column
                    this.startColResize(colIdx, e.clientX);
                    e.preventDefault();
                    e.stopPropagation();
                    return;
                } else if (e.clientX <= rect.left + EDGE_ZONE && colIdx > 0) {
                    // Resize previous column
                    this.startColResize(colIdx - 1, e.clientX);
                    e.preventDefault();
                    e.stopPropagation();
                    return;
                }
            }

            // Row resize
            if (target.classList.contains('row-header') && target.dataset.rowIndex) {
                const rect = target.getBoundingClientRect();
                const rowIdx = parseInt(target.dataset.rowIndex);

                if (e.clientY >= rect.bottom - EDGE_ZONE) {
                    this.startRowResize(rowIdx, e.clientY);
                    e.preventDefault();
                    e.stopPropagation();
                    return;
                }
            }
        });
    }

    startColResize(colIndex, startX) {
        this.isResizingCol = true;
        this.resizeIndex = colIndex;
        this.resizeStartPos = startX;
        this.resizeStartSize = this.colWidths[colIndex];
        document.body.style.cursor = 'col-resize';
        document.body.style.userSelect = 'none';

        // Show guide
        this.showResizeGuide('col', colIndex);
    }

    startRowResize(rowIndex, startY) {
        this.isResizingRow = true;
        this.resizeIndex = rowIndex;
        this.resizeStartPos = startY;
        this.resizeStartSize = this.rowHeights[rowIndex];
        document.body.style.cursor = 'row-resize';
        document.body.style.userSelect = 'none';

        // Show guide
        this.showResizeGuide('row', rowIndex);
    }

    handleResizeMove(e) {
        if (this.isResizingCol) {
            const delta = e.clientX - this.resizeStartPos;
            const newWidth = Math.max(30, this.resizeStartSize + delta);
            this.colWidths[this.resizeIndex] = newWidth;

            // Update colgroup
            const col = this.colgroup.children[this.resizeIndex + 1]; // +1 for row-header col
            if (col) col.style.width = `${newWidth}px`;

            this.updateResizeGuide('col', e.clientX);
        }

        if (this.isResizingRow) {
            const delta = e.clientY - this.resizeStartPos;
            const newHeight = Math.max(20, this.resizeStartSize + delta);
            this.rowHeights[this.resizeIndex] = newHeight;

            // Update row height
            const rows = this.table.querySelectorAll('tr');
            const tr = rows[this.resizeIndex]; // row 1 = tr index 1 (0 is header)
            if (tr) tr.style.height = `${newHeight}px`;

            this.updateResizeGuide('row', e.clientY);
        }
    }

    handleResizeEnd(e) {
        this.isResizingCol = false;
        this.isResizingRow = false;
        this.resizeIndex = -1;
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
        this.resizeGuide.style.display = 'none';

        // Refresh fill handle position
        this.updateFillHandlePosition();
    }

    showResizeGuide(type, index) {
        const containerRect = this.container.getBoundingClientRect();
        this.resizeGuide.style.display = 'block';

        if (type === 'col') {
            const col = this.colgroup.children[index + 1];
            // We'll position it in updateResizeGuide
            this.resizeGuide.style.top = '0px';
            this.resizeGuide.style.width = '2px';
            this.resizeGuide.style.height = `${this.container.scrollHeight}px`;
        } else {
            this.resizeGuide.style.left = '0px';
            this.resizeGuide.style.height = '2px';
            this.resizeGuide.style.width = `${this.container.scrollWidth}px`;
        }
    }

    updateResizeGuide(type, pos) {
        const containerRect = this.container.getBoundingClientRect();
        if (type === 'col') {
            this.resizeGuide.style.left = `${pos - containerRect.left + this.container.scrollLeft}px`;
        } else {
            this.resizeGuide.style.top = `${pos - containerRect.top + this.container.scrollTop}px`;
        }
    }

    // ─── Cell Focus / Input ────────────────────────────────
    handleCellFocus(cell) {
        this.selectedCell = cell;
        const cellId = cell.dataset.id;
        this.cellAddress.innerText = cellId;
        this.formulaInput.value = cell.innerText;
        this.highlightHeaders(cell);
    }

    clearHighlights() {
        document.querySelectorAll('.cell.header.active').forEach(h => h.classList.remove('active'));
    }

    highlightHeaders(cell) {
        this.clearHighlights();
        const cellId = cell.dataset.id;
        const col = cellId.match(/[A-Z]+/)[0];
        const row = cellId.match(/\d+/)[0];

        const colHeader = Array.from(document.querySelectorAll('.cell.header')).find(h => h.innerText === col);
        if (colHeader) colHeader.classList.add('active');

        const rowHeader = Array.from(document.querySelectorAll('.cell.row-header')).find(h => h.innerText === row);
        if (rowHeader) rowHeader.classList.add('active');
    }

    handleCellInput(cell) {
        const cellId = cell.dataset.id;
        this.data[cellId] = cell.innerText;
        this.formulaInput.value = cell.innerText;
        this.markDirty();
        this.updateItemCount();
    }

    handleCellBlur(cell) {
        // Save data on blur
    }

    markDirty() {
        if (this.isDirty) return;
        this.isDirty = true;
        const badge = document.querySelector('.status-badge');
        if (badge) {
            badge.innerText = 'Edited';
            badge.className = 'status-badge modified';
        }
    }

    markClean() {
        this.isDirty = false;
        const badge = document.querySelector('.status-badge');
        if (badge) {
            badge.innerText = 'Saved';
            badge.className = 'status-badge saved';
            // Optional: fade out the 'Saved' state after some time
            setTimeout(() => {
                if (!this.isDirty) {
                    badge.classList.remove('saved');
                }
            }, 3000);
        }
    }

    // ─── Range Selection ───────────────────────────────────
    handleCellMouseDown(cell, e) {
        // Don't start selection if clicking fill handle
        if (e.target.classList.contains('fill-handle')) return;

        const cellId = cell.dataset.id;
        const { col, row, colNum } = this.parseCellId(cellId);

        if (e.shiftKey && this.selectedCell) {
            // Shift+Click: extend selection from current cell
            e.preventDefault();
            const anchor = this.parseCellId(this.selectedCell.dataset.id);
            this.setSelectionRange(anchor.colNum, anchor.row, colNum, row);
            this.updateRangeVisual();
            this.updateFillHandlePosition();
            return;
        }

        // Normal click: start new selection
        this.clearRangeSelection();
        this.selectionAnchor = { colNum, row };
        this.isSelecting = true;
        this.setSelectionRange(colNum, row, colNum, row);
        this.updateRangeVisual();
    }

    extendRangeSelection(cell) {
        const { colNum, row } = this.parseCellId(cell.dataset.id);
        if (!this.selectionAnchor) return;

        this.setSelectionRange(
            this.selectionAnchor.colNum, this.selectionAnchor.row,
            colNum, row
        );
        this.updateRangeVisual();
    }

    endRangeSelection() {
        this.isSelecting = false;
        this.updateFillHandlePosition();
    }

    setSelectionRange(c1, r1, c2, r2) {
        this.selectionRange = {
            startCol: Math.min(c1, c2),
            startRow: Math.min(r1, r2),
            endCol: Math.max(c1, c2),
            endRow: Math.max(r1, r2)
        };
    }

    clearRangeSelection() {
        this.selectionRange = null;
        document.querySelectorAll('.cell.in-range').forEach(c => c.classList.remove('in-range'));
        if (this.rangeOverlay) this.rangeOverlay.style.display = 'none';
    }

    updateRangeVisual() {
        // Remove old highlights
        document.querySelectorAll('.cell.in-range').forEach(c => c.classList.remove('in-range'));

        if (!this.selectionRange) {
            if (this.rangeOverlay) this.rangeOverlay.style.display = 'none';
            return;
        }

        const { startCol, startRow, endCol, endRow } = this.selectionRange;
        const isSingleCell = (startCol === endCol && startRow === endRow);

        // Only show range visual for multi-cell selection
        if (isSingleCell) {
            if (this.rangeOverlay) this.rangeOverlay.style.display = 'none';
            return;
        }

        // Highlight cells in range
        for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
                const cell = this.getCellEl(c, r);
                if (cell) cell.classList.add('in-range');
            }
        }

        // Update border overlay
        const topLeftCell = this.getCellEl(startCol, startRow);
        const bottomRightCell = this.getCellEl(endCol, endRow);

        if (topLeftCell && bottomRightCell && this.rangeOverlay) {
            const tlRect = topLeftCell.getBoundingClientRect();
            const brRect = bottomRightCell.getBoundingClientRect();
            const containerRect = this.container.getBoundingClientRect();

            this.rangeOverlay.style.display = 'block';
            this.rangeOverlay.style.top = `${tlRect.top - containerRect.top + this.container.scrollTop}px`;
            this.rangeOverlay.style.left = `${tlRect.left - containerRect.left + this.container.scrollLeft}px`;
            this.rangeOverlay.style.width = `${brRect.right - tlRect.left}px`;
            this.rangeOverlay.style.height = `${brRect.bottom - tlRect.top}px`;
        }
    }

    getSelectedCells() {
        if (!this.selectionRange) {
            return this.selectedCell ? [this.selectedCell] : [];
        }
        const cells = [];
        const { startCol, startRow, endCol, endRow } = this.selectionRange;
        for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
                const cell = this.getCellEl(c, r);
                if (cell) cells.push(cell);
            }
        }
        return cells;
    }

    getEffectiveRange() {
        if (this.selectionRange) return this.selectionRange;
        if (this.selectedCell) {
            const { colNum, row } = this.parseCellId(this.selectedCell.dataset.id);
            return { startCol: colNum, startRow: row, endCol: colNum, endRow: row };
        }
        return null;
    }

    selectAll() {
        this.setSelectionRange(1, 1, this.cols, this.rows);
        this.updateRangeVisual();
        this.updateFillHandlePosition();
    }

    // ─── Range Overlay & Fill Handle DOM ───────────────────
    createRangeOverlay() {
        this.rangeOverlay = document.createElement('div');
        this.rangeOverlay.className = 'range-overlay';
        this.rangeOverlay.style.display = 'none';
        this.container.appendChild(this.rangeOverlay);
    }

    createFillHandle() {
        this.fillHandle = document.createElement('div');
        this.fillHandle.className = 'fill-handle';
        this.fillHandle.style.display = 'none';
        this.container.appendChild(this.fillHandle);
        
        this.fillHandle.addEventListener('mousedown', (e) => this.handleFillStart(e));
    }

    createSelectionOverlay() {
        this.selectionOverlay = document.createElement('div');
        this.selectionOverlay.className = 'selection-overlay';
        this.selectionOverlay.style.display = 'none';
        this.container.appendChild(this.selectionOverlay);
    }

    updateFillHandlePosition() {
        const range = this.getEffectiveRange();
        if (!range) {
            this.fillHandle.style.display = 'none';
            return;
        }

        const bottomRightCell = this.getCellEl(range.endCol, range.endRow);
        if (!bottomRightCell) {
            this.fillHandle.style.display = 'none';
            return;
        }

        const rect = bottomRightCell.getBoundingClientRect();
        const containerRect = this.container.getBoundingClientRect();

        this.fillHandle.style.display = 'block';
        this.fillHandle.style.left = `${rect.right - containerRect.left + this.container.scrollLeft - 5}px`;
        this.fillHandle.style.top = `${rect.bottom - containerRect.top + this.container.scrollTop - 5}px`;
    }

    // ─── Fill Handle Logic ─────────────────────────────────
    handleFillStart(e) {
        const range = this.getEffectiveRange();
        if (!range) return;
        this.isFilling = true;
        this.fillRange = { ...range };
        this.lastFillTargetCell = null;
        this.fillHandle.style.pointerEvents = 'none';
        this.selectionOverlay.style.display = 'block';
        e.preventDefault();
        e.stopPropagation();
    }

    handleFillMove(e) {
        if (!this.isFilling) return;
        
        const target = document.elementFromPoint(e.clientX, e.clientY);
        const cell = target ? target.closest('.cell') : null;
        
        if (cell && !cell.classList.contains('header') && cell.dataset.id) {
            this.lastFillTargetCell = cell;
            // Show overlay from fill range start to target
            const startCell = this.getCellEl(this.fillRange.startCol, this.fillRange.startRow);
            if (startCell) {
                this.updateSelectionOverlayBetween(startCell, cell);
            }
        }
    }

    updateSelectionOverlayBetween(startCell, endCell) {
        const startRect = startCell.getBoundingClientRect();
        const endRect = endCell.getBoundingClientRect();
        const containerRect = this.container.getBoundingClientRect();

        const top = Math.min(startRect.top, endRect.top) - containerRect.top + this.container.scrollTop;
        const left = Math.min(startRect.left, endRect.left) - containerRect.left + this.container.scrollLeft;
        const width = Math.max(startRect.right, endRect.right) - Math.min(startRect.left, endRect.left);
        const height = Math.max(startRect.bottom, endRect.bottom) - Math.min(startRect.top, endRect.top);

        this.selectionOverlay.style.top = `${top}px`;
        this.selectionOverlay.style.left = `${left}px`;
        this.selectionOverlay.style.width = `${width}px`;
        this.selectionOverlay.style.height = `${height}px`;
    }

    handleFillEnd(e) {
        if (!this.isFilling) return;
        this.isFilling = false;
        this.selectionOverlay.style.display = 'none';
        this.fillHandle.style.pointerEvents = 'auto';

        if (this.lastFillTargetCell) {
            this.fillFromRange(this.fillRange, this.lastFillTargetCell);
        }
        this.updateFillHandlePosition();
    }

    fillFromRange(sourceRange, targetCell) {
        const target = this.parseCellId(targetCell.dataset.id);
        const { startCol, startRow, endCol, endRow } = sourceRange;
        const rangeCols = endCol - startCol + 1;
        const rangeRows = endRow - startRow + 1;

        // Collect source values as 2D array
        const sourceValues = [];
        for (let r = startRow; r <= endRow; r++) {
            const rowVals = [];
            for (let c = startCol; c <= endCol; c++) {
                const cell = this.getCellEl(c, r);
                rowVals.push(cell ? cell.innerText : '');
            }
            sourceValues.push(rowVals);
        }

        // Determine fill direction
        if (target.colNum >= startCol && target.colNum <= endCol) {
            // Vertical fill
            const fillStart = target.row > endRow ? endRow + 1 : (target.row < startRow ? target.row : startRow);
            const fillEnd = target.row > endRow ? target.row : (target.row < startRow ? startRow - 1 : endRow);

            for (let r = fillStart; r <= fillEnd; r++) {
                const srcRowIdx = ((r - fillStart) % rangeRows);
                for (let c = startCol; c <= endCol; c++) {
                    const srcColIdx = c - startCol;
                    const value = sourceValues[srcRowIdx][srcColIdx];
                    const cell = this.getCellEl(c, r);
                    if (cell) {
                        cell.innerText = value;
                        this.data[cell.dataset.id] = value;
                    }
                }
            }
        } else if (target.row >= startRow && target.row <= endRow) {
            // Horizontal fill
            const fillStart = target.colNum > endCol ? endCol + 1 : (target.colNum < startCol ? target.colNum : startCol);
            const fillEnd = target.colNum > endCol ? target.colNum : (target.colNum < startCol ? startCol - 1 : endCol);

            for (let c = fillStart; c <= fillEnd; c++) {
                const srcColIdx = ((c - fillStart) % rangeCols);
                for (let r = startRow; r <= endRow; r++) {
                    const srcRowIdx = r - startRow;
                    const value = sourceValues[srcRowIdx][srcColIdx];
                    const cell = this.getCellEl(c, r);
                    if (cell) {
                        cell.innerText = value;
                        this.data[cell.dataset.id] = value;
                    }
                }
            }
        }
    }

    // ─── Clipboard ─────────────────────────────────────────
    copySelection() {
        const range = this.getEffectiveRange();
        if (!range) return;

        const { startCol, startRow, endCol, endRow } = range;
        const rows = [];
        this.clipboardData = [];

        for (let r = startRow; r <= endRow; r++) {
            const rowVals = [];
            const rowText = [];
            for (let c = startCol; c <= endCol; c++) {
                const cell = this.getCellEl(c, r);
                const val = cell ? cell.innerText : '';
                rowVals.push(val);
                rowText.push(val);
            }
            this.clipboardData.push(rowVals);
            rows.push(rowText.join('\t'));
        }

        this.isCut = false;
        this.cutRange = null;

        // Copy to system clipboard as TSV
        const text = rows.join('\n');
        navigator.clipboard.writeText(text).catch(() => {});

        // Flash visual feedback
        this.flashCopyBorder(range);
    }

    cutSelection() {
        const range = this.getEffectiveRange();
        if (!range) return;

        this.copySelection();
        this.isCut = true;
        this.cutRange = { ...range };

        // Add dashed border for cut visual
        this.flashCutBorder(range);
    }

    async pasteAtSelection() {
        if (!this.selectedCell) return;

        const anchor = this.parseCellId(this.selectedCell.dataset.id);
        let rowsData = [];
        let isInternalPaste = false;

        try {
            // Priority: Attempt to read from system clipboard first
            const text = await navigator.clipboard.readText();
            
            // Convert internal clipboard to TSV string for comparison if needed
            const internalTsv = this.clipboardData ? this.clipboardData.map(r => r.join('\t')).join('\n') : null;

            if (text && text.trim() !== '' && text !== internalTsv) {
                rowsData = this.parseTSV(text);
                isInternalPaste = false;
            } else if (this.clipboardData && this.clipboardData.length > 0) {
                rowsData = this.clipboardData;
                isInternalPaste = true;
            } else if (text && text.trim() !== '') {
                rowsData = this.parseTSV(text);
                isInternalPaste = false;
            }
        } catch (err) {
            console.warn('System clipboard access denied, using internal data:', err);
            if (this.clipboardData) {
                rowsData = this.clipboardData;
                isInternalPaste = true;
            }
        }

        if (rowsData.length === 0) return;

        // Perform Paste
        const numRows = rowsData.length;
        const numCols = rowsData[0].length;

        for (let r = 0; r < numRows; r++) {
            for (let c = 0; c < numCols; c++) {
                const targetRow = anchor.row + r;
                const targetColNum = anchor.colNum + c;

                // Auto-expand rows if pasting beyond current limit
                if (targetRow > this.rows) {
                    const rowsToAdd = Math.max(30, targetRow - this.rows);
                    this.createRowElements(this.rows + 1, this.rows + rowsToAdd);
                    this.rows += rowsToAdd;
                }

                if (targetColNum <= this.cols) {
                    const cell = this.getCellEl(targetColNum, targetRow);
                    if (cell) {
                        const val = rowsData[r][c] || '';
                        cell.innerText = val;
                        this.data[cell.dataset.id] = val;
                    }
                }
            }
        }

        // If it was an internal cut, clear source cells
        if (isInternalPaste && this.isCut && this.cutRange) {
            const { startCol, startRow, endCol, endRow } = this.cutRange;
            for (let r = startRow; r <= endRow; r++) {
                for (let c = startCol; c <= endCol; c++) {
                    const cell = this.getCellEl(c, r);
                    if (cell) {
                        cell.innerText = '';
                        this.data[cell.dataset.id] = '';
                    }
                }
            }
            this.isCut = false;
            this.cutRange = null;
            // Clear dashed border
            if (this.rangeOverlay) this.rangeOverlay.classList.remove('cut-dashed');
        }

        // Update State & UI
        this.markDirty();
        this.updateItemCount();

        // Select the pasted range
        this.setSelectionRange(anchor.colNum, anchor.row, anchor.colNum + numCols - 1, anchor.row + numRows - 1);
        this.updateRangeVisual();
        this.updateFillHandlePosition();
    }

    parseTSV(text) {
        const rows = [];
        let currentRow = [];
        let currentField = '';
        let inQuotes = false;

        // Normalize newlines
        const cleanText = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

        for (let i = 0; i < cleanText.length; i++) {
            const char = cleanText[i];
            const nextChar = cleanText[i + 1];

            if (inQuotes) {
                if (char === '"' && nextChar === '"') {
                    currentField += '"';
                    i++;
                } else if (char === '"') {
                    inQuotes = false;
                } else {
                    currentField += char;
                }
            } else {
                if (char === '"') {
                    inQuotes = true;
                } else if (char === '\t') {
                    currentRow.push(currentField);
                    currentField = '';
                } else if (char === '\n') {
                    currentRow.push(currentField);
                    rows.push(currentRow);
                    currentRow = [];
                    currentField = '';
                } else {
                    currentField += char;
                }
            }
        }

        if (currentField !== '' || currentRow.length > 0) {
            currentRow.push(currentField);
            rows.push(currentRow);
        }

        // Remove potential empty trailing row often added by Excel copies
        if (rows.length > 1 && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '') {
            rows.pop();
        }

        return rows;
    }

    deleteSelection() {
        const range = this.getEffectiveRange();
        if (!range) return;

        const { startCol, startRow, endCol, endRow } = range;
        let changed = false;
        for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
                const cell = this.getCellEl(c, r);
                if (cell) {
                    cell.innerText = '';
                    this.data[cell.dataset.id] = '';
                    changed = true;
                }
            }
        }
        if (changed) {
            this.markDirty();
            this.updateItemCount();
        }
    }

    flashCopyBorder(range) {
        if (!this.rangeOverlay) return;
        this.rangeOverlay.classList.add('copy-flash');
        setTimeout(() => this.rangeOverlay.classList.remove('copy-flash'), 500);
    }

    flashCutBorder(range) {
        if (!this.rangeOverlay) return;
        this.rangeOverlay.classList.add('cut-dashed');
    }

    // ─── Keyboard Navigation ───────────────────────────────
    handleKeyDown(e) {
        const activeCell = document.activeElement.closest('.cell');
        if (!activeCell || activeCell.classList.contains('header')) return;
        if (e.target.id === 'formula-input') return;

        // Ctrl/Meta shortcuts handled globally
        if (e.ctrlKey || e.metaKey) return;

        // Delete/Backspace for range
        if (e.key === 'Delete' || e.key === 'Backspace') {
            if (this.selectionRange) {
                const { startCol, startRow, endCol, endRow } = this.selectionRange;
                if (startCol !== endCol || startRow !== endRow) {
                    e.preventDefault();
                    this.deleteSelection();
                    return;
                }
            }
        }

        const cellId = activeCell.dataset.id;
        const colChar = cellId.match(/[A-Z]+/)[0];
        const colNum = this.colToNumber(colChar);
        const rowNum = parseInt(cellId.match(/\d+/)[0]);
        let nextCol = colNum;
        let nextRow = rowNum;
        let moved = false;

        switch(e.key) {
            case 'ArrowUp':
                if (e.shiftKey) {
                    e.preventDefault();
                    this.extendSelectionByKey(colNum, rowNum, colNum, rowNum - 1);
                    return;
                }
                if (rowNum > 1) nextRow--;
                moved = true;
                break;
            case 'ArrowDown':
                if (e.shiftKey) {
                    e.preventDefault();
                    this.extendSelectionByKey(colNum, rowNum, colNum, rowNum + 1);
                    return;
                }
                if (rowNum < this.rows) nextRow++;
                moved = true;
                break;
            case 'ArrowLeft':
                if (e.shiftKey) {
                    e.preventDefault();
                    this.extendSelectionByKey(colNum, rowNum, colNum - 1, rowNum);
                    return;
                }
                if (colNum > 1) nextCol = colNum - 1;
                moved = true;
                break;
            case 'ArrowRight':
                if (e.shiftKey) {
                    e.preventDefault();
                    this.extendSelectionByKey(colNum, rowNum, colNum + 1, rowNum);
                    return;
                }
                if (colNum < this.cols) nextCol = colNum + 1;
                moved = true;
                break;
            case 'Tab':
                e.preventDefault();
                if (e.shiftKey) {
                    if (colNum > 1) nextCol = colNum - 1;
                } else {
                    if (colNum < this.cols) nextCol = colNum + 1;
                }
                moved = true;
                break;
            case 'Enter':
                if (e.altKey) {
                    e.preventDefault();
                    e.stopPropagation();
                    document.execCommand('insertLineBreak');
                    return;
                }
                e.preventDefault();
                e.stopPropagation();
                if (e.shiftKey) {
                    if (rowNum > 1) nextRow--;
                } else {
                    if (rowNum < this.rows) nextRow++;
                }
                moved = true;
                break;
            case 'Escape':
                this.clearRangeSelection();
                this.updateFillHandlePosition();
                return;
        }

        if (moved) {
            this.clearRangeSelection();
            const nextCellId = `${this.numberToCol(nextCol)}${nextRow}`;
            const nextCell = document.querySelector(`[data-id="${nextCellId}"]`);
            if (nextCell) {
                e.preventDefault();
                e.stopPropagation();
                nextCell.focus();
                this.handleCellFocus(nextCell);
                // Set single-cell range
                this.setSelectionRange(nextCol, nextRow, nextCol, nextRow);
                this.updateFillHandlePosition();
            }
        }
    }

    extendSelectionByKey(anchorCol, anchorRow, targetCol, targetRow) {
        if (targetCol < 1 || targetCol > this.cols || targetRow < 1 || targetRow > this.rows) return;

        // Use existing anchor or current cell as anchor
        if (!this.selectionAnchor) {
            this.selectionAnchor = { colNum: anchorCol, row: anchorRow };
        }

        // Get the furthest extent already selected
        const currentRange = this.selectionRange;
        let extendCol = targetCol;
        let extendRow = targetRow;

        if (currentRange) {
            // Extend from the edge furthest from anchor
            if (targetRow !== anchorRow) {
                extendRow = targetRow > anchorRow
                    ? Math.max(currentRange.endRow, targetRow)
                    : Math.min(currentRange.startRow, targetRow);
                extendCol = currentRange.endCol !== this.selectionAnchor.colNum ? currentRange.endCol : currentRange.startCol;
            }
            if (targetCol !== anchorCol) {
                extendCol = targetCol > anchorCol
                    ? Math.max(currentRange.endCol, targetCol)
                    : Math.min(currentRange.startCol, targetCol);
                extendRow = currentRange.endRow !== this.selectionAnchor.row ? currentRange.endRow : currentRange.startRow;
            }
        }

        this.setSelectionRange(this.selectionAnchor.colNum, this.selectionAnchor.row, extendCol, extendRow);
        this.updateRangeVisual();
        this.updateFillHandlePosition();

        // Move focus to the target cell
        const targetCell = this.getCellEl(targetCol, targetRow);
        if (targetCell) {
            targetCell.focus();
            this.selectedCell = targetCell;
            this.cellAddress.innerText = targetCell.dataset.id;
            this.formulaInput.value = targetCell.innerText;
        }
    }

    // ─── File Import / Export ───────────────────────────────
    openFileDialog() {
        this.fileInput.value = ''; // Reset so same file can be reopened
        this.fileInput.click();
    }

    handleFileSelect(e) {
        const file = e.target.files[0];
        if (!file) return;

        // Check if there's existing data
        const hasData = Object.keys(this.data).some(k => this.data[k] && this.data[k].trim() !== '');
        if (hasData) {
            if (!confirm('기존 데이터가 있거나 작업 중인 내용이 덮어씌워질 수 있습니다. 계속할까요?\n(Existing data may be overwritten. Continue?)')) {
                return;
            }
        }

        const extension = file.name.split('.').pop().toLowerCase();
        const reader = new FileReader();

        reader.onload = (event) => {
            try {
                if (extension === 'xlsx' || extension === 'xls') {
                    this.importXLSX(event.target.result);
                } else if (extension === 'vsht') {
                    this.importVSHT(event.target.result);
                } else {
                    // Fallback to text parsing (CSV/TSV/TXT)
                    this.importFromText(event.target.result);
                }

                // Update filename display
                const filenameEl = document.querySelector('.filename');
                if (filenameEl) {
                    filenameEl.innerText = file.name.replace(/\.[^.]+$/, '');
                }
            } catch (err) {
                console.error('File import failed:', err);
                alert('파일을 불러오는 데 실패했습니다. 지원되는 형식인지 확인해 주세요.\n(Failed to load file. Please check the format.)');
            }
        };

        if (extension === 'xlsx' || extension === 'xls') {
            reader.readAsArrayBuffer(file);
        } else {
            reader.readAsText(file, 'UTF-8');
        }
    }

    // 1. .vsht Import (JSON based, full layout)
    importVSHT(jsonText) {
        const doc = JSON.parse(jsonText);
        
        // Clear current
        this.clearAllData(false);

        // Restore Metadata
        if (doc.colWidths) this.colWidths = doc.colWidths;
        if (doc.rowHeights) this.rowHeights = doc.rowHeights;
        this.data = doc.data || {};

        // Re-render or Refresh UI
        this.refreshGridUI();
        
        this.updateItemCount();
        this.markClean();
    }

    // 2. .xlsx Import (using SheetJS)
    importXLSX(buffer) {
        if (typeof XLSX === 'undefined') {
            alert('Excel 라이브러리를 불러오지 못했습니다. 네트워크 연결을 확인해 주세요.');
            return;
        }

        const workbook = XLSX.read(buffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert to 2D array
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        this.clearAllData(false);
        
        jsonData.forEach((row, i) => {
            row.forEach((cellValue, j) => {
                const cellId = `${this.numberToCol(j + 1)}${i + 1}`;
                const val = cellValue === null || cellValue === undefined ? '' : String(cellValue);
                this.data[cellId] = val;
            });
        });

        this.refreshGridUI();
        this.updateItemCount();
        this.markClean();
    }

    // 3. CSV/TSV Text Import
    importFromText(text) {
        // Clear existing
        this.clearAllData(false);

        // Remove BOM if present
        if (text.charCodeAt(0) === 0xFEFF) {
            text = text.substring(1);
        }

        const delimiter = this.detectDelimiter(text);
        
        // Advanced Parsing (handles multi-line cells)
        let row = 0;
        let col = 0;
        let currentField = '';
        let inQuotes = false;

        for (let i = 0; i < text.length; i++) {
            const ch = text[i];
            const nextCh = text[i + 1];

            if (inQuotes) {
                if (ch === '"') {
                    if (nextCh === '"') {
                        currentField += '"';
                        i++; // Skip escaped quote
                    } else {
                        inQuotes = false;
                    }
                } else {
                    currentField += ch;
                }
            } else {
                if (ch === '"') {
                    inQuotes = true;
                } else if (ch === delimiter) {
                    this.setInternalData(row + 1, col + 1, currentField);
                    currentField = '';
                    col++;
                } else if (ch === '\r' && nextCh === '\n') {
                    this.setInternalData(row + 1, col + 1, currentField);
                    currentField = '';
                    row++;
                    col = 0;
                    i++; // Skip \n
                } else if (ch === '\n' || ch === '\r') {
                    this.setInternalData(row + 1, col + 1, currentField);
                    currentField = '';
                    row++;
                    col = 0;
                } else {
                    currentField += ch;
                }
            }
        }
        
        if (currentField !== '' || col > 0) {
            this.setInternalData(row + 1, col + 1, currentField);
        }

        this.refreshGridUI();
        this.updateItemCount();
        this.markClean();
    }

    setInternalData(rowNum, colNum, value) {
        // Automatically grow if needed (just in case)
        if (rowNum > this.rows) {
            this.createRowElements(this.rows + 1, rowNum);
            this.rows = rowNum;
        }
        const cellId = `${this.numberToCol(colNum)}${rowNum}`;
        this.data[cellId] = value;
    }

    refreshGridUI() {
        // Clear current table content and re-render or just update exists
        // Easiest is to update the <tbody> content
        if (this.tbody) {
            this.tbody.innerHTML = '';
            this.createRowElements(1, Math.max(50, this.rows));
        }

        // Apply column widths
        if (this.colgroup) {
            Array.from(this.colgroup.children).forEach((colEl, idx) => {
                if (idx > 0) { // skip row header col
                    colEl.style.width = `${this.colWidths[idx - 1]}px`;
                }
            });
        }
    }

    detectDelimiter(text) {
        const firstLine = text.split('\n')[0] || '';
        const commas = (firstLine.match(/,/g) || []).length;
        const tabs = (firstLine.match(/\t/g) || []).length;
        const semis = (firstLine.match(/;/g) || []).length;

        if (tabs >= commas && tabs >= semis && tabs > 0) return '\t';
        if (semis > commas && semis > 0) return ';';
        return ',';
    }

    async exportFile() {
        // Prepare VSHT Data
        const vshtData = {
            version: "1.0",
            title: document.querySelector('.filename')?.innerText || 'Untitled',
            data: this.data,
            colWidths: this.colWidths,
            rowHeights: this.rowHeights,
            rows: this.rows,
            cols: this.cols
        };

        const jsonString = JSON.stringify(vshtData, null, 2);

        // Filename
        const filenameEl = document.querySelector('.filename');
        const defaultName = filenameEl ? filenameEl.innerText.trim() : 'VibrantSheets';

        // Try File System Access API
        if ('showSaveFilePicker' in window) {
            try {
                const handle = await window.showSaveFilePicker({
                    suggestedName: `${defaultName}.vsht`,
                    types: [
                        {
                            description: 'VibrantSheets Document (.vsht)',
                            accept: { 'application/json': ['.vsht'] },
                        },
                        {
                            description: 'CSV File (.csv)',
                            accept: { 'text/csv': ['.csv'] },
                        }
                    ],
                });
                
                const writable = await handle.createWritable();
                const fileName = handle.name;
                
                if (fileName.endsWith('.csv')) {
                    // Export as CSV
                    const csvContent = this.generateCSVContent();
                    await writable.write(csvContent);
                } else {
                    // Export as .vsht (JSON)
                    await writable.write(jsonString);
                }
                
                await writable.close();

                if (filenameEl) {
                    filenameEl.innerText = handle.name.replace(/\.[^.]+$/, '');
                }
                this.markClean();
                return;
            } catch (err) {
                if (err.name === 'AbortError') return;
                console.error('File System Access API failed:', err);
            }
        }

        // Fallback Download
        const blob = new Blob([jsonString], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${defaultName}.vsht`;
        link.click();
        URL.revokeObjectURL(url);
        this.markClean();
    }

    generateCSVContent() {
        // Find used range
        let maxRow = 0, maxCol = 0;
        for (const key in this.data) {
            if (this.data[key] && this.data[key].trim() !== '') {
                const { colNum, row } = this.parseCellId(key);
                maxRow = Math.max(maxRow, row);
                maxCol = Math.max(maxCol, colNum);
            }
        }

        const rows = [];
        for (let r = 1; r <= maxRow; r++) {
            const rowFields = [];
            for (let c = 1; c <= maxCol; c++) {
                const cellId = `${this.numberToCol(c)}${r}`;
                let val = this.data[cellId] || '';
                if (val.includes(',') || val.includes('"') || val.includes('\n')) {
                    val = '"' + val.replace(/"/g, '""') + '"';
                }
                rowFields.push(val);
            }
            rows.push(rowFields.join(','));
        }
        return '\uFEFF' + rows.join('\r\n');
    }

    clearAllData(shouldMarkDirty = true) {
        if (this.tbody) {
            this.tbody.querySelectorAll('.cell').forEach(cell => {
                cell.innerText = '';
            });
        }
        this.data = {};
        if (shouldMarkDirty) this.markDirty();
        this.updateItemCount();
    }

    updateItemCount() {
        const count = Object.keys(this.data).filter(k => this.data[k] && this.data[k].trim() !== '').length;
        const metricsSpan = document.querySelector('.metrics span:first-child');
        if (metricsSpan) {
            metricsSpan.innerText = `Items: ${count}`;
        }
    }
}

// Initialize on load
document.addEventListener('DOMContentLoaded', () => {
    window.sheets = new VibrantSheets();
});
