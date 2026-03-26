class VibrantSheets {
    constructor() {
        this.baseRows = 50;
        this.baseCols = 26; // A to Z
        this.selectedCell = null;
        this.isDirty = false;
        this.fileHandle = null; // Current working file handle
        this.isEditing = false; // State: Ready (false) vs Edit (true)
        this.originalValue = ""; // Backup for Esc key (cancel edit)
        this.needsOverwrite = false; // Enter mode overwrite flag
        this.isComposing = false; // IME composition state
        
        // Range selection state
        this.selectionRange = null; // { startCol, startRow, endCol, endRow }
        this.isSelecting = false;
        this.selectionAnchor = null; // Cell ID where selection started
        this.isHeaderSelecting = false;
        this.headerSelectType = null; // 'row' | 'col'
        this.headerSelectionAnchor = null; // { type, index }

        // Fill handle state
        this.isFilling = false;
        this.fillStartCell = null;
        this.lastFillTargetCell = null;
        this.fillPreviewMode = 'series';
        this.fillSkipBlanks = false;
        this.customLists = this.loadCustomLists();

        // Clipboard state
        this.clipboardData = null; // 2D array of copied values
        this.isCut = false;
        this.cutRange = null;

        // Resize state
        this.isResizingCol = false;
        this.isResizingRow = false;
        this.resizeIndex = -1;
        this.resizeStartPos = 0;
        this.resizeStartSize = 0;

        this.sheets = [this.createSheet('Sheet1')];
        this.activeSheetIndex = 0;
        this.findState = {
            query: '',
            replace: '',
            matchCase: false,
            exact: false,
            matches: [],
            currentIndex: -1
        };
        this.csvConfirmInProgress = false;
        this.xlsxStyleWarnInProgress = false;
        this.sheetClickTimer = null;
        this.formulaEngine = typeof FormulaEngine !== 'undefined' ? new FormulaEngine() : null;
        this.formulaCache = new Map();
        
        // Border selection state
        this.currentBorderStyle = 'solid-1';
        this.currentBorderType = 'all';

        this.init();
    }

    async confirmCsvSingleSheet() {
        return window.VSIO.confirmCsvSingleSheet(this);
    }

    hasAnyStyles() {
        return this.sheets.some(sheet => Object.keys(sheet.cellStyles || {}).length > 0);
    }

    async confirmXlsxStyleWarning() {
        return window.VSIO.confirmXlsxStyleWarning(this);
    }

    get activeSheet() {
        return this.sheets[this.activeSheetIndex];
    }

    get rows() {
        return this.activeSheet.rows;
    }

    set rows(value) {
        this.activeSheet.rows = value;
    }

    get cols() {
        return this.activeSheet.cols;
    }

    set cols(value) {
        this.activeSheet.cols = value;
    }

    get data() {
        return this.activeSheet.data;
    }

    set data(value) {
        this.activeSheet.data = value;
    }

    get cellStyles() {
        return this.activeSheet.cellStyles;
    }

    set cellStyles(value) {
        this.activeSheet.cellStyles = value;
    }

    get cellFormats() {
        return this.activeSheet.cellFormats;
    }

    set cellFormats(value) {
        this.activeSheet.cellFormats = value;
    }

    get cellFormulas() {
        return this.activeSheet.cellFormulas;
    }

    set cellFormulas(value) {
        this.activeSheet.cellFormulas = value;
    }

    get cellBorders() {
        return this.activeSheet.cellBorders;
    }

    set cellBorders(value) {
        this.activeSheet.cellBorders = value;
    }

    get mergedRanges() {
        return this.activeSheet.mergedRanges;
    }

    set mergedRanges(value) {
        this.activeSheet.mergedRanges = value;
    }

    get colWidths() {
        return this.activeSheet.colWidths;
    }

    set colWidths(value) {
        this.activeSheet.colWidths = value;
    }

    get rowHeights() {
        return this.activeSheet.rowHeights;
    }

    set rowHeights(value) {
        this.activeSheet.rowHeights = value;
    }

    createSheet(name) {
        return {
            name,
            rows: this.baseRows,
            cols: this.baseCols,
            data: {},
            cellStyles: {},
            cellFormats: {},
            cellFormulas: {},
            cellBorders: {},
            mergedRanges: [],
            colWidths: new Array(this.baseCols).fill(100),
            rowHeights: new Array(this.baseRows + 1).fill(25)
        };
    }

    init() {
        this.container = document.getElementById('grid-container');
        this.formulaInput = document.getElementById('formula-input');
        this.cellAddress = document.getElementById('selected-cell-id');
        this.sheetTabs = document.querySelector('.sheet-tabs');
        
        // Create Selection & Resize Visuals
        this.selectionOverlay = this.createOverlay('selection-overlay');
        this.rangeOverlay = this.createOverlay('range-overlay');
        
        this.fillHandle = this.createOverlay('fill-handle');
        this.fillHandle.addEventListener('mousedown', (e) => this.handleFillStart(e));
        
        this.resizeGuide = this.createOverlay('resize-guide');
        this.fillPreview = this.createOverlay('fill-preview');

        this.renderGrid();
        this.renderSheetTabs();
        this.setupEventListeners();
    }

    createOverlay(className) {
        const div = document.createElement('div');
        div.className = className;
        div.style.display = 'none';
        this.container.appendChild(div);
        if (className === 'fill-preview') {
            div.innerHTML = '<span class="fill-preview-label">시리즈 채우기</span><span class="fill-preview-hint">Alt: 값 복사</span>';
        }
        return div;
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

    getRawValue(cellId) {
        const normalizedId = this.normalizeMergedCellId(cellId);
        const formula = this.cellFormulas[normalizedId];
        if (formula) return '=' + formula;
        const val = this.data[normalizedId];
        return val === undefined || val === null ? '' : String(val);
    }

    setRawValue(cellId, value) {
        const normalizedId = this.normalizeMergedCellId(cellId);
        const text = value === undefined || value === null ? '' : String(value);
        if (text.trim().startsWith('=')) {
            this.cellFormulas[normalizedId] = text.trim().slice(1);
            delete this.data[normalizedId];
        } else {
            delete this.cellFormulas[normalizedId];
            this.data[normalizedId] = text;
        }
    }

    getRawValueForSheet(sheet, cellId) {
        const normalizedId = this.normalizeMergedCellId(cellId, sheet);
        const formula = sheet.cellFormulas?.[normalizedId];
        if (formula) return '=' + formula;
        const val = sheet.data[normalizedId];
        return val === undefined || val === null ? '' : String(val);
    }

    setRawValueForSheet(sheet, cellId, value) {
        const normalizedId = this.normalizeMergedCellId(cellId, sheet);
        const text = value === undefined || value === null ? '' : String(value);
        if (text.trim().startsWith('=')) {
            sheet.cellFormulas[normalizedId] = text.trim().slice(1);
            delete sheet.data[normalizedId];
        } else {
            delete sheet.cellFormulas[normalizedId];
            sheet.data[normalizedId] = text;
        }
    }

    getDefaultDecimalsByType(type) {
        if (type === 'currency' || type === 'percentage') return 2;
        return null;
    }

    normalizeDecimals(decimals) {
        if (decimals === null || decimals === undefined || decimals === '') return null;
        const n = Number(decimals);
        if (!Number.isFinite(n)) return null;
        return Math.max(0, Math.min(10, Math.round(n)));
    }

    normalizeFormat(format) {
        const type = ['general', 'currency', 'percentage', 'date'].includes(format?.type) ? format.type : 'general';
        let decimals = this.normalizeDecimals(format?.decimals);
        if (decimals === null) decimals = this.getDefaultDecimalsByType(type);
        if (type === 'date') decimals = null;
        return { type, decimals };
    }

    isDefaultFormat(format) {
        return format.type === 'general' && format.decimals === null;
    }

    getCellFormat(cellId) {
        const normalizedId = this.normalizeMergedCellId(cellId);
        const stored = this.cellFormats[normalizedId] || {};
        return this.normalizeFormat(stored);
    }

    getCellFormatForSheet(sheet, cellId) {
        const normalizedId = this.normalizeMergedCellId(cellId, sheet);
        const stored = sheet.cellFormats[normalizedId] || {};
        return this.normalizeFormat(stored);
    }

    setCellFormat(cellId, format) {
        const normalizedId = this.normalizeMergedCellId(cellId);
        const normalized = this.normalizeFormat(format);
        if (this.isDefaultFormat(normalized)) {
            delete this.cellFormats[normalizedId];
        } else {
            this.cellFormats[normalizedId] = normalized;
        }
    }

    parseNumberFromRaw(rawValue, type = 'general') {
        if (rawValue === null || rawValue === undefined) return null;
        const text = String(rawValue).trim();
        if (!text) return null;

        const normalized = text.replace(/,/g, '').replace(/\s+/g, '');
        const hasPercentSign = normalized.includes('%');
        const cleaned = normalized.replace(/[^0-9.\-+]/g, '');
        if (!cleaned || cleaned === '-' || cleaned === '+' || cleaned === '.' || cleaned === '-.' || cleaned === '+.') {
            return null;
        }
        const parsed = Number(cleaned);
        if (!Number.isFinite(parsed)) return null;

        if (type === 'percentage' && hasPercentSign) return parsed / 100;
        return parsed;
    }

    formatDate(rawValue) {
        if (rawValue === null || rawValue === undefined) return '';
        const text = String(rawValue).trim();
        if (!text) return '';
        const dt = new Date(text);
        if (Number.isNaN(dt.getTime())) return String(rawValue);
        return dt.toLocaleDateString('ko-KR');
    }

    parseDateFromRaw(rawValue) {
        if (rawValue === null || rawValue === undefined) return null;
        const text = String(rawValue).trim();
        if (!text) return null;
        const dt = new Date(text);
        if (Number.isNaN(dt.getTime())) return null;
        return dt;
    }

    toExcelDateSerial(dateObj) {
        const excelEpoch = Date.UTC(1899, 11, 30);
        const utc = Date.UTC(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
        return (utc - excelEpoch) / 86400000;
    }

    fromExcelDateSerial(serial) {
        if (!Number.isFinite(serial)) return null;
        const excelEpoch = Date.UTC(1899, 11, 30);
        const utcMillis = excelEpoch + Math.round(serial * 86400000);
        const dt = new Date(utcMillis);
        if (Number.isNaN(dt.getTime())) return null;
        const yyyy = dt.getUTCFullYear();
        const mm = String(dt.getUTCMonth() + 1).padStart(2, '0');
        const dd = String(dt.getUTCDate()).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}`;
    }

    getFormattedValue(cellId, rawValue) {
        if (rawValue === null || rawValue === undefined) return '';
        let raw = String(rawValue);
        if (this.formulaEngine && raw.trim().startsWith('=')) {
            raw = String(this.formulaEngine.evaluate(raw, this.getFormulaContext(), new Set()));
        }
        if (!raw) return '';

        const format = this.getCellFormat(cellId);
        const decimals = format.decimals ?? 0;

        if (format.type === 'date') {
            return this.formatDate(raw);
        }

        const numeric = this.parseNumberFromRaw(raw, format.type);
        if (numeric === null) return raw;

        if (format.type === 'currency') {
            return new Intl.NumberFormat('ko-KR', {
                style: 'currency',
                currency: 'KRW',
                minimumFractionDigits: decimals,
                maximumFractionDigits: decimals
            }).format(numeric);
        }

        if (format.type === 'percentage') {
            return new Intl.NumberFormat('ko-KR', {
                minimumFractionDigits: decimals,
                maximumFractionDigits: decimals
            }).format(numeric * 100) + '%';
        }

        if (format.decimals !== null) {
            return new Intl.NumberFormat('ko-KR', {
                minimumFractionDigits: decimals,
                maximumFractionDigits: decimals
            }).format(numeric);
        }

        return raw;
    }

    renderCellValue(cell) {
        if (!cell || !cell.dataset?.id) return;
        const cellId = cell.dataset.id;
        let rawValue = this.getRawValue(cellId);
        if (this.formulaEngine && rawValue.trim().startsWith('=')) {
            rawValue = String(this.formulaEngine.evaluate(rawValue, this.getFormulaContext(), new Set()));
        }
        cell.innerText = this.getFormattedValue(cellId, rawValue);
    }

    getFormulaContext() {
        return {
            getCellValue: (cellId, stack) => {
                const key = this.normalizeCellRef(cellId);
                const safeStack = stack || new Set();
                if (safeStack.has(key)) return '#CYCLE';
                if (this.formulaCache.has(key)) return this.formulaCache.get(key);

                safeStack.add(key);
                const val = this.getRawValue(key);
                let result = val;
                if (val.trim().startsWith('=') && this.formulaEngine) {
                    result = this.formulaEngine.evaluate(val, this.getFormulaContext(), safeStack);
                }
                safeStack.delete(key);
                this.formulaCache.set(key, result);
                return result;
            },
            getRangeValues: (range, stack) => {
                const [start, end] = range.split(':');
                if (!start || !end) return [];
                const a = this.parseCellId(this.normalizeCellRef(start));
                const b = this.parseCellId(this.normalizeCellRef(end));
                const values = [];
                for (let r = Math.min(a.row, b.row); r <= Math.max(a.row, b.row); r++) {
                    for (let c = Math.min(a.colNum, b.colNum); c <= Math.max(a.colNum, b.colNum); c++) {
                        const id = `${this.numberToCol(c)}${r}`;
                        values.push(this.getFormulaContext().getCellValue(id, stack));
                    }
                }
                return values;
            }
        };
    }

    normalizeCellRef(cellId) {
        if (!cellId) return '';
        return String(cellId).toUpperCase().replace(/\$/g, '');
    }

    normalizeMergedRangeEntry(range) {
        if (!range) return null;
        const startCol = range.startCol ?? range.c1 ?? range.col1 ?? range.left ?? range.start?.col ?? range.start?.c;
        const startRow = range.startRow ?? range.r1 ?? range.row1 ?? range.top ?? range.start?.row ?? range.start?.r;
        const endCol = range.endCol ?? range.c2 ?? range.col2 ?? range.right ?? range.end?.col ?? range.end?.c;
        const endRow = range.endRow ?? range.r2 ?? range.row2 ?? range.bottom ?? range.end?.row ?? range.end?.r;
        if (!startCol || !startRow || !endCol || !endRow) return null;
        return {
            startCol: Math.min(startCol, endCol),
            startRow: Math.min(startRow, endRow),
            endCol: Math.max(startCol, endCol),
            endRow: Math.max(startRow, endRow)
        };
    }

    getNormalizedMergedRanges(sheet = this.activeSheet) {
        const ranges = Array.isArray(sheet?.mergedRanges) ? sheet.mergedRanges : [];
        const normalized = [];
        ranges.forEach((range) => {
            const norm = this.normalizeMergedRangeEntry(range);
            if (norm) normalized.push(norm);
        });
        return normalized;
    }

    normalizeMergedCellId(cellId, sheet = this.activeSheet) {
        const safeId = this.normalizeCellRef(cellId);
        if (!safeId) return safeId;
        const { colNum, row } = this.parseCellId(safeId);
        const range = this.getMergedRangeAt(colNum, row, sheet);
        if (!range) return safeId;
        return `${this.numberToCol(range.startCol)}${range.startRow}`;
    }

    rangesIntersect(a, b) {
        return !(a.endCol < b.startCol || a.startCol > b.endCol || a.endRow < b.startRow || a.startRow > b.endRow);
    }

    getMergedRangeAt(colNum, row, sheet = this.activeSheet) {
        const ranges = this.getNormalizedMergedRanges(sheet);
        return ranges.find((range) => (
            colNum >= range.startCol &&
            colNum <= range.endCol &&
            row >= range.startRow &&
            row <= range.endRow
        )) || null;
    }

    getMergedAnchorForCell(colNum, row, sheet = this.activeSheet) {
        const range = this.getMergedRangeAt(colNum, row, sheet);
        if (!range) return null;
        return { colNum: range.startCol, row: range.startRow };
    }

    expandRangeToIncludeMerges(range, sheet = this.activeSheet) {
        let expanded = { ...range };
        let changed = true;
        while (changed) {
            changed = false;
            const ranges = this.getNormalizedMergedRanges(sheet);
            ranges.forEach((merge) => {
                if (this.rangesIntersect(expanded, merge)) {
                    const next = {
                        startCol: Math.min(expanded.startCol, merge.startCol),
                        startRow: Math.min(expanded.startRow, merge.startRow),
                        endCol: Math.max(expanded.endCol, merge.endCol),
                        endRow: Math.max(expanded.endRow, merge.endRow)
                    };
                    if (
                        next.startCol !== expanded.startCol ||
                        next.startRow !== expanded.startRow ||
                        next.endCol !== expanded.endCol ||
                        next.endRow !== expanded.endRow
                    ) {
                        expanded = next;
                        changed = true;
                    }
                }
            });
        }
        return expanded;
    }

    rangeIntersectsMerges(range, sheet = this.activeSheet) {
        const ranges = this.getNormalizedMergedRanges(sheet);
        return ranges.some((merge) => this.rangesIntersect(range, merge));
    }

    unmergeRange(range) {
        if (!range) return false;
        const normalized = this.expandRangeToIncludeMerges(range);
        const existing = this.getNormalizedMergedRanges();
        const next = existing.filter((merge) => !this.rangesIntersect(merge, normalized));
        if (next.length === existing.length) return false;
        this.mergedRanges = next;
        this.applyMergesToGrid();
        return true;
    }

    parseRangeRef(rangeRef) {
        if (!rangeRef || typeof rangeRef !== 'string') return null;
        const parts = rangeRef.split(':');
        const start = parts[0];
        const end = parts[1] || parts[0];
        if (!start) return null;
        const a = this.parseCellId(this.normalizeCellRef(start));
        const b = this.parseCellId(this.normalizeCellRef(end));
        return {
            startCol: Math.min(a.colNum, b.colNum),
            startRow: Math.min(a.row, b.row),
            endCol: Math.max(a.colNum, b.colNum),
            endRow: Math.max(a.row, b.row)
        };
    }

    getSelectableCell(colNum, row) {
        const anchor = this.getMergedAnchorForCell(colNum, row);
        if (anchor) {
            return this.getCellEl(anchor.colNum, anchor.row);
        }
        return this.getCellEl(colNum, row);
    }

    getCellRectForCoord(colNum, row) {
        const range = this.getMergedRangeAt(colNum, row);
        const anchor = range ? this.getCellEl(range.startCol, range.startRow) : this.getCellEl(colNum, row);
        return anchor ? anchor.getBoundingClientRect() : null;
    }

    applyMergesToGrid() {
        if (!this.tbody) return;
        const cells = this.tbody.querySelectorAll('.cell');
        cells.forEach((cell) => {
            if (cell.classList.contains('header')) return;
            cell.style.display = '';
            cell.classList.remove('merge-hidden');
            cell.classList.remove('merge-anchor');
            cell.removeAttribute('rowspan');
            cell.removeAttribute('colspan');
            if (cell.tabIndex < 0) cell.tabIndex = 0;
        });

        const ranges = this.getNormalizedMergedRanges();
        ranges.forEach((range) => {
            const anchorCell = this.getCellEl(range.startCol, range.startRow);
            if (!anchorCell) return;
            anchorCell.setAttribute('rowspan', String(range.endRow - range.startRow + 1));
            anchorCell.setAttribute('colspan', String(range.endCol - range.startCol + 1));
            anchorCell.classList.add('merge-anchor');
            for (let r = range.startRow; r <= range.endRow; r++) {
                for (let c = range.startCol; c <= range.endCol; c++) {
                    if (r === range.startRow && c === range.startCol) continue;
                    const cell = this.getCellEl(c, r);
                    if (!cell) continue;
                    cell.style.display = 'none';
                    cell.classList.add('merge-hidden');
                    cell.tabIndex = -1;
                }
            }
        });
    }

    renderCellById(cellId) {
        const normalizedId = this.normalizeMergedCellId(cellId);
        const cell = document.querySelector(`[data-id="${normalizedId}"]`);
        if (cell) this.renderCellValue(cell);
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
        emptyHeader.addEventListener('mousedown', (e) => {
            if (this.isNearResizeEdge(emptyHeader, e, 'corner')) return;
            e.preventDefault();
            this.selectAll();
            this.headerSelectionAnchor = null;
        });
        headerRow.appendChild(emptyHeader);
        
        for (let j = 0; j < this.cols; j++) {
            const th = document.createElement('th');
            th.className = 'cell header col-header';
            th.innerText = String.fromCharCode(65 + j);
            th.dataset.colIndex = j;
            th.addEventListener('mousedown', (e) => this.handleHeaderMouseDown('col', j + 1, e));
            th.addEventListener('mouseover', () => this.handleHeaderMouseOver('col', j + 1));
            headerRow.appendChild(th);
        }
        table.appendChild(headerRow);
        
        // Data Rows
        this.tbody = document.createElement('tbody');
        table.appendChild(this.tbody);
        this.createRowElements(1, this.rows);
        
        this.container.appendChild(table);
        this.applyMergesToGrid();
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
            rowHeader.addEventListener('mousedown', (e) => this.handleHeaderMouseDown('row', i, e));
            rowHeader.addEventListener('mouseover', () => this.handleHeaderMouseOver('row', i));
            tr.appendChild(rowHeader);
            
            for (let j = 0; j < this.cols; j++) {
                const td = document.createElement('td');
                td.className = 'cell';
                td.contentEditable = false; // Initially Ready Mode
                td.tabIndex = 0; // Make focusable for keyboard events in Ready mode
                const cellId = `${this.numberToCol(j + 1)}${i}`;
                td.dataset.id = cellId;
                
                this.renderCellValue(td);
                
                td.addEventListener('focus', () => this.handleCellFocus(td));
                td.addEventListener('input', () => this.handleCellInput(td));
                td.addEventListener('compositionstart', () => this.handleCompositionStart(td));
                td.addEventListener('compositionend', () => this.handleCompositionEnd(td));
                td.addEventListener('blur', () => this.handleCellBlur(td));
                td.addEventListener('keydown', (e) => this.handleKeyDown(e));
                td.addEventListener('mousedown', (e) => this.handleCellMouseDown(td, e));
                td.addEventListener('dblclick', (e) => this.enterEditMode(td));
                
                this.renderStyles(td);
                tr.appendChild(td);
            }
            this.tbody.appendChild(tr);
        }
    }

    // ─── Event Listeners ───────────────────────────────────
    setupEventListeners() {
        const fillSkip = document.getElementById('fill-skip-blanks');
        if (fillSkip) {
            fillSkip.checked = this.fillSkipBlanks;
            fillSkip.addEventListener('change', () => {
                this.fillSkipBlanks = fillSkip.checked;
            });
        }
        const customBtn = document.getElementById('btn-custom-lists');
        if (customBtn) {
            customBtn.addEventListener('click', () => this.openCustomListModal());
        }
        this.bindCustomListModal();
        document.addEventListener('keydown', (e) => {
            if (!this.isFilling) return;
            if (e.altKey) {
                this.fillPreviewMode = 'copy';
                this.updateFillPreview();
            }
        });
        document.addEventListener('keyup', (e) => {
            if (!this.isFilling) return;
            if (!e.altKey) {
                this.fillPreviewMode = 'series';
                this.updateFillPreview();
            }
        });
        // Styling buttons (Bold, Italic handled by execCommand AND toggleStyle for persistence)
        document.getElementById('btn-bold').addEventListener('click', () => {
            document.execCommand('bold', false, null);
            this.toggleStyle('fontWeight', 'bold', 'normal');
        });
        document.getElementById('btn-italic').addEventListener('click', () => {
            document.execCommand('italic', false, null);
            this.toggleStyle('fontStyle', 'italic', 'normal');
        });
        document.getElementById('btn-underline').addEventListener('click', () => {
            document.execCommand('underline', false, null);
            this.toggleStyle('textDecoration', 'underline', 'none');
        });
        document.getElementById('btn-strike').addEventListener('click', () => {
            document.execCommand('strikethrough', false, null);
            this.toggleStyle('textDecoration', 'line-through', 'none');
        });

        // Color pickers
        document.getElementById('text-color').addEventListener('input', (e) => this.applyStyle('color', e.target.value));
        document.getElementById('bg-color').addEventListener('input', (e) => this.applyStyle('backgroundColor', e.target.value));

        // Alignment
        document.getElementById('btn-align-left').addEventListener('click', () => this.applyStyle('textAlign', 'left'));
        document.getElementById('btn-align-center').addEventListener('click', () => this.applyStyle('textAlign', 'center'));
        document.getElementById('btn-align-right').addEventListener('click', () => this.applyStyle('textAlign', 'right'));

        // Font family & size
        document.getElementById('font-family').addEventListener('change', (e) => this.applyStyle('fontFamily', e.target.value));
        document.getElementById('font-size').addEventListener('input', (e) => this.applyStyle('fontSize', e.target.value + 'pt'));

        // Custom Border Style Dropdown Logic
        const bsTrigger = document.getElementById('border-style-trigger');
        const bsOptions = document.getElementById('border-style-options');
        if (bsTrigger && bsOptions) {
            bsTrigger.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                bsOptions.classList.toggle('show');
            });

            bsOptions.querySelectorAll('.style-option').forEach(opt => {
                opt.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const val = opt.getAttribute('data-value');
                    this.currentBorderStyle = val;
                    
                    // Update trigger UI preview
                    let previewHtml = '';
                    if (val === 'none') {
                        previewHtml = '<div style="font-size: 0.75rem; color:#94a3b8; width: 100%; text-align: center;">None ✕</div>';
                    } else {
                        previewHtml = opt.querySelector('svg').outerHTML;
                    }
                    
                    const preview = bsTrigger.querySelector('.style-preview');
                    if (preview) preview.innerHTML = previewHtml;
                    
                    // Mark selection visually
                    bsOptions.querySelectorAll('.style-option').forEach(o => o.classList.remove('selected'));
                    opt.classList.add('selected');
                    
                    bsOptions.classList.remove('show');
                });
            });

            // Close on any click outside
            document.addEventListener('click', (e) => {
                if (!bsTrigger.contains(e.target)) {
                    bsOptions.classList.remove('show');
                }
            });
        }

        // Border Type Icon Buttons
        document.querySelectorAll('.border-type-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                const type = btn.getAttribute('data-type');
                this.currentBorderType = type;
                const color = document.getElementById('border-color')?.value || '#000000';
                this.applyBorder(type, this.currentBorderStyle, color);
            });
        });

        // Phase 7: Data formatting
        document.getElementById('format-type').addEventListener('change', (e) => {
            const selectedType = e.target.value;
            this.applyFormat({ type: selectedType });
        });
        document.getElementById('format-decimals').addEventListener('change', (e) => {
            this.applyFormat({ decimals: e.target.value });
        });
        document.getElementById('btn-decimal-decrease').addEventListener('click', () => this.adjustDecimals(-1));
        document.getElementById('btn-decimal-increase').addEventListener('click', () => this.adjustDecimals(1));

        // Find / Replace
        this.findInput = document.getElementById('find-input');
        this.replaceInput = document.getElementById('replace-input');
        this.findCase = document.getElementById('find-case');
        this.findExact = document.getElementById('find-exact');

        const refreshFind = () => this.updateFindResults();
        if (this.findInput) this.findInput.addEventListener('input', refreshFind);
        if (this.replaceInput) {
            this.replaceInput.addEventListener('input', () => {
                this.findState.replace = this.replaceInput.value || '';
            });
        }
        if (this.findCase) this.findCase.addEventListener('change', refreshFind);
        if (this.findExact) this.findExact.addEventListener('change', refreshFind);
        document.getElementById('btn-find-prev').addEventListener('click', () => this.gotoFindMatch(-1));
        document.getElementById('btn-find-next').addEventListener('click', () => this.gotoFindMatch(1));
        document.getElementById('btn-replace').addEventListener('click', () => this.replaceCurrentMatch());
        document.getElementById('btn-replace-all').addEventListener('click', () => this.replaceAllMatches());

        // File Dialog buttons
        document.getElementById('btn-open').addEventListener('click', () => this.openFileDialog());
        document.getElementById('btn-save').addEventListener('click', () => this.saveFile());
        document.getElementById('btn-save-as').addEventListener('click', () => this.saveFileAs());

        // Table Operations
        document.getElementById('btn-insert-row').addEventListener('click', () => this.insertRow());
        document.getElementById('btn-delete-row').addEventListener('click', () => this.deleteRow());
        document.getElementById('btn-insert-col').addEventListener('click', () => this.insertColumn());
        document.getElementById('btn-delete-col').addEventListener('click', () => this.deleteColumn());

        const mergeToggleBtn = document.getElementById('btn-merge-toggle');
        if (mergeToggleBtn) mergeToggleBtn.addEventListener('click', () => this.toggleMergeSelection());

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
                this.applyMergesToGrid();
            }
        });

        // Formula bar input
        this.formulaInput.addEventListener('input', (e) => {
            if (this.selectedCell) {
                const id = this.selectedCell.dataset.id;
                const rawValue = e.target.value;
                this.setRawValue(id, rawValue);
                if (this.isEditing) {
                    this.selectedCell.innerText = rawValue;
                } else {
                    this.renderCellValue(this.selectedCell);
                }
                this.markDirty();
                this.updateItemCount();
                this.refreshFindIfActive();
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
            if (this.isHeaderSelecting) {
                this.endHeaderSelection();
            }
            if (this.isSelecting) {
                this.endRangeSelection();
            }
            if (this.isFilling) {
                this.handleFillEnd(e);
            }
        });

        // Maintain overlay alignment on window resize or container scroll
        window.addEventListener('resize', () => {
            this.updateSelectionOverlay();
            this.updateRangeVisual();
            this.updateFillHandlePosition();
        });
        this.container.addEventListener('scroll', () => {
             this.updateSelectionOverlay();
             this.updateRangeVisual();
             this.updateFillHandlePosition();
        }, { passive: true });
        // Global keyboard shortcuts (clipboard, delete)
        document.addEventListener('keydown', (e) => {
            // Ignore if typing in formula bar or in a cell edit
            if (e.target.id === 'formula-input' || this.isEditing) return;

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
                        this.saveFile();
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

        // Resize Handlers
        this.setupResizeHandlers();
    }

    renderSheetTabs() {
        if (!this.sheetTabs) return;
        this.sheetTabs.innerHTML = '';

        this.sheets.forEach((sheet, index) => {
            const btn = document.createElement('button');
            btn.className = 'sheet-tab' + (index === this.activeSheetIndex ? ' active' : '');
            btn.dataset.index = index;
            const label = document.createElement('span');
            label.className = 'sheet-tab-label';
            label.textContent = sheet.name || `Sheet${index + 1}`;
            btn.appendChild(label);

            const closeBtn = document.createElement('span');
            closeBtn.className = 'sheet-tab-close';
            closeBtn.textContent = '×';
            closeBtn.title = 'Delete sheet';
            closeBtn.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                this.deleteSheet(index);
            });
            btn.appendChild(closeBtn);
            btn.addEventListener('click', () => {
                if (this.sheetClickTimer) clearTimeout(this.sheetClickTimer);
                this.sheetClickTimer = setTimeout(() => {
                    this.switchSheet(index);
                }, 200);
            });
            btn.addEventListener('dblclick', (e) => {
                e.preventDefault();
                if (this.sheetClickTimer) clearTimeout(this.sheetClickTimer);
                this.startSheetRename(label, index);
            });
            btn.addEventListener('contextmenu', (e) => {
                e.preventDefault();
                this.deleteSheet(index);
            });

            btn.draggable = true;
            btn.addEventListener('dragstart', (e) => {
                e.dataTransfer.setData('text/plain', String(index));
            });
            btn.addEventListener('dragover', (e) => e.preventDefault());
            btn.addEventListener('drop', (e) => {
                e.preventDefault();
                const from = Number(e.dataTransfer.getData('text/plain'));
                const to = index;
                this.moveSheet(from, to);
            });

            this.sheetTabs.appendChild(btn);
        });

        const addBtn = document.createElement('button');
        addBtn.className = 'sheet-tab add-btn';
        addBtn.textContent = '+';
        addBtn.addEventListener('click', () => this.addSheet());
        this.sheetTabs.appendChild(addBtn);
    }

    addSheet() {
        const name = `Sheet${this.sheets.length + 1}`;
        this.sheets.push(this.createSheet(name));
        this.switchSheet(this.sheets.length - 1);
    }

    startSheetRename(tabEl, index) {
        const sheet = this.sheets[index];
        if (!sheet || !tabEl) return;
        const input = document.createElement('input');
        input.type = 'text';
        input.value = sheet.name || `Sheet${index + 1}`;
        input.className = 'sheet-rename-input';

        tabEl.textContent = '';
        tabEl.appendChild(input);
        input.focus();
        input.select();

        const finish = (commit) => {
            const raw = input.value || '';
            const nextName = commit ? raw.trim() : sheet.name;
            sheet.name = nextName || `Sheet${index + 1}`;
            this.renderSheetTabs();
        };

        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                finish(true);
            } else if (e.key === 'Escape') {
                e.preventDefault();
                finish(false);
            }
        });
        input.addEventListener('blur', () => finish(true));
    }

    async deleteSheet(index) {
        if (this.sheets.length <= 1) {
            alert('At least one sheet is required.');
            return;
        }
        const sheet = this.sheets[index];
        if (!sheet) return;
        const proceed = await this.showConfirmAsync(`"${sheet.name || `Sheet${index + 1}`}" 시트를 삭제할까요?`);
        if (!proceed) return;

        this.sheets.splice(index, 1);
        if (this.activeSheetIndex >= this.sheets.length) {
            this.activeSheetIndex = this.sheets.length - 1;
        }
        this.switchSheet(this.activeSheetIndex);
    }

    moveSheet(from, to) {
        if (from === to || from < 0 || to < 0 || from >= this.sheets.length || to >= this.sheets.length) return;
        const [sheet] = this.sheets.splice(from, 1);
        this.sheets.splice(to, 0, sheet);

        if (this.activeSheetIndex === from) {
            this.activeSheetIndex = to;
        } else if (from < this.activeSheetIndex && to >= this.activeSheetIndex) {
            this.activeSheetIndex -= 1;
        } else if (from > this.activeSheetIndex && to <= this.activeSheetIndex) {
            this.activeSheetIndex += 1;
        }

        this.renderSheetTabs();
    }

    switchSheet(index) {
        if (index < 0 || index >= this.sheets.length) return;
        if (this.isEditing) this.exitEditMode(true);

        this.activeSheetIndex = index;
        this.selectedCell = null;
        this.selectionRange = null;
        this.selectionAnchor = null;

        this.refreshGridUI();
        this.renderSheetTabs();
        this.updateItemCount();

        const cell = this.getCellEl(1, 1);
        if (cell) {
            cell.focus({ preventScroll: true });
            this.handleCellFocus(cell);
        }
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
        const cellId = cell.dataset.id;
        const normalizedId = this.normalizeMergedCellId(cellId);
        if (normalizedId !== cellId) {
            const { colNum, row } = this.parseCellId(normalizedId);
            const anchorCell = this.getCellEl(colNum, row);
            if (anchorCell && anchorCell !== cell) {
                anchorCell.focus({ preventScroll: true });
                return;
            }
        }
        this.selectedCell = cell;
        this.cellAddress.innerText = cellId;
        this.formulaInput.value = this.getRawValue(cellId);
        const { colNum, row } = this.parseCellId(cellId);
        this.highlightHeaders(this.selectionRange || { startCol: colNum, endCol: colNum, startRow: row, endRow: row });
        this.updateSelectionOverlay();
        this.scrollToVisible(cell);
        this.updateToolbarState(cell);

        // Always make focused cell editable to support seamless IME start
        if (!this.isEditing) {
            cell.contentEditable = true;
            
            // CRITICAL for IME: Select all content so first keystroke replaces it naturally
            // This prevents the "rㅏ" bug where the first IME character is lost during manual clearing.
            const range = document.createRange();
            range.selectNodeContents(cell);
            const sel = window.getSelection();
            sel.removeAllRanges();
            sel.addRange(range);
        }
    }

    scrollToVisible(cell) {
        if (!cell) return;
        
        const containerRect = this.container.getBoundingClientRect();
        const cellRect = cell.getBoundingClientRect();
        
        // Sticky boundary offsets (Matches CSS heights/widths)
        const headerHeight = 25; // Standard row height for col headers
        const rowHeaderWidth = 40; // Fixed width for row headers
        
        const buffer = 5; // Extra padding for comfort
        
        // Vertical check (Hidden by sticky header or below container)
        if (cellRect.top < containerRect.top + headerHeight) {
            // Scroll UP to show the cell beneath the header
            this.container.scrollTop -= (containerRect.top + headerHeight - cellRect.top + buffer);
        } else if (cellRect.bottom > containerRect.bottom) {
            // Scroll DOWN
            this.container.scrollTop += (cellRect.bottom - containerRect.bottom + buffer);
        }
        
        // Horizontal check (Hidden by sticky row header or right of container)
        if (cellRect.left < containerRect.left + rowHeaderWidth) {
            // Scroll LEFT
            this.container.scrollLeft -= (containerRect.left + rowHeaderWidth - cellRect.left + buffer);
        } else if (cellRect.right > containerRect.right) {
            // Scroll RIGHT
            this.container.scrollLeft += (cellRect.right - containerRect.right + buffer);
        }
    }

    // ─── Styling ───────────────────────────────────────────
    toggleStyle(prop, activeValue, defaultValue) {
        const targetIds = this.getSelectionTargetIds();
        if (targetIds.length === 0) return;

        // Determine if we should set or unset based on the first cell
        const firstCellId = targetIds[0];
        const currentStyle = this.cellStyles[firstCellId] || {};
        const newValue = currentStyle[prop] === activeValue ? defaultValue : activeValue;

        this.applyStyle(prop, newValue);
    }

    applyStyle(prop, value) {
        const targetIds = this.getSelectionTargetIds();
        if (targetIds.length === 0) return;

        targetIds.forEach(id => {
            if (!this.cellStyles[id]) this.cellStyles[id] = {};
            this.cellStyles[id][prop] = value;

            const el = document.querySelector(`[data-id="${id}"]`);
            if (el) {
                el.style[prop] = value;
            }
        });

        this.markDirty();
        if (this.selectedCell) this.updateToolbarState(this.selectedCell);
    }

    applyFormat(partialFormat) {
        const targetIds = this.getSelectionTargetIds();
        if (targetIds.length === 0) return;

        targetIds.forEach((id) => {
            const current = this.getCellFormat(id);
            const merged = { ...current, ...partialFormat };
            this.setCellFormat(id, merged);

            if (!(this.isEditing && this.selectedCell && this.selectedCell.dataset.id === id)) {
                this.renderCellById(id);
            }
        });

        if (this.selectedCell) {
            this.formulaInput.value = this.getRawValue(this.selectedCell.dataset.id);
            this.updateToolbarState(this.selectedCell);
        }
        this.markDirty();
    }

    adjustDecimals(delta) {
        if (!this.selectedCell) return;
        const current = this.getCellFormat(this.selectedCell.dataset.id);
        if (current.type === 'date') return;

        const base = current.decimals ?? this.getDefaultDecimalsByType(current.type) ?? 0;
        const next = Math.max(0, Math.min(10, base + delta));
        this.applyFormat({ decimals: next });
    }

    getSelectionTargetIds() {
        if (this.selectionRange) {
            const { startCol, startRow, endCol, endRow } = this.selectionRange;
            const ids = [];
            for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) {
                for (let c = Math.min(startCol, endCol); c <= Math.max(startCol, endCol); c++) {
                    ids.push(`${this.numberToCol(c)}${r}`);
                }
            }
            return ids;
        } else if (this.selectedCell) {
            return [this.selectedCell.dataset.id];
        }
        return [];
    }

    updateToolbarState(cell) {
        const id = cell.dataset.id;
        const style = this.cellStyles[id] || {};
        const format = this.getCellFormat(id);

        // Helper to toggle active class
        const setBtnActive = (btnId, isActive) => {
            const btn = document.getElementById(btnId);
            if (btn) btn.classList.toggle('active', isActive);
        };

        setBtnActive('btn-bold', style.fontWeight === 'bold');
        setBtnActive('btn-italic', style.fontStyle === 'italic');
        setBtnActive('btn-underline', style.textDecoration === 'underline');
        setBtnActive('btn-strike', style.textDecoration === 'line-through');

        // Color pickers
        document.getElementById('text-color').value = style.color || '#ffffff';
        document.getElementById('bg-color').value = style.backgroundColor || '#1e1e1e';

        // Alignment
        setBtnActive('btn-align-left', style.textAlign === 'left');
        setBtnActive('btn-align-center', style.textAlign === 'center');
        setBtnActive('btn-align-right', style.textAlign === 'right');

        // Font dropdowns
        if (style.fontFamily) document.getElementById('font-family').value = style.fontFamily;
        if (style.fontSize) document.getElementById('font-size').value = parseInt(style.fontSize);

        // Data format controls
        const formatType = document.getElementById('format-type');
        const decimalsInput = document.getElementById('format-decimals');
        if (formatType) formatType.value = format.type;
        if (decimalsInput) {
            decimalsInput.value = format.decimals ?? '';
            decimalsInput.disabled = format.type === 'date';
        }
        // Merge toggle state
        setBtnActive('btn-merge-toggle', cell.classList.contains('merge-anchor'));
    }

    renderStyles(cell) {
        const id = cell.dataset.id;
        const style = this.cellStyles[id];
        if (style) {
            Object.assign(cell.style, style);
        }
        this.renderBorders(cell);
    }

    renderBorders(cell) {
        const id = cell.dataset.id;
        const parsed = this.parseCellId(id);
        if (!parsed) return;
        
        const colNum = parsed.colNum;
        const rowNum = parsed.row;

        const getBorder = (c, r, s) => {
            const tempId = `${this.numberToCol(c)}${r}`;
            return this.cellBorders?.[tempId]?.[s];
        };

        // 자신의 테두리와 인접 셀의 테두리를 양방향으로 평가하여, 하나라도 있으면 그 테두리를 양쪽 셀 렌더링에 모두 적용함
        // (Webkit 계열의 border-collapse 충돌 시 특정 방향(bottom, right)이 우선시되어 지워지는 현상 방지)
        const topBorder = getBorder(colNum, rowNum, 'top') || (rowNum > 1 ? getBorder(colNum, rowNum - 1, 'bottom') : null);
        const rightBorder = getBorder(colNum, rowNum, 'right') || (colNum < this.cols ? getBorder(colNum + 1, rowNum, 'left') : null);
        const bottomBorder = getBorder(colNum, rowNum, 'bottom') || (rowNum < this.rows ? getBorder(colNum, rowNum + 1, 'top') : null);
        const leftBorder = getBorder(colNum, rowNum, 'left') || (colNum > 1 ? getBorder(colNum - 1, rowNum, 'right') : null);

        const toCss = (b) => b ? `${b.width || 1}px ${b.style || 'solid'} ${b.color || '#000000'}` : null;
        
        const applyOrRemove = (side, val) => {
            if (val) {
                cell.style.setProperty(`border-${side}`, val, 'important');
            } else {
                cell.style.removeProperty(`border-${side}`);
            }
        };

        applyOrRemove('top', toCss(topBorder));
        applyOrRemove('right', toCss(rightBorder));
        applyOrRemove('bottom', toCss(bottomBorder));
        applyOrRemove('left', toCss(leftBorder));
    }

    updateSelectionOverlay() {
        if (!this.selectedCell || !this.selectionOverlay) return;

        const cell = this.selectedCell;
        const rect = cell.getBoundingClientRect();
        const containerRect = this.container.getBoundingClientRect();

        this.selectionOverlay.style.display = 'block';
        this.selectionOverlay.style.top = `${rect.top - containerRect.top + this.container.scrollTop}px`;
        this.selectionOverlay.style.left = `${rect.left - containerRect.left + this.container.scrollLeft}px`;
        this.selectionOverlay.style.width = `${rect.width}px`;
        this.selectionOverlay.style.height = `${rect.height}px`;
    }

    clearHighlights() {
        document.querySelectorAll('.cell.header.active').forEach(h => h.classList.remove('active'));
    }

    highlightHeaders(range) {
        this.clearHighlights();
        if (!range) return;

        for (let c = range.startCol; c <= range.endCol; c++) {
            const colHeader = this.table?.querySelector(`.col-header[data-col-index="${c - 1}"]`);
            if (colHeader) colHeader.classList.add('active');
        }

        for (let r = range.startRow; r <= range.endRow; r++) {
            const rowHeader = this.table?.querySelector(`.row-header[data-row-index="${r}"]`);
            if (rowHeader) rowHeader.classList.add('active');
        }

        if (range.startCol === 1 && range.endCol === this.cols && range.startRow === 1 && range.endRow === this.rows) {
            const corner = this.table?.querySelector('.corner-header');
            if (corner) corner.classList.add('active');
        }
    }

    handleCellInput(cell) {
        // If we just entered Enter Mode, we want the first input to overwrite everything
        if (this.needsOverwrite) {
            this.needsOverwrite = false;
            // The browser already inserted the first character/composition. 
            // We just need to make sure it's the ONLY thing in the cell.
            // However, with IME, innerText might contain the composing character.
        }
        
        const cellId = cell.dataset.id;
        this.setRawValue(cellId, cell.innerText);
        this.formulaInput.value = this.getRawValue(cellId);
        this.markDirty();
        this.updateItemCount();
        this.refreshFindIfActive();
        this.formulaCache.clear();
    }

    handleCompositionStart(cell) {
        this.isComposing = true;
        if (!this.isEditing) {
            this.isEditing = true;
            this.originalValue = this.getRawValue(cell.dataset.id);
            cell.classList.add('editing');
        }
        this.needsOverwrite = false;
        // Ensure no selection remains that could be replaced on space.
        const range = document.createRange();
        const sel = window.getSelection();
        range.selectNodeContents(cell);
        range.collapse(false);
        sel.removeAllRanges();
        sel.addRange(range);
    }

    handleCompositionEnd(cell) {
        this.isComposing = false;
        // Input event will sync data; keep this for state clarity.
    }

    handleCellBlur(cell) {
        if (this.isEditing) {
            this.exitEditMode(true);
        }
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

    updateFindResults() {
        return window.VSFind.updateFindResults(this);
    }

    refreshFindIfActive() {
        return window.VSFind.refreshFindIfActive(this);
    }

    clearFindHighlights() {
        return window.VSFind.clearFindHighlights(this);
    }

    applyFindHighlights() {
        return window.VSFind.applyFindHighlights(this);
    }

    selectFindMatch(index) {
        return window.VSFind.selectFindMatch(this, index);
    }

    gotoFindMatch(direction) {
        return window.VSFind.gotoFindMatch(this, direction);
    }

    replaceCurrentMatch() {
        return window.VSFind.replaceCurrentMatch(this);
    }

    replaceAllMatches() {
        return window.VSFind.replaceAllMatches(this);
    }

    replaceInValue(rawValue) {
        return window.VSFind.replaceInValue(this, rawValue);
    }

    // Find/Replace methods are loaded from engines/vs_find.js

    // ─── Range Selection ───────────────────────────────────
    isNearResizeEdge(headerEl, e, type) {
        const EDGE_ZONE = 5;
        const rect = headerEl.getBoundingClientRect();
        if (type === 'col') {
            return e.clientX >= rect.right - EDGE_ZONE || e.clientX <= rect.left + EDGE_ZONE;
        }
        if (type === 'row') {
            return e.clientY >= rect.bottom - EDGE_ZONE;
        }
        return e.clientX >= rect.right - EDGE_ZONE || e.clientY >= rect.bottom - EDGE_ZONE;
    }

    handleHeaderMouseDown(type, index, e) {
        if (this.isResizingCol || this.isResizingRow) return;
        if (this.isNearResizeEdge(e.currentTarget, e, type)) return;

        e.preventDefault();

        if (this.isEditing) {
            this.exitEditMode(true);
        }

        this.isHeaderSelecting = true;
        this.headerSelectType = type;

        if (!e.shiftKey || !this.headerSelectionAnchor || this.headerSelectionAnchor.type !== type) {
            this.headerSelectionAnchor = { type, index };
        }

        const anchorIndex = (e.shiftKey && this.headerSelectionAnchor.type === type)
            ? this.headerSelectionAnchor.index
            : index;

        this.selectHeaderRange(type, anchorIndex, index);
    }

    handleHeaderMouseOver(type, index) {
        if (!this.isHeaderSelecting || this.headerSelectType !== type || !this.headerSelectionAnchor) return;
        this.selectHeaderRange(type, this.headerSelectionAnchor.index, index);
    }

    endHeaderSelection() {
        this.isHeaderSelecting = false;
        this.headerSelectType = null;
        this.updateFillHandlePosition();
    }

    selectHeaderRange(type, startIndex, endIndex) {
        if (type === 'col') {
            this.setSelectionRange(startIndex, 1, endIndex, this.rows);
        } else {
            this.setSelectionRange(1, startIndex, this.cols, endIndex);
        }

        const focusCol = type === 'col' ? Math.min(startIndex, endIndex) : 1;
        const focusRow = type === 'row' ? Math.min(startIndex, endIndex) : 1;
        const focusCell = this.getSelectableCell(focusCol, focusRow);
        if (focusCell) {
            this.selectedCell = focusCell;
            this.cellAddress.innerText = focusCell.dataset.id;
            this.formulaInput.value = this.getRawValue(focusCell.dataset.id);
        }

        this.updateRangeVisual();
        this.updateFillHandlePosition();
    }

    handleCellMouseDown(cell, e) {
        // Don't start selection if clicking fill handle
        if (e.target.classList.contains('fill-handle')) return;

        let cellId = cell.dataset.id;
        let { row, colNum } = this.parseCellId(cellId);
        const anchor = this.getMergedAnchorForCell(colNum, row);
        if (anchor) {
            colNum = anchor.colNum;
            row = anchor.row;
            const anchorCell = this.getCellEl(colNum, row);
            if (anchorCell) {
                cell = anchorCell;
                cellId = anchorCell.dataset.id;
            }
        }

        // Enter Edit mode if already selected (second click)
        if (this.selectedCell === cell && !this.isEditing && !e.shiftKey) {
            this.enterEditMode(cell);
            return;
        }

        if (e.shiftKey && this.selectedCell) {
            // Shift+Click: extend selection from current cell
            e.preventDefault();
            const anchor = this.parseCellId(this.selectedCell.dataset.id);
            this.setSelectionRange(anchor.colNum, anchor.row, colNum, row);
            this.updateRangeVisual();
            this.updateFillHandlePosition();
            return;
        }

        // Standard selection behavior
        if (this.isEditing && this.selectedCell !== cell) {
            this.exitEditMode(true);
        }

        this.clearRangeSelection();
        this.selectionAnchor = { colNum, row };
        this.isSelecting = true;
        this.setSelectionRange(colNum, row, colNum, row);
        this.updateRangeVisual();
        
        // Focus cell but don't edit yet
        cell.focus();
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
        const baseRange = {
            startCol: Math.min(c1, c2),
            startRow: Math.min(r1, r2),
            endCol: Math.max(c1, c2),
            endRow: Math.max(r1, r2)
        };
        this.selectionRange = this.expandRangeToIncludeMerges(baseRange);
    }

    clearRangeSelection() {
        this.selectionRange = null;
        document.querySelectorAll('.cell.in-range').forEach(c => c.classList.remove('in-range'));
        if (this.rangeOverlay) this.rangeOverlay.style.display = 'none';
        if (this.selectionOverlay) this.selectionOverlay.style.display = 'none';
        this.clearHighlights();
    }

    updateRangeVisual() {
        // Always update selection overlay for the active cell
        this.updateSelectionOverlay();

        // Remove old highlights
        document.querySelectorAll('.cell.in-range').forEach(c => c.classList.remove('in-range'));

        if (!this.selectionRange) {
            if (this.rangeOverlay) this.rangeOverlay.style.display = 'none';
            if (this.selectedCell) {
                const { colNum, row } = this.parseCellId(this.selectedCell.dataset.id);
                this.highlightHeaders({ startCol: colNum, endCol: colNum, startRow: row, endRow: row });
            } else {
                this.clearHighlights();
            }
            return;
        }

        const { startCol, startRow, endCol, endRow } = this.selectionRange;
        this.highlightHeaders(this.selectionRange);
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
        const tlRect = this.getCellRectForCoord(startCol, startRow);
        const brRect = this.getCellRectForCoord(endCol, endRow);

        if (tlRect && brRect && this.rangeOverlay) {
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

    toggleMergeSelection() {
        const range = this.getEffectiveRange();
        if (!range) return;
        
        const expanded = this.expandRangeToIncludeMerges(range);
        const isAlreadyMerged = this.isEntirelyMerged(expanded);
        
        if (isAlreadyMerged) {
            this.unmergeSelection();
        } else {
            this.mergeSelection();
        }
    }

    isEntirelyMerged(range) {
        const existing = this.getNormalizedMergedRanges();
        return existing.some(m => 
            m.startCol === range.startCol && 
            m.startRow === range.startRow && 
            m.endCol === range.endCol && 
            m.endRow === range.endRow
        );
    }

    mergeSelection() {
        const range = this.getEffectiveRange();
        if (!range) return;
        const normalized = this.expandRangeToIncludeMerges(range);
        const isSingleCell = normalized.startCol === normalized.endCol && normalized.startRow === normalized.endRow;
        if (isSingleCell) return;

        const existing = this.getNormalizedMergedRanges();
        const keep = [];
        existing.forEach((merge) => {
            if (!this.rangesIntersect(merge, normalized)) {
                keep.push(merge);
            }
        });
        this.mergedRanges = keep;

        for (let r = normalized.startRow; r <= normalized.endRow; r++) {
            for (let c = normalized.startCol; c <= normalized.endCol; c++) {
                if (r === normalized.startRow && c === normalized.startCol) continue;
                const cellId = `${this.numberToCol(c)}${r}`;
                delete this.data[cellId];
                delete this.cellFormulas[cellId];
            }
        }

        this.mergedRanges = [
            ...this.getNormalizedMergedRanges(),
            normalized
        ];

        this.applyMergesToGrid();
        this.setSelectionRange(normalized.startCol, normalized.startRow, normalized.endCol, normalized.endRow);
        const anchorCell = this.getCellEl(normalized.startCol, normalized.startRow);
        if (anchorCell) {
            anchorCell.focus({ preventScroll: true });
            this.handleCellFocus(anchorCell);
        }
        this.updateRangeVisual();
        this.updateFillHandlePosition();
        this.markDirty();
    }

    applyBorder(type, styleStr, color) {
        const range = this.getEffectiveRange();
        if (!range) return;

        // Parse style and width (e.g., 'solid-2' -> style: solid, width: 2)
        let style = 'solid';
        let width = 1;
        if (styleStr && styleStr.includes('-')) {
            const parts = styleStr.split('-');
            style = parts[0];
            width = parseInt(parts[1]) || 1;
        } else if (styleStr) {
            style = styleStr;
        }

        const expanded = this.expandRangeToIncludeMerges(range);
        const border = {
            style: style,
            color: color || '#000000',
            width: width
        };

        const applySide = (cellId, side, value) => {
            if (!this.cellBorders[cellId]) this.cellBorders[cellId] = {};
            if (value) {
                this.cellBorders[cellId][side] = value;
            } else {
                delete this.cellBorders[cellId][side];
                if (Object.keys(this.cellBorders[cellId]).length === 0) {
                    delete this.cellBorders[cellId];
                }
            }
        };

        const sidesForType = () => {
            if (type === 'none') return ['top', 'right', 'bottom', 'left'];
            if (type === 'all') return ['top', 'right', 'bottom', 'left'];
            if (type === 'outer') return ['top', 'right', 'bottom', 'left'];
            if (type === 'inner') return ['top', 'left'];
            if (type === 'inner-v') return ['left'];
            if (type === 'inner-h') return ['top'];
            return [type];
        };

        const shouldApply = (side, r, c, merge) => {
            if (type === 'all') return true;
            if (type === 'outer') {
                if (side === 'top') return r === expanded.startRow;
                if (side === 'bottom') return r === expanded.endRow;
                if (side === 'left') return c === expanded.startCol;
                if (side === 'right') return c === expanded.endCol;
            }
            if (type === 'inner') {
                if (side === 'top') return r > expanded.startRow;
                if (side === 'left') return c > expanded.startCol;
                return false;
            }
            if (type === 'inner-v') {
                if (side === 'left') return c > expanded.startCol;
                return false;
            }
            if (type === 'inner-h') {
                if (side === 'top') return r > expanded.startRow;
                return false;
            }
            // 특정 모서리 단일 테두리에 대해서는 전체 선택 영역의 외곽 가장자리만 적용
            if (type === 'top') return r === expanded.startRow;
            if (type === 'bottom') return r === expanded.endRow;
            if (type === 'left') return c === expanded.startCol;
            if (type === 'right') return c === expanded.endCol;

            return true;
        };

        for (let r = expanded.startRow; r <= expanded.endRow; r++) {
            for (let c = expanded.startCol; c <= expanded.endCol; c++) {
                const merge = this.getMergedRangeAt(c, r);
                const targetId = merge
                    ? `${this.numberToCol(merge.startCol)}${merge.startRow}`
                    : `${this.numberToCol(c)}${r}`;
                const sides = sidesForType();
                sides.forEach((side) => {
                    const isBoundary = merge
                        ? (
                            (side === 'top' && r === merge.startRow) ||
                            (side === 'bottom' && r === merge.endRow) ||
                            (side === 'left' && c === merge.startCol) ||
                            (side === 'right' && c === merge.endCol)
                        )
                        : true;
                    if (isBoundary && shouldApply(side, r, c, merge)) {
                        applySide(targetId, side, (type === 'none' || style === 'none') ? null : border);
                    }
                });
            }
        }

        this.refreshGridUI();
        this.markDirty();
    }

    unmergeSelection() {
        const range = this.getEffectiveRange();
        if (!range) return;
        const normalized = this.expandRangeToIncludeMerges(range);

        const existing = this.getNormalizedMergedRanges();
        const next = existing.filter((merge) => !this.rangesIntersect(merge, normalized));
        if (next.length === existing.length) return;

        this.mergedRanges = next;
        this.applyMergesToGrid();
        this.updateRangeVisual();
        this.updateFillHandlePosition();
        this.markDirty();
    }

    shiftMergedRanges(type, threshold, delta) {
        const ranges = this.getNormalizedMergedRanges();
        if (ranges.length === 0) return;

        const isRow = type === 'row';
        const removedStart = delta < 0 ? threshold + delta : null;
        const removedEnd = delta < 0 ? threshold - 1 : null;
        const shiftCoord = (coord) => (coord >= threshold ? coord + delta : coord);

        const next = [];
        ranges.forEach((range) => {
            if (delta < 0) {
                const start = isRow ? range.startRow : range.startCol;
                const end = isRow ? range.endRow : range.endCol;
                if (start <= removedEnd && end >= removedStart) {
                    return;
                }
            }

            const updated = {
                startCol: isRow ? range.startCol : shiftCoord(range.startCol),
                endCol: isRow ? range.endCol : shiftCoord(range.endCol),
                startRow: isRow ? shiftCoord(range.startRow) : range.startRow,
                endRow: isRow ? shiftCoord(range.endRow) : range.endRow
            };

            if (updated.startCol < 1 || updated.startRow < 1) return;
            next.push(updated);
        });

        this.mergedRanges = next;
    }

    selectAll() {
        this.setSelectionRange(1, 1, this.cols, this.rows);
        this.updateRangeVisual();
        this.updateFillHandlePosition();
    }

    // Simplified Overlay methods (Already handled in init)
    createRangeOverlay() {}
    createFillHandle() {}
    createSelectionOverlay() {}

    updateFillHandlePosition() {
        const range = this.getEffectiveRange();
        if (!range) {
            this.fillHandle.style.display = 'none';
            return;
        }

        const rect = this.getCellRectForCoord(range.endCol, range.endRow);
        if (!rect) {
            this.fillHandle.style.display = 'none';
            return;
        }
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
        this.fillPreviewMode = 'series';
        this.fillHandle.style.pointerEvents = 'none';
        this.selectionOverlay.style.display = 'block';
        if (this.fillPreview) this.fillPreview.style.display = 'none';
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
            this.updateFillPreview();
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

    async handleFillEnd(e) {
        if (!this.isFilling) return;
        this.isFilling = false;
        this.selectionOverlay.style.display = 'none';
        if (this.fillPreview) this.fillPreview.style.display = 'none';
        this.fillHandle.style.pointerEvents = 'auto';

        if (this.lastFillTargetCell) {
            await this.fillFromRange(this.fillRange, this.lastFillTargetCell);
        }
        this.updateFillHandlePosition();
    }

    async fillFromRange(sourceRange, targetCell) {
        const target = this.parseCellId(targetCell.dataset.id);
        const { startCol, startRow, endCol, endRow } = sourceRange;
        const rangeCols = endCol - startCol + 1;
        const rangeRows = endRow - startRow + 1;

        // Collect source values as 2D array
        const sourceValues = [];
        for (let r = startRow; r <= endRow; r++) {
            const rowVals = [];
            for (let c = startCol; c <= endCol; c++) {
                const cellId = `${this.numberToCol(c)}${r}`;
                rowVals.push(this.getRawValue(cellId));
            }
            sourceValues.push(rowVals);
        }
        const blankMask = sourceValues.map(row => row.map(val => this.isBlankValue(val)));

        const useSeriesPref = this.fillPreviewMode === 'series';
        const forceCopy = this.fillPreviewMode === 'copy';

        // Determine fill direction
        if (target.colNum >= startCol && target.colNum <= endCol) {
            // Vertical fill
            const fillStart = target.row > endRow ? endRow + 1 : (target.row < startRow ? target.row : startRow);
            const fillEnd = target.row > endRow ? target.row : (target.row < startRow ? startRow - 1 : endRow);
            if (fillStart > fillEnd) return;
            const fillCount = fillEnd - fillStart + 1;
            const directionSign = target.row > endRow ? 1 : -1;
            const forward = target.row > endRow;
            const deltaRow = forward ? 1 : -1;

            const seriesByCol = [];
            let hasSeries = false;
            for (let c = 0; c < rangeCols; c++) {
                const colValues = sourceValues.map(row => row[c]);
                const seedValues = forward ? colValues : colValues.slice().reverse();
                const seriesDir = forward ? directionSign : 1;
                const series = this.buildSeriesFromValues(seedValues, fillCount, seriesDir);
                seriesByCol.push(series);
                if (series) hasSeries = true;
            }

            const useSeries = hasSeries ? (forceCopy ? false : (useSeriesPref ? true : await this.confirmFillSeries())) : false;

            for (let r = fillStart; r <= fillEnd; r++) {
                for (let c = startCol; c <= endCol; c++) {
                    const srcColIdx = c - startCol;
                    const series = seriesByCol[srcColIdx];
                    const stepIndex = forward ? (r - fillStart) : (fillEnd - r);
                    if (this.fillSkipBlanks && blankMask[(stepIndex % rangeRows)][srcColIdx]) continue;
                    const baseValue = useSeries && series
                        ? series[stepIndex]
                        : sourceValues[(stepIndex % rangeRows)][srcColIdx];
                    const srcRow = startRow + (stepIndex % rangeRows);
                    const srcCol = startCol + srcColIdx;
                    const rowDelta = deltaRow * (Math.floor(stepIndex / rangeRows) + 1);
                    const value = this.adjustFormulaForFill(baseValue, rowDelta, c - srcCol);
                    const cell = this.getCellEl(c, r);
                    if (cell) {
                        this.setRawValue(cell.dataset.id, value);
                        this.renderCellValue(cell);
                    }
                }
            }
        } else if (target.row >= startRow && target.row <= endRow) {
            // Horizontal fill
            const fillStart = target.colNum > endCol ? endCol + 1 : (target.colNum < startCol ? target.colNum : startCol);
            const fillEnd = target.colNum > endCol ? target.colNum : (target.colNum < startCol ? startCol - 1 : endCol);
            if (fillStart > fillEnd) return;
            const fillCount = fillEnd - fillStart + 1;
            const directionSign = target.colNum > endCol ? 1 : -1;
            const forward = target.colNum > endCol;
            const deltaCol = forward ? 1 : -1;

            const seriesByRow = [];
            let hasSeries = false;
            for (let r = 0; r < rangeRows; r++) {
                const rowValues = sourceValues[r];
                const seedValues = forward ? rowValues : rowValues.slice().reverse();
                const seriesDir = forward ? directionSign : 1;
                const series = this.buildSeriesFromValues(seedValues, fillCount, seriesDir);
                seriesByRow.push(series);
                if (series) hasSeries = true;
            }

            const useSeries = hasSeries ? (forceCopy ? false : (useSeriesPref ? true : await this.confirmFillSeries())) : false;

            for (let c = fillStart; c <= fillEnd; c++) {
                for (let r = startRow; r <= endRow; r++) {
                    const srcRowIdx = r - startRow;
                    const series = seriesByRow[srcRowIdx];
                    const stepIndex = forward ? (c - fillStart) : (fillEnd - c);
                    if (this.fillSkipBlanks && blankMask[srcRowIdx][(stepIndex % rangeCols)]) continue;
                    const baseValue = useSeries && series
                        ? series[stepIndex]
                        : sourceValues[srcRowIdx][(stepIndex % rangeCols)];
                    const srcRow = startRow + srcRowIdx;
                    const srcCol = startCol + (stepIndex % rangeCols);
                    const colDelta = deltaCol * (Math.floor(stepIndex / rangeCols) + 1);
                    const value = this.adjustFormulaForFill(baseValue, r - srcRow, colDelta);
                    const cell = this.getCellEl(c, r);
                    if (cell) {
                        this.setRawValue(cell.dataset.id, value);
                        this.renderCellValue(cell);
                    }
                }
            }
        }

        this.markDirty();
        this.updateItemCount();
        this.refreshFindIfActive();
    }

    async confirmFillSeries() {
        return await this.showConfirmAsyncNextTick('패턴이 감지되었습니다. 시리즈 채우기를 적용할까요?\n취소를 누르면 값 복사로 채웁니다.');
    }

    showConfirmAsyncNextTick(message) {
        return new Promise((resolve) => {
            requestAnimationFrame(() => {
                this.showConfirmAsync(message).then(resolve);
            });
        });
    }

    updateFillPreview() {
        if (!this.isFilling || !this.lastFillTargetCell || !this.fillPreview) return;
        const range = this.fillRange;
        if (!range) return;
        const target = this.parseCellId(this.lastFillTargetCell.dataset.id);
        let direction = null;
        if (target.colNum >= range.startCol && target.colNum <= range.endCol) direction = 'v';
        if (target.row >= range.startRow && target.row <= range.endRow) direction = direction ? null : 'h';
        if (!direction) return;

        const startCell = this.getCellEl(range.startCol, range.startRow);
        if (!startCell) return;
        const rect = startCell.getBoundingClientRect();
        const containerRect = this.container.getBoundingClientRect();
        const left = rect.left - containerRect.left + this.container.scrollLeft;
        const top = rect.top - containerRect.top + this.container.scrollTop;
        this.fillPreview.style.left = `${left}px`;
        this.fillPreview.style.top = `${top - 26}px`;
        this.fillPreview.style.display = 'block';
        const label = this.fillPreview.querySelector('.fill-preview-label');
        if (label) label.textContent = this.fillPreviewMode === 'series' ? '시리즈 채우기' : '값 복사';
    }

    buildSeriesFromValues(values, count, directionSign) {
        if (!Array.isArray(values) || values.length === 0 || count <= 0) return null;
        const trimmed = values.map(v => (v === null || v === undefined) ? '' : String(v));
        if (trimmed.some(v => v.trim().startsWith('='))) return null;

        const numeric = this.buildNumericSeries(trimmed, count, directionSign);
        if (numeric) return numeric;

        const dateSeries = this.buildDateSeries(trimmed, count, directionSign);
        if (dateSeries) return dateSeries;

        const listSeries = this.buildListSeries(trimmed, count, directionSign);
        if (listSeries) return listSeries;

        const textNumSeries = this.buildTextNumberSeries(trimmed, count, directionSign);
        if (textNumSeries) return textNumSeries;

        return null;
    }

    buildNumericSeries(values, count, directionSign) {
        const nums = [];
        for (const raw of values) {
            if (!/^\s*[+-]?(\d{1,3}(,\d{3})*|\d+)(\.\d+)?\s*$/.test(raw)) {
                return null;
            }
            const trimmed = raw.trim();
            if (/^[+-]?0\d+/.test(trimmed) && !/^0(\.\d+)?$/.test(trimmed)) {
                return null;
            }
            const n = Number(String(raw).replace(/,/g, '').trim());
            if (!Number.isFinite(n)) return null;
            nums.push(n);
        }
        const decimals = this.getDecimalPlaces(values[values.length - 1]);
        if (nums.length >= 2) {
            const diffs = [];
            for (let i = 1; i < nums.length; i++) {
                diffs.push(nums[i] - nums[i - 1]);
            }
            const allSameDiff = diffs.every(d => Math.abs(d - diffs[0]) < 1e-9);
            const ratioValid = nums.every(n => n !== 0);
            let ratio = null;
            if (ratioValid) {
                const ratios = [];
                for (let i = 1; i < nums.length; i++) {
                    ratios.push(nums[i] / nums[i - 1]);
                }
                if (ratios.every(r => Math.abs(r - ratios[0]) < 1e-9)) ratio = ratios[0];
            }

            if (ratio !== null && ratio !== 1 && !allSameDiff) {
                const stepRatio = directionSign < 0 ? 1 / ratio : ratio;
                const series = [];
                let current = nums[nums.length - 1];
                for (let i = 0; i < count; i++) {
                    current *= stepRatio;
                    series.push(this.formatSeriesNumber(current, decimals));
                }
                return series;
            }

            if (allSameDiff) {
                let step = diffs[0];
                if (directionSign < 0) step = -step;
                const series = [];
                let current = nums[nums.length - 1];
                for (let i = 0; i < count; i++) {
                    current += step;
                    series.push(this.formatSeriesNumber(current, decimals));
                }
                return series;
            }
        }

        const stepBase = 1;
        const step = directionSign < 0 ? -stepBase : stepBase;
        const series = [];
        let current = nums[nums.length - 1];
        for (let i = 0; i < count; i++) {
            current += step;
            series.push(this.formatSeriesNumber(current, decimals));
        }
        return series;
    }

    buildDateSeries(values, count, directionSign) {
        const dates = [];
        for (const raw of values) {
            if (!/^\s*\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}\s*$/.test(raw)) return null;
            const dt = this.parseDateFromRaw(raw);
            if (!dt) return null;
            dates.push(dt);
        }
        const last = dates[dates.length - 1];
        const prev = dates.length >= 2 ? dates[dates.length - 2] : null;
        let stepDays = 1;
        if (prev) {
            stepDays = Math.round((last - prev) / 86400000) || 1;
        }
        if (directionSign < 0) stepDays = -stepDays;
        const series = [];
        let current = new Date(last.getTime());
        for (let i = 0; i < count; i++) {
            current = new Date(current.getTime() + stepDays * 86400000);
            series.push(this.formatDateIso(current));
        }
        return series;
    }

    buildListSeries(values, count, directionSign) {
        const lists = [
            ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'],
            ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
            ['SUNDAY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY'],
            ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
            ['일', '월', '화', '수', '목', '금', '토'],
            ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'],
            ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
            ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'],
            ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
            ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월']
        ];
        this.customLists.forEach(list => lists.push(list));

        const upperValues = values.map(v => v.trim().toUpperCase());
        for (const list of lists) {
            const upperList = list.map(v => v.toUpperCase());
            const indices = [];
            for (const v of upperValues) {
                const idx = upperList.indexOf(v);
                if (idx === -1) {
                    indices.length = 0;
                    break;
                }
                indices.push(idx);
            }
            if (indices.length === 0) continue;

            let step = indices.length >= 2 ? (indices[indices.length - 1] - indices[indices.length - 2]) : 1;
            if (directionSign < 0) step = -step;
            const series = [];
            let current = indices[indices.length - 1];
            for (let i = 0; i < count; i++) {
                current = (current + step) % list.length;
                if (current < 0) current += list.length;
                series.push(list[current]);
            }
            return series;
        }
        return null;
    }

    buildTextNumberSeries(values, count, directionSign) {
        const parts = values.map(v => v.match(/^(.*?)(-?\d+)$/));
        if (parts.some(p => !p)) return null;
        const prefix = parts[0][1];
        if (!parts.every(p => p[1] === prefix)) return null;
        const nums = parts.map(p => Number(p[2]));
        if (nums.some(n => !Number.isFinite(n))) return null;
        const lastNumStr = parts[parts.length - 1][2];
        const padWidth = lastNumStr.startsWith('-') ? lastNumStr.length - 1 : lastNumStr.length;
        const stepBase = nums.length >= 2 ? (nums[nums.length - 1] - nums[nums.length - 2]) : 1;
        const step = directionSign < 0 ? -stepBase : stepBase;
        const series = [];
        let current = nums[nums.length - 1];
        for (let i = 0; i < count; i++) {
            current += step;
            series.push(prefix + this.formatPaddedNumber(current, padWidth));
        }
        return series;
    }

    getDecimalPlaces(value) {
        const text = String(value).trim();
        const match = text.match(/\.(\d+)\s*$/);
        return match ? match[1].length : 0;
    }

    formatSeriesNumber(num, decimals) {
        if (decimals > 0) return num.toFixed(decimals);
        return String(Number.isFinite(num) ? num : 0);
    }

    formatDateIso(date) {
        const yyyy = date.getFullYear();
        const mm = String(date.getMonth() + 1).padStart(2, '0');
        const dd = String(date.getDate()).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}`;
    }

    formatPaddedNumber(num, width) {
        if (!Number.isFinite(num)) return String(num);
        const sign = num < 0 ? '-' : '';
        const abs = Math.abs(Math.trunc(num));
        const str = String(abs).padStart(width, '0');
        return sign + str;
    }

    loadCustomLists() {
        try {
            const raw = localStorage.getItem('vs_custom_lists');
            if (!raw) return [];
            const parsed = JSON.parse(raw);
            if (!Array.isArray(parsed)) return [];
            return parsed
                .map(line => Array.isArray(line) ? line : String(line).split(','))
                .map(list => list.map(item => String(item).trim()).filter(Boolean))
                .filter(list => list.length > 0);
        } catch (err) {
            return [];
        }
    }

    saveCustomLists() {
        try {
            const serialized = this.customLists.map(list => list.join(','));
            localStorage.setItem('vs_custom_lists', JSON.stringify(serialized));
        } catch (err) {
            console.warn('Failed to save custom lists', err);
        }
    }

    bindCustomListModal() {
        const modal = document.getElementById('custom-list-modal');
        const input = document.getElementById('custom-list-input');
        const btnSave = document.getElementById('custom-list-save');
        const btnCancel = document.getElementById('custom-list-cancel');
        const btnClose = document.getElementById('custom-list-close');
        if (!modal || !input || !btnSave || !btnCancel || !btnClose) return;

        const close = () => { modal.style.display = 'none'; };
        btnCancel.addEventListener('click', close);
        btnClose.addEventListener('click', close);

        btnSave.addEventListener('click', () => {
            const lines = input.value.split('\n').map(line => line.trim()).filter(Boolean);
            const lists = lines.map(line => line.split(',').map(item => item.trim()).filter(Boolean)).filter(list => list.length > 0);
            this.customLists = lists;
            this.saveCustomLists();
            close();
        });
    }

    openCustomListModal() {
        const modal = document.getElementById('custom-list-modal');
        const input = document.getElementById('custom-list-input');
        if (!modal || !input) return;
        const lines = (this.customLists || []).map(list => list.join(', '));
        input.value = lines.join('\n');
        modal.style.display = 'flex';
        input.focus();
    }

    isBlankValue(value) {
        if (value === null || value === undefined) return true;
        return String(value).trim() === '';
    }

    adjustFormulaForFill(rawValue, rowDelta, colDelta) {
        if (rawValue === null || rawValue === undefined) return rawValue;
        const text = String(rawValue);
        const trimmed = text.trim();
        if (!trimmed.startsWith('=')) return rawValue;

        const body = trimmed.slice(1);
        let out = '';
        let i = 0;
        while (i < body.length) {
            const ch = body[i];
            if (ch === '"') {
                out += ch;
                i++;
                while (i < body.length) {
                    const c = body[i];
                    const next = body[i + 1];
                    out += c;
                    i++;
                    if (c === '"' && next === '"') {
                        out += '"';
                        i++;
                        continue;
                    }
                    if (c === '"') break;
                }
                continue;
            }
            let j = i;
            while (j < body.length && body[j] !== '"') j++;
            const segment = body.slice(i, j);
            out += segment.replace(/(\$?[A-Z]+)(\$?\d+)/gi, (match, colPart, rowPart) => {
                const colAbs = colPart.startsWith('$');
                const rowAbs = rowPart.startsWith('$');
                const colLetters = colPart.replace('$', '').toUpperCase();
                const rowNum = parseInt(rowPart.replace('$', ''), 10);
                if (!Number.isFinite(rowNum)) return match;
                let newColNum = this.colToNumber(colLetters);
                let newRowNum = rowNum;
                if (!colAbs) newColNum += colDelta;
                if (!rowAbs) newRowNum += rowDelta;
                if (newColNum < 1 || newRowNum < 1) return match;
                const newCol = (colAbs ? '$' : '') + this.numberToCol(newColNum);
                const newRow = (rowAbs ? '$' : '') + newRowNum;
                return `${newCol}${newRow}`;
            });
            i = j;
        }
        return '=' + out;
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

        // Determine paste range and handle merged cells policy
        const numRows = rowsData.length;
        const numCols = rowsData[0].length;
        const pasteRange = {
            startCol: anchor.colNum,
            startRow: anchor.row,
            endCol: anchor.colNum + numCols - 1,
            endRow: anchor.row + numRows - 1
        };

        const isSingleCellPaste = numRows === 1 && numCols === 1;
        if (!isSingleCellPaste && this.rangeIntersectsMerges(pasteRange)) {
            this.unmergeRange(pasteRange);
        }

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
                        this.setRawValue(cell.dataset.id, val);
                        this.renderCellValue(cell);
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
                        this.setRawValue(cell.dataset.id, '');
                        this.renderCellValue(cell);
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
                    this.setRawValue(cell.dataset.id, '');
                    this.renderCellValue(cell);
                    changed = true;
                }
            }
        }
        if (changed) {
            this.markDirty();
            this.updateItemCount();
            this.refreshFindIfActive();
            this.formulaCache.clear();
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

        // Global shortcuts like Ctrl+C/V are handled globally (setupEventListeners)
        if (e.ctrlKey || e.metaKey) return;

        // F2 for Edit Mode
        if (e.key === 'F2' && !this.isEditing) {
            e.preventDefault();
            this.enterEditMode(activeCell);
            return;
        }

        // Esc for Canceletion
        if (e.key === 'Escape') {
            if (this.isEditing) {
                e.preventDefault();
                this.exitEditMode(false);
            } else {
                this.clearRangeSelection();
                this.updateFillHandlePosition();
            }
            return;
        }

        // If in Edit mode, limit keyboard navigation
        if (this.isEditing) {
            if (e.key === 'Enter') {
                e.preventDefault();
                this.exitEditMode(true);
                // Move down after commit
                this.moveSelection(0, 1);
            } else if (e.key === 'Tab') {
                e.preventDefault();
                this.exitEditMode(true);
                this.moveSelection(e.shiftKey ? -1 : 1, 0);
            }
            e.stopPropagation(); // Prevent bubbling to global listeners (like Delete)
            return; // Arrows move cursor normally in Edit mode
        }

        // If in Ready mode (not editing)
        // Check for direct typing (start overwriting)
        if (e.key.length === 1 && !e.ctrlKey && !e.altKey && !e.metaKey) {
            const isImeKey = e.isComposing || e.key === 'Process' || e.key === 'Unidentified';
            if (isImeKey) {
                this.isEditing = true; // Flag we are now in edit state
                return;
            }
            
            // Transition to editing mode without manual clearing (the selection from handleCellFocus will handle overwriting)
            this.isEditing = true;
            this.originalValue = this.getRawValue(activeCell.dataset.id);
            activeCell.classList.add('editing');
            
            // Note: We don't call prepareEnterMode here because that clears the text manually,
            // which breaks IME hook. The "Select All" in handleCellFocus handles the clean overwrite.
            return;
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
                e.preventDefault();
                break;
            case 'ArrowDown':
                if (e.shiftKey) {
                    e.preventDefault();
                    this.extendSelectionByKey(colNum, rowNum, colNum, rowNum + 1);
                    return;
                }
                if (rowNum < this.rows) nextRow++;
                moved = true;
                e.preventDefault();
                break;
            case 'ArrowLeft':
                if (e.shiftKey) {
                    e.preventDefault();
                    this.extendSelectionByKey(colNum, rowNum, colNum - 1, rowNum);
                    return;
                }
                if (colNum > 1) nextCol = colNum - 1;
                moved = true;
                e.preventDefault();
                break;
            case 'ArrowRight':
                if (e.shiftKey) {
                    e.preventDefault();
                    this.extendSelectionByKey(colNum, rowNum, colNum + 1, rowNum);
                    return;
                }
                if (colNum < this.cols) nextCol = colNum + 1;
                moved = true;
                e.preventDefault(); // Stop default browser scroll
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
                e.preventDefault();
                if (e.shiftKey) {
                    if (rowNum > 1) nextRow--;
                } else {
                    if (rowNum < this.rows) nextRow++;
                }
                moved = true;
                break;
            case 'Delete':
            case 'Backspace':
                if (this.selectionRange) {
                    const { startCol, startRow, endCol, endRow } = this.selectionRange;
                    if (startCol !== endCol || startRow !== endRow) {
                        e.preventDefault();
                        this.deleteSelection();
                        return;
                    }
                }
                // Single cell delete
                this.setRawValue(activeCell.dataset.id, '');
                this.renderCellValue(activeCell);
                this.formulaInput.value = '';
                this.markDirty();
                this.updateItemCount();
                this.refreshFindIfActive();
                this.formulaCache.clear();
                break;
        }

        if (moved) {
            this.clearRangeSelection();
            this.setSelectionRange(nextCol, nextRow, nextCol, nextRow);
            const nextCell = this.getSelectableCell(nextCol, nextRow);
            if (nextCell) {
                nextCell.focus({ preventScroll: true });
                // Note: focus event will trigger handleCellFocus(nextCell) automatically
                this.updateFillHandlePosition();
            }
        }
    }

    // ─── Mode Handlers ─────────────────────────────────────
    enterEditMode(cell) {
        if (this.isEditing) return;
        const cellId = cell.dataset.id;
        const normalizedId = this.normalizeMergedCellId(cellId);
        if (normalizedId !== cellId) {
            const { colNum, row } = this.parseCellId(normalizedId);
            const anchorCell = this.getCellEl(colNum, row);
            if (anchorCell) {
                cell = anchorCell;
            }
        }
        if (cell.classList.contains('merge-hidden')) return;
        this.isEditing = true;
        this.needsOverwrite = false;
        this.originalValue = this.getRawValue(cellId);
        cell.contentEditable = true;
        cell.classList.add('editing');
        cell.innerText = this.originalValue;
        cell.focus();
        
        // Move cursor to end
        const range = document.createRange();
        const sel = window.getSelection();
        range.selectNodeContents(cell);
        range.collapse(false);
        sel.removeAllRanges();
        sel.addRange(range);
    }

    prepareEnterMode(cell, overwrite = true) {
        if (this.isEditing) return;
        const cellId = cell.dataset.id;
        const normalizedId = this.normalizeMergedCellId(cellId);
        if (normalizedId !== cellId) {
            const { colNum, row } = this.parseCellId(normalizedId);
            const anchorCell = this.getCellEl(colNum, row);
            if (anchorCell) {
                cell = anchorCell;
            }
        }
        if (cell.classList.contains('merge-hidden')) return;
        this.isEditing = true;
        this.originalValue = this.getRawValue(cellId);
        this.needsOverwrite = overwrite;
        
        cell.classList.add('editing');
        cell.innerText = overwrite ? '' : this.originalValue;
        
        // CRITICAL: Setting innerText = '' often clears the caret in some browsers.
        // We must re-establish a selection within the cell so typing works.
        const range = document.createRange();
        const sel = window.getSelection();
        range.selectNodeContents(cell);
        range.collapse(false); // End of empty is same as start
        sel.removeAllRanges();
        sel.addRange(range);
        
        this.markDirty();
    }

    enterEnterMode(cell) {
        this.prepareEnterMode(cell);
    }

    exitEditMode(commit = true) {
        if (!this.isEditing || !this.selectedCell) return;
        
        const cell = this.selectedCell;
        const cellId = cell.dataset.id;
        if (!commit) {
            this.setRawValue(cellId, this.originalValue);
        } else {
            this.setRawValue(cellId, cell.innerText);
            this.formulaInput.value = this.getRawValue(cellId);
            this.markDirty();
            this.updateItemCount();
        }
        
        cell.contentEditable = false;
        cell.classList.remove('editing');
        this.isEditing = false;
        this.needsOverwrite = false;
        this.renderCellValue(cell);
        this.formulaInput.value = this.getRawValue(cellId);
        cell.focus(); // Keep focus for Ready mode navigation
    }

    moveSelection(deltaCol, deltaRow) {
        const activeCell = this.selectedCell;
        if (!activeCell) return;
        
        const { colNum, row } = this.parseCellId(activeCell.dataset.id);
        const nextCol = Math.max(1, Math.min(this.cols, colNum + deltaCol));
        const nextRow = Math.max(1, Math.min(this.rows, row + deltaRow));
        
        const nextCell = this.getSelectableCell(nextCol, nextRow);
        if (nextCell) {
            nextCell.focus();
            this.handleCellFocus(nextCell);
            this.setSelectionRange(nextCol, nextRow, nextCol, nextRow);
            this.updateFillHandlePosition();
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
        const targetCell = this.getSelectableCell(targetCol, targetRow);
        if (targetCell) {
            targetCell.focus();
            this.selectedCell = targetCell;
            this.cellAddress.innerText = targetCell.dataset.id;
            this.formulaInput.value = this.getRawValue(targetCell.dataset.id);
        }
    }

    // ─── File Import / Export ───────────────────────────────
    // File Import / Export methods are loaded from engines/vs_io.js
    async openFileDialog() {
        return window.VSIO.openFileDialog(this);
    }

    handleFileSelect(e) {
        return window.VSIO.handleFileSelect(this, e);
    }

    // File Import / Export methods are loaded from engines/vs_io.js

    processFile(file, handle = null) {
        return window.VSIO.processFile(this, file, handle);
        // Check if there's existing data
        const hasData = Object.keys(this.data).some(k => this.data[k] && this.data[k].trim() !== '');
        if (hasData) {
            if (!confirm('작업 중인 내용이 덮어씌워질 수 있습니다. 계속할까요?\n(Unsaved changes will be lost. Continue?)')) {
                return;
            }
        }

        // Confirmation passed: update handle and process
        this.fileHandle = handle;

        const extension = file.name.split('.').pop().toLowerCase();
        const reader = new FileReader();

        reader.onload = async (event) => {
            try {
                if (extension === 'xlsx' || extension === 'xls') {
                    await this.importXLSX(event.target.result);
                } else if (extension === 'vsht') {
                    this.importVSHT(event.target.result);
                } else {
                    this.importFromText(event.target.result);
                }

                // Update filename display
                const filenameEl = document.querySelector('.filename');
                if (filenameEl) {
                    filenameEl.innerText = file.name.replace(/\.[^.]+$/, '');
                }
            } catch (err) {
                console.error('File import failed:', err);
                alert('파일을 불러오지 못했습니다.');
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
        return window.VSIO.importVSHT(this, jsonText);
    }

    // 2. .xlsx Import (using SheetJS)
    async importXLSX(buffer) {
        return window.VSIO.importXLSX(this, buffer);
        return this.importXLSXExcelJS(buffer);
    }

    async importXLSXExcelJS(buffer) {
        return window.VSIO.importXLSXExcelJS(this, buffer);
        if (typeof ExcelJS === 'undefined') {
            alert('Excel 라이브러리를 불러오지 못했습니다. 네트워크 연결을 확인해 주세요.');
            return;
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);

        this.sheets = [];
        workbook.eachSheet((worksheet, sheetIndex) => {
            const name = worksheet.name || `Sheet${sheetIndex}`;
            const sheet = this.createSheet(name);

            sheet.rows = Math.max(this.baseRows, worksheet.rowCount || 0);
            sheet.cols = Math.max(this.baseCols, worksheet.columnCount || 0);
            sheet.colWidths = new Array(sheet.cols).fill(100);
            sheet.rowHeights = new Array(sheet.rows + 1).fill(25);

            worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                    const cellId = `${this.numberToCol(colNumber)}${rowNumber}`;
                    let raw = '';

                    if (cell.type === ExcelJS.ValueType.Formula) {
                        this.setRawValueForSheet(sheet, cellId, '=' + cell.formula);
                    } else if (cell.type === ExcelJS.ValueType.Date && cell.value instanceof Date) {
                        const yyyy = cell.value.getFullYear();
                        const mm = String(cell.value.getMonth() + 1).padStart(2, '0');
                        const dd = String(cell.value.getDate()).padStart(2, '0');
                        raw = `${yyyy}-${mm}-${dd}`;
                        this.setRawValueForSheet(sheet, cellId, raw);
                    } else {
                        raw = cell.value === null || cell.value === undefined ? '' : String(cell.value);
                        this.setRawValueForSheet(sheet, cellId, raw);
                    }
                    if (cell.numFmt) {
                        sheet.cellFormats[cellId] = this.getExceljsFormatFromNumFmt(cell.numFmt, cell.type);
                    }
                    if (cell.font || cell.alignment || cell.fill) {
                        this.mapExceljsStyleToSheet(sheet, cellId, cell);
                    }
                });
            });

            const mergeRefs = new Set();
            if (worksheet.model && Array.isArray(worksheet.model.merges)) {
                worksheet.model.merges.forEach(ref => mergeRefs.add(ref));
            }
            if (worksheet._merges) {
                if (worksheet._merges instanceof Map) {
                    Array.from(worksheet._merges.keys()).forEach(ref => mergeRefs.add(ref));
                } else if (Array.isArray(worksheet._merges)) {
                    worksheet._merges.forEach(ref => mergeRefs.add(ref));
                } else {
                    Object.keys(worksheet._merges).forEach(ref => mergeRefs.add(ref));
                }
            }

            sheet.mergedRanges = Array.from(mergeRefs)
                .map((ref) => this.parseRangeRef(ref))
                .filter(Boolean);

            this.sheets.push(sheet);
        });

        if (this.sheets.length === 0) {
            this.sheets = [this.createSheet('Sheet1')];
        }
        this.activeSheetIndex = 0;
        this.refreshGridUI();
        this.renderSheetTabs();
        this.updateItemCount();
        this.markClean();
    }

    // 3. CSV/TSV Text Import
    importFromText(text) {
        return window.VSIO.importFromText(this, text);
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
        return window.VSIO.setInternalData(this, rowNum, colNum, value);
        // Automatically grow if needed (just in case)
        if (rowNum > this.rows) {
            this.createRowElements(this.rows + 1, rowNum);
            this.rows = rowNum;
        }
        const cellId = `${this.numberToCol(colNum)}${rowNum}`;
        this.setRawValue(cellId, value);
    }

    refreshGridUI() {
        if (this.container) {
            this.container.innerHTML = '';
            
            // Re-create overlays that were inside the container
            this.selectionOverlay = this.createOverlay('selection-overlay');
            this.rangeOverlay = this.createOverlay('range-overlay');
            this.fillHandle = this.createOverlay('fill-handle');
            this.resizeGuide = this.createOverlay('resize-guide');
            
            this.renderGrid();
            
            // Re-attach fill handle listener (since we just created a new one)
            this.fillHandle.addEventListener('mousedown', (e) => this.handleFillStart(e));
            
            // Refresh overlays
            this.updateSelectionOverlay();
            this.updateRangeVisual();
            this.updateFillHandlePosition();
            this.refreshFindIfActive();
        }
    }

    detectDelimiter(text) {
        return window.VSIO.detectDelimiter(this, text);
        const firstLine = text.split('\n')[0] || '';
        const commas = (firstLine.match(/,/g) || []).length;
        const tabs = (firstLine.match(/\t/g) || []).length;
        const semis = (firstLine.match(/;/g) || []).length;

        if (tabs >= commas && tabs >= semis && tabs > 0) return '\t';
        if (semis > commas && semis > 0) return ';';
        return ',';
    }

    async saveFile() {
        return window.VSIO.saveFile(this);
        if (!this.fileHandle) {
            return this.saveFileAs();
        }

        try {
            const fileName = (this.fileHandle && this.fileHandle.name) ? String(this.fileHandle.name) : '';
            const lowerName = fileName.toLowerCase();
            if (lowerName.endsWith('.csv')) {
                if (!await this.confirmCsvSingleSheet()) return;
            }

            const writable = await this.fileHandle.createWritable();

            if (lowerName.endsWith('.xlsx')) {
                if (!await this.confirmXlsxStyleWarning()) return;
                const buffer = await this.generateXLSXBufferExcelJS();
                if (buffer) await writable.write(buffer);
            } else if (lowerName.endsWith('.csv')) {
                const csvContent = this.generateCSVContent();
                await writable.write(csvContent);
            } else {
                // Default to .vsht (JSON)
                const vshtData = this.generateVSHTData();
                await writable.write(JSON.stringify(vshtData, null, 2));
            }
            
            await writable.close();
            this.markClean();
        } catch (err) {
            console.error('Save failed, using Save As:', err);
            this.saveFileAs();
        }
    }

    async saveFileAs() {
        return window.VSIO.saveFileAs(this);
        const defaultName = (document.querySelector('.filename')?.innerText || 'VibrantSheets').trim();

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
                            description: 'Excel Workbook (.xlsx)',
                            accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
                        },
                        {
                            description: 'CSV File (.csv)',
                            accept: { 'text/csv': ['.csv'] },
                        }
                    ],
                });
                
                this.fileHandle = handle;
                await this.saveFile(); // Overwrite with new handle

                const filenameEl = document.querySelector('.filename');
                if (filenameEl) {
                    filenameEl.innerText = handle.name.replace(/\.[^.]+$/, '');
                }
            } catch (err) {
                if (err.name === 'AbortError') return;
                console.error('Save As failed:', err);
            }
        } else {
            // Legacy Fallback (Download .vsht)
            const ext = (defaultName.split('.').pop() || '').toLowerCase();
            if (ext !== 'csv') {
                const vshtData = this.generateVSHTData();
                const blob = new Blob([JSON.stringify(vshtData, null, 2)], { type: 'application/json' });
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = `${defaultName}.vsht`;
                link.click();
                URL.revokeObjectURL(url);
                this.markClean();
            }
        }
    }


    generateVSHTData() {
        return window.VSIO.generateVSHTData(this);
        return {
            version: "1.0",
            title: document.querySelector('.filename')?.innerText || 'Untitled',
            sheets: this.sheets.map(sheet => ({
                name: sheet.name,
                rows: sheet.rows,
                cols: sheet.cols,
                data: sheet.data,
                colWidths: sheet.colWidths,
                rowHeights: sheet.rowHeights,
                cellStyles: sheet.cellStyles,
                cellFormats: sheet.cellFormats,
                cellFormulas: sheet.cellFormulas,
                mergedRanges: sheet.mergedRanges || []
            })),
            activeSheetIndex: this.activeSheetIndex
        };
    }

    getUsedRangeForSheet(sheet) {
        return window.VSIO.getUsedRangeForSheet(this, sheet);
        let maxRow = 0;
        let maxCol = 0;
        const scan = (key) => {
            const { colNum, row } = this.parseCellId(key);
            maxRow = Math.max(maxRow, row);
            maxCol = Math.max(maxCol, colNum);
        };
        Object.keys(sheet.data || {}).forEach(scan);
        Object.keys(sheet.cellFormulas || {}).forEach(scan);
        Object.keys(sheet.cellStyles || {}).forEach(scan);
        Object.keys(sheet.cellFormats || {}).forEach(scan);
        (sheet.mergedRanges || []).forEach((range) => {
            const normalized = this.normalizeMergedRangeEntry(range);
            if (!normalized) return;
            maxRow = Math.max(maxRow, normalized.endRow);
            maxCol = Math.max(maxCol, normalized.endCol);
        });
        return { maxRow, maxCol };
    }

    async generateXLSXBufferExcelJS() {
        return window.VSIO.generateXLSXBufferExcelJS(this);
        if (typeof ExcelJS === 'undefined') {
            alert('Excel 라이브러리를 불러오지 못했습니다. 네트워크 연결을 확인해 주세요.');
            return null;
        }

        const wb = new ExcelJS.Workbook();
        this.sheets.forEach((sheet, index) => {
            const ws = wb.addWorksheet(sheet.name || `Sheet${index + 1}`);
            const { maxRow, maxCol } = this.getUsedRangeForSheet(sheet);

            for (let r = 1; r <= maxRow; r++) {
                for (let c = 1; c <= maxCol; c++) {
                    const cellId = `${this.numberToCol(c)}${r}`;
                    const rawValue = this.getRawValueForSheet(sheet, cellId);
                    const format = this.getCellFormatForSheet(sheet, cellId);
                    const numericValue = this.parseNumberFromRaw(rawValue, format.type);
                    const dateValue = this.parseDateFromRaw(rawValue);

                    const cell = ws.getCell(r, c);
                    if (rawValue.trim().startsWith('=')) {
                        const formula = rawValue.trim().slice(1);
                        const result = this.formulaEngine
                            ? this.formulaEngine.evaluate(rawValue, this.getFormulaContext(), new Set())
                            : undefined;
                        cell.value = { formula, result: result === '#ERROR' ? undefined : result };
                    } else if (format.type === 'date' && dateValue) {
                        cell.value = dateValue;
                        cell.numFmt = 'yyyy-mm-dd';
                    } else if (numericValue !== null && format.type !== 'date') {
                        cell.value = numericValue;
                        const numFmt = this.getNumFmtFromFormat(format);
                        if (numFmt) cell.numFmt = numFmt;
                    } else {
                        cell.value = rawValue;
                    }

                    const style = sheet.cellStyles[cellId];
                    if (style) {
                        this.mapInternalStyleToExceljs(cell, style);
                    }
                }
            }

            const merges = this.getNormalizedMergedRanges(sheet);
            merges.forEach((range) => {
                ws.mergeCells(range.startRow, range.startCol, range.endRow, range.endCol);
            });
        });

        return wb.xlsx.writeBuffer();
    }

    generateXLSXBuffer() {
        return window.VSIO.generateXLSXBuffer(this);
        if (typeof XLSX === 'undefined') {
            alert('Excel 라이브러리를 찾을 수 없습니다.');
            return null;
        }

        const wb = XLSX.utils.book_new();
        this.sheets.forEach((sheet, index) => {
            // Find used range
            let maxRow = 0, maxCol = 0;
            for (const key in sheet.data) {
                if (sheet.data[key] && sheet.data[key].trim() !== '') {
                    const { colNum, row } = this.parseCellId(key);
                    maxRow = Math.max(maxRow, row);
                    maxCol = Math.max(maxCol, colNum);
                }
            }

            const aoa = [];
            for (let r = 1; r <= maxRow; r++) {
                const rowArr = [];
                for (let c = 1; c <= maxCol; c++) {
                    const cellId = `${this.numberToCol(c)}${r}`;
                    const rawValue = this.getRawValueForSheet(sheet, cellId);
                    const format = this.getCellFormatForSheet(sheet, cellId);
                    const numericValue = this.parseNumberFromRaw(rawValue, format.type);
                    const dateValue = this.parseDateFromRaw(rawValue);

                    let cellObj = { v: rawValue, t: 's' };
                    if (rawValue !== '') {
                        if (format.type === 'date' && dateValue) {
                            cellObj = { v: this.toExcelDateSerial(dateValue), t: 'n', z: 'yyyy-mm-dd' };
                        } else if (format.type === 'currency' && numericValue !== null) {
                            const d = format.decimals ?? 2;
                            cellObj = { v: numericValue, t: 'n', z: `"₩"#,##0${d > 0 ? '.' + '0'.repeat(d) : ''}` };
                        } else if (format.type === 'percentage' && numericValue !== null) {
                            const d = format.decimals ?? 2;
                            cellObj = { v: numericValue, t: 'n', z: `0${d > 0 ? '.' + '0'.repeat(d) : ''}%` };
                        } else if (format.decimals !== null && numericValue !== null) {
                            const d = format.decimals;
                            cellObj = { v: numericValue, t: 'n', z: `#,##0${d > 0 ? '.' + '0'.repeat(d) : ''}` };
                        } else if (numericValue !== null && /^\s*[+-]?\d+(\.\d+)?\s*$/.test(rawValue)) {
                            cellObj = { v: numericValue, t: 'n' };
                        }
                    }

                    const style = sheet.cellStyles[cellId];
                    if (style) {
                        cellObj.s = this.mapInternalStyleToXlsx(style);
                    }
                    rowArr.push(cellObj);
                }
                aoa.push(rowArr);
            }

            const ws = XLSX.utils.aoa_to_sheet(aoa);
            const sheetName = sheet.name || `Sheet${index + 1}`;
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
        });
        
        return XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellStyles: true });
    }

    generateCSVContent() {
        return window.VSIO.generateCSVContent(this);
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

    hexToArgb(hex) {
        if (!hex) return null;
        const clean = hex.replace('#', '').toUpperCase();
        if (clean.length === 6) return 'FF' + clean;
        if (clean.length === 8) return clean;
        return null;
    }

    argbToHex(argb) {
        if (!argb) return null;
        const clean = String(argb).replace('#', '').toUpperCase();
        if (clean.length === 8) return '#' + clean.slice(2);
        if (clean.length === 6) return '#' + clean;
        return null;
    }

    mapInternalStyleToExceljs(cell, style) {
        const font = {};
        if (style.fontWeight === 'bold') font.bold = true;
        if (style.fontStyle === 'italic') font.italic = true;
        if (style.textDecoration === 'underline') font.underline = true;
        if (style.fontSize) font.size = parseInt(style.fontSize, 10);
        if (style.fontFamily) font.name = style.fontFamily;
        if (style.color) {
            const argb = this.hexToArgb(style.color);
            if (argb) font.color = { argb };
        }
        if (Object.keys(font).length > 0) cell.font = font;

        if (style.textAlign) {
            cell.alignment = { horizontal: style.textAlign };
        }

        if (style.backgroundColor) {
            const argb = this.hexToArgb(style.backgroundColor);
            if (argb) {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb }
                };
            }
        }
    }

    mapExceljsStyleToSheet(sheet, cellId, cell) {
        const style = {};
        const font = cell.font || {};
        if (font.bold) style.fontWeight = 'bold';
        if (font.italic) style.fontStyle = 'italic';
        if (font.underline) style.textDecoration = 'underline';
        if (font.size) style.fontSize = font.size + 'pt';
        if (font.name) style.fontFamily = font.name;
        if (font.color && font.color.argb) style.color = this.argbToHex(font.color.argb);

        const alignment = cell.alignment || {};
        if (alignment.horizontal) style.textAlign = alignment.horizontal;

        const fill = cell.fill || {};
        if (fill.fgColor && fill.fgColor.argb) {
            style.backgroundColor = this.argbToHex(fill.fgColor.argb);
        }

        if (Object.keys(style).length > 0) {
            sheet.cellStyles[cellId] = style;
        }

        if (cell.border) {
            const toInternal = (side) => {
                const b = cell.border?.[side];
                if (!b || !b.style) return null;
                return {
                    style: b.style,
                    color: this.argbToHex(b.color?.argb),
                    width: 1
                };
            };
            const border = {
                top: toInternal('top'),
                right: toInternal('right'),
                bottom: toInternal('bottom'),
                left: toInternal('left')
            };
            const hasBorder = Object.values(border).some(Boolean);
            if (hasBorder) {
                if (!sheet.cellBorders) sheet.cellBorders = {};
                sheet.cellBorders[cellId] = border;
            }
        }
    }

    getNumFmtFromFormat(format) {
        if (format.type === 'currency') {
            const d = format.decimals ?? 2;
            return `"₩"#,##0${d > 0 ? '.' + '0'.repeat(d) : ''}`;
        }
        if (format.type === 'percentage') {
            const d = format.decimals ?? 2;
            return `0${d > 0 ? '.' + '0'.repeat(d) : ''}%`;
        }
        if (format.type === 'date') {
            return 'yyyy-mm-dd';
        }
        if (format.decimals !== null) {
            const d = format.decimals;
            return `#,##0${d > 0 ? '.' + '0'.repeat(d) : ''}`;
        }
        return null;
    }

    getExceljsFormatFromNumFmt(numFmt, cellType) {
        const formatCode = String(numFmt || '').toLowerCase();
        let type = 'general';
        let decimals = null;

        if (formatCode.includes('%')) {
            type = 'percentage';
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
        } else if (/\b(y|m|d|h|s)+\b/.test(formatCode) || cellType === ExcelJS.ValueType.Date) {
            type = 'date';
            decimals = null;
        } else if (formatCode.includes('[$') || formatCode.includes('₩') || formatCode.includes('$')) {
            type = 'currency';
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
        } else if (cellType === ExcelJS.ValueType.Number) {
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
        }

        return this.normalizeFormat({ type, decimals });
    }

    // ─── Style Mapping Helpers ─────────────────────────────
    mapInternalStyleToXlsx(style) {
        const s = { font: {}, alignment: {}, fill: {} };
        if (style.fontWeight === 'bold') s.font.bold = true;
        if (style.fontStyle === 'italic') s.font.italic = true;
        if (style.textDecoration === 'underline') s.font.underline = true;
        
        if (style.color) {
            s.font.color = { rgb: style.color.replace('#', '') };
        }
        if (style.backgroundColor) {
            s.fill.fgColor = { rgb: style.backgroundColor.replace('#', '') };
            s.fill.patternType = 'solid';
        }
        if (style.textAlign) {
            s.alignment.horizontal = style.textAlign;
        }
        return s;
    }

    mapXlsxStyleToInternal(cellId, s) {
        const style = {};
        if (s.font) {
            if (s.font.bold) style.fontWeight = 'bold';
            if (s.font.italic) style.fontStyle = 'italic';
            if (s.font.underline) style.textDecoration = 'underline';
            if (s.font.color && s.font.color.rgb) style.color = '#' + s.font.color.rgb;
        }
        if (s.fill && s.fill.fgColor && s.fill.fgColor.rgb) {
            style.backgroundColor = '#' + s.fill.fgColor.rgb;
        }
        if (s.alignment && s.alignment.horizontal) {
            style.textAlign = s.alignment.horizontal;
        }
        if (Object.keys(style).length > 0) {
            this.cellStyles[cellId] = style;
        }
    }

    mapXlsxStyleToSheet(sheet, cellId, s) {
        const style = {};
        if (s.font) {
            if (s.font.bold) style.fontWeight = 'bold';
            if (s.font.italic) style.fontStyle = 'italic';
            if (s.font.underline) style.textDecoration = 'underline';
            if (s.font.color && s.font.color.rgb) style.color = '#' + s.font.color.rgb;
        }
        if (s.fill && s.fill.fgColor && s.fill.fgColor.rgb) {
            style.backgroundColor = '#' + s.fill.fgColor.rgb;
        }
        if (s.alignment && s.alignment.horizontal) {
            style.textAlign = s.alignment.horizontal;
        }
        if (Object.keys(style).length > 0) {
            sheet.cellStyles[cellId] = style;
        }
    }

    extractDecimalPlacesFromFormat(formatCode) {
        if (!formatCode) return null;
        const match = formatCode.match(/\.(0+|#+)/);
        return match ? match[1].length : null;
    }

    getXlsxFormatFromCell(wsCell) {
        const formatCode = String(wsCell.z || '').toLowerCase();
        let type = 'general';
        let decimals = null;

        if (formatCode.includes('%')) {
            type = 'percentage';
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
        } else if (/\b[ymdhis]+\b/.test(formatCode) || wsCell.t === 'd') {
            type = 'date';
            decimals = null;
        } else if (formatCode.includes('[$') || formatCode.includes('₩') || formatCode.includes('$')) {
            type = 'currency';
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
        } else if (wsCell.t === 'n') {
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
        }

        return this.normalizeFormat({ type, decimals });
    }

    mapXlsxFormatToInternal(cellId, wsCell) {
        const formatCode = String(wsCell.z || '').toLowerCase();
        let type = 'general';
        let decimals = null;

        if (formatCode.includes('%')) {
            type = 'percentage';
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
        } else if (/\b[ymdhis]+\b/.test(formatCode) || wsCell.t === 'd') {
            type = 'date';
            decimals = null;
        } else if (formatCode.includes('[$') || formatCode.includes('₩') || formatCode.includes('$')) {
            type = 'currency';
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
        } else if (wsCell.t === 'n') {
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
        }

        if (type !== 'general' || decimals !== null) {
            this.setCellFormat(cellId, { type, decimals });
        }
    }

    clearAllData(shouldMarkDirty = true) {
        if (this.tbody) {
            this.tbody.querySelectorAll('.cell').forEach(cell => {
                cell.innerText = '';
                cell.style = ''; // Clear styles
            });
        }
        this.data = {};
        this.cellStyles = {};
        this.cellFormats = {};
        this.cellFormulas = {};
        this.cellBorders = {};
        this.mergedRanges = [];
        if (shouldMarkDirty) this.markDirty();
        this.updateItemCount();
        this.refreshFindIfActive();
    }

    updateItemCount() {
        const dataCount = Object.keys(this.data).filter(k => this.data[k] && this.data[k].trim() !== '').length;
        const formulaCount = Object.keys(this.cellFormulas).filter(k => this.cellFormulas[k] && this.cellFormulas[k].trim() !== '').length;
        const count = dataCount + formulaCount;
        const metricsSpan = document.querySelector('.metrics span:first-child');
        if (metricsSpan) {
            metricsSpan.innerText = `Items: ${count}`;
        }
    }

    // ─── Phase 6: Table Operations ────────────────────────
    getEffectiveRange() {
        if (this.selectionRange) {
            return this.selectionRange;
        }
        if (this.selectedCell) {
            const { colNum, row } = this.parseCellId(this.selectedCell.dataset.id);
            return { startCol: colNum, startRow: row, endCol: colNum, endRow: row };
        }
        return null;
    }

    insertRow() {
        const range = this.getEffectiveRange();
        const index = range ? range.startRow : 1;
        this.shiftData('row', index, 1);
        this.rowHeights.splice(index, 0, 25);
        this.rows++;
        this.refreshGridUI();
        this.markDirty();
        
        // Select the newly inserted row's first cell
        const nextCell = this.getCellEl(range ? range.startCol : 1, index);
        if (nextCell) {
            nextCell.focus();
            this.handleCellFocus(nextCell);
        }
    }

    deleteRow() {
        const range = this.getEffectiveRange();
        if (!range) return;
        const index = range.startRow;
        const count = range.endRow - range.startRow + 1;

        this.showConfirm(`${count}개 행을 삭제하시겠습니까?`, () => {
            this.shiftData('row', index + count, -count);
            this.rowHeights.splice(index, count);
            this.rows -= count;
            this.refreshGridUI();
            this.markDirty();

            const targetRow = Math.min(index, this.rows);
            const nextCell = this.getCellEl(range.startCol, targetRow);
            if (nextCell) {
                nextCell.focus();
                this.handleCellFocus(nextCell);
            }
        });
    }

    insertColumn() {
        const range = this.getEffectiveRange();
        const index = range ? range.startCol : 1;
        this.shiftData('col', index, 1);
        this.colWidths.splice(index - 1, 0, 100);
        this.cols++;
        this.refreshGridUI();
        this.markDirty();

        const nextCell = this.getCellEl(index, range ? range.startRow : 1);
        if (nextCell) {
            nextCell.focus();
            this.handleCellFocus(nextCell);
        }
    }

    deleteColumn() {
        const range = this.getEffectiveRange();
        if (!range) return;
        const index = range.startCol;
        const count = range.endCol - range.startCol + 1;

        this.showConfirm(`${count}개 열을 삭제하시겠습니까?`, () => {
            this.shiftData('col', index + count, -count);
            this.colWidths.splice(index - 1, count);
            this.cols -= count;
            this.refreshGridUI();
            this.markDirty();

            const targetCol = Math.min(index, this.cols);
            const nextCell = this.getCellEl(targetCol, range.startRow);
            if (nextCell) {
                nextCell.focus();
                this.handleCellFocus(nextCell);
            }
        });
    }

    showConfirm(message, onConfirm) {
        const modal = document.getElementById('confirm-modal');
        const msg = document.getElementById('confirm-message');
        const btnOk = document.getElementById('confirm-ok');
        const btnCancel = document.getElementById('confirm-cancel');
        if (!modal) { if (confirm(message)) onConfirm(); return; }
        msg.textContent = message;
        modal.style.display = 'flex';
        const close = () => { modal.style.display = 'none'; btnOk.onclick = null; btnCancel.onclick = null; };
        btnOk.onclick = () => { close(); onConfirm(); };
        btnCancel.onclick = () => close();
    }

    showConfirmAsync(message) {
        return new Promise((resolve) => {
            const modal = document.getElementById('confirm-modal');
            const msg = document.getElementById('confirm-message');
            const btnOk = document.getElementById('confirm-ok');
            const btnCancel = document.getElementById('confirm-cancel');

            if (!modal) {
                resolve(confirm(message));
                return;
            }

            msg.textContent = message;
            if (btnCancel) btnCancel.textContent = '취소';
            if (btnOk) btnOk.textContent = '확인';
            modal.style.display = 'flex';
            const close = (result) => {
                modal.style.display = 'none';
                btnOk.onclick = null;
                btnCancel.onclick = null;
                resolve(result);
            };
            btnOk.onclick = () => close(true);
            btnCancel.onclick = () => close(false);
        });
    }

    shiftData(type, threshold, delta) {
        const newData = {};
        const newStyles = {};
        const newFormats = {};
        const newFormulas = {};
        const newBorders = {};

        // Helper to shift a single coordinate and filter out deleted ones
        const shiftCoord = (coord, t, d) => {
            if (d < 0 && coord >= t + d && coord < t) return -1;
            return coord >= t ? coord + d : coord;
        };

        // Process data
        for (const key in this.data) {
            const parsed = this.parseCellId(key);
            if (!parsed) continue;
            const { col, row, colNum } = parsed;
            let nRow = row;
            let nColNum = colNum;

            if (type === 'row') {
                nRow = shiftCoord(row, threshold, delta);
            } else {
                nColNum = shiftCoord(colNum, threshold, delta);
            }

            if (nRow > 0 && nColNum > 0) {
                const newKey = `${this.numberToCol(nColNum)}${nRow}`;
                newData[newKey] = this.data[key];
            }
        }

        // Process styles
        for (const key in this.cellStyles) {
            const { col, row, colNum } = this.parseCellId(key);
            let nRow = row;
            let nColNum = colNum;

            if (type === 'row') {
                nRow = shiftCoord(row, threshold, delta);
            } else {
                nColNum = shiftCoord(colNum, threshold, delta);
            }

            if (nRow > 0 && nColNum > 0) {
                const newKey = `${this.numberToCol(nColNum)}${nRow}`;
                newStyles[newKey] = this.cellStyles[key];
            }
        }

        // Process formats
        for (const key in this.cellFormats) {
            const { row, colNum } = this.parseCellId(key);
            let nRow = row;
            let nColNum = colNum;

            if (type === 'row') {
                nRow = shiftCoord(row, threshold, delta);
            } else {
                nColNum = shiftCoord(colNum, threshold, delta);
            }

            if (nRow > 0 && nColNum > 0) {
                const newKey = `${this.numberToCol(nColNum)}${nRow}`;
                newFormats[newKey] = this.cellFormats[key];
            }
        }

        // Process formulas
        for (const key in this.cellFormulas) {
            const { row, colNum } = this.parseCellId(key);
            let nRow = row;
            let nColNum = colNum;

            if (type === 'row') {
                nRow = shiftCoord(row, threshold, delta);
            } else {
                nColNum = shiftCoord(colNum, threshold, delta);
            }

            if (nRow > 0 && nColNum > 0) {
                const newKey = `${this.numberToCol(nColNum)}${nRow}`;
                newFormulas[newKey] = this.cellFormulas[key];
            }
        }

        for (const key in this.cellBorders) {
            const { row, colNum } = this.parseCellId(key);
            let nRow = row;
            let nColNum = colNum;

            if (type === 'row') {
                nRow = shiftCoord(row, threshold, delta);
            } else {
                nColNum = shiftCoord(colNum, threshold, delta);
            }

            if (nRow > 0 && nColNum > 0) {
                const newKey = `${this.numberToCol(nColNum)}${nRow}`;
                newBorders[newKey] = this.cellBorders[key];
            }
        }

        this.data = newData;
        this.cellStyles = newStyles;
        this.cellFormats = newFormats;
        this.cellFormulas = newFormulas;
        this.cellBorders = newBorders;
        this.shiftMergedRanges(type, threshold, delta);
    }
}
