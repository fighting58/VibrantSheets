class VibrantSheets {
    constructor() {
        this.baseRows = 50;
        this.baseCols = 40; // A to AN
        this.baseColWidth = 64;
        this.baseRowHeight = 22;
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
        this.clipboardMerges = null; // Relative merged ranges for internal paste
        this.isCut = false;
        this.cutRange = null;

        // Resize state
        this.isResizingCol = false;
        this.isResizingRow = false;
        this.resizeIndex = -1;
        this.resizeStartPos = 0;
        this.resizeStartSize = 0;
        
        // Image insertion state
        this.imageLayer = null;
        this.activeImageId = null;
        this.isDraggingImage = false;
        this.isResizingImage = false;
        this.imageDragState = null;
        this.imageContextMenu = null;
        this.imageContextTargetId = null;
        this.rowHeaderWidth = 40;

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
        this.borderStamp = 0;
        this.showPageBreakPreview = false;

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
            images: [],
            printSettings: this.defaultPrintSettings(),
            colWidths: new Array(this.baseCols).fill(this.baseColWidth),
            rowHeights: new Array(this.baseRows + 1).fill(this.baseRowHeight)
        };
    }

    defaultPrintSettings() {
        return {
            range: 'current',
            printArea: null,
            paper: 'A4',
            orientation: 'portrait',
            margins: { top: 10, right: 10, bottom: 10, left: 10 },
            scale: 100,
            fitTo: { enabled: false, widthPages: 1, heightPages: 1 },
            headerFooter: {
                enabled: false,
                show: {
                    headerLeft: true,
                    headerCenter: true,
                    headerRight: true,
                    footerLeft: true,
                    footerCenter: true,
                    footerRight: true
                },
                header: { left: '', center: '', right: '' },
                footer: { left: '', center: '', right: '' }
            },
            repeat: { rows: '', cols: '' }
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
        this.pageBreakOverlay = this.createOverlay('page-break-overlay');
        this.ensurePrintPagesContainer();

        this.normalizeDefaultDimensions();
        this.renderGrid();
        this.renderSheetTabs();
        this.setupEventListeners();
        this.updatePrintOrientationIndicator();
        this.updatePrintAreaToggleUI();
        this.updatePrintControlsState();
    }

    createOverlay(className) {
        const div = document.createElement('div');
        div.className = className;
        div.style.display = 'none';
        this.container.appendChild(div);
        if (className === 'fill-preview') {
            div.innerHTML = '<span class="fill-preview-label">?쒕━利?梨꾩슦湲?/span><span class="fill-preview-hint">Alt: 媛?蹂듭궗</span>';
        }
        return div;
    }

    ensurePrintPagesContainer() {
        let container = document.getElementById('vs-print-pages');
        if (!container) {
            container = document.createElement('div');
            container.id = 'vs-print-pages';
            container.className = 'vs-print-pages';
            document.body.appendChild(container);
        }
        this.printPagesContainer = container;
        return container;
    }

    ensureImageLayer() {
        if (!this.container) return null;
        let layer = this.container.querySelector('.image-layer');
        if (!layer) {
            layer = document.createElement('div');
            layer.className = 'image-layer';
            this.container.appendChild(layer);
        }
        this.imageLayer = layer;
        return layer;
    }

    normalizeDefaultDimensions() {
        const isOldDefaultCols = (cols) => cols.every(w => w === 100 || w == null || w === 62);
        const isOldDefaultRows = (rows) => rows.slice(1).every(h => h === 25 || h == null || h === 23);
        this.sheets.forEach((sheet) => {
            if (sheet.colWidths && isOldDefaultCols(sheet.colWidths)) {
                sheet.colWidths = new Array(sheet.cols).fill(this.baseColWidth);
            }
            if (sheet.rowHeights && isOldDefaultRows(sheet.rowHeights)) {
                sheet.rowHeights = new Array(sheet.rows + 1).fill(this.baseRowHeight);
            }
        });
    }

    calcImageAnchorFromPixels(img, sheet = this.activeSheet) {
        if (!img || !sheet) return null;
        const colWidths = sheet.colWidths || [];
        const rowHeights = sheet.rowHeights || [];
        const headerRowHeight = rowHeights?.[1] || this.baseRowHeight;
        const start = this.calcCellFromOffset(img.x || 0, img.y || 0, colWidths, rowHeights, headerRowHeight);
        const end = this.calcCellFromOffset((img.x || 0) + (img.w || 0), (img.y || 0) + (img.h || 0), colWidths, rowHeights, headerRowHeight);
        if (!start || !end) return null;
        return {
            startCell: `${this.numberToCol(start.col)}${start.row}`,
            endCell: `${this.numberToCol(end.col)}${end.row}`,
            offsetStart: { x: start.offsetX, y: start.offsetY },
            offsetEnd: { x: end.offsetX, y: end.offsetY }
        };
    }

    calcCellFromOffset(x, y, colWidths, rowHeights, headerRowHeight) {
        const adjX = Math.max(0, x - this.rowHeaderWidth);
        const adjY = Math.max(0, y - headerRowHeight);

        let col = 1;
        let xRemain = adjX;
        for (let i = 0; i < colWidths.length; i++) {
            const w = colWidths[i] || this.baseColWidth;
            if (xRemain < w) { col = i + 1; break; }
            xRemain -= w;
            col = i + 2;
        }

        let row = 1;
        let yRemain = adjY;
        for (let i = 1; i < rowHeights.length; i++) {
            const h = rowHeights[i] || this.baseRowHeight;
            if (yRemain < h) { row = i; break; }
            yRemain -= h;
            row = i + 1;
        }

        return { col, row, offsetX: Math.max(0, Math.round(xRemain)), offsetY: Math.max(0, Math.round(yRemain)) };
    }

    updateImageAnchorsForSheet(sheet) {
        if (!sheet || !Array.isArray(sheet.images)) return;
        sheet.images.forEach((img) => {
            img.anchor = this.calcImageAnchorFromPixels(img, sheet);
        });
    }

    updateAllImageAnchors() {
        this.sheets.forEach((sheet) => this.updateImageAnchorsForSheet(sheet));
    }

    getImageById(id) {
        const images = this.activeSheet.images || [];
        return images.find(img => img.id === id);
    }

    selectImage(id) {
        this.activeImageId = id;
        this.renderImages();
    }

    clearImageSelection() {
        if (!this.activeImageId) return;
        this.activeImageId = null;
        this.renderImages();
    }

    // ??? Utility ???????????????????????????????????????????
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
        if (type === 'currency' || type === 'percentage' || type === 'number') return 2;
        return null;
    }

    normalizeDecimals(decimals) {
        if (decimals === null || decimals === undefined || decimals === '') return null;
        const n = Number(decimals);
        if (!Number.isFinite(n)) return null;
        return Math.max(0, Math.min(10, Math.round(n)));
    }

    normalizeFormat(format) {
        const type = ['general', 'number', 'currency', 'percentage', 'date', 'text'].includes(format?.type) ? format.type : 'general';
        let decimals = this.normalizeDecimals(format?.decimals);
        if (decimals === null) decimals = this.getDefaultDecimalsByType(type);
        if (type === 'date' || type === 'text') decimals = null;
        let currency = null;
        if (type === 'currency') {
            const inputCurrency = String(format?.currency || '').toUpperCase();
            currency = ['KRW', 'USD'].includes(inputCurrency) ? inputCurrency : 'KRW';
        }
        return { type, decimals, currency };
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
        
        // Text format returns raw string immediately
        if (format.type === 'text') {
            return raw;
        }

        const decimals = format.decimals ?? 0;

        if (format.type === 'date') {
            return this.formatDate(raw);
        }

        const numeric = this.parseNumberFromRaw(raw, format.type);
        if (numeric === null) return raw;

        if (format.type === 'currency') {
            const currency = format.currency || 'KRW';
            return new Intl.NumberFormat('ko-KR', {
                style: 'currency',
                currency,
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

        // Handle 'number' or 'general' with decimals
        if (format.type === 'number' || format.decimals !== null) {
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
        const formatted = this.getFormattedValue(cellId, rawValue);
        cell.classList.remove('cell-overflow');
        cell.style.removeProperty('--overflow-width');
        cell.innerHTML = '';
        const span = document.createElement('span');
        span.className = 'cell-text';
        span.textContent = formatted;
        cell.appendChild(span);
        this.applyOverflowForCell(cell, formatted);
    }

    applyOverflowForCell(cell, textOverride = null) {
        if (!cell || !cell.dataset?.id) return;
        if (cell.classList.contains('editing') || cell.classList.contains('merge-hidden')) return;
        const cellId = cell.dataset.id;
        const text = textOverride !== null ? textOverride : (cell.querySelector('.cell-text')?.textContent || '');
        if (!text) {
            cell.classList.remove('cell-overflow');
            cell.style.removeProperty('--overflow-width');
            cell.style.removeProperty('--overflow-left');
            cell.style.removeProperty('--text-offset');
            return;
        }

        const style = this.cellStyles[cellId] || {};
        const align = style.textAlign || 'left';

        const parsed = this.parseCellId(cellId);
        if (!parsed) return;
        const { colNum, row } = parsed;

        const merge = this.getMergedRangeAt(colNum, row);
        const startCol = merge ? merge.startCol : colNum;
        const baseEndCol = merge ? merge.endCol : colNum;

        let baseWidth = 0;
        for (let c = startCol; c <= baseEndCol; c++) {
            baseWidth += this.colWidths[c - 1] || this.baseColWidth;
        }

        let totalWidth = baseWidth;
        let offsetLeft = 0;
        let textOffset = 0;

        if (align === 'left') {
            let extraWidth = 0;
            for (let c = baseEndCol + 1; c <= this.cols; c++) {
                if (this.isOverflowBlocked(c, row)) break;
                extraWidth += this.colWidths[c - 1] || this.baseColWidth;
            }
            totalWidth += extraWidth;
            offsetLeft = 0;
        } else if (align === 'right') {
            let extraWidth = 0;
            for (let c = startCol - 1; c >= 1; c--) {
                if (this.isOverflowBlocked(c, row)) break;
                extraWidth += this.colWidths[c - 1] || this.baseColWidth;
            }
            totalWidth += extraWidth;
            offsetLeft = -extraWidth;
        }

        if (totalWidth > baseWidth) {
            cell.classList.add('cell-overflow');
            cell.style.setProperty('--overflow-width', `${totalWidth}px`);
            cell.style.setProperty('--overflow-left', `${offsetLeft}px`);
            if (textOffset !== 0) {
                cell.style.setProperty('--text-offset', `${textOffset}px`);
            } else {
                cell.style.removeProperty('--text-offset');
            }
        } else {
            cell.classList.remove('cell-overflow');
            cell.style.removeProperty('--overflow-width');
            cell.style.removeProperty('--overflow-left');
            cell.style.removeProperty('--text-offset');
        }
    }

    isOverflowBlocked(colNum, row) {
        const id = `${this.numberToCol(colNum)}${row}`;
        const raw = this.getRawValue(id);
        if (raw && raw.trim() !== '') return true;

        const style = this.cellStyles?.[id];
        if (style && style.backgroundColor && style.backgroundColor !== 'transparent') return true;

        if (this.cellBorders?.[id]) return true;

        if (this.getMergedRangeAt(colNum, row)) return true;

        return false;
    }

    refreshOverflowForRow(row) {
        if (!this.tbody) return;
        for (let c = 1; c <= this.cols; c++) {
            const cell = this.getCellEl(c, row);
            if (cell && !cell.classList.contains('header')) {
                this.applyOverflowForCell(cell);
            }
        }
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

    // ??? Grid Rendering ????????????????????????????????????
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
            th.innerText = this.numberToCol(j + 1);
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
        this.ensureImageLayer();
        this.renderImages();
    }

    createRowElements(startRow, endRow) {
        for (let i = startRow; i <= endRow; i++) {
            if (!this.rowHeights[i]) this.rowHeights[i] = this.baseRowHeight;
            const tr = document.createElement('tr');
            tr.style.height = `${this.rowHeights[i]}px`;
            tr.dataset.rowIndex = String(i);
            
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

    renderImages() {
        const layer = this.ensureImageLayer();
        if (!layer) return;
        const images = this.activeSheet.images || [];
        const anchorCell = this.getCellEl(1, 1);
        const actualTop = anchorCell?.offsetTop || 0;
        const actualLeft = anchorCell?.offsetLeft || 0;
        const expectedTop = (this.activeSheet?.rowHeights?.[1] || this.baseRowHeight);
        const expectedLeft = this.rowHeaderWidth;
        layer.innerHTML = '';
        if (this.table) {
            layer.style.width = `${this.table.offsetWidth}px`;
            layer.style.height = `${this.table.offsetHeight}px`;
        }
        images.forEach((img) => {
            if (img && img._importedOffsetPending && (actualTop || actualLeft)) {
                const baseY = Math.max(0, Number(img.y) || 0);
                const baseX = Math.max(0, Number(img.x) || 0);
                const deltaY = (actualTop - expectedTop) - 1;
                const deltaX = (actualLeft - expectedLeft);
                img.y = Math.max(0, baseY + deltaY);
                img.x = Math.max(0, baseX + deltaX);
                img._importedOffsetPending = false;
                if (typeof this.calcImageAnchorFromPixels === 'function') {
                    img.anchor = this.calcImageAnchorFromPixels(img, this.activeSheet);
                }
            }
            const wrapper = document.createElement('div');
            wrapper.className = 'sheet-image' + (img.id === this.activeImageId ? ' selected' : '');
            if (img.locked) wrapper.classList.add('locked');
            wrapper.dataset.imageId = img.id;
            wrapper.style.left = `${img.x}px`;
            wrapper.style.top = `${img.y}px`;
            wrapper.style.width = `${img.w}px`;
            wrapper.style.height = `${img.h}px`;
            wrapper.style.zIndex = String(img.z || 1);
            wrapper.addEventListener('mousedown', (e) => this.handleImageMouseDown(e, img.id));
            wrapper.addEventListener('contextmenu', (e) => this.showImageContextMenu(e, img.id));

            const imageEl = document.createElement('img');
            imageEl.src = img.src;
            imageEl.alt = img.name || 'image';
            wrapper.appendChild(imageEl);

            if (!img.locked) {
                ['nw', 'ne', 'sw', 'se'].forEach((handle) => {
                    const h = document.createElement('div');
                    h.className = 'image-handle';
                    h.dataset.handle = handle;
                    wrapper.appendChild(h);
                });
            }

            layer.appendChild(wrapper);
        });
    }

    bindImageContextMenu() {
        this.imageContextMenu = document.getElementById('image-context-menu');
        if (!this.imageContextMenu) return;
        this.imageContextMenu.addEventListener('click', (e) => {
            const btn = e.target.closest('button[data-action]');
            if (!btn) return;
            const action = btn.dataset.action;
            const id = this.imageContextTargetId;
            if (!id) return;
            if (action === 'bring-front') this.bringImageToFront(id);
            if (action === 'send-back') this.sendImageToBack(id);
            if (action === 'delete') this.deleteActiveImage();
            if (action === 'original-size') this.resetImageOriginalSize(id);
            if (action === 'toggle-lock') this.toggleImageLock(id);
            this.hideImageContextMenu();
        });
        document.addEventListener('click', () => this.hideImageContextMenu());
        window.addEventListener('scroll', () => this.hideImageContextMenu());
    }

    showImageContextMenu(e, id) {
        e.preventDefault();
        e.stopPropagation();
        this.selectImage(id);
        this.imageContextTargetId = id;
        if (!this.imageContextMenu) return;
        const img = this.getImageById(id);
        const lockBtn = this.imageContextMenu.querySelector('[data-action="toggle-lock"]');
        if (lockBtn) lockBtn.textContent = img && img.locked ? 'Unlock' : 'Lock';
        this.imageContextMenu.style.display = 'block';
        const pad = 8;
        const x = Math.min(window.innerWidth - this.imageContextMenu.offsetWidth - pad, e.clientX);
        const y = Math.min(window.innerHeight - this.imageContextMenu.offsetHeight - pad, e.clientY);
        this.imageContextMenu.style.left = `${Math.max(pad, x)}px`;
        this.imageContextMenu.style.top = `${Math.max(pad, y)}px`;
    }

    hideImageContextMenu() {
        if (!this.imageContextMenu) return;
        this.imageContextMenu.style.display = 'none';
        this.imageContextTargetId = null;
    }

    normalizeImageZ() {
        const images = this.activeSheet.images || [];
        images.sort((a, b) => (a.z || 0) - (b.z || 0));
        images.forEach((img, idx) => {
            img.z = idx + 1;
        });
    }

    bringImageToFront(id) {
        const img = this.getImageById(id);
        if (!img) return;
        const images = this.activeSheet.images || [];
        const maxZ = images.reduce((acc, cur) => Math.max(acc, cur.z || 0), 0);
        img.z = maxZ + 1;
        this.normalizeImageZ();
        this.renderImages();
        this.markDirty();
    }

    sendImageToBack(id) {
        const img = this.getImageById(id);
        if (!img) return;
        const images = this.activeSheet.images || [];
        const minZ = images.reduce((acc, cur) => Math.min(acc, cur.z || 1), img.z || 1);
        img.z = minZ - 1;
        this.normalizeImageZ();
        this.renderImages();
        this.markDirty();
    }

    resetImageOriginalSize(id) {
        const img = this.getImageById(id);
        if (!img || !img.src) return;
        const probe = new Image();
        probe.onload = () => {
            img.w = Math.max(40, probe.naturalWidth || img.w);
            img.h = Math.max(40, probe.naturalHeight || img.h);
            img.anchor = this.calcImageAnchorFromPixels(img, this.activeSheet);
            this.renderImages();
            this.markDirty();
        };
        probe.src = img.src;
    }

    toggleImageLock(id) {
        const img = this.getImageById(id);
        if (!img) return;
        img.locked = !img.locked;
        this.renderImages();
        this.markDirty();
    }

    openImageDialog() {
        if (!this.imageInput) return;
        this.imageInput.value = '';
        this.imageInput.click();
    }

    insertImageFromSrc(src, name = '') {
        if (!src) return;
        const img = new Image();
        img.onload = () => {
            const maxDim = 240;
            let width = img.naturalWidth || maxDim;
            let height = img.naturalHeight || maxDim;
            if (width > maxDim) {
                const scale = maxDim / width;
                width = Math.round(width * scale);
                height = Math.round(height * scale);
            }
            if (height > maxDim) {
                const scale = maxDim / height;
                width = Math.round(width * scale);
                height = Math.round(height * scale);
            }
            width = Math.max(40, width);
            height = Math.max(40, height);

            const cell = this.selectedCell || this.getCellEl(1, 1);
            const x = cell ? cell.offsetLeft : 0;
            const y = cell ? cell.offsetTop : 0;
            const id = `img_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
            const images = this.activeSheet.images || [];
            const maxZ = images.reduce((acc, cur) => Math.max(acc, cur.z || 0), 0);
            const payload = { id, src, x, y, w: width, h: height, name, z: maxZ + 1 };
            payload.anchor = this.calcImageAnchorFromPixels(payload, this.activeSheet);
            if (!this.activeSheet.images) this.activeSheet.images = [];
            this.activeSheet.images.push(payload);
            this.selectImage(id);
            this.renderImages();
            this.markDirty();
        };
        img.src = src;
    }

    async handleImageFileSelect(e) {
        const file = e.target.files && e.target.files[0];
        if (!file || !file.type || !file.type.startsWith('image/')) return;

        const reader = new FileReader();
        reader.onload = () => {
            const src = String(reader.result || '');
            this.insertImageFromSrc(src, file.name || '');
        };
        reader.readAsDataURL(file);
    }

    handleImagePaste(e) {
        const items = e.clipboardData && e.clipboardData.items ? Array.from(e.clipboardData.items) : [];
        const imageItem = items.find(item => item.type && item.type.startsWith('image/'));
        if (!imageItem) return false;
        const file = imageItem.getAsFile();
        if (!file) return false;
        const reader = new FileReader();
        reader.onload = () => {
            const src = String(reader.result || '');
            this.insertImageFromSrc(src, file.name || 'clipboard.png');
        };
        reader.readAsDataURL(file);
        return true;
    }

    handleImageMouseDown(e, id) {
        e.preventDefault();
        e.stopPropagation();
        this.selectImage(id);
        const handle = e.target.closest('.image-handle');
        const img = this.getImageById(id);
        if (!img) return;
        if (img.locked) return;

        this.imageDragState = {
            id,
            startX: e.clientX,
            startY: e.clientY,
            startLeft: img.x,
            startTop: img.y,
            startW: img.w,
            startH: img.h,
            handle: handle ? handle.dataset.handle : null
        };

        if (handle) {
            this.isResizingImage = true;
        } else {
            this.isDraggingImage = true;
        }
    }

    handleImageDragMove(e) {
        if (!this.imageDragState) return;
        const img = this.getImageById(this.imageDragState.id);
        if (!img) return;

        const dx = e.clientX - this.imageDragState.startX;
        const dy = e.clientY - this.imageDragState.startY;
        const minSize = 20;

        if (this.isDraggingImage) {
            img.x = Math.max(0, this.imageDragState.startLeft + dx);
            img.y = Math.max(0, this.imageDragState.startTop + dy);
        } else if (this.isResizingImage) {
            let left = this.imageDragState.startLeft;
            let top = this.imageDragState.startTop;
            let width = this.imageDragState.startW;
            let height = this.imageDragState.startH;
            const handle = this.imageDragState.handle || 'se';

            if (handle.includes('e')) {
                width = Math.max(minSize, this.imageDragState.startW + dx);
            }
            if (handle.includes('s')) {
                height = Math.max(minSize, this.imageDragState.startH + dy);
            }
            if (handle.includes('w')) {
                width = Math.max(minSize, this.imageDragState.startW - dx);
                left = this.imageDragState.startLeft + (this.imageDragState.startW - width);
            }
            if (handle.includes('n')) {
                height = Math.max(minSize, this.imageDragState.startH - dy);
                top = this.imageDragState.startTop + (this.imageDragState.startH - height);
            }

            img.x = Math.max(0, left);
            img.y = Math.max(0, top);
            img.w = width;
            img.h = height;
        }

        const el = this.imageLayer?.querySelector(`[data-image-id="${img.id}"]`);
        if (el) {
            el.style.left = `${img.x}px`;
            el.style.top = `${img.y}px`;
            el.style.width = `${img.w}px`;
            el.style.height = `${img.h}px`;
        }
    }

    handleImageDragEnd() {
        if (this.isDraggingImage || this.isResizingImage) {
            this.isDraggingImage = false;
            this.isResizingImage = false;
            this.imageDragState = null;
            if (this.activeImageId) {
                const img = this.getImageById(this.activeImageId);
                if (img) img.anchor = this.calcImageAnchorFromPixels(img, this.activeSheet);
            }
            this.markDirty();
        }
    }

    deleteActiveImage() {
        if (!this.activeImageId) return;
        const images = this.activeSheet.images || [];
        const idx = images.findIndex(img => img.id === this.activeImageId);
        if (idx >= 0) {
            images.splice(idx, 1);
            this.activeImageId = null;
            this.renderImages();
            this.markDirty();
        }
    }

    // ??? Event Listeners ???????????????????????????????????
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
        this.bindPrintSettingsModal();
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
                        previewHtml = '<div style="font-size: 0.75rem; color:#94a3b8; width: 100%; text-align: center;">None</div>';
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

        const printBtn = document.getElementById('btn-print-settings');
        if (printBtn) {
            printBtn.addEventListener('click', () => this.openPrintSettingsModal());
        }
        const orientationBtn = document.getElementById('btn-print-orientation');
        if (orientationBtn) {
            orientationBtn.addEventListener('click', () => this.togglePrintOrientationFromRibbon());
        }
        const printAreaToggleBtn = document.getElementById('btn-print-area-toggle');
        if (printAreaToggleBtn) {
            printAreaToggleBtn.addEventListener('click', () => this.togglePrintAreaFromRibbon());
        }
        const pageBreakBtn = document.getElementById('btn-page-break-preview');
        if (pageBreakBtn) {
            pageBreakBtn.addEventListener('click', () => this.togglePageBreakPreview());
        }

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

        const imageBtn = document.getElementById('btn-insert-image');
        if (imageBtn) imageBtn.addEventListener('click', () => this.openImageDialog());
        this.imageInput = document.getElementById('image-file-input');
        if (this.imageInput) this.imageInput.addEventListener('change', (e) => this.handleImageFileSelect(e));
        if (this.container) {
            this.container.addEventListener('paste', (e) => {
                if (this.handleImagePaste(e)) {
                    e.preventDefault();
                }
            });
        }
        this.bindImageContextMenu();

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
            if (this.isDraggingImage || this.isResizingImage) {
                this.handleImageDragMove(e);
                return;
            }
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
            if (this.isDraggingImage || this.isResizingImage) {
                this.handleImageDragEnd(e);
                return;
            }
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
            this.updatePageBreakPreview();
        });
        this.container.addEventListener('scroll', () => {
             this.updateSelectionOverlay();
             this.updateRangeVisual();
             this.updateFillHandlePosition();
             this.updatePageBreakPreview();
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
                if (this.activeImageId) {
                    e.preventDefault();
                    this.deleteActiveImage();
                    return;
                }
                if (this.selectionRange) {
                    e.preventDefault();
                    this.deleteSelection();
                }
            }
        });

        // Resize Handlers
        this.setupResizeHandlers();

        // Print visibility/layout handling
        window.addEventListener('beforeprint', () => this.handleBeforePrint());
        window.addEventListener('afterprint', () => this.handleAfterPrint());
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
        this.updatePrintOrientationIndicator();
        this.updatePrintAreaToggleUI();
        this.updatePrintControlsState();
        this.updatePageBreakPreview();
    }

    // ??? Resize Handlers ?????????????????????????????????
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

    // ??? Cell Focus / Input ????????????????????????????????
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
            // This prevents the IME first-character loss bug during manual clearing.
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
        const headerHeight = this.baseRowHeight; // Standard row height for col headers
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

    // ??? Styling ???????????????????????????????????????????
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
        const rowsToRefresh = new Set();

        targetIds.forEach(id => {
            if (!this.cellStyles[id]) this.cellStyles[id] = {};
            this.cellStyles[id][prop] = value;

            const el = document.querySelector(`[data-id="${id}"]`);
            if (el) {
                el.style[prop] = value;
            }
            const parsed = this.parseCellId(id);
            if (parsed) rowsToRefresh.add(parsed.row);
        });

        this.markDirty();
        if (this.selectedCell) this.updateToolbarState(this.selectedCell);
        if (prop === 'textAlign' || prop === 'backgroundColor') {
            rowsToRefresh.forEach((row) => this.refreshOverflowForRow(row));
        }
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
        const getTs = (b) => (b && typeof b._ts === 'number' ? b._ts : 0);
        const pickLatest = (...borders) => {
            let best = null;
            let bestTs = -1;
            borders.forEach((b) => {
                if (!b) return;
                const ts = getTs(b);
                if (ts >= bestTs) {
                    best = b;
                    bestTs = ts;
                }
            });
            return best;
        };
        
        const merge = cell.classList.contains('merge-anchor') ? this.getMergedRangeAt(colNum, rowNum) : null;
        const range = merge ? merge : { startCol: colNum, endCol: colNum, startRow: rowNum, endRow: rowNum };

        const getAnyNeighborBorder = (side) => {
            let best = null;
            if (!merge) return null;
            if (side === 'top' && range.startRow > 1) {
                const r = range.startRow - 1;
                for (let c = range.startCol; c <= range.endCol; c++) {
                    const b = getBorder(c, r, 'bottom');
                    best = pickLatest(best, b);
                }
            }
            if (side === 'bottom' && range.endRow < this.rows) {
                const r = range.endRow + 1;
                for (let c = range.startCol; c <= range.endCol; c++) {
                    const b = getBorder(c, r, 'top');
                    best = pickLatest(best, b);
                }
            }
            if (side === 'left' && range.startCol > 1) {
                const c = range.startCol - 1;
                for (let r = range.startRow; r <= range.endRow; r++) {
                    const b = getBorder(c, r, 'right');
                    best = pickLatest(best, b);
                }
            }
            if (side === 'right' && range.endCol < this.cols) {
                const c = range.endCol + 1;
                for (let r = range.startRow; r <= range.endRow; r++) {
                    const b = getBorder(c, r, 'left');
                    best = pickLatest(best, b);
                }
            }
            return best;
        };

        const neighborTop = getAnyNeighborBorder('top');
        const neighborRight = getAnyNeighborBorder('right');
        const neighborBottom = getAnyNeighborBorder('bottom');
        const neighborLeft = getAnyNeighborBorder('left');

        const topBorder = pickLatest(
            getBorder(colNum, rowNum, 'top'),
            rowNum > 1 ? getBorder(colNum, rowNum - 1, 'bottom') : null,
            neighborTop
        );
        const rightBorder = pickLatest(
            getBorder(colNum, rowNum, 'right'),
            colNum < this.cols ? getBorder(colNum + 1, rowNum, 'left') : null,
            neighborRight
        );
        const bottomBorder = pickLatest(
            getBorder(colNum, rowNum, 'bottom'),
            rowNum < this.rows ? getBorder(colNum, rowNum + 1, 'top') : null,
            neighborBottom
        );
        const leftBorder = pickLatest(
            getBorder(colNum, rowNum, 'left'),
            colNum > 1 ? getBorder(colNum - 1, rowNum, 'right') : null,
            neighborLeft
        );

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

        // For merged cells, thin borders can be lost due to border-collapse. Reinforce with inset shadow.
        if (merge) {
            const shadows = [];
            const addShadow = (side, b) => {
                if (!b || b.style !== 'solid' || (b.width || 1) !== 1) return;
                const color = b.color || '#000000';
                if (side === 'top') shadows.push(`inset 0 1px 0 0 ${color}`);
                if (side === 'bottom') shadows.push(`inset 0 -1px 0 0 ${color}`);
                if (side === 'left') shadows.push(`inset 1px 0 0 0 ${color}`);
                if (side === 'right') shadows.push(`inset -1px 0 0 0 ${color}`);
            };
            addShadow('top', topBorder);
            addShadow('right', rightBorder);
            addShadow('bottom', bottomBorder);
            addShadow('left', leftBorder);
            if (shadows.length > 0) {
                cell.style.boxShadow = shadows.join(', ');
            } else if (cell.style.boxShadow) {
                cell.style.boxShadow = '';
            }
        } else if (cell.style.boxShadow) {
            cell.style.boxShadow = '';
        }
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

    // ??? Range Selection ???????????????????????????????????
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
        if (this.activeImageId) this.clearImageSelection();
        if (cell.classList.contains('cell-overflow')) {
            const rect = cell.getBoundingClientRect();
            const outside = e.clientX < rect.left || e.clientX > rect.right || e.clientY < rect.top || e.clientY > rect.bottom;
            if (outside) {
                const prev = cell.style.pointerEvents;
                cell.style.pointerEvents = 'none';
                const target = document.elementFromPoint(e.clientX, e.clientY);
                cell.style.pointerEvents = prev;
                const targetCell = target?.closest?.('.cell');
                if (targetCell && targetCell !== cell) {
                    this.handleCellMouseDown(targetCell, e);
                    return;
                }
            }
        }

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

        // Enter Edit mode only on double click (single click should allow drag selection)
        if (this.selectedCell === cell && !this.isEditing && !e.shiftKey && e.detail >= 2) {
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
        this.updatePageBreakPreview();
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

    mergeRangeSilently(range) {
        if (!range) return;
        const normalized = this.expandRangeToIncludeMerges(range);
        const isSingleCell = normalized.startCol === normalized.endCol && normalized.startRow === normalized.endRow;
        if (isSingleCell) return;

        const existing = this.getNormalizedMergedRanges();
        const keep = existing.filter((merge) => !this.rangesIntersect(merge, normalized));
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
            width: width,
            _ts: ++this.borderStamp
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
        const applyMirrorThin = (side, r, c, value) => {
            if (!value || value.style !== 'solid' || (value.width || 1) !== 1) return;
            let mr = r;
            let mc = c;
            let mirrorSide = null;
            if (side === 'top') { mr = r - 1; mirrorSide = 'bottom'; }
            if (side === 'bottom') { mr = r + 1; mirrorSide = 'top'; }
            if (side === 'left') { mc = c - 1; mirrorSide = 'right'; }
            if (side === 'right') { mc = c + 1; mirrorSide = 'left'; }
            if (mr < 1 || mc < 1 || mr > this.rows || mc > this.cols) return;
            const mirrorId = `${this.numberToCol(mc)}${mr}`;
            applySide(mirrorId, mirrorSide, value);
        };
        const applyToAdjacentMerged = (side, r, c, value) => {
            if (!value) return;
            let nr = r;
            let nc = c;
            let targetSide = null;
            if (side === 'top') { nr = r - 1; targetSide = 'bottom'; }
            if (side === 'bottom') { nr = r + 1; targetSide = 'top'; }
            if (side === 'left') { nc = c - 1; targetSide = 'right'; }
            if (side === 'right') { nc = c + 1; targetSide = 'left'; }
            if (nr < 1 || nc < 1 || nr > this.rows || nc > this.cols) return;
            const merge = this.getMergedRangeAt(nc, nr);
            if (!merge) return;
            // Ensure this border is on the merged range boundary
            if (targetSide === 'top' && nr !== merge.startRow) return;
            if (targetSide === 'bottom' && nr !== merge.endRow) return;
            if (targetSide === 'left' && nc !== merge.startCol) return;
            if (targetSide === 'right' && nc !== merge.endCol) return;
            const anchorId = `${this.numberToCol(merge.startCol)}${merge.startRow}`;
            applySide(anchorId, targetSide, value);
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
            // Apply only the boundary edge for single-side borders.
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
                        const value = (type === 'none' || style === 'none') ? null : border;
                        applySide(targetId, side, value);
                        if (value) {
                            applyMirrorThin(side, r, c, value);
                            applyToAdjacentMerged(side, r, c, value);
                        }
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

    // ??? Fill Handle Logic ?????????????????????????????????
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
        return await this.showConfirmAsyncNextTick('연속된 값이 감지되었습니다. 시리즈 채우기를 적용할까요?\n취소를 누르면 값 복사로 채웁니다.');
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

    getPrintSettings() {
        if (!this.activeSheet.printSettings) {
            this.activeSheet.printSettings = this.defaultPrintSettings();
        }
        return this.activeSheet.printSettings;
    }

    bindPrintSettingsModal() {
        const modal = document.getElementById('print-settings-modal');
        const btnClose = document.getElementById('print-close');
        const btnCancel = document.getElementById('print-cancel');
        const btnApply = document.getElementById('print-apply');
        const btnPrint = document.getElementById('print-run');
        if (!modal || !btnClose || !btnCancel || !btnApply || !btnPrint) return;

        const close = (restore = false) => {
            if (restore && this.printSettingsSnapshot) {
                this.activeSheet.printSettings = this.printSettingsSnapshot;
                this.printSettingsSnapshot = null;
                this.updatePrintOrientationIndicator();
                this.updatePrintRunButtonState();
                this.renderPrintPreview();
                this.updatePageBreakPreview();
            }
            modal.style.display = 'none';
        };
        btnClose.addEventListener('click', () => close(true));
        btnCancel.addEventListener('click', () => close(true));

        btnApply.addEventListener('click', () => {
            this.applyPrintSettingsFromUI();
            this.printSettingsSnapshot = null;
        });
        btnPrint.addEventListener('click', () => {
            this.applyPrintSettingsFromUI();
            const settings = this.getPrintSettings();
            if (!settings.printArea) {
                alert('인쇄영역이 설정되지 않았습니다. 리본의 Print Area 버튼으로 인쇄영역을 먼저 설정해 주세요.');
                this.updateStatusBadge('Error: Print area is required');
                return;
            }
            close(false);
            window.print();
        });

        if (!this.printModalLiveBound) {
            modal.addEventListener('input', () => this.applyPrintSettingsFromUI());
            modal.addEventListener('change', () => this.applyPrintSettingsFromUI());
            this.printModalLiveBound = true;
        }
    }

    openPrintSettingsModal() {
        const modal = document.getElementById('print-settings-modal');
        if (!modal) return;
        const settings = this.getPrintSettings();
        this.printSettingsSnapshot = JSON.parse(JSON.stringify(settings));
        this.updatePrintOrientationIndicator();
        this.updatePrintRunButtonState();

        const setRadio = (name, value) => {
            const el = document.querySelector(`input[name="${name}"][value="${value}"]`);
            if (el) el.checked = true;
        };

        setRadio('print-range', settings.range);
        setRadio('print-orientation', settings.orientation);
        setRadio('print-scale-mode', settings.fitTo.enabled ? 'fit' : 'scale');

        const paper = document.getElementById('print-paper');
        if (paper) paper.value = settings.paper;

        const marginMap = [
            ['print-margin-top', settings.margins.top],
            ['print-margin-right', settings.margins.right],
            ['print-margin-bottom', settings.margins.bottom],
            ['print-margin-left', settings.margins.left]
        ];
        marginMap.forEach(([id, val]) => {
            const el = document.getElementById(id);
            if (el) el.value = val;
        });

        const scale = document.getElementById('print-scale');
        if (scale) scale.value = settings.scale;
        const fitW = document.getElementById('print-fit-width');
        const fitH = document.getElementById('print-fit-height');
        if (fitW) fitW.value = settings.fitTo.widthPages;
        if (fitH) fitH.value = settings.fitTo.heightPages;

        const hfEnabled = document.getElementById('print-hf-enabled');
        if (hfEnabled) hfEnabled.checked = settings.headerFooter.enabled;
        const show = settings.headerFooter.show || {};
        const setChecked = (id, fallback = true) => {
            const el = document.getElementById(id);
            if (el) el.checked = show[id.replace('print-show-', '').replace(/-([a-z])/g, (_, c) => c.toUpperCase())] ?? fallback;
        };
        setChecked('print-show-header-left', true);
        setChecked('print-show-header-center', true);
        setChecked('print-show-header-right', true);
        setChecked('print-show-footer-left', true);
        setChecked('print-show-footer-center', true);
        setChecked('print-show-footer-right', true);
        const hfMap = [
            ['print-header-left', settings.headerFooter.header.left],
            ['print-header-center', settings.headerFooter.header.center],
            ['print-header-right', settings.headerFooter.header.right],
            ['print-footer-left', settings.headerFooter.footer.left],
            ['print-footer-center', settings.headerFooter.footer.center],
            ['print-footer-right', settings.headerFooter.footer.right]
        ];
        hfMap.forEach(([id, val]) => {
            const el = document.getElementById(id);
            if (el) el.value = val;
        });

        const repeatRows = document.getElementById('print-repeat-rows');
        const repeatCols = document.getElementById('print-repeat-cols');
        if (repeatRows) repeatRows.value = settings.repeat.rows;
        if (repeatCols) repeatCols.value = settings.repeat.cols;

        modal.style.display = 'flex';
        // Defer preview render so layout sizes are available after display
        requestAnimationFrame(() => this.renderPrintPreview());
    }

    applyPrintSettingsFromUI() {
        const settings = this.getPrintSettings();
        const getRadio = (name, fallback) => {
            const el = document.querySelector(`input[name="${name}"]:checked`);
            return el ? el.value : fallback;
        };
        settings.range = getRadio('print-range', settings.range);
        if (settings.range === 'selection' && this.selectionRange) {
            settings.selectionRange = { ...this.selectionRange };
        }
        settings.orientation = getRadio('print-orientation', settings.orientation);
        const scaleMode = getRadio('print-scale-mode', settings.fitTo.enabled ? 'fit' : 'scale');
        settings.fitTo.enabled = scaleMode === 'fit';

        const paper = document.getElementById('print-paper');
        if (paper) settings.paper = paper.value;

        const marginVal = (id, fallback) => {
            const el = document.getElementById(id);
            const v = el ? Number(el.value) : fallback;
            return Number.isFinite(v) ? Math.max(0, Math.min(50, v)) : fallback;
        };
        settings.margins.top = marginVal('print-margin-top', settings.margins.top);
        settings.margins.right = marginVal('print-margin-right', settings.margins.right);
        settings.margins.bottom = marginVal('print-margin-bottom', settings.margins.bottom);
        settings.margins.left = marginVal('print-margin-left', settings.margins.left);

        const scale = document.getElementById('print-scale');
        if (scale) {
            const v = Number(scale.value);
            if (Number.isFinite(v)) settings.scale = Math.max(10, Math.min(200, v));
        }
        const fitW = document.getElementById('print-fit-width');
        const fitH = document.getElementById('print-fit-height');
        if (fitW) settings.fitTo.widthPages = Math.max(1, Math.min(99, Number(fitW.value) || settings.fitTo.widthPages));
        if (fitH) settings.fitTo.heightPages = Math.max(1, Math.min(99, Number(fitH.value) || settings.fitTo.heightPages));

        const hfEnabled = document.getElementById('print-hf-enabled');
        settings.headerFooter.enabled = hfEnabled ? hfEnabled.checked : settings.headerFooter.enabled;
        const setText = (id, fallback) => {
            const el = document.getElementById(id);
            return el ? String(el.value || '') : fallback;
        };
        settings.headerFooter.header.left = setText('print-header-left', settings.headerFooter.header.left);
        settings.headerFooter.header.center = setText('print-header-center', settings.headerFooter.header.center);
        settings.headerFooter.header.right = setText('print-header-right', settings.headerFooter.header.right);
        settings.headerFooter.footer.left = setText('print-footer-left', settings.headerFooter.footer.left);
        settings.headerFooter.footer.center = setText('print-footer-center', settings.headerFooter.footer.center);
        settings.headerFooter.footer.right = setText('print-footer-right', settings.headerFooter.footer.right);
        if (!settings.headerFooter.show) {
            settings.headerFooter.show = {
                headerLeft: true,
                headerCenter: true,
                headerRight: true,
                footerLeft: true,
                footerCenter: true,
                footerRight: true
            };
        }
        const getChecked = (id, fallback = true) => {
            const el = document.getElementById(id);
            return el ? !!el.checked : fallback;
        };
        settings.headerFooter.show.headerLeft = getChecked('print-show-header-left', true);
        settings.headerFooter.show.headerCenter = getChecked('print-show-header-center', true);
        settings.headerFooter.show.headerRight = getChecked('print-show-header-right', true);
        settings.headerFooter.show.footerLeft = getChecked('print-show-footer-left', true);
        settings.headerFooter.show.footerCenter = getChecked('print-show-footer-center', true);
        settings.headerFooter.show.footerRight = getChecked('print-show-footer-right', true);

        const repeatRows = document.getElementById('print-repeat-rows');
        const repeatCols = document.getElementById('print-repeat-cols');
        settings.repeat.rows = repeatRows ? String(repeatRows.value || '') : settings.repeat.rows;
        settings.repeat.cols = repeatCols ? String(repeatCols.value || '') : settings.repeat.cols;
        this.updatePrintOrientationIndicator();
        this.updatePrintRunButtonState();
        this.renderPrintPreview();
        this.updatePageBreakPreview();
    }

    updatePrintRunButtonState() {
        const btnPrint = document.getElementById('print-run');
        if (!btnPrint) return;
        const hasPrintArea = !!this.getPrintSettings().printArea;
        btnPrint.disabled = !hasPrintArea;
        btnPrint.title = hasPrintArea
            ? 'Print'
            : 'Set Print Area first';
    }

    updatePrintControlsState() {
        const hasPrintArea = !!this.getPrintSettings().printArea;
        const settingsBtn = document.getElementById('btn-print-settings');
        if (settingsBtn) {
            settingsBtn.disabled = !hasPrintArea;
            settingsBtn.title = hasPrintArea
                ? 'Print Settings'
                : 'Set Print Area first';
        }
        this.updatePrintRunButtonState();
    }

    togglePrintOrientationFromRibbon() {
        const settings = this.getPrintSettings();
        settings.orientation = settings.orientation === 'landscape' ? 'portrait' : 'landscape';
        this.markDirty();
        this.updatePrintOrientationIndicator();
        const selected = document.querySelector(`input[name="print-orientation"][value="${settings.orientation}"]`);
        if (selected) selected.checked = true;
        const modal = document.getElementById('print-settings-modal');
        if (modal && modal.style.display !== 'none') this.renderPrintPreview();
        this.updatePageBreakPreview();
    }

    updatePrintOrientationIndicator() {
        const chip = document.getElementById('print-orientation-indicator');
        const btn = document.getElementById('btn-print-orientation');
        const settings = this.getPrintSettings();
        const label = settings.orientation === 'landscape' ? 'Landscape' : 'Portrait';
        if (chip) chip.innerText = label;
        if (btn) {
            btn.title = `Orientation: ${label}`;
            btn.classList.toggle('active', settings.orientation === 'landscape');
        }
    }

    handleBeforePrint() {
        this.applyDynamicPageRule(this.getPrintSettings());
        this.applyPrintVisibility(true);
        this.buildPrintPages();
        document.body.classList.add('custom-print-active');
    }

    handleAfterPrint() {
        document.body.classList.remove('custom-print-active');
        this.clearPrintPages();
        this.applyPrintVisibility(false);
        this.applyDynamicPageRule(null);
    }

    clearPrintPages() {
        const container = this.ensurePrintPagesContainer();
        container.innerHTML = '';
    }

    buildPrintPages() {
        const container = this.ensurePrintPagesContainer();
        container.innerHTML = '';
        if (!this.table) return;

        const range = this.getPrintRange();
        const settings = this.getPrintSettings();
        const pagination = this.getPrintPagination(range, settings);
        const scaleFactor = this.getPrintScaleFactor(range, settings);
        const paperSizes = { A4: { w: 210, h: 297 }, Letter: { w: 216, h: 279 } };
        const basePaper = paperSizes[settings.paper] || paperSizes.A4;
        const paper = settings.orientation === 'landscape'
            ? { w: basePaper.h, h: basePaper.w }
            : basePaper;
        const pxPerMm = this.getPxPerMm();
        const printableW = Math.max(1, (paper.w - settings.margins.left - settings.margins.right) * pxPerMm);
        const printableH = Math.max(1, (paper.h - settings.margins.top - settings.margins.bottom) * pxPerMm);
        const totalPages = pagination.pagesX * pagination.pagesY;
        const offsetsX = [0, ...pagination.breaksX];
        const offsetsY = [0, ...pagination.breaksY];

        let pageNo = 1;
        offsetsY.forEach((offsetY) => {
            offsetsX.forEach((offsetX) => {
                const page = document.createElement('div');
                page.className = 'vs-print-page';

                if (settings.headerFooter?.enabled) {
                    const header = document.createElement('div');
                    header.className = 'vs-print-page-header';
                    header.textContent = this.expandPrintTokens(settings.headerFooter.header.center, pageNo, totalPages);
                    page.appendChild(header);
                }

                const viewport = document.createElement('div');
                viewport.className = 'vs-print-page-viewport';
                viewport.style.width = `${Math.ceil(printableW)}px`;
                viewport.style.height = `${Math.ceil(printableH)}px`;

                const scaleWrap = document.createElement('div');
                scaleWrap.className = 'vs-print-page-scale';
                scaleWrap.style.transform = `scale(${scaleFactor})`;
                scaleWrap.style.transformOrigin = 'top left';

                const clone = this.table.cloneNode(true);
                clone.classList.add('vs-print-page-clone');
                clone.style.zoom = '1';
                clone.style.transform = 'none';
                clone.style.width = '0';
                clone.style.position = 'absolute';
                clone.style.left = `${-offsetX}px`;
                clone.style.top = `${-offsetY}px`;

                scaleWrap.appendChild(clone);
                const imageLayer = this.buildPrintImageLayer(offsetX, offsetY);
                if (imageLayer) {
                    scaleWrap.appendChild(imageLayer);
                }
                viewport.appendChild(scaleWrap);
                page.appendChild(viewport);

                if (settings.headerFooter?.enabled) {
                    const footer = document.createElement('div');
                    footer.className = 'vs-print-page-footer-lite';
                    footer.textContent = this.expandPrintTokens(settings.headerFooter.footer.center, pageNo, totalPages);
                    page.appendChild(footer);
                }

                container.appendChild(page);
                pageNo++;
            });
        });
    }

    buildPrintImageLayer(offsetX, offsetY) {
        const images = this.activeSheet.images || [];
        if (!images.length) return null;
        const layer = document.createElement('div');
        layer.className = 'vs-print-image-layer';
        layer.style.left = `${-offsetX}px`;
        layer.style.top = `${-offsetY}px`;
        images.forEach((img) => {
            const imageEl = document.createElement('img');
            imageEl.src = img.src;
            imageEl.style.left = `${img.x}px`;
            imageEl.style.top = `${img.y}px`;
            imageEl.style.width = `${img.w}px`;
            imageEl.style.height = `${img.h}px`;
            layer.appendChild(imageEl);
        });
        return layer;
    }

    applyPrintPageLayout(enable) {
        const grid = document.getElementById('grid-container');
        if (!this.table || !grid) return;

        if (!this.printLayoutState) this.printLayoutState = {};

        if (!enable) {
            this.applyDynamicPageRule(null);
            if (this.printLayoutState.tableZoom !== undefined) this.table.style.zoom = this.printLayoutState.tableZoom;
            if (this.printLayoutState.tableTransformOrigin !== undefined) this.table.style.transformOrigin = this.printLayoutState.tableTransformOrigin;
            if (this.printLayoutState.tableWidth !== undefined) this.table.style.width = this.printLayoutState.tableWidth;
            if (this.printLayoutState.gridOverflow !== undefined) grid.style.overflow = this.printLayoutState.gridOverflow;
            if (this.printLayoutState.gridHeight !== undefined) grid.style.height = this.printLayoutState.gridHeight;
            return;
        }

        this.printLayoutState.tableZoom = this.table.style.zoom;
        this.printLayoutState.tableTransformOrigin = this.table.style.transformOrigin;
        this.printLayoutState.tableWidth = this.table.style.width;
        this.printLayoutState.gridOverflow = grid.style.overflow;
        this.printLayoutState.gridHeight = grid.style.height;

        const settings = this.getPrintSettings();
        const range = this.getPrintRange();
        this.applyDynamicPageRule(settings);

        const scaleFactor = this.getPrintScaleFactor(range, settings);
        this.table.style.zoom = String(scaleFactor);
        this.table.style.transformOrigin = 'top left';
        this.table.style.width = '0';
        grid.style.overflow = 'visible';
        grid.style.height = 'auto';
    }

    applyDynamicPageRule(settings) {
        const styleId = 'vs-print-page-style';
        let styleEl = document.getElementById(styleId);
        if (!styleEl) {
            styleEl = document.createElement('style');
            styleEl.id = styleId;
            document.head.appendChild(styleEl);
        }

        if (!settings) {
            styleEl.textContent = '';
            return;
        }

        const paper = settings.paper === 'Letter' ? 'Letter' : 'A4';
        const orientation = settings.orientation === 'landscape' ? 'landscape' : 'portrait';
        const top = Math.max(0, Math.min(50, Number(settings.margins?.top) || 0));
        const right = Math.max(0, Math.min(50, Number(settings.margins?.right) || 0));
        const bottom = Math.max(0, Math.min(50, Number(settings.margins?.bottom) || 0));
        const left = Math.max(0, Math.min(50, Number(settings.margins?.left) || 0));

        styleEl.textContent = `
@media print {
    @page {
        size: ${paper} ${orientation};
        margin: ${top}mm ${right}mm ${bottom}mm ${left}mm;
    }
}
`;
    }

    getPxPerMm() {
        const testEl = document.createElement('div');
        testEl.style.width = '100mm';
        testEl.style.height = '0';
        testEl.style.position = 'absolute';
        testEl.style.visibility = 'hidden';
        testEl.style.pointerEvents = 'none';
        document.body.appendChild(testEl);
        const px = testEl.getBoundingClientRect().width;
        document.body.removeChild(testEl);
        return Math.max(1, px / 100);
    }

    getPrintScaleFactor(range, settings = this.getPrintSettings()) {
        const paperSizes = { A4: { w: 210, h: 297 }, Letter: { w: 216, h: 279 } };
        const basePaper = paperSizes[settings.paper] || paperSizes.A4;
        const paper = settings.orientation === 'landscape'
            ? { w: basePaper.h, h: basePaper.w }
            : basePaper;
        const pxPerMm = this.getPxPerMm();
        const contentWidth = Math.max(1, this.getRangePixelSize(range).width);
        const contentHeight = Math.max(1, this.getRangePixelSize(range).height);
        const printableW = Math.max(1, (paper.w - settings.margins.left - settings.margins.right) * pxPerMm);
        const printableH = Math.max(1, (paper.h - settings.margins.top - settings.margins.bottom) * pxPerMm);

        let scaleFactor = settings.scale / 100;
        if (settings.fitTo.enabled) {
            const targetW = printableW * settings.fitTo.widthPages;
            const targetH = printableH * settings.fitTo.heightPages;
            scaleFactor = Math.min(targetW / contentWidth, targetH / contentHeight);
        }
        if (!Number.isFinite(scaleFactor) || scaleFactor <= 0) return 1;
        return Math.max(0.1, Math.min(4, scaleFactor));
    }

    applyPrintVisibility(enable) {
        if (!this.table) return;
        const range = this.getPrintRange();
        const hideClass = 'print-hide';
        const emptyClass = 'print-empty';
        this.togglePrintRowHeaderColumn(enable);
        this.togglePrintColumnVisibility(enable, range);
        this.togglePrintRowVisibility(enable, range);
        const cells = this.table.querySelectorAll('.cell');
        cells.forEach(cell => {
            const isHeader = cell.classList.contains('header');
            if (isHeader) {
                if (!enable) {
                    cell.classList.remove(hideClass);
                    cell.classList.remove(emptyClass);
                    return;
                }
                cell.classList.add(hideClass);
                cell.classList.remove(emptyClass);
                return;
            }

            if (!enable) {
                cell.classList.remove(hideClass);
                cell.classList.remove(emptyClass);
                return;
            }
            const id = cell.dataset.id;
            if (!id) return;
            const parsed = this.parseCellId(id);
            if (parsed.colNum < range.startCol || parsed.colNum > range.endCol || parsed.row < range.startRow || parsed.row > range.endRow) {
                cell.classList.add(hideClass);
                cell.classList.remove(emptyClass);
            } else {
                if (this.isCellPrintable(id)) {
                    cell.classList.remove(hideClass);
                    cell.classList.remove(emptyClass);
                } else {
                    cell.classList.remove(hideClass);
                    cell.classList.add(emptyClass);
                }
            }
        });
        this.applyPrintHeaderFooter(enable);
    }

    togglePrintRowHeaderColumn(enable) {
        if (!this.colgroup || !this.colgroup.children || this.colgroup.children.length === 0) return;
        const rowHeaderCol = this.colgroup.children[0];
        if (!rowHeaderCol) return;

        if (enable) {
            if (rowHeaderCol.dataset.printPrevWidth == null) {
                rowHeaderCol.dataset.printPrevWidth = rowHeaderCol.style.width || '';
            }
            rowHeaderCol.style.width = '0px';
            rowHeaderCol.style.minWidth = '0px';
            rowHeaderCol.style.maxWidth = '0px';
            rowHeaderCol.style.display = 'none';
            return;
        }

        rowHeaderCol.style.display = '';
        rowHeaderCol.style.minWidth = '';
        rowHeaderCol.style.maxWidth = '';
        if (rowHeaderCol.dataset.printPrevWidth != null) {
            rowHeaderCol.style.width = rowHeaderCol.dataset.printPrevWidth;
            delete rowHeaderCol.dataset.printPrevWidth;
        } else {
            rowHeaderCol.style.width = '40px';
        }
    }

    togglePrintColumnVisibility(enable, range) {
        if (!this.colgroup || !this.colgroup.children) return;
        const cols = this.colgroup.children;
        for (let i = 1; i < cols.length; i++) {
            const col = cols[i];
            const colNum = i;
            if (!col) continue;
            if (!enable) {
                this.restorePrintVisibilityStyles(col);
                continue;
            }
            if (colNum < range.startCol || colNum > range.endCol) {
                this.applyPrintHiddenStyles(col);
            } else {
                this.restorePrintVisibilityStyles(col);
            }
        }
    }

    togglePrintRowVisibility(enable, range) {
        if (!this.tbody) return;
        const rows = this.tbody.querySelectorAll('tr');
        rows.forEach((tr) => {
            const rowNum = Number(tr.dataset.rowIndex || '0');
            if (!rowNum) return;
            if (!enable) {
                this.restorePrintVisibilityStyles(tr);
                return;
            }
            if (rowNum < range.startRow || rowNum > range.endRow) {
                this.applyPrintHiddenStyles(tr);
            } else {
                this.restorePrintVisibilityStyles(tr);
            }
        });
    }

    applyPrintHiddenStyles(el) {
        if (!el) return;
        if (el.dataset.printPrevDisplay == null) el.dataset.printPrevDisplay = el.style.display || '';
        if (el.dataset.printPrevWidth == null) el.dataset.printPrevWidth = el.style.width || '';
        if (el.dataset.printPrevMinWidth == null) el.dataset.printPrevMinWidth = el.style.minWidth || '';
        if (el.dataset.printPrevMaxWidth == null) el.dataset.printPrevMaxWidth = el.style.maxWidth || '';
        if (el.dataset.printPrevHeight == null) el.dataset.printPrevHeight = el.style.height || '';
        if (el.dataset.printPrevMinHeight == null) el.dataset.printPrevMinHeight = el.style.minHeight || '';
        if (el.dataset.printPrevMaxHeight == null) el.dataset.printPrevMaxHeight = el.style.maxHeight || '';
        el.style.display = 'none';
        el.style.width = '0px';
        el.style.minWidth = '0px';
        el.style.maxWidth = '0px';
        el.style.height = '0px';
        el.style.minHeight = '0px';
        el.style.maxHeight = '0px';
    }

    restorePrintVisibilityStyles(el) {
        if (!el) return;
        if (el.dataset.printPrevDisplay != null) {
            el.style.display = el.dataset.printPrevDisplay;
            delete el.dataset.printPrevDisplay;
        }
        if (el.dataset.printPrevWidth != null) {
            el.style.width = el.dataset.printPrevWidth;
            delete el.dataset.printPrevWidth;
        }
        if (el.dataset.printPrevMinWidth != null) {
            el.style.minWidth = el.dataset.printPrevMinWidth;
            delete el.dataset.printPrevMinWidth;
        }
        if (el.dataset.printPrevMaxWidth != null) {
            el.style.maxWidth = el.dataset.printPrevMaxWidth;
            delete el.dataset.printPrevMaxWidth;
        }
        if (el.dataset.printPrevHeight != null) {
            el.style.height = el.dataset.printPrevHeight;
            delete el.dataset.printPrevHeight;
        }
        if (el.dataset.printPrevMinHeight != null) {
            el.style.minHeight = el.dataset.printPrevMinHeight;
            delete el.dataset.printPrevMinHeight;
        }
        if (el.dataset.printPrevMaxHeight != null) {
            el.style.maxHeight = el.dataset.printPrevMaxHeight;
            delete el.dataset.printPrevMaxHeight;
        }
    }

    isCellPrintable(cellId) {
        const anchorId = this.normalizeMergedCellId(cellId);
        const hasData = !!(this.data?.[anchorId] && String(this.data[anchorId]).trim() !== '');
        const hasFormula = !!(this.cellFormulas?.[anchorId] && String(this.cellFormulas[anchorId]).trim() !== '');
        const hasStyle = !!this.cellStyles?.[anchorId];
        const hasFormat = !!this.cellFormats?.[anchorId];
        const hasBorder = !!this.cellBorders?.[anchorId];
        return hasData || hasFormula || hasStyle || hasFormat || hasBorder;
    }

    applyPrintHeaderFooter(enable) {
        const header = document.getElementById('print-header');
        const footer = document.getElementById('print-footer');
        if (!header || !footer) return;
        const settings = this.getPrintSettings();
        const show = enable && settings.headerFooter?.enabled;
        const showSlots = settings.headerFooter?.show || {};
        header.classList.toggle('print-hf-visible', !!show);
        footer.classList.toggle('print-hf-visible', !!show);
        header.style.display = show ? 'flex' : 'none';
        footer.style.display = show ? 'flex' : 'none';
        if (!show) return;
        const pages = this.lastPrintPageCount || 1;
        header.querySelector('.print-h-left').textContent = showSlots.headerLeft === false ? '' : this.expandPrintTokens(settings.headerFooter.header.left, 1, pages);
        header.querySelector('.print-h-center').textContent = showSlots.headerCenter === false ? '' : this.expandPrintTokens(settings.headerFooter.header.center, 1, pages);
        header.querySelector('.print-h-right').textContent = showSlots.headerRight === false ? '' : this.expandPrintTokens(settings.headerFooter.header.right, 1, pages);
        footer.querySelector('.print-f-left').textContent = showSlots.footerLeft === false ? '' : this.expandPrintTokens(settings.headerFooter.footer.left, 1, pages);
        footer.querySelector('.print-f-center').textContent = showSlots.footerCenter === false ? '' : this.expandPrintTokens(settings.headerFooter.footer.center, 1, pages);
        footer.querySelector('.print-f-right').textContent = showSlots.footerRight === false ? '' : this.expandPrintTokens(settings.headerFooter.footer.right, 1, pages);
    }

    expandPrintTokens(text, page, pages) {
        const now = new Date();
        const yyyy = now.getFullYear();
        const mm = String(now.getMonth() + 1).padStart(2, '0');
        const dd = String(now.getDate()).padStart(2, '0');
        const date = `${yyyy}-${mm}-${dd}`;
        const title = (document.querySelector('.filename')?.innerText || 'Untitled').trim();
        return String(text || '')
            .replace(/\{date\}/gi, date)
            .replace(/\{title\}/gi, title)
            .replace(/\{page\}/gi, String(page))
            .replace(/\{pages\}/gi, String(pages));
    }

    renderPrintPreview() {
        const preview = document.querySelector('.print-preview-sheet');
        if (!preview) return;
        const range = this.getPrintRange();
        const settings = this.getPrintSettings();
        const paperSizes = { A4: { w: 210, h: 297 }, Letter: { w: 216, h: 279 } };
        const basePaper = paperSizes[settings.paper] || paperSizes.A4;
        const paper = settings.orientation === 'landscape'
            ? { w: basePaper.h, h: basePaper.w }
            : basePaper;
        const pxPerMm = this.getPxPerMm();
        const contentWidth = this.getRangePixelSize(range).width;
        const contentHeight = this.getRangePixelSize(range).height;
        const printableW = Math.max(1, (paper.w - settings.margins.left - settings.margins.right) * pxPerMm);
        const printableH = Math.max(1, (paper.h - settings.margins.top - settings.margins.bottom) * pxPerMm);

        const scaleFactor = this.getPrintScaleFactor(range, settings);
        const pagination = this.getPrintPagination(range, settings);
        const pagesX = pagination.pagesX;
        const pagesY = pagination.pagesY;
        this.lastPrintPageCount = pagesX * pagesY;

        preview.innerHTML = '';
        const canvas = document.createElement('div');
        canvas.className = 'print-preview-canvas';
        preview.appendChild(canvas);

        // Mini content preview (scaled grid + text)
        const grid = document.createElement('div');
        grid.className = 'print-preview-grid';
        canvas.appendChild(grid);
        const canvasWidth = Math.max(1, grid.clientWidth || grid.getBoundingClientRect().width || 1);
        const canvasHeight = Math.max(1, grid.clientHeight || grid.getBoundingClientRect().height || 1);
        const totalPrintableW = printableW * pagesX;
        const totalPrintableH = printableH * pagesY;
        const scaleToPreview = Math.min(canvasWidth / Math.max(1, totalPrintableW), canvasHeight / Math.max(1, totalPrintableH));
        const finalScale = scaleFactor * scaleToPreview;
        const previewW = totalPrintableW * scaleToPreview;
        const previewH = totalPrintableH * scaleToPreview;
        grid.style.width = `${previewW}px`;
        grid.style.height = `${previewH}px`;
        const gridOffsetX = (canvasWidth - previewW) / 2;
        const gridOffsetY = (canvasHeight - previewH) / 2;
        grid.style.left = `${gridOffsetX}px`;
        grid.style.top = `${gridOffsetY}px`;

        let x = 0;
        let y = 0;
        let cellCount = 0;
        const maxCells = 600;
        for (let r = range.startRow; r <= range.endRow; r++) {
            const h = (this.rowHeights[r] || this.baseRowHeight) * finalScale;
            if (y + h > previewH) break;
            x = 0;
            for (let c = range.startCol; c <= range.endCol; c++) {
                const w = (this.colWidths[c - 1] || this.baseColWidth) * finalScale;
                if (x + w > previewW) break;
                const id = `${this.numberToCol(c)}${r}`;
                if (!this.isCellPrintable(id)) {
                    x += w;
                    continue;
                }
                const cell = document.createElement('div');
                cell.className = 'preview-cell';
                cell.style.left = `${x}px`;
                cell.style.top = `${y}px`;
                cell.style.width = `${Math.max(1, w)}px`;
                cell.style.height = `${Math.max(1, h)}px`;
                const raw = this.getRawValue(id);
                if (raw) cell.textContent = raw;
                const style = this.cellStyles[id];
                if (style?.backgroundColor) cell.style.backgroundColor = style.backgroundColor;
                if (style?.color) cell.style.color = style.color;
                grid.appendChild(cell);
                cellCount++;
                if (cellCount >= maxCells) break;
                x += w;
            }
            if (cellCount >= maxCells) break;
            y += h;
        }

        // Page break lines
        pagination.breaksX.forEach((breakX) => {
            const line = document.createElement('div');
            line.className = 'print-preview-break-v';
            line.style.left = `${gridOffsetX + (breakX * finalScale)}px`;
            line.style.top = `${gridOffsetY}px`;
            line.style.height = `${previewH}px`;
            canvas.appendChild(line);
        });
        pagination.breaksY.forEach((breakY) => {
            const line = document.createElement('div');
            line.className = 'print-preview-break-h';
            line.style.top = `${gridOffsetY + (breakY * finalScale)}px`;
            line.style.left = `${gridOffsetX}px`;
            line.style.width = `${previewW}px`;
            canvas.appendChild(line);
        });

        // Page labels
        const pages = document.createElement('div');
        pages.className = 'print-preview-pages';
        pages.style.gridTemplateColumns = `repeat(${pagesX}, 1fr)`;
        pages.style.gridTemplateRows = `repeat(${pagesY}, 1fr)`;
        preview.appendChild(pages);
        const totalPages = pagesX * pagesY;
        for (let i = 0; i < totalPages; i++) {
            const tile = document.createElement('div');
            tile.className = 'print-preview-page-tile';
            const label = document.createElement('span');
            label.textContent = `Page ${i + 1}`;
            if (settings.headerFooter?.enabled) {
                const slot = settings.headerFooter.show || {};
                const header = document.createElement('div');
                header.className = 'print-preview-page-header';
                const hL = document.createElement('span');
                const hC = document.createElement('span');
                const hR = document.createElement('span');
                hL.className = 'left';
                hC.className = 'center';
                hR.className = 'right';
                hL.textContent = slot.headerLeft === false ? '' : this.expandPrintTokens(settings.headerFooter.header.left, i + 1, totalPages);
                hC.textContent = slot.headerCenter === false ? '' : this.expandPrintTokens(settings.headerFooter.header.center, i + 1, totalPages);
                hR.textContent = slot.headerRight === false ? '' : this.expandPrintTokens(settings.headerFooter.header.right, i + 1, totalPages);
                header.appendChild(hL);
                header.appendChild(hC);
                header.appendChild(hR);

                const footer = document.createElement('div');
                footer.className = 'print-preview-page-footer';
                const fL = document.createElement('span');
                const fC = document.createElement('span');
                const fR = document.createElement('span');
                fL.className = 'left';
                fC.className = 'center';
                fR.className = 'right';
                fL.textContent = slot.footerLeft === false ? '' : this.expandPrintTokens(settings.headerFooter.footer.left, i + 1, totalPages);
                fC.textContent = slot.footerCenter === false ? '' : this.expandPrintTokens(settings.headerFooter.footer.center, i + 1, totalPages);
                fR.textContent = slot.footerRight === false ? '' : this.expandPrintTokens(settings.headerFooter.footer.right, i + 1, totalPages);
                footer.appendChild(fL);
                footer.appendChild(fC);
                footer.appendChild(fR);
                tile.appendChild(header);
                tile.appendChild(footer);
            }
            tile.appendChild(label);
            pages.appendChild(tile);
        }
    }

    getPrintPagination(range, settings = this.getPrintSettings()) {
        const paperSizes = { A4: { w: 210, h: 297 }, Letter: { w: 216, h: 279 } };
        const basePaper = paperSizes[settings.paper] || paperSizes.A4;
        const paper = settings.orientation === 'landscape'
            ? { w: basePaper.h, h: basePaper.w }
            : basePaper;
        const pxPerMm = this.getPxPerMm();
        const content = this.getRangePixelSize(range);
        const contentWidth = Math.max(1, content.width);
        const contentHeight = Math.max(1, content.height);
        const printableW = Math.max(1, (paper.w - settings.margins.left - settings.margins.right) * pxPerMm);
        const printableH = Math.max(1, (paper.h - settings.margins.top - settings.margins.bottom) * pxPerMm);
        const scaleFactor = this.getPrintScaleFactor(range, settings);
        const scaledW = contentWidth * scaleFactor;
        const scaledH = contentHeight * scaleFactor;
        const breaksX = [];
        const breaksY = [];
        const pagesX = Math.max(1, Math.ceil(scaledW / printableW));
        const pagesY = Math.max(1, Math.ceil(scaledH / printableH));
        for (let i = 1; i < pagesX; i++) {
            breaksX.push((i * printableW) / scaleFactor);
        }
        for (let j = 1; j < pagesY; j++) {
            breaksY.push((j * printableH) / scaleFactor);
        }
        const pageSpanWidth = (printableW * pagesX) / scaleFactor;
        const pageSpanHeight = (printableH * pagesY) / scaleFactor;
        return {
            pagesX,
            pagesY,
            breaksX,
            breaksY,
            contentWidth,
            contentHeight,
            pageSpanWidth,
            pageSpanHeight
        };
    }

    syncPageBreakPreviewButton() {
        const btn = document.getElementById('btn-page-break-preview');
        if (!btn) return;
        btn.classList.toggle('active', this.showPageBreakPreview);
    }

    togglePageBreakPreview(forceValue = null) {
        this.showPageBreakPreview = forceValue == null ? !this.showPageBreakPreview : !!forceValue;
        this.syncPageBreakPreviewButton();
        this.updatePageBreakPreview();
    }

    updatePageBreakPreview() {
        if (!this.pageBreakOverlay) return;
        this.syncPageBreakPreviewButton();
        if (!this.showPageBreakPreview || !this.table) {
            this.pageBreakOverlay.style.display = 'none';
            this.pageBreakOverlay.innerHTML = '';
            return;
        }

        const range = this.getPrintRange();
        if (!range) {
            this.pageBreakOverlay.style.display = 'none';
            this.pageBreakOverlay.innerHTML = '';
            return;
        }

        const tlRect = this.getCellRectForCoord(range.startCol, range.startRow);
        const brRect = this.getCellRectForCoord(range.endCol, range.endRow);
        const containerRect = this.container.getBoundingClientRect();
        if (!tlRect || !brRect || !containerRect) {
            this.pageBreakOverlay.style.display = 'none';
            this.pageBreakOverlay.innerHTML = '';
            return;
        }

        const left = tlRect.left - containerRect.left + this.container.scrollLeft;
        const top = tlRect.top - containerRect.top + this.container.scrollTop;
        const content = this.getRangePixelSize(range);
        const contentWidthPx = Math.max(1, content.width);
        const contentHeightPx = Math.max(1, content.height);
        const pagination = this.getPrintPagination(range, this.getPrintSettings());
        const width = Math.max(contentWidthPx, Math.ceil(pagination.pageSpanWidth || 0));
        const height = Math.max(contentHeightPx, Math.ceil(pagination.pageSpanHeight || 0));

        this.pageBreakOverlay.style.display = 'block';
        this.pageBreakOverlay.style.left = `${left}px`;
        this.pageBreakOverlay.style.top = `${top}px`;
        this.pageBreakOverlay.style.width = `${width}px`;
        this.pageBreakOverlay.style.height = `${height}px`;
        this.pageBreakOverlay.innerHTML = '';

        pagination.breaksX.forEach((x) => {
            const line = document.createElement('div');
            line.className = 'page-break-line-v';
            line.style.left = `${Math.max(0, Math.min(width, x))}px`;
            this.pageBreakOverlay.appendChild(line);
        });
        pagination.breaksY.forEach((y) => {
            const line = document.createElement('div');
            line.className = 'page-break-line-h';
            line.style.top = `${Math.max(0, Math.min(height, y))}px`;
            this.pageBreakOverlay.appendChild(line);
        });
    }

    getRangePixelSize(range) {
        let width = 0;
        let height = 0;
        for (let c = range.startCol; c <= range.endCol; c++) {
            width += this.colWidths[c - 1] || this.baseColWidth;
        }
        for (let r = range.startRow; r <= range.endRow; r++) {
            height += this.rowHeights[r] || this.baseRowHeight;
        }
        return { width, height };
    }

    getUsedRangeForSheet(sheet = this.activeSheet) {
        let maxRow = 0;
        let maxCol = 0;
        const scan = (key) => {
            const parsed = this.parseCellId(key);
            if (!parsed) return;
            maxRow = Math.max(maxRow, parsed.row);
            maxCol = Math.max(maxCol, parsed.colNum);
        };
        Object.keys(sheet.data || {}).forEach(scan);
        Object.keys(sheet.cellFormulas || {}).forEach(scan);
        Object.keys(sheet.cellStyles || {}).forEach(scan);
        Object.keys(sheet.cellFormats || {}).forEach(scan);
        Object.keys(sheet.cellBorders || {}).forEach(scan);
        (sheet.mergedRanges || []).forEach((range) => {
            const normalized = this.normalizeMergedRangeEntry(range);
            if (!normalized) return;
            maxRow = Math.max(maxRow, normalized.endRow);
            maxCol = Math.max(maxCol, normalized.endCol);
        });
        if (maxRow === 0) maxRow = 1;
        if (maxCol === 0) maxCol = 1;
        return { startCol: 1, startRow: 1, endCol: maxCol, endRow: maxRow };
    }

    formatRangeRef(range) {
        if (!range) return '';
        const start = `${this.numberToCol(range.startCol)}${range.startRow}`;
        const end = `${this.numberToCol(range.endCol)}${range.endRow}`;
        return start === end ? start : `${start}:${end}`;
    }

    togglePrintAreaFromRibbon() {
        const settings = this.getPrintSettings();
        if (settings.printArea) {
            this.clearPrintArea();
            return;
        }
        this.setPrintAreaFromSelection();
    }

    updatePrintAreaToggleUI() {
        const btn = document.getElementById('btn-print-area-toggle');
        if (!btn) return;
        const settings = this.getPrintSettings();
        const hasArea = !!settings.printArea;
        btn.classList.toggle('active', hasArea);
        if (hasArea) {
            btn.title = `Print Area Set: ${this.formatRangeRef(settings.printArea)} (Click to Clear)`;
        } else {
            btn.title = 'Set Print Area (from selection)';
        }
    }

    setPrintAreaFromSelection() {
        const range = this.getEffectiveRange();
        if (!range) return;
        const normalized = this.expandRangeToIncludeMerges(range);
        const settings = this.getPrintSettings();
        settings.printArea = {
            startCol: normalized.startCol,
            startRow: normalized.startRow,
            endCol: normalized.endCol,
            endRow: normalized.endRow
        };
        this.markDirty();
        this.updateStatusBadge(`Print area set: ${this.formatRangeRef(settings.printArea)}`);
        this.updatePrintAreaToggleUI();
        this.updatePrintControlsState();
        const modal = document.getElementById('print-settings-modal');
        if (modal && modal.style.display !== 'none') this.renderPrintPreview();
        this.updatePageBreakPreview();
    }

    clearPrintArea() {
        const settings = this.getPrintSettings();
        settings.printArea = null;
        this.markDirty();
        this.updateStatusBadge('Print area cleared');
        this.updatePrintAreaToggleUI();
        this.updatePrintControlsState();
        const modal = document.getElementById('print-settings-modal');
        if (modal && modal.style.display !== 'none') this.renderPrintPreview();
        this.updatePageBreakPreview();
    }

    updateStatusBadge(message) {
        const badge = document.querySelector('.status-badge');
        if (!badge) return;
        const prev = badge.innerText;
        const prevClass = badge.className;
        badge.innerText = message;
        badge.className = 'status-badge modified';
        setTimeout(() => {
            if (!badge.isConnected) return;
            badge.innerText = prev;
            badge.className = prevClass;
        }, 1800);
    }

    getPrintRange() {
        const settings = this.getPrintSettings();
        if (settings.printArea) {
            return this.expandRangeToIncludeMerges(settings.printArea);
        }
        let baseRange = null;
        if (settings.range === 'selection') {
            baseRange = this.selectionRange || settings.selectionRange || null;
        } else if (settings.range === 'used') {
            baseRange = this.getUsedRangeForSheet();
        } else {
            baseRange = { startCol: 1, startRow: 1, endCol: this.cols, endRow: this.rows };
        }

        if (!baseRange) return { startCol: 1, startRow: 1, endCol: 1, endRow: 1 };

        if (settings.range === 'selection' || settings.range === 'current') {
            const used = this.getUsedRangeForSheet();
            const intersected = this.intersectRanges(baseRange, used);
            if (intersected) {
                return this.expandRangeToIncludeMerges(intersected);
            }
        }

        return this.expandRangeToIncludeMerges(baseRange);
    }

    intersectRanges(a, b) {
        if (!a || !b) return null;
        const startCol = Math.max(a.startCol, b.startCol);
        const startRow = Math.max(a.startRow, b.startRow);
        const endCol = Math.min(a.endCol, b.endCol);
        const endRow = Math.min(a.endRow, b.endRow);
        if (startCol > endCol || startRow > endRow) return null;
        return { startCol, startRow, endCol, endRow };
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

    // ??? Clipboard ?????????????????????????????????????????
    copySelection() {
        const range = this.getEffectiveRange();
        if (!range) return;

        const expandedRange = this.expandRangeToIncludeMerges(range);
        const { startCol, startRow, endCol, endRow } = expandedRange;
        const rows = [];
        this.clipboardData = [];
        this.clipboardMerges = [];

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

        // Capture merged ranges within the copied area (relative offsets)
        const merges = this.getNormalizedMergedRanges();
        merges.forEach((merge) => {
            if (
                merge.startCol >= startCol &&
                merge.endCol <= endCol &&
                merge.startRow >= startRow &&
                merge.endRow <= endRow
            ) {
                this.clipboardMerges.push({
                    startColOffset: merge.startCol - startCol,
                    endColOffset: merge.endCol - startCol,
                    startRowOffset: merge.startRow - startRow,
                    endRowOffset: merge.endRow - startRow
                });
            }
        });

        // Copy to system clipboard as TSV
        const text = rows.join('\n');
        navigator.clipboard.writeText(text).catch(() => {});

        // Flash visual feedback
        this.flashCopyBorder(expandedRange);
    }

    cutSelection() {
        const range = this.getEffectiveRange();
        if (!range) return;

        this.copySelection();
        this.isCut = true;
        this.cutRange = { ...this.expandRangeToIncludeMerges(range) };

        // Add dashed border for cut visual
        this.flashCutBorder(this.cutRange);
    }

    async pasteAtSelection() {
        if (!this.selectedCell) return;

        const anchor = this.parseCellId(this.selectedCell.dataset.id);
        let rowsData = [];
        let isInternalPaste = false;

        // Prefer internal clipboard to avoid triggering clipboard-read permission prompts.
        if (this.clipboardData && this.clipboardData.length > 0) {
            rowsData = this.clipboardData;
            isInternalPaste = true;
        } else {
            try {
                const text = await navigator.clipboard.readText();
                if (text && text.trim() !== '') {
                    rowsData = this.parseTSV(text);
                    isInternalPaste = false;
                }
            } catch (err) {
                console.warn('System clipboard access denied:', err);
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
            const ok = await this.showConfirmAsync('병합된 셀이 포함된 영역입니다. 붙여넣기를 위해 병합을 해제할까요?');
            if (!ok) return;
            this.unmergeRange(pasteRange);
        }

        let pasteEndCol = anchor.colNum + numCols - 1;
        let pasteEndRow = anchor.row + numRows - 1;
        if (pasteEndCol > this.cols) {
            for (let cAdd = this.cols + 1; cAdd <= pasteEndCol; cAdd++) {
                this.colWidths[cAdd - 1] = this.baseColWidth;
            }
            this.cols = pasteEndCol;
            this.refreshGridUI();
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

                const cell = this.getCellEl(targetColNum, targetRow);
                if (cell) {
                    const val = rowsData[r][c] || '';
                    this.setRawValue(cell.dataset.id, val);
                    this.renderCellValue(cell);
                }
            }
        }

        // Reapply merged ranges when pasting internal clipboard data
        if (isInternalPaste && Array.isArray(this.clipboardMerges) && this.clipboardMerges.length > 0) {
            this.clipboardMerges.forEach((merge) => {
                const startCol = anchor.colNum + merge.startColOffset;
                const endCol = anchor.colNum + merge.endColOffset;
                const startRow = anchor.row + merge.startRowOffset;
                const endRow = anchor.row + merge.endRowOffset;
                pasteEndCol = Math.max(pasteEndCol, endCol);
                pasteEndRow = Math.max(pasteEndRow, endRow);
                if (startCol < 1 || startRow < 1) return;
                if (endRow > this.rows) {
                    const rowsToAdd = Math.max(30, endRow - this.rows);
                    this.createRowElements(this.rows + 1, this.rows + rowsToAdd);
                    this.rows += rowsToAdd;
                }
                if (endCol > this.cols) {
                    for (let cAdd = this.cols + 1; cAdd <= endCol; cAdd++) {
                        this.colWidths[cAdd - 1] = this.baseColWidth;
                    }
                    this.cols = endCol;
                    this.refreshGridUI();
                }
                this.mergeRangeSilently({ startCol, startRow, endCol, endRow });
            });
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
        this.setSelectionRange(anchor.colNum, anchor.row, pasteEndCol, pasteEndRow);
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

    normalizeClipboardText(text) {
        if (text === null || text === undefined) return '';
        return String(text)
            .replace(/\r\n/g, '\n')
            .replace(/\r/g, '\n')
            .replace(/\n+$/g, '');
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

    // ??? Keyboard Navigation ???????????????????????????????
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

            // For merged cells or range selections, ensure overwrite by explicitly preparing enter mode.
            if (this.selectionRange || activeCell.classList.contains('merge-anchor')) {
                this.prepareEnterMode(activeCell, true);
                return;
            }

            // Transition to editing mode without manual clearing (selection from handleCellFocus will overwrite)
            this.isEditing = true;
            this.originalValue = this.getRawValue(activeCell.dataset.id);
            activeCell.classList.add('editing');
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
                if (rowNum < this.rows) {
                    nextRow++;
                } else {
                    const rowsToAdd = 30;
                    this.createRowElements(this.rows + 1, this.rows + rowsToAdd);
                    this.rows += rowsToAdd;
                    nextRow = rowNum + 1;
                }
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
                if (colNum < this.cols) {
                    nextCol = colNum + 1;
                } else {
                    const colsToAdd = 10;
                    for (let c = this.cols + 1; c <= this.cols + colsToAdd; c++) {
                        this.colWidths[c - 1] = this.baseColWidth;
                    }
                    this.cols += colsToAdd;
                    this.refreshGridUI();
                    nextCol = colNum + 1;
                }
                moved = true;
                e.preventDefault(); // Stop default browser scroll
                break;
            case 'Tab':
                e.preventDefault();
                if (e.shiftKey) {
                    if (colNum > 1) nextCol = colNum - 1;
                } else {
                    if (colNum < this.cols) {
                        nextCol = colNum + 1;
                    } else {
                        const colsToAdd = 10;
                        for (let c = this.cols + 1; c <= this.cols + colsToAdd; c++) {
                            this.colWidths[c - 1] = this.baseColWidth;
                        }
                        this.cols += colsToAdd;
                        this.refreshGridUI();
                        nextCol = colNum + 1;
                    }
                }
                moved = true;
                break;
            case 'Enter':
                e.preventDefault();
                if (e.shiftKey) {
                    if (rowNum > 1) nextRow--;
                } else {
                    if (rowNum < this.rows) {
                        nextRow++;
                    } else {
                        const rowsToAdd = 30;
                        this.createRowElements(this.rows + 1, this.rows + rowsToAdd);
                        this.rows += rowsToAdd;
                        nextRow = rowNum + 1;
                    }
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

    // ??? Mode Handlers ?????????????????????????????????????
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
        const parsed = this.parseCellId(cellId);
        if (parsed) this.refreshOverflowForRow(parsed.row);
        cell.focus(); // Keep focus for Ready mode navigation
    }

    moveSelection(deltaCol, deltaRow) {
        const activeCell = this.selectedCell;
        if (!activeCell) return;
        
        const { colNum, row } = this.parseCellId(activeCell.dataset.id);
        const desiredCol = colNum + deltaCol;
        const desiredRow = row + deltaRow;

        if (desiredRow > this.rows) {
            const rowsToAdd = Math.max(30, desiredRow - this.rows);
            this.createRowElements(this.rows + 1, this.rows + rowsToAdd);
            this.rows += rowsToAdd;
        }
        if (desiredCol > this.cols) {
            for (let c = this.cols + 1; c <= desiredCol; c++) {
                this.colWidths[c - 1] = this.baseColWidth;
            }
            this.cols = desiredCol;
            this.refreshGridUI();
        }

        const nextCol = Math.max(1, Math.min(this.cols, desiredCol));
        const nextRow = Math.max(1, Math.min(this.rows, desiredRow));
        
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

    // ??? File Import / Export ???????????????????????????????
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
            if (!confirm('작업 중인 내용이 사라집니다. 계속할까요?\n(Unsaved changes will be lost. Continue?)')) {
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
            alert('Excel 라이브러리를 불러오지 못했습니다. 스크립트 연결을 확인해 주세요.');
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
            sheet.colWidths = new Array(sheet.cols).fill(this.baseColWidth);
            sheet.rowHeights = new Array(sheet.rows + 1).fill(this.baseRowHeight);

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
            this.pageBreakOverlay = this.createOverlay('page-break-overlay');
            
            this.renderGrid();
            
            // Re-attach fill handle listener (since we just created a new one)
            this.fillHandle.addEventListener('mousedown', (e) => this.handleFillStart(e));
            
            // Refresh overlays
            this.updateSelectionOverlay();
            this.updateRangeVisual();
            this.updateFillHandlePosition();
            this.refreshFindIfActive();
            this.updatePrintAreaToggleUI();
            this.updatePrintControlsState();
            this.updatePageBreakPreview();
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
            alert('Excel 라이브러리를 불러오지 못했습니다. 스크립트 연결을 확인해 주세요.');
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
                            const currencySymbol = format.currency === 'USD' ? '$' : '₩';
                            cellObj = { v: numericValue, t: 'n', z: `"${currencySymbol}"#,##0${d > 0 ? '.' + '0'.repeat(d) : ''}` };
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

    getExcelThemeBaseHex(themeIndex) {
        const theme = Number(themeIndex);
        const palette = {
            0: '#FFFFFF', // lt1
            1: '#000000', // dk1
            2: '#E7E6E6', // lt2
            3: '#44546A', // dk2
            4: '#4472C4', // accent1
            5: '#ED7D31', // accent2
            6: '#A5A5A5', // accent3
            7: '#FFC000', // accent4
            8: '#5B9BD5', // accent5
            9: '#70AD47', // accent6
            10: '#0000FF', // hyperlink
            11: '#800080' // followed hyperlink
        };
        return palette[theme] || null;
    }

    applyExcelTint(hex, tint) {
        if (!hex || tint === null || tint === undefined || !Number.isFinite(Number(tint))) return hex;
        const t = Number(tint);
        const clean = hex.replace('#', '');
        if (clean.length !== 6) return hex;
        const toChannel = (v) => {
            if (t < 0) return Math.round(v * (1 + t));
            return Math.round(v * (1 - t) + 255 * t);
        };
        const clamp = (n) => Math.max(0, Math.min(255, n));
        const r = clamp(toChannel(parseInt(clean.slice(0, 2), 16)));
        const g = clamp(toChannel(parseInt(clean.slice(2, 4), 16)));
        const b = clamp(toChannel(parseInt(clean.slice(4, 6), 16)));
        const toHex = (n) => n.toString(16).padStart(2, '0').toUpperCase();
        return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
    }

    excelColorToHex(color) {
        if (!color) return null;
        if (color.argb) return this.argbToHex(color.argb);
        if (color.rgb) return this.argbToHex(color.rgb);
        if (color.theme !== undefined && color.theme !== null) {
            const base = this.getExcelThemeBaseHex(color.theme);
            return this.applyExcelTint(base, color.tint);
        }
        return null;
    }

    mapInternalStyleToExceljs(cell, style) {
        const font = {};
        if (style.fontWeight === 'bold') font.bold = true;
        if (style.fontStyle === 'italic') font.italic = true;
        if (style.textDecoration && String(style.textDecoration).includes('underline')) font.underline = true;
        if (style.textDecoration && String(style.textDecoration).includes('line-through')) font.strike = true;
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
        if (font.strike) {
            style.textDecoration = style.textDecoration
                ? `${style.textDecoration} line-through`
                : 'line-through';
        }
        if (font.size) style.fontSize = font.size + 'pt';
        if (font.name) style.fontFamily = font.name;
        if (font.color) {
            const fontHex = this.excelColorToHex(font.color);
            if (fontHex) style.color = fontHex;
        }

        const alignment = cell.alignment || {};
        if (alignment.horizontal) style.textAlign = alignment.horizontal;

        const fill = cell.fill || {};
        if (fill.pattern === 'solid' || fill.patternType === 'solid' || fill.type === 'pattern') {
            const fgHex = this.excelColorToHex(fill.fgColor);
            const bgHex = this.excelColorToHex(fill.bgColor);
            if (fgHex) {
                style.backgroundColor = fgHex;
            } else if (bgHex) {
                style.backgroundColor = bgHex;
            }
        }

        if (Object.keys(style).length > 0) {
            sheet.cellStyles[cellId] = style;
        }

        if (cell.border) {
            const mapStyle = (excelStyle) => {
                if (!excelStyle) return { style: 'solid', width: 1 };
                switch (excelStyle) {
                    case 'thin':
                        return { style: 'solid', width: 1 };
                    case 'medium':
                        return { style: 'solid', width: 2 };
                    case 'thick':
                        return { style: 'solid', width: 3 };
                    case 'dashed':
                        return { style: 'dashed', width: 1 };
                    case 'dotted':
                        return { style: 'dotted', width: 1 };
                    case 'mediumDashed':
                        return { style: 'dashed', width: 2 };
                    case 'dashDot':
                    case 'dashDotDot':
                    case 'slantDashDot':
                        return { style: 'dashed', width: 1 };
                    case 'mediumDashDot':
                    case 'mediumDashDotDot':
                        return { style: 'dashed', width: 2 };
                    case 'double':
                        return { style: 'solid', width: 2 };
                    case 'hair':
                        return { style: 'solid', width: 1 };
                    default:
                        return { style: 'solid', width: 1 };
                }
            };

            const toInternal = (side) => {
                const b = cell.border?.[side];
                if (!b || !b.style) return null;
                const mapped = mapStyle(b.style);
                return {
                    style: mapped.style,
                    color: this.argbToHex(b.color?.argb),
                    width: mapped.width
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

    detectCurrencyFromFormatCode(formatCode) {
        const raw = String(formatCode || '');
        if (!raw) return null;
        if (raw.includes('₩')) return 'KRW';
        if (raw.includes('$')) return 'USD';
        const bracketTokens = raw.match(/\[\$[^\]]+\]/g) || [];
        for (const token of bracketTokens) {
            if (token.includes('₩')) return 'KRW';
            if (token.includes('$')) return 'USD';
        }
        return null;
    }

    getNumFmtFromFormat(format) {
        if (format.type === 'currency') {
            const d = format.decimals ?? 2;
            const currencySymbol = format.currency === 'USD' ? '$' : '₩';
            return `"${currencySymbol}"#,##0${d > 0 ? '.' + '0'.repeat(d) : ''}`;
        }
        if (format.type === 'percentage') {
            const d = format.decimals ?? 2;
            return `0${d > 0 ? '.' + '0'.repeat(d) : ''}%`;
        }
        if (format.type === 'date') {
            return 'yyyy-mm-dd';
        }
        if (format.type === 'text') {
            return '@';
        }
        if (format.type === 'number' || format.decimals !== null) {
            const d = format.decimals ?? 0;
            return `#,##0${d > 0 ? '.' + '0'.repeat(d) : ''}`;
        }
        return null;
    }

    getExceljsFormatFromNumFmt(numFmt, cellType) {
        const rawCode = String(numFmt || '');
        const formatCode = rawCode.toLowerCase();
        let type = 'general';
        let decimals = null;
        let currency = null;

        if (formatCode === '@' || formatCode === 'text') {
            type = 'text';
            decimals = null;
        } else if (formatCode.includes('%')) {
            type = 'percentage';
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
        } else if (/\b(y|m|d|h|s)+\b/.test(formatCode) || cellType === 4 /* Date */) {
            type = 'date';
            decimals = null;
        } else if (formatCode.includes('[$') || formatCode.includes('₩') || formatCode.includes('$')) {
            type = 'currency';
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
            currency = this.detectCurrencyFromFormatCode(rawCode) || 'KRW';
        } else if (formatCode.includes('#') || formatCode.includes('0') || cellType === 2 /* Number */) {
            type = 'number';
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
        }

        return this.normalizeFormat({ type, decimals, currency });
    }

    // Style Mapping Helpers
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
        let currency = null;

        if (formatCode.includes('%')) {
            type = 'percentage';
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
        } else if (/\b[ymdhis]+\b/.test(formatCode) || wsCell.t === 'd') {
            type = 'date';
            decimals = null;
        } else if (formatCode.includes('[$') || formatCode.includes('₩') || formatCode.includes('$')) {
            type = 'currency';
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
            currency = this.detectCurrencyFromFormatCode(wsCell.z || '') || 'KRW';
        } else if (wsCell.t === 'n') {
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
        }

        return this.normalizeFormat({ type, decimals, currency });
    }

    mapXlsxFormatToInternal(cellId, wsCell) {
        const formatCode = String(wsCell.z || '').toLowerCase();
        let type = 'general';
        let decimals = null;
        let currency = null;

        if (formatCode.includes('%')) {
            type = 'percentage';
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
        } else if (/\b[ymdhis]+\b/.test(formatCode) || wsCell.t === 'd') {
            type = 'date';
            decimals = null;
        } else if (formatCode.includes('[$') || formatCode.includes('₩') || formatCode.includes('$')) {
            type = 'currency';
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
            currency = this.detectCurrencyFromFormatCode(wsCell.z || '') || 'KRW';
        } else if (wsCell.t === 'n') {
            decimals = this.extractDecimalPlacesFromFormat(formatCode);
        }

        if (type !== 'general' || decimals !== null) {
            this.setCellFormat(cellId, { type, decimals, currency });
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

    // Phase 6: Table Operations
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
        this.rowHeights.splice(index, 0, this.baseRowHeight);
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
        this.colWidths.splice(index - 1, 0, this.baseColWidth);
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
            if (btnCancel) btnCancel.textContent = '痍⑥냼';
            if (btnOk) btnOk.textContent = '?뺤씤';
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



