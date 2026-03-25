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
        this.init();
    }

    async confirmCsvSingleSheet() {
        if (this.sheets.length <= 1) return true;
        if (this.csvConfirmInProgress) return false;
        this.csvConfirmInProgress = true;
        const proceed = await this.showConfirmAsync(
            'CSV는 활성 시트만 저장할 수 있습니다. 전체 시트를 저장하려면 다른 형식을 선택하세요. 계속 저장할까요?'
        );
        this.csvConfirmInProgress = false;
        return proceed;
    }

    hasAnyStyles() {
        return this.sheets.some(sheet => Object.keys(sheet.cellStyles || {}).length > 0);
    }

    async confirmXlsxStyleWarning() {
        return true;
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
        this.resizeGuide = this.createOverlay('resize-guide');

        this.renderGrid();
        this.renderSheetTabs();
        this.setupEventListeners();
    }

    createOverlay(className) {
        const div = document.createElement('div');
        div.className = className;
        div.style.display = 'none';
        this.container.appendChild(div);
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
        const val = this.data[cellId];
        return val === undefined || val === null ? '' : String(val);
    }

    setRawValue(cellId, value) {
        this.data[cellId] = value === undefined || value === null ? '' : String(value);
    }

    getRawValueForSheet(sheet, cellId) {
        const val = sheet.data[cellId];
        return val === undefined || val === null ? '' : String(val);
    }

    setRawValueForSheet(sheet, cellId, value) {
        sheet.data[cellId] = value === undefined || value === null ? '' : String(value);
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
        const stored = this.cellFormats[cellId] || {};
        return this.normalizeFormat(stored);
    }

    getCellFormatForSheet(sheet, cellId) {
        const stored = sheet.cellFormats[cellId] || {};
        return this.normalizeFormat(stored);
    }

    setCellFormat(cellId, format) {
        const normalized = this.normalizeFormat(format);
        if (this.isDefaultFormat(normalized)) {
            delete this.cellFormats[cellId];
        } else {
            this.cellFormats[cellId] = normalized;
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
                const key = cellId.toUpperCase();
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
                const a = this.parseCellId(start);
                const b = this.parseCellId(end);
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

    renderCellById(cellId) {
        const cell = document.querySelector(`[data-id="${cellId}"]`);
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
        this.findInput.addEventListener('input', refreshFind);
        this.replaceInput.addEventListener('input', () => {
            this.findState.replace = this.replaceInput.value || '';
        });
        this.findCase.addEventListener('change', refreshFind);
        this.findExact.addEventListener('change', refreshFind);
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

        // Fill Handle
        this.createFillHandle();
        this.createSelectionOverlay();
        this.createRangeOverlay();

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
        this.selectedCell = cell;
        const cellId = cell.dataset.id;
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
    }

    renderStyles(cell) {
        const id = cell.dataset.id;
        const style = this.cellStyles[id];
        if (style) {
            Object.assign(cell.style, style);
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
            // Enter edit mode for IME without overwrite selection.
            this.prepareEnterMode(cell, false);
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
        const query = this.findInput ? this.findInput.value : '';
        const replace = this.replaceInput ? this.replaceInput.value : '';
        const matchCase = this.findCase ? this.findCase.checked : false;
        const exact = this.findExact ? this.findExact.checked : false;

        this.findState.query = query;
        this.findState.replace = replace;
        this.findState.matchCase = matchCase;
        this.findState.exact = exact;

        this.clearFindHighlights();

        if (!query) {
            this.findState.matches = [];
            this.findState.currentIndex = -1;
            return;
        }

        const target = matchCase ? query : query.toLowerCase();
        const matches = [];

        for (const key in this.data) {
            const rawValue = this.getRawValue(key);
            if (!rawValue) continue;
            const hay = matchCase ? rawValue : rawValue.toLowerCase();
            const isMatch = exact ? hay === target : hay.includes(target);
            if (isMatch) matches.push(key);
        }

        matches.sort((a, b) => {
            const pa = this.parseCellId(a);
            const pb = this.parseCellId(b);
            if (pa.row !== pb.row) return pa.row - pb.row;
            return pa.colNum - pb.colNum;
        });

        this.findState.matches = matches;
        if (this.findState.currentIndex >= matches.length) {
            this.findState.currentIndex = matches.length - 1;
        }
        if (matches.length === 0) {
            this.findState.currentIndex = -1;
        }

        this.applyFindHighlights();
    }

    refreshFindIfActive() {
        if (this.findState.query) {
            this.updateFindResults();
        }
    }

    clearFindHighlights() {
        document.querySelectorAll('.cell.match').forEach(cell => cell.classList.remove('match'));
    }

    applyFindHighlights() {
        this.findState.matches.forEach((cellId) => {
            const cell = this.getCellEl(cellId.match(/[A-Z]+/)[0], parseInt(cellId.match(/\d+/)[0]));
            if (cell) cell.classList.add('match');
        });
    }

    selectFindMatch(index) {
        const cellId = this.findState.matches[index];
        if (!cellId) return;
        const { colNum, row } = this.parseCellId(cellId);
        const cell = this.getCellEl(colNum, row);
        if (!cell) return;

        this.clearRangeSelection();
        this.setSelectionRange(colNum, row, colNum, row);
        cell.focus({ preventScroll: true });
        this.handleCellFocus(cell);
        this.updateRangeVisual();
        this.updateFillHandlePosition();
    }

    gotoFindMatch(direction) {
        const total = this.findState.matches.length;
        if (total === 0) return;

        let nextIndex = this.findState.currentIndex + direction;
        if (nextIndex < 0) nextIndex = total - 1;
        if (nextIndex >= total) nextIndex = 0;

        this.findState.currentIndex = nextIndex;
        this.selectFindMatch(nextIndex);
    }

    replaceCurrentMatch() {
        const idx = this.findState.currentIndex;
        if (idx < 0) return;
        const cellId = this.findState.matches[idx];
        if (!cellId) return;

        const rawValue = this.getRawValue(cellId);
        const updated = this.replaceInValue(rawValue);
        this.setRawValue(cellId, updated);
        this.renderCellById(cellId);
        this.markDirty();
        this.updateItemCount();
        this.updateFindResults();
    }

    replaceAllMatches() {
        if (this.findState.matches.length === 0) return;
        this.findState.matches.forEach((cellId) => {
            const rawValue = this.getRawValue(cellId);
            const updated = this.replaceInValue(rawValue);
            this.setRawValue(cellId, updated);
            this.renderCellById(cellId);
        });
        this.markDirty();
        this.updateItemCount();
        this.updateFindResults();
    }

    replaceInValue(rawValue) {
        const query = this.findState.query || '';
        if (!query) return rawValue;

        const replaceValue = this.findState.replace || '';
        if (this.findState.exact) {
            const a = this.findState.matchCase ? rawValue : rawValue.toLowerCase();
            const b = this.findState.matchCase ? query : query.toLowerCase();
            return a === b ? replaceValue : rawValue;
        }

        const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const flags = this.findState.matchCase ? 'g' : 'gi';
        const re = new RegExp(escaped, flags);
        return rawValue.replace(re, replaceValue);
    }

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
        const focusCell = this.getCellEl(focusCol, focusRow);
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

        const cellId = cell.dataset.id;
        const { col, row, colNum } = this.parseCellId(cellId);

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
                const cellId = `${this.numberToCol(c)}${r}`;
                rowVals.push(this.getRawValue(cellId));
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
                        this.setRawValue(cell.dataset.id, value);
                        this.renderCellValue(cell);
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
                // Avoid enter-mode here; let IME composition start on the focused cell.
                return;
            }
            this.prepareEnterMode(activeCell, true);
            // DO NOT e.preventDefault() -> Let the browser insert the first char
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
            const nextCell = this.getCellEl(nextCol, nextRow);
            if (nextCell) {
                nextCell.focus({ preventScroll: true });
                this.handleCellFocus(nextCell);
                this.updateFillHandlePosition();
            }
        }
    }

    // ─── Mode Handlers ─────────────────────────────────────
    enterEditMode(cell) {
        if (this.isEditing) return;
        const cellId = cell.dataset.id;
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
        this.isEditing = true;
        this.originalValue = this.getRawValue(cellId);
        this.needsOverwrite = overwrite;
        
        cell.classList.add('editing');
        cell.innerText = this.originalValue;
        if (overwrite) {
            // Select all text in the cell so the next keystroke replaces it (Overwrite behavior)
            const range = document.createRange();
            const sel = window.getSelection();
            range.selectNodeContents(cell);
            sel.removeAllRanges();
            sel.addRange(range);
        } else {
            // Place caret at end for IME-friendly composition
            const range = document.createRange();
            const sel = window.getSelection();
            range.selectNodeContents(cell);
            range.collapse(false);
            sel.removeAllRanges();
            sel.addRange(range);
        }
        
        this.markDirty();
    }

    enterEnterMode(cell) {
        this.prepareEnterMode(cell);
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
        
        const nextCell = this.getCellEl(nextCol, nextRow);
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
        const targetCell = this.getCellEl(targetCol, targetRow);
        if (targetCell) {
            targetCell.focus();
            this.selectedCell = targetCell;
            this.cellAddress.innerText = targetCell.dataset.id;
            this.formulaInput.value = this.getRawValue(targetCell.dataset.id);
        }
    }

    // ─── File Import / Export ───────────────────────────────
    async openFileDialog() {
        if ('showOpenFilePicker' in window) {
            try {
                const [handle] = await window.showOpenFilePicker({
                    types: [
                        {
                            description: 'VibrantSheets / Excel Files',
                            accept: {
                                'application/json': ['.vsht'],
                                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                                'text/csv': ['.csv', '.tsv', '.txt']
                            }
                        }
                    ],
                    multiple: false
                });
                const file = await handle.getFile();
                this.processFile(file, handle);
            } catch (err) {
                if (err.name === 'AbortError') return;
                console.error('File Picker failed, falling back:', err);
                this.fileInput.click();
            }
        } else {
            this.fileInput.click();
        }
    }

    handleFileSelect(e) {
        const file = e.target.files[0];
        if (!file) return;
        this.processFile(file, null);
    }

    processFile(file, handle = null) {
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
        const doc = JSON.parse(jsonText);
        
        // Clear current
        this.clearAllData(false);

        // Restore Metadata
        if (doc.sheets && Array.isArray(doc.sheets) && doc.sheets.length > 0) {
            this.sheets = doc.sheets.map((sheet, idx) => ({
                name: sheet.name || `Sheet${idx + 1}`,
                rows: sheet.rows || this.baseRows,
                cols: sheet.cols || this.baseCols,
                data: sheet.data || {},
                cellStyles: sheet.cellStyles || {},
                cellFormats: sheet.cellFormats || {},
                colWidths: sheet.colWidths || new Array(this.baseCols).fill(100),
                rowHeights: sheet.rowHeights || new Array(this.baseRows + 1).fill(25)
            }));
            this.activeSheetIndex = Math.min(Math.max(doc.activeSheetIndex || 0, 0), this.sheets.length - 1);
        } else {
            if (doc.colWidths) this.colWidths = doc.colWidths;
            if (doc.rowHeights) this.rowHeights = doc.rowHeights;
            this.data = doc.data || {};
            this.cellStyles = doc.cellStyles || {};
            this.cellFormats = doc.cellFormats || {};
        }

        // Re-render or Refresh UI
        this.refreshGridUI();
        this.renderSheetTabs();
        
        this.updateItemCount();
        this.markClean();
    }

    // 2. .xlsx Import (using SheetJS)
    async importXLSX(buffer) {
        return this.importXLSXExcelJS(buffer);
    }

    async importXLSXExcelJS(buffer) {
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
                        raw = '=' + cell.formula;
                    } else if (cell.type === ExcelJS.ValueType.Date && cell.value instanceof Date) {
                        const yyyy = cell.value.getFullYear();
                        const mm = String(cell.value.getMonth() + 1).padStart(2, '0');
                        const dd = String(cell.value.getDate()).padStart(2, '0');
                        raw = `${yyyy}-${mm}-${dd}`;
                    } else {
                        raw = cell.value === null || cell.value === undefined ? '' : String(cell.value);
                    }

                    this.setRawValueForSheet(sheet, cellId, raw);
                    if (cell.numFmt) {
                        sheet.cellFormats[cellId] = this.getExceljsFormatFromNumFmt(cell.numFmt, cell.type);
                    }
                    if (cell.font || cell.alignment || cell.fill) {
                        this.mapExceljsStyleToSheet(sheet, cellId, cell);
                    }
                });
            });

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
        const firstLine = text.split('\n')[0] || '';
        const commas = (firstLine.match(/,/g) || []).length;
        const tabs = (firstLine.match(/\t/g) || []).length;
        const semis = (firstLine.match(/;/g) || []).length;

        if (tabs >= commas && tabs >= semis && tabs > 0) return '\t';
        if (semis > commas && semis > 0) return ';';
        return ',';
    }

    async saveFile() {
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
                cellFormats: sheet.cellFormats
            })),
            activeSheetIndex: this.activeSheetIndex
        };
    }

    getUsedRangeForSheet(sheet) {
        let maxRow = 0;
        let maxCol = 0;
        const scan = (key) => {
            const { colNum, row } = this.parseCellId(key);
            maxRow = Math.max(maxRow, row);
            maxCol = Math.max(maxCol, colNum);
        };
        Object.keys(sheet.data || {}).forEach(scan);
        Object.keys(sheet.cellStyles || {}).forEach(scan);
        Object.keys(sheet.cellFormats || {}).forEach(scan);
        return { maxRow, maxCol };
    }

    async generateXLSXBufferExcelJS() {
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
        });

        return wb.xlsx.writeBuffer();
    }

    generateXLSXBuffer() {
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
        if (shouldMarkDirty) this.markDirty();
        this.updateItemCount();
        this.refreshFindIfActive();
    }

    updateItemCount() {
        const count = Object.keys(this.data).filter(k => this.data[k] && this.data[k].trim() !== '').length;
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

        // Helper to shift a single coordinate
        const shiftCoord = (coord, t, d) => (coord >= t ? coord + d : coord);

        // Process data
        for (const key in this.data) {
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

        this.data = newData;
        this.cellStyles = newStyles;
        this.cellFormats = newFormats;
    }
}

// Initialize immediately (script is at bottom of body)
try {
    window.sheets = new VibrantSheets();
} catch (err) {
    console.error('VibrantSheets initialization failed:', err);
}
