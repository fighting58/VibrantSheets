/* File IO module for VibrantSheets */
(() => {
    const VSIO = {
        async confirmCsvSingleSheet(ctx) {
            if (ctx.sheets.length <= 1) return true;
            if (ctx.csvConfirmInProgress) return false;
            ctx.csvConfirmInProgress = true;
            const proceed = await ctx.showConfirmAsync(
                'CSV는 활성 시트만 저장할 수 있습니다. 전체 시트를 저장하려면 다른 형식을 선택하세요. 계속 저장할까요?'
            );
            ctx.csvConfirmInProgress = false;
            return proceed;
        },

        async confirmXlsxStyleWarning(ctx) {
            return true;
        },

        async openFileDialog(ctx) {
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
                    VSIO.processFile(ctx, file, handle);
                } catch (err) {
                    if (err.name === 'AbortError') return;
                    console.error('File Picker failed, falling back:', err);
                    ctx.fileInput.click();
                }
            } else {
                ctx.fileInput.click();
            }
        },

        handleFileSelect(ctx, e) {
            const file = e.target.files[0];
            if (!file) return;
            VSIO.processFile(ctx, file, null);
        },

        processFile(ctx, file, handle = null) {
            const hasData = Object.keys(ctx.data).some(k => ctx.data[k] && ctx.data[k].trim() !== '');
            if (hasData) {
                if (!confirm('작업 중인 내용을 덮어씁니다. 계속할까요?\n(Unsaved changes will be lost. Continue?)')) {
                    return;
                }
            }

            ctx.fileHandle = handle;
            const extension = file.name.split('.').pop().toLowerCase();
            const reader = new FileReader();

            reader.onload = async (event) => {
                try {
                    if (extension === 'xlsx' || extension === 'xls') {
                        await VSIO.importXLSX(ctx, event.target.result);
                    } else if (extension === 'vsht') {
                        VSIO.importVSHT(ctx, event.target.result);
                    } else {
                        VSIO.importFromText(ctx, event.target.result);
                    }

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
        },

        importVSHT(ctx, jsonText) {
            const doc = JSON.parse(jsonText);
            ctx.clearAllData(false);

            if (doc.sheets && Array.isArray(doc.sheets) && doc.sheets.length > 0) {
                ctx.sheets = doc.sheets.map((sheet, idx) => ({
                    name: sheet.name || `Sheet${idx + 1}`,
                    rows: sheet.rows || ctx.baseRows,
                    cols: sheet.cols || ctx.baseCols,
                    data: sheet.data || {},
                    cellStyles: sheet.cellStyles || {},
                    cellFormats: sheet.cellFormats || {},
                    cellFormulas: sheet.cellFormulas || {},
                    cellBorders: sheet.cellBorders || {},
                    mergedRanges: sheet.mergedRanges || [],
                    images: sheet.images || [],
                    printSettings: sheet.printSettings || ctx.defaultPrintSettings(),
                    colWidths: sheet.colWidths || new Array(ctx.baseCols).fill(ctx.baseColWidth),
                    rowHeights: sheet.rowHeights || new Array(ctx.baseRows + 1).fill(ctx.baseRowHeight)
                }));
                ctx.activeSheetIndex = Math.min(Math.max(doc.activeSheetIndex || 0, 0), ctx.sheets.length - 1);
            } else {
                if (doc.colWidths) ctx.colWidths = doc.colWidths;
                if (doc.rowHeights) ctx.rowHeights = doc.rowHeights;
                ctx.data = doc.data || {};
                ctx.cellStyles = doc.cellStyles || {};
                ctx.cellFormats = doc.cellFormats || {};
                ctx.cellFormulas = doc.cellFormulas || {};
                ctx.cellBorders = doc.cellBorders || {};
                ctx.mergedRanges = doc.mergedRanges || [];
                ctx.activeSheet.images = doc.images || [];
                if (!ctx.activeSheet.printSettings) {
                    ctx.activeSheet.printSettings = ctx.defaultPrintSettings();
                }
            }

            ctx.refreshGridUI();
            ctx.renderSheetTabs();
            ctx.updateItemCount();
            if (typeof ctx.normalizeDefaultDimensions === 'function') {
                ctx.normalizeDefaultDimensions();
            }
            ctx.markClean();
        },

        async importXLSX(ctx, buffer) {
            return VSIO.importXLSXExcelJS(ctx, buffer);
        },

        async importXLSXExcelJS(ctx, buffer) {
            if (typeof ExcelJS === 'undefined') {
                alert('Excel 라이브러리를 불러오지 못했습니다. 네트워크 연결을 확인하세요.');
                return;
            }

            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);

            const normalizeCellValue = (value) => {
                if (value === null || value === undefined) return '';
                if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
                    return String(value);
                }
                if (value instanceof Date) return value;
                if (typeof value === 'object') {
                    if (Array.isArray(value.richText)) {
                        return value.richText.map((part) => String(part && part.text !== undefined ? part.text : '')).join('');
                    }
                    if (value.text !== undefined && value.text !== null) {
                        return String(value.text);
                    }
                    if (value.result !== undefined && value.result !== null) {
                        return String(value.result);
                    }
                }
                return String(value);
            };

            const emuToPx = (v) => Math.round(Number(v || 0) / 9525);
            const normalizeOffset = (v) => {
                if (v === null || v === undefined) return 0;
                const n = Number(v) || 0;
                return n > 10000 ? emuToPx(n) : Math.round(n);
            };
            const normalizeIndex = (v) => {
                const n = Number(v);
                if (!Number.isFinite(n)) return 1;
                if (n <= 0) return 1;
                return Math.round(n);
            };
            const getAnchor = (pt) => {
                if (!pt) return { col: 1, row: 1, offsetX: 0, offsetY: 0 };
                let col = pt.col !== undefined ? pt.col : (pt.nativeCol !== undefined ? pt.nativeCol + 1 : 1);
                let row = pt.row !== undefined ? pt.row : (pt.nativeRow !== undefined ? pt.nativeRow + 1 : 1);
                col = normalizeIndex(col);
                row = normalizeIndex(row);
                const offsetX = normalizeOffset(pt.offsetX ?? (pt.offset ? pt.offset.x : undefined) ?? pt.nativeColOff);
                const offsetY = normalizeOffset(pt.offsetY ?? (pt.offset ? pt.offset.y : undefined) ?? pt.nativeRowOff);
                return { col, row, offsetX, offsetY };
            };
            const sumColWidths = (colWidths, col) => {
                let sum = 0;
                for (let i = 1; i < col; i++) sum += colWidths[i - 1] || ctx.baseColWidth;
                return sum;
            };
            const sumRowHeights = (rowHeights, row) => {
                let sum = 0;
                for (let i = 1; i < row; i++) sum += rowHeights[i] || ctx.baseRowHeight;
                return sum;
            };
            const rangeToPixels = (range, sheet) => {
                if (!range || !sheet) return null;
                const headerRowHeight = sheet.rowHeights?.[1] || ctx.baseRowHeight;
                const tl = getAnchor(range.tl || range);
                const br = getAnchor(range.br || range);
                const startX = ctx.rowHeaderWidth + sumColWidths(sheet.colWidths, tl.col) + tl.offsetX;
                const startY = headerRowHeight + sumRowHeights(sheet.rowHeights, tl.row) + tl.offsetY;
                let width = 0;
                let height = 0;
                if (range.ext) {
                    const w = range.ext.width ?? range.ext.cx;
                    const h = range.ext.height ?? range.ext.cy;
                    width = normalizeOffset(w);
                    height = normalizeOffset(h);
                } else if (range.br || (range.tl && range.br)) {
                    const endX = ctx.rowHeaderWidth + sumColWidths(sheet.colWidths, br.col) + br.offsetX;
                    const endY = headerRowHeight + sumRowHeights(sheet.rowHeights, br.row) + br.offsetY;
                    width = Math.max(1, Math.round(endX - startX));
                    height = Math.max(1, Math.round(endY - startY));
                } else {
                    width = 1;
                    height = 1;
                }
                return { x: Math.round(startX), y: Math.round(startY), w: Math.max(1, Math.round(width)), h: Math.max(1, Math.round(height)) };
            };
            const bufferToBase64 = (buffer) => {
                const bytes = new Uint8Array(buffer);
                let binary = '';
                for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
                return btoa(binary);
            };

            ctx.sheets = [];
            workbook.eachSheet((worksheet, sheetIndex) => {
                const name = worksheet.name || `Sheet${sheetIndex}`;
                const sheet = ctx.createSheet(name);

                sheet.rows = Math.max(ctx.baseRows, worksheet.rowCount || 0);
                sheet.cols = Math.max(ctx.baseCols, worksheet.columnCount || 0);
                sheet.colWidths = new Array(sheet.cols).fill(ctx.baseColWidth);
                sheet.rowHeights = new Array(sheet.rows + 1).fill(ctx.baseRowHeight);

                // Column widths / row heights from Excel (approx px conversion)
                for (let c = 1; c <= sheet.cols; c++) {
                    const col = worksheet.getColumn(c);
                    if (col && col.width) {
                        // Excel width is roughly in character units; convert to px.
                        sheet.colWidths[c - 1] = Math.max(30, Math.round(col.width * 7.105263 + 2.094));
                    }
                }
                for (let r = 1; r <= sheet.rows; r++) {
                    const row = worksheet.getRow(r);
                    if (row && row.height) {
                        sheet.rowHeights[r] = Math.max(15, Math.round(row.height / 0.75));
                    }
                }

                worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                        const cellId = `${ctx.numberToCol(colNumber)}${rowNumber}`;
                        let raw = '';
                        const hasStyle = !!(cell.font || cell.alignment || cell.fill || cell.border);
                        const hasNumFmt = !!cell.numFmt;
                        const hasValue = cell.value !== null && cell.value !== undefined && cell.value !== '';
                        if (!hasValue && !hasStyle && !hasNumFmt) {
                            return;
                        }

                        if (cell.type === ExcelJS.ValueType.Formula) {
                            ctx.setRawValueForSheet(sheet, cellId, '=' + cell.formula);
                        } else if (cell.type === ExcelJS.ValueType.Date && cell.value instanceof Date) {
                            const yyyy = cell.value.getFullYear();
                            const mm = String(cell.value.getMonth() + 1).padStart(2, '0');
                            const dd = String(cell.value.getDate()).padStart(2, '0');
                            raw = `${yyyy}-${mm}-${dd}`;
                            ctx.setRawValueForSheet(sheet, cellId, raw);
                        } else if (hasValue) {
                            const normalized = normalizeCellValue(cell.value);
                            if (normalized instanceof Date) {
                                const yyyy = normalized.getFullYear();
                                const mm = String(normalized.getMonth() + 1).padStart(2, '0');
                                const dd = String(normalized.getDate()).padStart(2, '0');
                                raw = `${yyyy}-${mm}-${dd}`;
                            } else {
                                raw = String(normalized);
                            }
                            ctx.setRawValueForSheet(sheet, cellId, raw);
                        }
                        if (cell.numFmt) {
                            sheet.cellFormats[cellId] = ctx.getExceljsFormatFromNumFmt(cell.numFmt, cell.type);
                        }
                        if (cell.font || cell.alignment || cell.fill || cell.border) {
                            ctx.mapExceljsStyleToSheet(sheet, cellId, cell);
                        }
                    });
                });

                const normalizeMergeEntry = (entry) => {
                    if (!entry) return null;
                    if (typeof entry === 'string') return ctx.parseRangeRef(entry);
                    if (entry.range) return ctx.parseRangeRef(entry.range);
                    if (entry.address) return ctx.parseRangeRef(entry.address);
                    if (entry.model && entry.model.top !== undefined) {
                        const m = entry.model;
                        return {
                            startCol: m.left,
                            startRow: m.top,
                            endCol: m.right,
                            endRow: m.bottom
                        };
                    }
                    if (entry.top !== undefined && entry.left !== undefined) {
                        return {
                            startCol: entry.left,
                            startRow: entry.top,
                            endCol: entry.right,
                            endRow: entry.bottom
                        };
                    }
                    if (entry.tl && entry.br) {
                        return {
                            startCol: entry.tl.col,
                            startRow: entry.tl.row,
                            endCol: entry.br.col,
                            endRow: entry.br.row
                        };
                    }
                    return null;
                };

                const mergeList = [];
                if (worksheet.model && Array.isArray(worksheet.model.merges)) {
                    worksheet.model.merges.forEach(ref => mergeList.push(ref));
                }
                if (worksheet._merges) {
                    if (worksheet._merges instanceof Map) {
                        Array.from(worksheet._merges.values()).forEach(ref => mergeList.push(ref));
                    } else if (Array.isArray(worksheet._merges)) {
                        worksheet._merges.forEach(ref => mergeList.push(ref));
                    } else {
                        Object.keys(worksheet._merges).forEach(key => mergeList.push(worksheet._merges[key]));
                    }
                }

                const dedupe = new Set();
                sheet.mergedRanges = mergeList
                    .map(normalizeMergeEntry)
                    .filter(Boolean)
                    .filter((m) => {
                        const key = `${m.startCol},${m.startRow},${m.endCol},${m.endRow}`;
                        if (dedupe.has(key)) return false;
                        dedupe.add(key);
                        return true;
                    });

                // Ensure grid size covers merged ranges
                let maxMergeRow = 0;
                let maxMergeCol = 0;
                sheet.mergedRanges.forEach((range) => {
                    const normalized = ctx.normalizeMergedRangeEntry(range);
                    if (!normalized) return;
                    maxMergeRow = Math.max(maxMergeRow, normalized.endRow);
                    maxMergeCol = Math.max(maxMergeCol, normalized.endCol);
                });
                if (maxMergeRow > sheet.rows) {
                    for (let r = sheet.rows + 1; r <= maxMergeRow; r++) {
                        sheet.rowHeights[r] = ctx.baseRowHeight;
                    }
                    sheet.rows = maxMergeRow;
                }
                if (maxMergeCol > sheet.cols) {
                    for (let c = sheet.cols + 1; c <= maxMergeCol; c++) {
                        sheet.colWidths[c - 1] = ctx.baseColWidth;
                    }
                    sheet.cols = maxMergeCol;
                }

                if (sheet.printSettings == null) {
                    sheet.printSettings = ctx.defaultPrintSettings();
                }

                // Import images from ExcelJS worksheet
                if (typeof worksheet.getImages === 'function') {
                    const imageEntries = worksheet.getImages() || [];
                    if (imageEntries.length) {
                        sheet.images = [];
                        imageEntries.forEach((entry) => {
                            const info = workbook.getImage(entry.imageId);
                            if (!info) return;
                            const ext = String(info.extension || 'png').toLowerCase();
                            let src = '';
                            if (info.base64) {
                                src = `data:image/${ext};base64,${info.base64}`;
                            } else if (info.buffer) {
                                src = `data:image/${ext};base64,${bufferToBase64(info.buffer)}`;
                            }
                            if (!src) return;
                            const range = entry.range || entry;
                            const px = rangeToPixels(range, sheet);
                            if (!px) return;
                            const img = {
                                id: `img_${Date.now()}_${Math.floor(Math.random() * 10000)}`,
                                src,
                                x: px.x,
                                y: px.y,
                                w: px.w,
                                h: px.h,
                                name: info.name || `image.${ext}`,
                                z: sheet.images.length,
                                locked: false
                            };
                            if (typeof ctx.calcImageAnchorFromPixels === 'function') {
                                img.anchor = ctx.calcImageAnchorFromPixels(img, sheet);
                            }
                            sheet.images.push(img);
                        });
                    }
                }

                ctx.sheets.push(sheet);
            });

            if (ctx.sheets.length === 0) {
                ctx.sheets = [ctx.createSheet('Sheet1')];
            }
            ctx.activeSheetIndex = 0;
            ctx.refreshGridUI();
            ctx.renderSheetTabs();
            ctx.updateItemCount();
            ctx.markClean();
        },

        importFromText(ctx, text) {
            ctx.clearAllData(false);

            if (text.charCodeAt(0) === 0xFEFF) {
                text = text.substring(1);
            }

            const delimiter = VSIO.detectDelimiter(ctx, text);

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
                            i++;
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
                        VSIO.setInternalData(ctx, row + 1, col + 1, currentField);
                        currentField = '';
                        col++;
                    } else if (ch === '\r' && nextCh === '\n') {
                        VSIO.setInternalData(ctx, row + 1, col + 1, currentField);
                        currentField = '';
                        row++;
                        col = 0;
                        i++;
                    } else if (ch === '\n' || ch === '\r') {
                        VSIO.setInternalData(ctx, row + 1, col + 1, currentField);
                        currentField = '';
                        row++;
                        col = 0;
                    } else {
                        currentField += ch;
                    }
                }
            }

            if (currentField !== '' || col > 0) {
                VSIO.setInternalData(ctx, row + 1, col + 1, currentField);
            }

            ctx.refreshGridUI();
            ctx.updateItemCount();
            ctx.markClean();
        },

        setInternalData(ctx, rowNum, colNum, value) {
            if (rowNum > ctx.rows) {
                ctx.createRowElements(ctx.rows + 1, rowNum);
                ctx.rows = rowNum;
            }
            if (colNum > ctx.cols) {
                for (let c = ctx.cols + 1; c <= colNum; c++) {
                    ctx.colWidths[c - 1] = ctx.baseColWidth;
                }
                ctx.cols = colNum;
            }
            const cellId = `${ctx.numberToCol(colNum)}${rowNum}`;
            ctx.setRawValue(cellId, value);
        },

        detectDelimiter(ctx, text) {
            const firstLine = text.split('\n')[0] || '';
            const commas = (firstLine.match(/,/g) || []).length;
            const tabs = (firstLine.match(/\t/g) || []).length;
            const semis = (firstLine.match(/;/g) || []).length;

            if (tabs >= commas && tabs >= semis && tabs > 0) return '\t';
            if (semis > commas && semis > 0) return ';';
            return ',';
        },

        async saveFile(ctx) {
            if (!ctx.fileHandle) {
                return VSIO.saveFileAs(ctx);
            }

            try {
                const fileName = (ctx.fileHandle && ctx.fileHandle.name) ? String(ctx.fileHandle.name) : '';
                const lowerName = fileName.toLowerCase();
                if (lowerName.endsWith('.csv')) {
                    if (!await VSIO.confirmCsvSingleSheet(ctx)) return;
                }

                const writable = await ctx.fileHandle.createWritable();

                if (lowerName.endsWith('.xlsx')) {
                    if (!await VSIO.confirmXlsxStyleWarning(ctx)) return;
                    const buffer = await VSIO.generateXLSXBufferExcelJS(ctx);
                    if (buffer) await writable.write(buffer);
                } else if (lowerName.endsWith('.csv')) {
                    const csvContent = VSIO.generateCSVContent(ctx);
                    await writable.write(csvContent);
                } else {
                    const vshtData = VSIO.generateVSHTData(ctx);
                    await writable.write(JSON.stringify(vshtData, null, 2));
                }

                await writable.close();
                ctx.markClean();
            } catch (err) {
                console.error('Save failed, using Save As:', err);
                VSIO.saveFileAs(ctx);
            }
        },

        async saveFileAs(ctx) {
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

                    ctx.fileHandle = handle;
                    await VSIO.saveFile(ctx);

                    const filenameEl = document.querySelector('.filename');
                    if (filenameEl) {
                        filenameEl.innerText = handle.name.replace(/\.[^.]+$/, '');
                    }
                } catch (err) {
                    if (err.name === 'AbortError') return;
                    console.error('Save As failed:', err);
                }
            } else {
                const ext = (defaultName.split('.').pop() || '').toLowerCase();
                if (ext !== 'csv') {
                    const vshtData = VSIO.generateVSHTData(ctx);
                    const blob = new Blob([JSON.stringify(vshtData, null, 2)], { type: 'application/json' });
                    const url = URL.createObjectURL(blob);
                    const link = document.createElement('a');
                    link.href = url;
                    link.download = `${defaultName}.vsht`;
                    link.click();
                    URL.revokeObjectURL(url);
                    ctx.markClean();
                }
            }
        },

        generateVSHTData(ctx) {
            return {
                version: "1.0",
                title: document.querySelector('.filename')?.innerText || 'Untitled',
                sheets: ctx.sheets.map(sheet => ({
                    name: sheet.name,
                    rows: sheet.rows,
                    cols: sheet.cols,
                    data: sheet.data,
                    colWidths: sheet.colWidths,
                    rowHeights: sheet.rowHeights,
                    cellStyles: sheet.cellStyles,
                    cellFormats: sheet.cellFormats,
                    cellFormulas: sheet.cellFormulas,
                    cellBorders: sheet.cellBorders || {},
                    mergedRanges: sheet.mergedRanges || [],
                    images: sheet.images || [],
                    printSettings: sheet.printSettings || ctx.defaultPrintSettings()
                })),
                activeSheetIndex: ctx.activeSheetIndex
            };
        },

        getUsedRangeForSheet(ctx, sheet) {
            let maxRow = 0;
            let maxCol = 0;
            const scan = (key) => {
                const { colNum, row } = ctx.parseCellId(key);
                maxRow = Math.max(maxRow, row);
                maxCol = Math.max(maxCol, colNum);
            };
            Object.keys(sheet.data || {}).forEach(scan);
            Object.keys(sheet.cellFormulas || {}).forEach(scan);
            Object.keys(sheet.cellStyles || {}).forEach(scan);
            Object.keys(sheet.cellFormats || {}).forEach(scan);
            Object.keys(sheet.cellBorders || {}).forEach(scan);
            (sheet.mergedRanges || []).forEach((range) => {
                const normalized = ctx.normalizeMergedRangeEntry(range);
                if (!normalized) return;
                maxRow = Math.max(maxRow, normalized.endRow);
                maxCol = Math.max(maxCol, normalized.endCol);
            });
            return { maxRow, maxCol };
        },

        async generateXLSXBufferExcelJS(ctx) {
            if (typeof ExcelJS === 'undefined') {
                alert('Excel 라이브러리를 불러오지 못했습니다. 네트워크 연결을 확인하세요.');
                return null;
            }

            if (typeof ctx.updateAllImageAnchors === 'function') {
                ctx.updateAllImageAnchors();
            }

            const wb = new ExcelJS.Workbook();
            ctx.sheets.forEach((sheet, index) => {
                const ws = wb.addWorksheet(sheet.name || `Sheet${index + 1}`);
                if (ws.properties) {
                    // Keep worksheet descent metadata in a safe Excel-compatible range.
                    ws.properties.dyDescent = 0.25;
                }
                let { maxRow, maxCol } = VSIO.getUsedRangeForSheet(ctx, sheet);

                const rowHeaderWidth = 40;
                const headerRowHeight = sheet.rowHeights?.[1] || ctx.baseRowHeight;
                const colWidths = sheet.colWidths || [];
                const rowHeights = sheet.rowHeights || [];
                const posFromPixels = (x, y) => {
                    const adjX = Math.max(0, x - rowHeaderWidth);
                    const adjY = Math.max(0, y - headerRowHeight);

                    let col = 1;
                    let xRemain = adjX;
                    for (let i = 0; i < Math.max(colWidths.length, maxCol); i++) {
                        const w = colWidths[i] || ctx.baseColWidth;
                        if (xRemain < w) { col = i + 1; break; }
                        xRemain -= w;
                        col = i + 2;
                    }

                    let row = 1;
                    let yRemain = adjY;
                    for (let i = 1; i < Math.max(rowHeights.length, maxRow + 1); i++) {
                        const h = rowHeights[i] || ctx.baseRowHeight;
                        if (yRemain < h) { row = i; break; }
                        yRemain -= h;
                        row = i + 1;
                    }

                    return { col, row, offsetX: Math.max(0, Math.round(xRemain)), offsetY: Math.max(0, Math.round(yRemain)) };
                };

                const imagesForRange = sheet.images || [];
                imagesForRange.forEach((img) => {
                    if (!img) return;
                    if (img.anchor && img.anchor.startCell) {
                        const parsed = ctx.parseCellId(img.anchor.startCell);
                        maxCol = Math.max(maxCol, parsed.colNum);
                        maxRow = Math.max(maxRow, parsed.row);
                        return;
                    }
                    const end = posFromPixels((Number(img.x) || 0) + (Number(img.w) || 0), (Number(img.y) || 0) + (Number(img.h) || 0));
                    maxCol = Math.max(maxCol, end.col);
                    maxRow = Math.max(maxRow, end.row);
                });

                // Apply column widths / row heights (px -> Excel units)
                for (let c = 1; c <= maxCol; c++) {
                    const px = colWidths[c - 1] || ctx.baseColWidth;
                    // Calibrated linear mapping (px -> Excel width)
                    // Fit using measured pairs: 59->7.38, 73->9.38, 329->45.38
                    const width = Math.max(3, Math.round((px * 0.1407407 - 0.9247 + 0.63) * 100) / 100);
                    ws.getColumn(c).width = width;
                }
                for (let r = 1; r <= maxRow; r++) {
                    const px = rowHeights[r] || ctx.baseRowHeight;
                    // Excel row height uses points. 1pt ≈1.333px.
                    const height = Math.max(10, Math.round(px * 0.75));
                    ws.getRow(r).height = height;
                }

                for (let r = 1; r <= maxRow; r++) {
                    for (let c = 1; c <= maxCol; c++) {
                        const cellId = `${ctx.numberToCol(c)}${r}`;
                        const rawValue = ctx.getRawValueForSheet(sheet, cellId);
                        const rawText = rawValue == null ? '' : String(rawValue);
                        const rawTrim = rawText.trim();
                        const format = ctx.getCellFormatForSheet(sheet, cellId);
                        const numericValue = ctx.parseNumberFromRaw(rawText, format.type);
                        const dateValue = ctx.parseDateFromRaw(rawText);
                        const isStrictNumericLiteral = (() => {
                            if (!rawTrim) return false;
                            const normalized = rawTrim.replace(/,/g, '');
                            return /^[+-]?(?:\d+(?:\.\d+)?|\.\d+)$/.test(normalized);
                        })();
                        const allowNumericByFormat =
                            format.type === 'number' ||
                            format.type === 'currency' ||
                            format.type === 'percentage' ||
                            (format.type === 'general' && isStrictNumericLiteral);

                        const cell = ws.getCell(r, c);
                        if (format.type === 'text') {
                            if (rawText !== '') cell.value = rawText;
                            cell.numFmt = '@';
                        } else if (rawTrim.startsWith('=')) {
                            const formula = rawTrim.slice(1);
                            const result = ctx.formulaEngine
                                ? ctx.formulaEngine.evaluate(rawTrim, ctx.getFormulaContext(), new Set())
                                : undefined;
                            cell.value = { formula, result: result === '#ERROR' ? undefined : result };
                        } else if (format.type === 'date' && dateValue) {
                            cell.value = dateValue;
                            cell.numFmt = 'yyyy-mm-dd';
                        } else if (numericValue !== null && allowNumericByFormat) {
                            cell.value = numericValue;
                            const numFmt = ctx.getNumFmtFromFormat(format);
                            if (numFmt) cell.numFmt = numFmt;
                        } else {
                            if (rawText !== '') cell.value = rawText;
                        }

                        const style = sheet.cellStyles[cellId];
                        if (style) {
                            ctx.mapInternalStyleToExceljs(cell, style);
                        }
                        const border = sheet.cellBorders?.[cellId];
                        if (border) {
                            const toExcel = (b) => {
                                if (!b || !b.style) return null;
                                const argb = ctx.hexToArgb(b.color || '#000000');
                                let style = b.style;
                                if (style === 'solid') {
                                    if (b.width >= 3) style = 'thick';
                                    else if (b.width >= 2) style = 'medium';
                                    else style = 'thin';
                                } else if (style === 'dashed') {
                                    style = b.width >= 2 ? 'mediumDashed' : 'dashed';
                                } else if (style === 'dotted') {
                                    style = 'dotted';
                                }
                                return { style, color: argb ? { argb } : undefined };
                            };
                            cell.border = {
                                top: toExcel(border.top),
                                right: toExcel(border.right),
                                bottom: toExcel(border.bottom),
                                left: toExcel(border.left)
                            };
                        }
                    }
                }

                const merges = ctx.getNormalizedMergedRanges(sheet);
                merges.forEach((range) => {
                    ws.mergeCells(range.startRow, range.startCol, range.endRow, range.endCol);
                });

                // Images (export from VSHT image metadata)
                const images = sheet.images || [];
                if (images.length) {
                    images.forEach((img) => {
                        if (!img || !img.src) return;
                        const match = String(img.src).match(/^data:image\/([^;]+);/i);
                        const ext = match ? match[1].toLowerCase() : 'png';
                        const imageId = wb.addImage({ base64: img.src, extension: ext });
                        if (img.anchor && img.anchor.startCell) {
                            const start = ctx.parseCellId(img.anchor.startCell);
                            const offsetStart = img.anchor.offsetStart || { x: 0, y: 0 };
                            ws.addImage(imageId, {
                                tl: { col: Math.max(0, start.colNum - 1), row: Math.max(0, start.row - 1), offsetX: offsetStart.x || 0, offsetY: offsetStart.y || 0 },
                                ext: {
                                    width: Math.max(1, Math.round(img.w || 1)),
                                    height: Math.max(1, Math.round(img.h || 1))
                                }
                            });
                        } else {
                            const pos = posFromPixels(Number(img.x) || 0, Number(img.y) || 0);
                            ws.addImage(imageId, {
                                tl: { col: Math.max(0, pos.col - 1), row: Math.max(0, pos.row - 1), offsetX: pos.offsetX, offsetY: pos.offsetY },
                                ext: {
                                    width: Math.max(1, Math.round(img.w || 1)),
                                    height: Math.max(1, Math.round(img.h || 1))
                                }
                            });
                        }
                    });
                }
            });

            return wb.xlsx.writeBuffer();
        },

        generateXLSXBuffer(ctx) {
            if (typeof XLSX === 'undefined') {
                alert('Excel 라이브러리를 찾을 수 없습니다.');
                return null;
            }

            const wb = XLSX.utils.book_new();
            ctx.sheets.forEach((sheet, index) => {
                let maxRow = 0, maxCol = 0;
                for (const key in sheet.data) {
                    if (sheet.data[key] && sheet.data[key].trim() !== '') {
                        const { colNum, row } = ctx.parseCellId(key);
                        maxRow = Math.max(maxRow, row);
                        maxCol = Math.max(maxCol, colNum);
                    }
                }

                const aoa = [];
                for (let r = 1; r <= maxRow; r++) {
                    const rowArr = [];
                    for (let c = 1; c <= maxCol; c++) {
                        const cellId = `${ctx.numberToCol(c)}${r}`;
                        const rawValue = ctx.getRawValueForSheet(sheet, cellId);
                        const format = ctx.getCellFormatForSheet(sheet, cellId);
                        const numericValue = ctx.parseNumberFromRaw(rawValue, format.type);
                        const dateValue = ctx.parseDateFromRaw(rawValue);

                        let cellObj = { v: rawValue, t: 's' };
                        if (rawValue !== '') {
                            if (format.type === 'date' && dateValue) {
                                cellObj = { v: ctx.toExcelDateSerial(dateValue), t: 'n', z: 'yyyy-mm-dd' };
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
                            cellObj.s = ctx.mapInternalStyleToXlsx(style);
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
        },

        generateCSVContent(ctx) {
            let maxRow = 0, maxCol = 0;
            for (const key in ctx.data) {
                if (ctx.data[key] && ctx.data[key].trim() !== '') {
                    const { colNum, row } = ctx.parseCellId(key);
                    maxRow = Math.max(maxRow, row);
                    maxCol = Math.max(maxCol, colNum);
                }
            }

            const rows = [];
            for (let r = 1; r <= maxRow; r++) {
                const rowFields = [];
                for (let c = 1; c <= maxCol; c++) {
                    const cellId = `${ctx.numberToCol(c)}${r}`;
                    let val = ctx.data[cellId] || '';
                    if (val.includes(',') || val.includes('"') || val.includes('\n')) {
                        val = '"' + val.replace(/"/g, '""') + '"';
                    }
                    rowFields.push(val);
                }
                rows.push(rowFields.join(','));
            }
            return '\uFEFF' + rows.join('\r\n');
        }
    };

    window.VSIO = VSIO;
})();
