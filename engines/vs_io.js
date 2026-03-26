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
                    colWidths: sheet.colWidths || new Array(ctx.baseCols).fill(100),
                    rowHeights: sheet.rowHeights || new Array(ctx.baseRows + 1).fill(25)
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
            }

            ctx.refreshGridUI();
            ctx.renderSheetTabs();
            ctx.updateItemCount();
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

            ctx.sheets = [];
            workbook.eachSheet((worksheet, sheetIndex) => {
                const name = worksheet.name || `Sheet${sheetIndex}`;
                const sheet = ctx.createSheet(name);

                sheet.rows = Math.max(ctx.baseRows, worksheet.rowCount || 0);
                sheet.cols = Math.max(ctx.baseCols, worksheet.columnCount || 0);
                sheet.colWidths = new Array(sheet.cols).fill(100);
                sheet.rowHeights = new Array(sheet.rows + 1).fill(25);

                worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                        const cellId = `${ctx.numberToCol(colNumber)}${rowNumber}`;
                        let raw = '';

                        if (cell.type === ExcelJS.ValueType.Formula) {
                            ctx.setRawValueForSheet(sheet, cellId, '=' + cell.formula);
                        } else if (cell.type === ExcelJS.ValueType.Date && cell.value instanceof Date) {
                            const yyyy = cell.value.getFullYear();
                            const mm = String(cell.value.getMonth() + 1).padStart(2, '0');
                            const dd = String(cell.value.getDate()).padStart(2, '0');
                            raw = `${yyyy}-${mm}-${dd}`;
                            ctx.setRawValueForSheet(sheet, cellId, raw);
                        } else {
                            raw = cell.value === null || cell.value === undefined ? '' : String(cell.value);
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
                    .map((ref) => ctx.parseRangeRef(ref))
                    .filter(Boolean);

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
                    mergedRanges: sheet.mergedRanges || []
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

            const wb = new ExcelJS.Workbook();
            ctx.sheets.forEach((sheet, index) => {
                const ws = wb.addWorksheet(sheet.name || `Sheet${index + 1}`);
                const { maxRow, maxCol } = VSIO.getUsedRangeForSheet(ctx, sheet);

                for (let r = 1; r <= maxRow; r++) {
                    for (let c = 1; c <= maxCol; c++) {
                        const cellId = `${ctx.numberToCol(c)}${r}`;
                        const rawValue = ctx.getRawValueForSheet(sheet, cellId);
                        const format = ctx.getCellFormatForSheet(sheet, cellId);
                        const numericValue = ctx.parseNumberFromRaw(rawValue, format.type);
                        const dateValue = ctx.parseDateFromRaw(rawValue);

                        const cell = ws.getCell(r, c);
                        if (format.type === 'text') {
                            cell.value = rawValue;
                            cell.numFmt = '@';
                        } else if (rawValue.trim().startsWith('=')) {
                            const formula = rawValue.trim().slice(1);
                            const result = ctx.formulaEngine
                                ? ctx.formulaEngine.evaluate(rawValue, ctx.getFormulaContext(), new Set())
                                : undefined;
                            cell.value = { formula, result: result === '#ERROR' ? undefined : result };
                        } else if (format.type === 'date' && dateValue) {
                            cell.value = dateValue;
                            cell.numFmt = 'yyyy-mm-dd';
                        } else if (numericValue !== null && format.type !== 'date' && format.type !== 'text') {
                            cell.value = numericValue;
                            const numFmt = ctx.getNumFmtFromFormat(format);
                            if (numFmt) cell.numFmt = numFmt;
                        } else {
                            cell.value = rawValue;
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
                                return { style: b.style, color: argb ? { argb } : undefined };
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
