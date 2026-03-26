/* Find/Replace module for VibrantSheets */
(() => {
    const VSFind = {
        updateFindResults(ctx) {
            const query = ctx.findInput ? ctx.findInput.value : '';
            const replace = ctx.replaceInput ? ctx.replaceInput.value : '';
            const matchCase = ctx.findCase ? ctx.findCase.checked : false;
            const exact = ctx.findExact ? ctx.findExact.checked : false;

            ctx.findState.query = query;
            ctx.findState.replace = replace;
            ctx.findState.matchCase = matchCase;
            ctx.findState.exact = exact;

            VSFind.clearFindHighlights(ctx);

            if (!query) {
                ctx.findState.matches = [];
                ctx.findState.currentIndex = -1;
                return;
            }

            const target = matchCase ? query : query.toLowerCase();
            const matches = [];

            const keys = new Set([
                ...Object.keys(ctx.data),
                ...Object.keys(ctx.cellFormulas)
            ]);
            for (const key of keys) {
                const rawValue = ctx.getRawValue(key);
                if (!rawValue) continue;
                const hay = matchCase ? rawValue : rawValue.toLowerCase();
                const isMatch = exact ? hay === target : hay.includes(target);
                if (isMatch) matches.push(key);
            }

            matches.sort((a, b) => {
                const pa = ctx.parseCellId(a);
                const pb = ctx.parseCellId(b);
                if (pa.row !== pb.row) return pa.row - pb.row;
                return pa.colNum - pb.colNum;
            });

            ctx.findState.matches = matches;
            if (ctx.findState.currentIndex >= matches.length) {
                ctx.findState.currentIndex = matches.length - 1;
            }
            if (matches.length === 0) {
                ctx.findState.currentIndex = -1;
            }

            VSFind.applyFindHighlights(ctx);
        },

        refreshFindIfActive(ctx) {
            if (ctx.findState.query) {
                VSFind.updateFindResults(ctx);
            }
        },

        clearFindHighlights(ctx) {
            document.querySelectorAll('.cell.match').forEach(cell => cell.classList.remove('match'));
        },

        applyFindHighlights(ctx) {
            ctx.findState.matches.forEach((cellId) => {
                const normalizedId = ctx.normalizeMergedCellId(cellId);
                const col = normalizedId.match(/[A-Z]+/)[0];
                const row = parseInt(normalizedId.match(/\d+/)[0]);
                const cell = ctx.getCellEl(col, row);
                if (cell) cell.classList.add('match');
            });
        },

        selectFindMatch(ctx, index) {
            const cellId = ctx.findState.matches[index];
            if (!cellId) return;
            const { colNum, row } = ctx.parseCellId(cellId);
            const cell = ctx.getSelectableCell(colNum, row);
            if (!cell) return;

            ctx.clearRangeSelection();
            ctx.setSelectionRange(colNum, row, colNum, row);
            cell.focus({ preventScroll: true });
            ctx.handleCellFocus(cell);
            ctx.updateRangeVisual();
            ctx.updateFillHandlePosition();
        },

        gotoFindMatch(ctx, direction) {
            const total = ctx.findState.matches.length;
            if (total === 0) return;

            let nextIndex = ctx.findState.currentIndex + direction;
            if (nextIndex < 0) nextIndex = total - 1;
            if (nextIndex >= total) nextIndex = 0;

            ctx.findState.currentIndex = nextIndex;
            VSFind.selectFindMatch(ctx, nextIndex);
        },

        replaceCurrentMatch(ctx) {
            const idx = ctx.findState.currentIndex;
            if (idx < 0) return;
            const cellId = ctx.findState.matches[idx];
            if (!cellId) return;

            const rawValue = ctx.getRawValue(cellId);
            const updated = VSFind.replaceInValue(ctx, rawValue);
            ctx.setRawValue(cellId, updated);
            ctx.renderCellById(cellId);
            ctx.markDirty();
            ctx.updateItemCount();
            VSFind.updateFindResults(ctx);
        },

        replaceAllMatches(ctx) {
            if (ctx.findState.matches.length === 0) return;
            ctx.findState.matches.forEach((cellId) => {
                const rawValue = ctx.getRawValue(cellId);
                const updated = VSFind.replaceInValue(ctx, rawValue);
                ctx.setRawValue(cellId, updated);
                ctx.renderCellById(cellId);
            });
            ctx.markDirty();
            ctx.updateItemCount();
            VSFind.updateFindResults(ctx);
        },

        replaceInValue(ctx, rawValue) {
            const query = ctx.findState.query || '';
            if (!query) return rawValue;

            const replaceValue = ctx.findState.replace || '';
            if (ctx.findState.exact) {
                const a = ctx.findState.matchCase ? rawValue : rawValue.toLowerCase();
                const b = ctx.findState.matchCase ? query : query.toLowerCase();
                return a === b ? replaceValue : rawValue;
            }

            const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const flags = ctx.findState.matchCase ? 'g' : 'gi';
            const re = new RegExp(escaped, flags);
            return rawValue.replace(re, replaceValue);
        }
    };

    window.VSFind = VSFind;
})();
