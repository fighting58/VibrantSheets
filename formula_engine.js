class FormulaEngine {
    constructor() {
        this.functions = {
            SUM: (args) => {
                const nums = this.flattenArgs(args).map(a => this.toNumber(a));
                return nums.reduce((acc, n) => acc + n, 0);
            },
            CONCAT: (args) => this.flattenArgs(args).map(a => this.toString(a)).join(''),
            CONCATENATE: (args) => this.flattenArgs(args).map(a => this.toString(a)).join(''),
            LEFT: (args) => {
                const text = this.toString(args[0]);
                const n = this.toNumber(args[1], 1);
                return text.substring(0, Math.max(0, n));
            },
            RIGHT: (args) => {
                const text = this.toString(args[0]);
                const n = this.toNumber(args[1], 1);
                return text.substring(Math.max(0, text.length - n));
            },
            MID: (args) => {
                const text = this.toString(args[0]);
                const start = this.toNumber(args[1], 1) - 1;
                const len = this.toNumber(args[2], 0);
                return text.substring(Math.max(0, start), Math.max(0, start) + Math.max(0, len));
            }
        };
    }

    evaluate(input, context, stack) {
        if (!input || typeof input !== 'string') return input;
        if (!input.startsWith('=')) return input;
        const expr = input.slice(1).trim();
        if (!expr) return '';

        const parsed = this.parseFunctionCall(expr);
        if (!parsed) return '#ERROR';

        const fnName = parsed.name.toUpperCase();
        const fn = this.functions[fnName];
        if (!fn) return '#NAME?';

        const args = parsed.args.map(arg => this.resolveArg(arg, context, stack));
        try {
            return fn(args);
        } catch (err) {
            return '#ERROR';
        }
    }

    parseFunctionCall(expr) {
        const match = expr.match(/^([A-Za-z_][A-Za-z0-9_]*)\((.*)\)$/);
        if (!match) return null;
        const name = match[1];
        const argsStr = match[2].trim();
        const args = this.splitArgs(argsStr);
        return { name, args };
    }

    splitArgs(argsStr) {
        if (argsStr === '') return [];
        const args = [];
        let current = '';
        let inQuotes = false;

        for (let i = 0; i < argsStr.length; i++) {
            const ch = argsStr[i];
            const next = argsStr[i + 1];
            if (ch === '"' && next === '"') {
                current += '"';
                i++;
                continue;
            }
            if (ch === '"') {
                inQuotes = !inQuotes;
                current += ch;
                continue;
            }
            if ((ch === ',' || ch === ';') && !inQuotes) {
                args.push(current.trim());
                current = '';
                continue;
            }
            current += ch;
        }
        if (current.trim() !== '') {
            args.push(current.trim());
        } else if (current === '') {
            args.push('');
        }
        return args;
    }

    resolveArg(raw, context, stack) {
        if (raw === '') return '';
        if (raw.startsWith('"') && raw.endsWith('"')) {
            return raw.slice(1, -1).replace(/""/g, '"');
        }
        if (context && typeof context.getRangeValues === 'function' && this.isRangeRef(raw)) {
            return context.getRangeValues(raw.toUpperCase(), stack);
        }
        if (context && typeof context.getCellValue === 'function' && this.isCellRef(raw)) {
            return context.getCellValue(raw.toUpperCase(), stack);
        }
        return raw;
    }

    isCellRef(token) {
        return /^[A-Za-z]+[0-9]+$/.test(token);
    }

    isRangeRef(token) {
        return /^[A-Za-z]+[0-9]+:[A-Za-z]+[0-9]+$/.test(token);
    }

    flattenArgs(args) {
        const flat = [];
        args.forEach(arg => {
            if (Array.isArray(arg)) {
                arg.forEach(v => flat.push(v));
            } else {
                flat.push(arg);
            }
        });
        return flat;
    }

    toNumber(value, fallback = 0) {
        const n = Number(String(value).replace(/,/g, '').trim());
        return Number.isFinite(n) ? n : fallback;
    }

    toString(value) {
        return value === null || value === undefined ? '' : String(value);
    }
}

window.FormulaEngine = FormulaEngine;
