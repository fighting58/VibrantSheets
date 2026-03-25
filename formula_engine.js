class FormulaEngine {
    constructor() {
        this.functions = {
            SUM: (args) => {
                const nums = this.flattenArgs(args).map(a => this.toNumber(a));
                return nums.reduce((acc, n) => acc + n, 0);
            },
            AVG: (args) => {
                const nums = this.flattenArgs(args).map(a => this.toNumber(a)).filter(n => Number.isFinite(n));
                if (nums.length === 0) return 0;
                return nums.reduce((acc, n) => acc + n, 0) / nums.length;
            },
            COUNT: (args) => {
                const nums = this.flattenArgs(args).map(a => this.toNumber(a, NaN)).filter(n => Number.isFinite(n));
                return nums.length;
            },
            MIN: (args) => {
                const nums = this.flattenArgs(args).map(a => this.toNumber(a, NaN)).filter(n => Number.isFinite(n));
                return nums.length === 0 ? 0 : Math.min(...nums);
            },
            MAX: (args) => {
                const nums = this.flattenArgs(args).map(a => this.toNumber(a, NaN)).filter(n => Number.isFinite(n));
                return nums.length === 0 ? 0 : Math.max(...nums);
            },
            IF: (args) => {
                const cond = this.toBoolean(args[0]);
                return cond ? args[1] : (args.length >= 3 ? args[2] : '');
            },
            AND: (args) => this.flattenArgs(args).every(a => this.toBoolean(a)) ? 'TRUE' : 'FALSE',
            OR: (args) => this.flattenArgs(args).some(a => this.toBoolean(a)) ? 'TRUE' : 'FALSE',
            NOT: (args) => this.toBoolean(args[0]) ? 'FALSE' : 'TRUE',
            LEN: (args) => this.toString(args[0]).length,
            LOWER: (args) => this.toString(args[0]).toLowerCase(),
            UPPER: (args) => this.toString(args[0]).toUpperCase(),
            TRIM: (args) => this.toString(args[0]).trim(),
            ROUND: (args) => {
                const num = this.toNumber(args[0], 0);
                const digits = this.toNumber(args[1], 0);
                const factor = Math.pow(10, Math.max(0, digits));
                return Math.round(num * factor) / factor;
            },
            ABS: (args) => Math.abs(this.toNumber(args[0], 0)),
            VALUE: (args) => {
                const n = Number(String(args[0]).replace(/,/g, '').trim());
                if (!Number.isFinite(n)) throw this.makeError('#VALUE!');
                return n;
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

        try {
            const tokens = this.tokenize(expr);
            const parser = new Parser(tokens);
            const ast = parser.parseExpression();
            if (!parser.isAtEnd()) return '#ERROR';
            return this.evalNode(ast, context, stack);
        } catch (err) {
            if (err && err.code) return err.code;
            return '#ERROR';
        }
    }

    evalNode(node, context, stack) {
        switch (node.type) {
            case 'Number':
                return node.value;
            case 'String':
                return node.value;
            case 'Unary': {
                const raw = this.evalNode(node.expr, context, stack);
                const err = this.findError(raw);
                if (err) return err;
                const val = this.toNumber(raw);
                return node.op === '-' ? -val : val;
            }
            case 'Binary': {
                const left = this.evalNode(node.left, context, stack);
                const right = this.evalNode(node.right, context, stack);
                const err = this.findError(left) || this.findError(right);
                if (err) return err;
                if (node.op === '&') return this.toString(left) + this.toString(right);
                if (node.op === '=' || node.op === '<>' || node.op === '<' || node.op === '<=' || node.op === '>' || node.op === '>=') {
                    const a = this.coerceComparable(left);
                    const b = this.coerceComparable(right);
                    const res = this.compare(a, b, node.op);
                    return res ? 'TRUE' : 'FALSE';
                }
                const a = this.toNumber(left);
                const b = this.toNumber(right);
                if (node.op === '+') return a + b;
                if (node.op === '-') return a - b;
                if (node.op === '*') return a * b;
                if (node.op === '/') {
                    if (b === 0) throw this.makeError('#DIV/0!');
                    return a / b;
                }
                return '#ERROR';
            }
            case 'Cell': {
                if (context && typeof context.getCellValue === 'function') {
                    return context.getCellValue(node.ref, stack);
                }
                return '';
            }
            case 'Range': {
                if (context && typeof context.getRangeValues === 'function') {
                    return context.getRangeValues(node.ref, stack);
                }
                return [];
            }
            case 'Call': {
                const fnName = node.name.toUpperCase();
                const fn = this.functions[fnName];
                if (!fn) throw this.makeError('#NAME?');
                const args = node.args.map(arg => this.evalNode(arg, context, stack));
                const argError = this.findErrorInArgs(args);
                if (argError) return argError;
                return fn(args);
            }
            default:
                return '#ERROR';
        }
    }

    tokenize(input) {
        const tokens = [];
        let i = 0;

        while (i < input.length) {
            const ch = input[i];

            if (/\s/.test(ch)) {
                i++;
                continue;
            }
            if (ch === '"' ) {
                let str = '';
                i++;
                while (i < input.length) {
                    const c = input[i];
                    const next = input[i + 1];
                    if (c === '"' && next === '"') {
                        str += '"';
                        i += 2;
                        continue;
                    }
                    if (c === '"') {
                        i++;
                        break;
                    }
                    str += c;
                    i++;
                }
                tokens.push({ type: 'String', value: str });
                continue;
            }
            if (/[0-9.]/.test(ch)) {
                let num = ch;
                i++;
                while (i < input.length && /[0-9.]/.test(input[i])) {
                    num += input[i++];
                }
                tokens.push({ type: 'Number', value: Number(num) });
                continue;
            }
            if (/[A-Za-z_]/.test(ch)) {
                let ident = ch;
                i++;
                while (i < input.length && /[A-Za-z0-9_$]/.test(input[i])) {
                    ident += input[i++];
                }

                // Possibly a cell/range like A1 or A1:B2
                if (/^\$?[A-Za-z]+\$?[0-9]+$/.test(ident) && input[i] === ':') {
                    i++;
                    let end = '';
                    if (input[i] === '$') {
                        end += '$';
                        i++;
                    }
                    while (i < input.length && /[A-Za-z0-9$]/.test(input[i])) {
                        end += input[i++];
                    }
                    tokens.push({ type: 'Range', value: `${ident}:${end}`.toUpperCase() });
                    continue;
                }

                // Cell reference like A1
                if (/^\$?[A-Za-z]+\$?[0-9]+$/.test(ident)) {
                    tokens.push({ type: 'Cell', value: ident.toUpperCase() });
                    continue;
                }

                tokens.push({ type: 'Ident', value: ident });
                continue;
            }

            if (ch === '$') {
                const next = input[i + 1];
                if (next && /[A-Za-z]/.test(next)) {
                    let ident = '$' + next;
                    i += 2;
                    while (i < input.length && /[A-Za-z0-9_$]/.test(input[i])) {
                        ident += input[i++];
                    }

                    // Possibly a cell/range like $A1 or $A1:$B2
                    if (/^\$?[A-Za-z]+\$?[0-9]+$/.test(ident) && input[i] === ':') {
                        i++;
                        let end = '';
                        if (input[i] === '$') {
                            end += '$';
                            i++;
                        }
                        while (i < input.length && /[A-Za-z0-9$]/.test(input[i])) {
                            end += input[i++];
                        }
                        tokens.push({ type: 'Range', value: `${ident}:${end}`.toUpperCase() });
                        continue;
                    }

                    if (/^\$?[A-Za-z]+\$?[0-9]+$/.test(ident)) {
                        tokens.push({ type: 'Cell', value: ident.toUpperCase() });
                        continue;
                    }

                    tokens.push({ type: 'Ident', value: ident });
                    continue;
                }
            }

            if (ch === '<' || ch === '>' || ch === '=') {
                const next = input[i + 1];
                if ((ch === '<' || ch === '>') && next === '=') {
                    tokens.push({ type: 'Symbol', value: ch + '=' });
                    i += 2;
                    continue;
                }
                if (ch === '<' && next === '>') {
                    tokens.push({ type: 'Symbol', value: '<>' });
                    i += 2;
                    continue;
                }
                tokens.push({ type: 'Symbol', value: ch });
                i++;
                continue;
            }
            if (ch === '&' || ch === '+' || ch === '-' || ch === '*' || ch === '/' || ch === '(' || ch === ')' || ch === ',' || ch === ';') {
                tokens.push({ type: 'Symbol', value: ch });
                i++;
                continue;
            }

            throw new Error('Invalid token');
        }

        return tokens;
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
        if (this.isError(value)) return fallback;
        const n = Number(String(value).replace(/,/g, '').trim());
        return Number.isFinite(n) ? n : fallback;
    }

    toString(value) {
        return value === null || value === undefined ? '' : String(value);
    }

    isError(value) {
        return typeof value === 'string' && value.startsWith('#');
    }

    findError(value) {
        if (this.isError(value)) return value;
        if (Array.isArray(value)) {
            for (const item of value) {
                const err = this.findError(item);
                if (err) return err;
            }
        }
        return null;
    }

    findErrorInArgs(args) {
        for (const arg of args) {
            const err = this.findError(arg);
            if (err) return err;
        }
        return null;
    }

    makeError(code) {
        const err = new Error(code);
        err.code = code;
        return err;
    }

    toBoolean(value) {
        if (value === true || value === false) return value;
        if (typeof value === 'string') {
            const v = value.trim().toUpperCase();
            if (v === 'TRUE') return true;
            if (v === 'FALSE') return false;
        }
        const n = this.toNumber(value, NaN);
        if (Number.isFinite(n)) return n !== 0;
        return Boolean(value);
    }

    coerceComparable(value) {
        if (value === null || value === undefined) return '';
        const num = this.toNumber(value, NaN);
        if (Number.isFinite(num)) return num;
        return String(value);
    }

    compare(a, b, op) {
        if (op === '=') return a === b;
        if (op === '<>') return a !== b;
        if (op === '<') return a < b;
        if (op === '<=') return a <= b;
        if (op === '>') return a > b;
        if (op === '>=') return a >= b;
        return false;
    }
}

class Parser {
    constructor(tokens) {
        this.tokens = tokens;
        this.current = 0;
    }

    isAtEnd() {
        return this.current >= this.tokens.length;
    }

    peek() {
        return this.tokens[this.current];
    }

    advance() {
        if (!this.isAtEnd()) this.current++;
        return this.tokens[this.current - 1];
    }

    matchSymbol(...symbols) {
        if (this.isAtEnd()) return false;
        const t = this.peek();
        if (t.type === 'Symbol' && symbols.includes(t.value)) {
            this.advance();
            return true;
        }
        return false;
    }

    parseExpression() {
        return this.parseComparison();
    }

    parseComparison() {
        let expr = this.parseConcat();
        while (this.matchSymbol('=', '<>', '<', '<=', '>', '>=')) {
            const op = this.tokens[this.current - 1].value;
            const right = this.parseConcat();
            expr = { type: 'Binary', op, left: expr, right };
        }
        return expr;
    }

    parseConcat() {
        let expr = this.parseAddition();
        while (this.matchSymbol('&')) {
            const op = this.tokens[this.current - 1].value;
            const right = this.parseAddition();
            expr = { type: 'Binary', op, left: expr, right };
        }
        return expr;
    }

    parseAddition() {
        let expr = this.parseTerm();
        while (this.matchSymbol('+', '-')) {
            const op = this.tokens[this.current - 1].value;
            const right = this.parseTerm();
            expr = { type: 'Binary', op, left: expr, right };
        }
        return expr;
    }

    parseTerm() {
        let expr = this.parseFactor();
        while (this.matchSymbol('*', '/')) {
            const op = this.tokens[this.current - 1].value;
            const right = this.parseFactor();
            expr = { type: 'Binary', op, left: expr, right };
        }
        return expr;
    }

    parseFactor() {
        if (this.matchSymbol('+')) {
            return { type: 'Unary', op: '+', expr: this.parseFactor() };
        }
        if (this.matchSymbol('-')) {
            return { type: 'Unary', op: '-', expr: this.parseFactor() };
        }
        return this.parsePrimary();
    }

    parsePrimary() {
        if (this.isAtEnd()) return { type: 'Number', value: 0 };
        const token = this.advance();

        if (token.type === 'Number') return { type: 'Number', value: token.value };
        if (token.type === 'String') return { type: 'String', value: token.value };
        if (token.type === 'Cell') return { type: 'Cell', ref: token.value };
        if (token.type === 'Range') return { type: 'Range', ref: token.value };

        if (token.type === 'Ident') {
            if (this.matchSymbol('(')) {
                const args = [];
                if (!this.matchSymbol(')')) {
                    do {
                        args.push(this.parseExpression());
                    } while (this.matchSymbol(',', ';'));
                    if (!this.matchSymbol(')')) throw this.makeParseError('#ERROR');
                }
                return { type: 'Call', name: token.value, args };
            }
            return { type: 'String', value: token.value };
        }

        if (token.type === 'Symbol' && token.value === '(') {
            const expr = this.parseExpression();
            if (!this.matchSymbol(')')) throw this.makeParseError('#ERROR');
            return expr;
        }

        throw this.makeParseError('#ERROR');
    }

    makeParseError(code) {
        const err = new Error(code);
        err.code = code;
        return err;
    }
}

window.FormulaEngine = FormulaEngine;
