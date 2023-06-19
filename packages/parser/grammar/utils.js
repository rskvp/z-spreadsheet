/* eslint-disable no-throw-literal */
const FormulaError = require('../formulas/error');
const { Address } = require('../formulas/helpers');
const { Prefix, Postfix, Infix, Operators } = require('../formulas/operators');
const Collection = require('./type/collection');

const MAX_ROW = 1048576;
const MAX_COLUMN = 16384;
const { NotAllInputParsedException } = require('chevrotain');
const lexer = require('./lexing');
const dayjs = require('dayjs');

const days = '30|31|(1|2)[0-9]||0[1-9]|[1-9]';
const month_nums = '1(0|1|2)|0[1-9]|[1-9]';
const month_names_long =
  'january|february|march|april|may(?!sbes)|june|july|august|september|october|november|december';
const month_names_short =
  'jan|feb|mar|apr|may(?!sbes)|jun|jul|aug|sep|oct|nov|dec';
const month_names = `${month_names_long}|${month_names_short}`;
const year = '(17|18|19|20)\\d{2}|^(17|18|19|20)\\d{2}';
const year_relaxed = `${year}|\\d{2}`;
const separator = '(\\,?)(\\s{1,}|\\-|\\/)';
const DATE_PATTERNS = [
  // 2 Jan 2012, 2-Jan-2012, 2,Jan,2012
  [
    `^(?<day>${days})${separator}(?<month>${month_names})${separator}(?<year>${year})$`,
  ],
  // Jan 12 2012, Jan-12-2012, Jan,12,2012
  [
    `^(?<month>${month_names})${separator}(?<day>${days})${separator}(?<year>${year})$`,
  ],
  // 2/12/2012, 2-12-2012, 2,2,2012
  [
    `^(?<month>${month_nums})${separator}(?<day>${days})${separator}(?<year>${year_relaxed})$`,
  ],
  // 2012/2/12, 2012-2-12, 2012,2,12
  [
    `^(?<year>${year})${separator}(?<month>${month_nums})${separator}(?<day>${days})$`,
  ],
];

const ValueFunctions = {
  date: new Set(['DATE', 'TODAY', 'DATEDIF']),
  datetime: new Set(['NOW']),
  number: new Set([
    'YEAR',
    'MONTH',
    'DAY',
    'HOUR',
    'MINUTE',
    'SECOND',
    'DATEVALUE',
  ]),
  currency: new Set([]),
};
class Utils {
  constructor(context) {
    this.context = context;
  }

  columnNameToNumber(columnName) {
    return Address.columnNameToNumber(columnName);
  }

  parseAddress(address) {
    const range = address.split(':');
    const cells = range.map((address) => this.parseCellAddress(address));
    if (cells.length == 1) {
      return cells[0];
    }
    if (cells.length == 2) {
      return {
        ref: {
          from: {
            col: Math.min(cells[0].ref.col, cells[1].ref.col),
            row: Math.min(cells[0].ref.row, cells[1].ref.row),
          },
          to: {
            col: Math.max(cells[0].ref.col, cells[1].ref.col),
            row: Math.max(cells[0].ref.row, cells[1].ref.row),
          },
        },
      };
    }
    throw `Unreachable Code Error number of arguments: = ${cells.length}`;
  }

  parseR1C1(currPosition, R1C1) {
    const R1C1Formula = /R(\[-?\d+\])C(\[-?\d+\])|R(\[-?\d+\])C|RC(\[-?\d+\])/;
    if (!R1C1Formula.test(R1C1)) throw FormulaError.VALUE;
    const bothRowCol = R1C1.split(':');
    const cells = bothRowCol.map((addy) =>
      this.parseCellR1C1(currPosition, addy)
    );
    if (cells.length == 1) {
      return cells[0];
    }
    if (cells.length == 2) {
      return {
        ref: {
          from: {
            col: Math.min(cells[0].ref.col, cells[1].ref.col),
            row: Math.min(cells[0].ref.row, cells[0].ref.row),
          },
          to: {
            col: Math.max(cells[0].ref.col, cells[0].ref.col),
            row: Math.max(cells[0].ref.row, cells[1].ref.row),
          },
        },
      };
    }
    throw 'Unreachable Code Error';
  }

  parseCellR1C1(currPosition, R1C1) {
    const basicSplit = R1C1.split(/[RC]+/);
    let splitRC = [];

    if (basicSplit.length > 2) {
      splitRC = basicSplit.splice(1, basicSplit.length);
    } else if (basicSplit.length <= 2) {
      splitRC = basicSplit;
    } else {
      throw "basicSplit array's length is too long";
    }

    const fullySplit = splitRC.map((str) => {
      if (str.length == 0) {
        return 0;
      }
      return parseInt(str.slice(1, str.length - 1));
    });

    return {
      ref: {
        col: currPosition.col + fullySplit[1],
        row: currPosition.row + fullySplit[0],
      },
    };
  }

  /**
   * Parse the cell address only.
   * @param {string} cellAddress
   * @return {{ref: {col: number, address: string, row: number}}}
   */
  parseCellAddress(cellAddress) {
    const res = cellAddress.match(/([$]?)([A-Za-z]{1,3})([$]?)([1-9][0-9]*)/);
    // console.log('parseCellAddress', cellAddress);
    return {
      ref: {
        address: res[0],
        col: this.columnNameToNumber(res[2]),
        row: +res[4],
      },
    };
  }

  parseRow(row) {
    const rowNum = +row;
    if (!Number.isInteger(rowNum)) {
      throw Error('Row number must be integer.');
    }
    return {
      ref: {
        col: undefined,
        row: +row,
      },
    };
  }

  parseCol(col) {
    return {
      ref: {
        col: this.columnNameToNumber(col),
        row: undefined,
      },
    };
  }

  parseColRange(col1, col2) {
    // const res = colRange.match(/([$]?)([A-Za-z]{1,3}):([$]?)([A-Za-z]{1,4})/);
    col1 = this.columnNameToNumber(col1);
    col2 = this.columnNameToNumber(col2);
    return {
      ref: {
        from: {
          col: Math.min(col1, col2),
          row: null,
        },
        to: {
          col: Math.max(col1, col2),
          row: null,
        },
      },
    };
  }

  parseRowRange(row1, row2) {
    // const res = rowRange.match(/([$]?)([1-9][0-9]*):([$]?)([1-9][0-9]*)/);
    return {
      ref: {
        from: {
          col: null,
          row: Math.min(row1, row2),
        },
        to: {
          col: null,
          row: Math.max(row1, row2),
        },
      },
    };
  }

  _applyPrefix(prefixes, val, isArray) {
    if (this.isFormulaError(val)) {
      return val;
    }
    return Prefix.unaryOp(prefixes, val, isArray);
  }

  async applyPrefixAsync(prefixes, value) {
    const { val, isArray } = this.extractRefValue(await value);
    return this._applyPrefix(prefixes, val, isArray);
  }

  /**
   * Apply + or - unary prefix.
   * @param {Array.<string>} prefixes
   * @param {*} value
   * @return {*}
   */
  applyPrefix(prefixes, value) {
    // console.log('applyPrefix', prefixes, value);
    if (this.context.async) {
      return this.applyPrefixAsync(prefixes, value);
    }
    const { val, isArray } = this.extractRefValue(value);
    return this._applyPrefix(prefixes, val, isArray);
  }

  _applyPostfix(val, isArray, postfix) {
    if (this.isFormulaError(val)) {
      return val;
    }
    return Postfix.percentOp(val, postfix, isArray);
  }

  async applyPostfixAsync(value, postfix) {
    const { val, isArray } = this.extractRefValue(await value);
    return this._applyPostfix(val, isArray, postfix);
  }

  applyPostfix(value, postfix) {
    // console.log('applyPostfix', value, postfix);
    if (this.context.async) {
      return this.applyPostfixAsync(value, postfix);
    }
    const { val, isArray } = this.extractRefValue(value);
    return this._applyPostfix(val, isArray, postfix);
  }

  _applyInfix(res1, infix, res2) {
    const val1 = res1.val;
    const isArray1 = res1.isArray;
    const val2 = res2.val;
    const isArray2 = res2.isArray;
    if (this.isFormulaError(val1)) return val1;
    if (this.isFormulaError(val2)) return val2;
    if (Operators.compareOp.includes(infix)) {
      return Infix.compareOp(val1, infix, val2, isArray1, isArray2);
    }
    if (Operators.concatOp.includes(infix)) {
      return Infix.concatOp(val1, infix, val2, isArray1, isArray2);
    }
    if (Operators.mathOp.includes(infix)) {
      return Infix.mathOp(val1, infix, val2, isArray1, isArray2);
    }
    throw new Error(`Unrecognized infix: ${infix}`);
  }

  async applyInfixAsync(value1, infix, value2) {
    const res1 = this.extractRefValue(await value1);
    const res2 = this.extractRefValue(await value2);
    return this._applyInfix(res1, infix, res2);
  }

  applyInfix(value1, infix, value2) {
    if (this.context.async) {
      return this.applyInfixAsync(value1, infix, value2);
    }
    const res1 = this.extractRefValue(value1);
    const res2 = this.extractRefValue(value2);
    return this._applyInfix(res1, infix, res2);
  }

  applyIntersect(refs) {
    // console.log('applyIntersect', refs);
    if (this.isFormulaError(refs[0])) {
      return refs[0];
    }
    if (!refs[0].ref) {
      throw Error(`Expecting a reference, but got ${refs[0]}.`);
    }
    // a intersection will keep track of references, value won't be retrieved here.
    let maxRow;
    let maxCol;
    let minRow;
    let minCol;
    let sheet;
    let res; // index start from 1
    // first time setup
    const { ref } = refs.shift();
    sheet = ref.sheet;
    if (!ref.from) {
      // check whole row/col reference
      if (ref.row === undefined || ref.col === undefined) {
        throw Error('Cannot intersect the whole row or column.');
      }

      // cell ref
      maxRow = minRow = ref.row;
      maxCol = minCol = ref.col;
    } else {
      // range ref
      // update
      maxRow = Math.max(ref.from.row, ref.to.row);
      minRow = Math.min(ref.from.row, ref.to.row);
      maxCol = Math.max(ref.from.col, ref.to.col);
      minCol = Math.min(ref.from.col, ref.to.col);
    }

    let err;
    refs.forEach((ref) => {
      if (this.isFormulaError(ref)) {
        return ref;
      }
      ref = ref.ref;
      if (!ref) throw Error(`Expecting a reference, but got ${ref}.`);
      if (!ref.from) {
        if (ref.row === undefined || ref.col === undefined) {
          throw Error('Cannot intersect the whole row or column.');
        }
        // cell ref
        if (
          ref.row > maxRow ||
          ref.row < minRow ||
          ref.col > maxCol ||
          ref.col < minCol ||
          sheet !== ref.sheet
        ) {
          err = FormulaError.NULL;
        }
        maxRow = minRow = ref.row;
        maxCol = minCol = ref.col;
      } else {
        // range ref
        const refMaxRow = Math.max(ref.from.row, ref.to.row);
        const refMinRow = Math.min(ref.from.row, ref.to.row);
        const refMaxCol = Math.max(ref.from.col, ref.to.col);
        const refMinCol = Math.min(ref.from.col, ref.to.col);
        if (
          refMinRow > maxRow ||
          refMaxRow < minRow ||
          refMinCol > maxCol ||
          refMaxCol < minCol ||
          sheet !== ref.sheet
        ) {
          err = FormulaError.NULL;
        }
        // update
        maxRow = Math.min(maxRow, refMaxRow);
        minRow = Math.max(minRow, refMinRow);
        maxCol = Math.min(maxCol, refMaxCol);
        minCol = Math.max(minCol, refMinCol);
      }
    });
    if (err) return err;
    // check if the ref can be reduced to cell reference
    if (maxRow === minRow && maxCol === minCol) {
      res = {
        ref: {
          sheet,
          row: maxRow,
          col: maxCol,
        },
      };
    } else {
      res = {
        ref: {
          sheet,
          from: { row: minRow, col: minCol },
          to: { row: maxRow, col: maxCol },
        },
      };
    }

    if (!res.ref.sheet) {
      delete res.ref.sheet;
    }
    return res;
  }

  applyUnion(refs) {
    const collection = new Collection();
    for (let i = 0; i < refs.length; i++) {
      if (this.isFormulaError(refs[i])) {
        return refs[i];
      }
      collection.add(this.extractRefValue(refs[i]).val, refs[i]);
    }

    // console.log('applyUnion', unions);
    return collection;
  }

  /**
     * Apply multiple references, e.g. A1:B3:C8:A:1:.....
     * @param refs
     // * @return {{ref: {from: {col: number, row: number}, to: {col: number, row: number}}}}
     */
  applyRange(refs) {
    let res;
    let maxRow = -1;
    let maxCol = -1;
    let minRow = MAX_ROW + 1;
    let minCol = MAX_COLUMN + 1;
    refs.forEach((ref) => {
      if (this.isFormulaError(ref)) {
        return ref;
      }
      // row ref is saved as number, parse the number to row ref here
      if (typeof ref === 'number') {
        ref = this.parseRow(ref);
      }
      ref = ref.ref;
      // check whole row/col reference
      if (ref.row === undefined) {
        minRow = 1;
        maxRow = MAX_ROW;
      }
      if (ref.col === undefined) {
        minCol = 1;
        maxCol = MAX_COLUMN;
      }

      if (ref.row > maxRow) {
        maxRow = ref.row;
      }
      if (ref.row < minRow) {
        minRow = ref.row;
      }
      if (ref.col > maxCol) {
        maxCol = ref.col;
      }
      if (ref.col < minCol) {
        minCol = ref.col;
      }
    });
    if (maxRow === minRow && maxCol === minCol) {
      res = {
        ref: {
          row: maxRow,
          col: maxCol,
        },
      };
    } else {
      res = {
        ref: {
          from: { row: minRow, col: minCol },
          to: { row: maxRow, col: maxCol },
        },
      };
    }
    return res;
  }

  /**
   * Throw away the refs, and retrieve the value.
   * @return {{val: *, isArray: boolean}}
   */
  extractRefValue(obj) {
    const res = obj;
    let isArray = false;
    if (
      Array.isArray(res) ||
      (res.ref != null && res.ref.from != null && res.ref.to != null)
    ) {
      isArray = true;
    }
    if (obj.ref) {
      // can be number or array
      return { val: this.context.retrieveRef(obj), isArray };
    }
    return { val: res, isArray };
  }

  /**
   *
   * @param array
   * @return {Array}
   */
  toArray(array) {
    // TODO: check if array is valid
    // console.log('toArray', array);
    return array;
  }

  /**
   * @param {string} number
   * @return {number}
   */
  toNumber(number) {
    return Number(number);
  }

  /**
   * @param {string} string
   * @return {string}
   */
  toString(string) {
    return string.substring(1, string.length - 1).replace(/""/g, '"');
  }

  /**
   * @param {string} bool
   * @return {boolean}
   */
  toBoolean(bool) {
    return bool === 'TRUE';
  }

  /**
   * Parse an error.
   * @param {string} error
   * @return {string}
   */
  toError(error) {
    return new FormulaError(error.toUpperCase());
  }

  isFormulaError(obj) {
    return obj instanceof FormulaError;
  }

  static isNumber(s) {
    return s !== '' && !isNaN(s);
  }

  static serialFromDatestring(datestring) {
    const day = dayjs(datestring);
    const bugOffset = day > dayjs('2/28/1900') ? 1 : 0;
    return day.diff('1900-01-01', 'days') + 1 + bugOffset;
  }

  static isValidDate(value) {
    for (const [pattern] of DATE_PATTERNS) {
      const result = new RegExp(pattern, 'gi').test(value);
      if (result) {
        return result;
      }
    }
    return undefined;
  }

  static formatChevrotainError(error, inputText) {
    let line;
    let column;
    let msg = '';
    // e.g. SUM(1))
    if (error instanceof NotAllInputParsedException) {
      line = error.token.startLine;
      column = error.token.startColumn;
    } else {
      line = error.previousToken.startLine;
      column = error.previousToken.startColumn + 1;
    }

    msg += `\n${inputText.split('\n')[line - 1]}\n`;
    msg += `${Array(column - 1)
      .fill(' ')
      .join('')}^\n`;
    msg += `Error at position ${line}:${column}\n${error.message}`;
    error.errorLocation = { line, column };
    return FormulaError.ERROR(msg, error);
  }

  static cleanFunctionToken(text) {
    return text.replace(new RegExp(/\(|\)/, 'gi'), '');
  }

  static isTokenInList(tokens, set) {
    return tokens.length > 0 && tokens.some((token) => set.has(token));
  }

  static isType(result, text, dependencies, type) {
    if (text === null || typeof result === 'string') {
      return false;
    }
    const { tokens } = lexer.lex(text);
    const normalizedTokens = tokens
      .filter((token) => token.tokenType.name === 'Function')
      .map((token) => Utils.cleanFunctionToken(token.image).toUpperCase());

    const allDependenciesAreOfType =
      dependencies.length > 0 &&
      dependencies.every((d) => d.resultType === type || d.datatype === type);

    const numberFunction = Utils.isTokenInList(
      normalizedTokens,
      ValueFunctions.number
    );
    const typeFunction = Utils.isTokenInList(
      normalizedTokens,
      ValueFunctions[type]
    );
    return !numberFunction && (typeFunction || allDependenciesAreOfType);
  }

  static resultType(result, inputText, dependencies) {
    if (Array.isArray(result)) {
      return 'array';
    }
    if (typeof result === 'string' && result.length > 0 && result[0] === '$') {
      return 'currency';
    }
    if (typeof result === 'string') {
      return 'string';
    }
    if (result === null || result === undefined) {
      return undefined;
    }
    if (typeof result === 'boolean') {
      return 'boolean';
    }

    for (const type of ['datetime', 'date', 'currency']) {
      if (
        typeof result === 'number' &&
        !isNaN(Number(result)) &&
        Utils.isType(result, inputText, dependencies, type)
      ) {
        return type;
      }
    }

    if (!isNaN(Number(result))) {
      return 'number';
    }
    return 'string';
  }

  static addType(rawResult, inputText, dependencies) {
    let result;
    if (typeof rawResult === 'string') {
      try {
        const parsedResult = JSON.parse(rawResult);
        if (!Array.isArray(parsedResult) && typeof parsedResult === 'object') {
          result = JSON.parse(rawResult);
        } else {
          result = rawResult;
        }
        // if not json, it's fine, just treat like normal string
      } catch (e) {
        result = rawResult;
      }
    } else {
      result = rawResult;
    }

    if (Array.isArray(result)) {
      if (result.length === 0) {
        result.push([]);
      }
      // {1, 2, 3} is a horizontal array, so too are all plain arrays
      if (!Array.isArray(result[0])) {
        result = [result];
      }
      const baseData = dependencies.find(
        (d) =>
          Array.isArray(d) &&
          Array.isArray(d[0]) &&
          d[0].length === result[0].length
      );
      for (let i = 0; i < result.length; i++) {
        for (let j = 0; j < result[i].length; j++) {
          if (typeof result[i][j] !== 'object') {
            let resultType;
            // Hack to get dates to work when making filter tables
            if (
              baseData &&
              typeof result[i][j] === 'number' &&
              baseData[0][j].resultType === 'date'
            ) {
              resultType = 'date';
            } else {
              resultType = Utils.resultType(
                result[i][j],
                inputText,
                dependencies
              );
            }

            result[i][j] = {
              result: result[i][j],
              resultType,
            };
          }
        }
      }
    }
    // result already is typed
    if (
      !Array.isArray(result) &&
      typeof result === 'object' &&
      result !== null &&
      'result' in result &&
      'resultType' in result
    ) {
      return result;
    }
    if (!Array.isArray(result) && typeof result === 'object') {
      return {
        result: rawResult,
        resultType: 'string',
      };
    }
    return {
      result,
      resultType: Utils.resultType(result, inputText, dependencies),
    };
  }

  static expandActionMacro(tokens) {
    return `SUM(${tokens
      .slice(1, tokens.length - 1)
      .map((t) => t.image)
      .join('')})`;
  }

  static findAllIndicies(a, f) {
    const b = [];
    b.push(a.findIndex(f));
    while (b[b.length - 1] !== -1) {
      b.push(a.findIndex((e, i) => i > b[b.length - 1] && f(e)));
    }
    return b.slice(0, b.length - 1);
  }

  static computeColumnComma(tokens, delayComma) {
    let i = delayComma;
    while (
      i >= 0 &&
      tokens[i].tokenType.name !== 'Comma' &&
      tokens[i].tokenType.name !== 'CloseCurlyParen'
    ) {
      i--;
    }
    if (i === -1) {
      throw new Error("Can't find column comma in computed column macro");
    }

    if (tokens[i].tokenType.name === 'Comma') {
      return i;
    }

    if (tokens[i].tokenType.name === 'CloseCurlyParen') {
      while (i >= 0 && tokens[i].tokenType.name !== 'OpenCurlyParen') {
        i--;
      }

      if (i === -1) {
        throw new Error("Can't find column comma in computed column macro");
      }
      return i - 1;
    }
    throw new Error(`Unreachable code exception: ${i}`);
  }

  static computeTableComma(tokens, columnComma) {
    let i = columnComma - 1;
    while (i >= 0 && tokens[i].tokenType.name !== 'Comma') {
      i--;
    }

    if (i === -1) {
      throw new Error("Can't find table comma in computed column macro");
    }

    return i;
  }

  static expandComputedColumnMacro(tokens) {
    const commaLocations = Utils.findAllIndicies(
      tokens,
      (t) => t.tokenType.name === 'Comma'
    );
    const columnComma = Utils.computeColumnComma(tokens, tokens.length - 1);
    const tableComma = Utils.computeTableComma(tokens, columnComma);

    const tableName = tokens
      .slice(tableComma + 1, columnComma)
      .map((t) => t.image)
      .join(' ')
      .replaceAll('"', '');
    const columnNames = tokens
      .slice(columnComma + 1, tokens.length - 1)
      .map((t) => t.image)
      .join(' ');
    const rArgs = [
      `"=${tokens
        .slice(1, tableComma)
        .map((t) => t.image)
        .join('')
        .replaceAll('"', '""')}"`,
      `ROWS(${tableName}[])`,
    ];
    const repeat = `repeat(${rArgs[0]},${rArgs[1]},1)`;
    const formula = `extendTable(${repeat}, "${tableName}", ${columnNames})`;

    return formula;
  }

  static expandDelayedComputedColumnMacro(tokens) {
    const commaLocations = Utils.findAllIndicies(
      tokens,
      (t) => t.tokenType.name === 'Comma'
    );
    const columnComma = Utils.computeColumnComma(tokens, tokens.length - 1);
    const tableComma = Utils.computeTableComma(tokens, columnComma);

    const tableName = tokens
      .slice(tableComma + 1, columnComma)
      .map((t) => t.image)
      .join(' ')
      .replaceAll('"', '');
    const columnNames = tokens
      .slice(columnComma + 1, tokens.length - 1)
      .map((t) => t.image)
      .join(' ');
    const rArgs = [
      `"=${tokens
        .slice(1, tableComma)
        .map((t) => t.image)
        .join('')
        .replaceAll('"', '""')}"`,
      `ROWS(${tableName}[])`,
    ];
    const repeat = `repeat(${rArgs[0]},${rArgs[1]},1, 1000)`;
    const formula = `extendTable(${repeat}, "${tableName}", ${columnNames})`;

    return formula;
  }

  static isMacro(tokens, macroName) {
    return (
      tokens.length > 0 &&
      tokens[0].image.toUpperCase() === `${macroName}(` &&
      tokens[tokens.length - 1].tokenType.name === 'CloseParen'
    );
  }
}

module.exports = Utils;
