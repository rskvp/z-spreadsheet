const TextFunctions = require('../formulas/functions/text');
const MathFunctions = require('../formulas/functions/math');
const TrigFunctions = require('../formulas/functions/trigonometry');
const LogicalFunctions = require('../formulas/functions/logical');
const EngFunctions = require('../formulas/functions/engineering');
const ReferenceFunctions = require('../formulas/functions/reference');
const InformationFunctions = require('../formulas/functions/information');
const StatisticalFunctions = require('../formulas/functions/statistical');
const DateFunctions = require('../formulas/functions/date');
const WebFunctions = require('../formulas/functions/web');
const FormulaError = require('../formulas/error');
const { FormulaHelpers } = require('../formulas/helpers');
const { Parser, allTokens } = require('./parsing');
const lexer = require('./lexing');
const Utils = require('./utils');
const { DepParser } = require('./dependency/hooks');

/**
 * A Excel Formula Parser & Evaluator
 */
class FormulaParser {
  /**
   * @param {{functions: {}, functionsNeedContext: {}, onVariable: function, onCell: function, onRange: function}} [config]
   * @param isTest - is in testing environment
   */
  constructor(config, isTest = false) {
    this.logs = [];
    this.isTest = isTest;
    this.utils = new Utils(this);
    config = {
      functions: {},
      functionsNeedContext: {},
      onVariable: () => null,
      onCell: () => 0,
      onRange: () => [[0]],
      onStructuredReference: () => ({ ref: null }),
      ...config,
    };

    this.onVariable = config.onVariable;
    this.onStructuredReference = config.onStructuredReference;
    this.functions = {
      ...DateFunctions,
      ...StatisticalFunctions,
      ...InformationFunctions,
      ...ReferenceFunctions,
      ...EngFunctions,
      ...LogicalFunctions,
      ...TextFunctions,
      ...MathFunctions,
      ...TrigFunctions,
      ...WebFunctions,
      ...config.functions,
      ...config.functionsNeedContext,
    };
    this.onRange = config.onRange;
    this.onCell = config.onCell;
    this.isRunningAction = config.isRunningAction;
    this.onFullCell = config.onFullCell;
    this.onFullRange = config.onFullRange;
    this.formulaParser = config.formulaParser;

    // functions treat null as 0, other functions treats null as ""
    this.funsNullAs0 = Object.keys(MathFunctions)
      .concat(Object.keys(TrigFunctions))
      .concat(Object.keys(LogicalFunctions))
      .concat(Object.keys(EngFunctions))
      .concat(Object.keys(ReferenceFunctions))
      .concat(Object.keys(StatisticalFunctions))
      .concat(Object.keys(DateFunctions));

    // functions need context and don't need to retrieve references
    this.funsNeedContextAndNoDataRetrieve = [
      'ROW',
      'ROWS',
      'COLUMN',
      'COLUMNS',
      'SUMIF',
      'SUMIFS',
      'INDEX',
      'AVERAGEIF',
    ];

    // functions need parser context
    this.funsNeedContext = [
      ...Object.keys(config.functionsNeedContext),
      ...this.funsNeedContextAndNoDataRetrieve,
      'INDEX',
      'OFFSET',
      'INDIRECT',
      'CHOOSE',
      'WEBSERVICE',
    ];

    // functions preserve reference in arguments
    this.funsPreserveRef = Object.keys(InformationFunctions);

    this.parser = new Parser(this, this.utils);
    this.depParser = new DepParser({
      onStructuredReference: this.onStructuredReference,
    });
  }

  /**
   * Get all lexing token names. Webpack needs this.
   * @return {Array.<string>} - All token names that should not be minimized.
   */
  static get allTokens() {
    return allTokens;
  }

  /**
   * Get value from the cell reference
   * @param ref
   * @return {*}
   */
  getCell(ref) {
    // console.log('get cell', JSON.stringify(ref));
    if (ref.sheet == null) {
      ref.sheet = this.position ? this.position.sheet : undefined;
    }
    return this.onCell(ref);
  }

  /**
   * Get values from the range reference.
   * @param ref
   * @return {*}
   */
  getRange(ref) {
    // console.log('get range', JSON.stringify(ref));
    if (ref.sheet == null) {
      ref.sheet = this.position ? this.position.sheet : undefined;
    }
    return this.onRange(ref);
  }

  /**
   * TODO:
   * Get references or values from a user defined variable.
   * @param name
   * @return {*}
   */
  getVariable(name) {
    // console.log('get variable', name);
    const res = {
      ref: this.onVariable(name, this.position.sheet, this.position),
    };
    if (res.ref == null) {
      return FormulaError.NAME;
    }
    return res;
  }

  /**
   * Get references or values for a structured referte
   * @param {string} tableName
   * @param {string} columnName
   * @param {boolean} thisRow
   */
  getStructuredReference(tableName, columnName, thisRow, specialItem) {
    const res = {
      ref: this.onStructuredReference(
        tableName,
        columnName,
        thisRow,
        specialItem,
        this.position.sheet,
        this.position
      ),
    };
    if (res.ref == null) {
      return FormulaError.NAME;
    }
    return res;
  }

  /**
   * Retrieve values from the given reference.
   * @param valueOrRef
   * @return {*}
   */
  retrieveRef(valueOrRef) {
    if (FormulaHelpers.isRangeRef(valueOrRef)) {
      return this.getRange(valueOrRef.ref);
    }
    if (FormulaHelpers.isCellRef(valueOrRef)) {
      return this.getCell(valueOrRef.ref);
    }
    return valueOrRef;
  }

  /**
   * Call an excel function.
   * @param name - Function name.
   * @param args - Arguments that pass to the function.
   * @return {*}
   */
  _callFunction(name, args) {
    if (name.indexOf('_xlfn.') === 0) name = name.slice(6);
    name = name.toUpperCase();
    // if one arg is null, it means 0 or "" depends on the function it calls
    const nullValue = this.funsNullAs0.includes(name) ? 0 : '';

    if (!this.funsNeedContextAndNoDataRetrieve.includes(name)) {
      // retrieve reference
      args = args.map((arg) => {
        if (arg === null)
          return { value: nullValue, isArray: false, omitted: true };
        const res = this.utils.extractRefValue(arg);

        if (this.funsPreserveRef.includes(name)) {
          return { value: res.val, isArray: res.isArray, ref: arg.ref };
        }
        return {
          ref: arg.ref,
          value: res.val,
          isArray: res.isArray,
          isRangeRef: !!FormulaHelpers.isRangeRef(arg),
          isCellRef: !!FormulaHelpers.isCellRef(arg),
        };
      });
    }
    // console.log('callFunction', name, args)

    if (this.functions[name]) {
      let res;
      try {
        if (
          !this.funsNeedContextAndNoDataRetrieve.includes(name) &&
          !this.funsNeedContext.includes(name)
        )
          res = this.functions[name](...args);
        else res = this.functions[name](this, ...args);
      } catch (e) {
        // allow functions throw FormulaError, this make functions easier to implement!
        if (e instanceof FormulaError) {
          return e;
        }
        throw e;
      }
      if (res === undefined) {
        // console.log(`Function ${name} may be not implemented.`);
        if (this.isTest) {
          if (!this.logs.includes(name)) this.logs.push(name);
          return { value: 0, ref: {} };
        }
        throw FormulaError.NOT_IMPLEMENTED(name);
      }
      return res;
    }
    // console.log(`Function ${name} is not implemented`);
    if (this.isTest) {
      if (!this.logs.includes(name)) this.logs.push(name);
      return { value: 0, ref: {} };
    }
    throw FormulaError.NOT_IMPLEMENTED(name);
  }

  async callFunctionAsync(name, args) {
    const awaitedArgs = [];
    for (const arg of args) {
      try {
        awaitedArgs.push(await arg);
      } catch (e) {
        if (e instanceof FormulaError) {
          awaitedArgs.push(e);
        } else {
          throw e;
        }
      }
    }
    const res = await this._callFunction(name, awaitedArgs);
    return FormulaHelpers.checkFunctionResult(res);
  }

  callFunction(name, args) {
    if (this.async) {
      return this.callFunctionAsync(name, args);
    }
    const res = this._callFunction(name, args);
    return FormulaHelpers.checkFunctionResult(res);
  }

  /**
   * Return currently supported functions.
   * @return {this}
   */
  supportedFunctions() {
    const supported = [];
    for (const fn in this.functions) {
      supported.push(fn);
    }
    return supported;
  }

  /**
   * Check and return the appropriate formula result.
   * @param result
   * @param {boolean} [allowReturnArray] - If the formula can return an array
   * @return {*}
   */
  checkFormulaResult(result, allowReturnArray = false) {
    const type = typeof result;
    // number
    if (type === 'number') {
      if (isNaN(result)) {
        return FormulaError.VALUE;
      }
      if (!isFinite(result)) {
        return FormulaError.NUM;
      }
      result += 0; // make -0 to 0
    } else if (type === 'object') {
      if (result instanceof FormulaError) return result;
      if (allowReturnArray) {
        if (result.ref) {
          result = this.retrieveRef(result);
        }
        // Disallow union, and other unknown data types.
        // e.g. `=(A1:C1, A2:E9)` -> #VALUE!
        if (
          typeof result === 'object' &&
          !Array.isArray(result) &&
          result != null
        ) {
          return FormulaError.VALUE;
        }
      } else if (result.ref && result.ref.row && !result.ref.from) {
        // single cell reference
        result = this.retrieveRef(result);
      } else if (
        result.ref &&
        result.ref.from &&
        result.ref.from.col === result.ref.to.col
      ) {
        // single Column reference
        result = this.retrieveRef({
          ref: {
            row: result.ref.from.row,
            col: result.ref.from.col,
          },
        });
      } else if (Array.isArray(result)) {
        result = result[0][0];
      } else {
        // array, range reference, union collections
        return FormulaError.VALUE;
      }
    }
    return result;
  }

  /**
   * Parse an excel formula.
   * @param {string} inputText
   * @param {{row: number, col: number}} [position] - The position of the parsed formula
   *              e.g. {row: 1, col: 1}
   * @param {boolean} [allowReturnArray] - If the formula can return an array. Useful when parsing array formulas,
   *                                      or data validation formulas.
   * @returns {*}
   */
  parseWithType(inputText, position, allowReturnArray = false) {
    const result = this.parse(inputText, position, allowReturnArray);
    const rawDeps = this.depParser.parse(inputText, position);
    const dependencies = rawDeps.map((e) => {
      if ('from' in e && 'to' in e) {
        return this.onFullRange(e);
      }
      if ('row' in e && 'col' in e) {
        return this.onFullCell(e);
      }
      throw new Error(`Invalid dependency: ${JSON.stringify(e)}`);
    });
    return Utils.addType(result, inputText, dependencies);
  }

  parse(rawInputText, position, allowReturnArray = false) {
    if (rawInputText.length === 0) throw Error('Input must not be empty.');
    const inputText =
      rawInputText[0] === '=' ? rawInputText.slice(1) : rawInputText;

    this.position = position;
    this.async = false;
    const { tokens } = lexer.lex(inputText);

    if (Utils.isMacro(tokens, 'ACTION') && this.isRunningAction) {
      return this.parse(
        Utils.expandActionMacro(tokens),
        position,
        allowReturnArray
      );
    }

    if (Utils.isMacro(tokens, 'ACTION') && !this.isRunningAction) {
      return inputText;
    }

    if (Utils.isMacro(tokens, 'DELAYEDCOMPUTEDCOLUMN')) {
      return this.parse(
        Utils.expandDelayedComputedColumnMacro(tokens),
        position,
        allowReturnArray
      );
    }

    if (Utils.isMacro(tokens, 'COMPUTEDCOLUMN')) {
      return this.parse(
        Utils.expandComputedColumnMacro(tokens),
        position,
        allowReturnArray
      );
    }

    this.parser.input = tokens;
    let res;
    try {
      res = this.parser.formulaWithBinaryOp();
      res = this.checkFormulaResult(res, allowReturnArray);
      if (res instanceof FormulaError) {
        return res;
      }
    } catch (e) {
      console.error(e);
      throw FormulaError.ERROR(e.message, e);
    }
    if (this.parser.errors.length > 0) {
      const error = this.parser.errors[0];
      throw Utils.formatChevrotainError(error, inputText);
    }
    return res;
  }

  /**
   * Parse an excel formula asynchronously.
   * Use when providing custom async functions.
   * @param {string} inputText
   * @param {{row: number, col: number}} [position] - The position of the parsed formula
   *              e.g. {row: 1, col: 1}
   * @param {boolean} [allowReturnArray] - If the formula can return an array. Useful when parsing array formulas,
   *                                      or data validation formulas.
   * @returns {*}
   */
  async parseAsyncWithType(inputText, position, allowReturnArray = false) {
    if (typeof inputText !== 'string' || inputText[0] !== '=') {
      return await this.getResult(inputText);
    }
    const result = await this.parseAsync(inputText, position, allowReturnArray);
    const rawDeps = this.depParser.parse(inputText, position);
    const dependencies = rawDeps.map((e) => {
      if ('from' in e && 'to' in e) {
        return this.onFullRange(e);
      }
      if ('row' in e && 'col' in e) {
        return this.onFullCell(e);
      }
      throw new Error(`Invalid dependency: ${JSON.stringify(e)}`);
    });

    if (result.resultType === 'string' && Utils.isValidDate(result.result)) {
      (result.result = Utils.serialFromDatestring(result.result)),
        (result.resultType = 'date');
    }

    return Utils.addType(result, inputText, dependencies);
  }

  async parseAsync(rawInputText, position, allowReturnArray = false) {
    if (rawInputText.length === 0) throw Error('Input must not be empty.');
    const inputText =
      rawInputText[0] === '=' ? rawInputText.slice(1) : rawInputText;

    this.position = position;
    this.async = true;
    const { tokens } = lexer.lex(inputText);

    if (Utils.isMacro(tokens, 'ACTION') && this.isRunningAction) {
      return this.parseAsync(
        Utils.expandActionMacro(tokens),
        position,
        allowReturnArray
      );
    }

    if (Utils.isMacro(tokens, 'ACTION') && !this.isRunningAction) {
      return '='.concat(inputText);
    }

    if (Utils.isMacro(tokens, 'DELAYEDCOMPUTEDCOLUMN')) {
      return this.parseAsync(
        Utils.expandDelayedComputedColumnMacro(tokens),
        position,
        allowReturnArray
      );
    }

    if (Utils.isMacro(tokens, 'COMPUTEDCOLUMN')) {
      return this.parseAsync(
        Utils.expandComputedColumnMacro(tokens),
        position,
        allowReturnArray
      );
    }

    if (Utils.isMacro(tokens, 'IFDO')) {
      const commaLocations = Utils.findAllIndicies(
        tokens,
        (t) => t.tokenType.name === 'Comma'
      );
      const refComma = commaLocations[0];
      const predicate = tokens
        .slice(refComma - 1, refComma)
        .map((t) => t.image)
        .join('');
      const refVal = await this.parseAsync(
        predicate,
        position,
        allowReturnArray
      );
      if (refVal) {
        const consequent = tokens
          .slice(refComma + 1, tokens.length - 1)
          .map((t) => t.image)
          .join('');
        return await this.parseAsync(consequent, position, allowReturnArray);
      }
      return false;
    }

    this.parser.input = tokens;
    let res;
    try {
      res = await this.parser.formulaWithBinaryOp();
      res = this.checkFormulaResult(res, allowReturnArray);
    } catch (e) {
      console.error(e);
      throw FormulaError.ERROR(e.message, e);
    }
    if (res instanceof Error) {
      throw res;
    }
    if (this.parser.errors.length > 0) {
      const error = this.parser.errors[0];
      throw Utils.formatChevrotainError(error, inputText);
    }
    return res;
  }

  async getResult(text) {
    let result;
    if (typeof text === 'string' && Utils.isValidDate(text)) {
      result = {
        result: Utils.serialFromDatestring(text),
        resultType: 'date',
        timestamp: Date.now(),
      };
    } else if (
      typeof text === 'string' &&
      !Utils.isValidDate(text) &&
      Utils.isNumber(text)
    ) {
      result = {
        result: Number(text),
        resultType: 'number',
        timestamp: Date.now(),
      };
    } else if (
      typeof text === 'string' &&
      text[0] === '$' &&
      Utils.isNumber(text.slice(1))
    ) {
      result = {
        result: Number(text.slice(1)),
        resultType: 'currency',
        timestamp: Date.now(),
      };
    } else if (
      typeof text === 'string' &&
      text[text.length - 1] === '%' &&
      Utils.isNumber(text.slice(0, text.length - 1))
    ) {
      result = {
        result: Number(text.slice(0, text.length - 1)) / 100,
        resultType: 'percentage',
        timestamp: Date.now(),
      };
    } else if (text === undefined) {
      result = {
        result: '',
        resultType: 'string',
        timestamp: Date.now(),
      };
    } else if (typeof text === 'number') {
      result = {
        result: text,
        resultType: 'number',
        timestamp: Date.now(),
      };
    } else if (
      typeof text === 'string' &&
      ['FALSE', 'TRUE'].includes(text.toUpperCase())
    ) {
      result = {
        result: text.toUpperCase() === 'TRUE',
        resultType: 'boolean',
        timestamp: Date.now(),
      };
    } else if (typeof text === 'string') {
      result = {
        result: text,
        resultType: 'string',
        timestamp: Date.now(),
      };
    } else {
      console.error({
        S: 'Unexpected formula, defaulting to empty string',
        formula,
      });
      result = {
        result: '',
        resultType: 'string',
        timestamp: Date.now(),
      };
    }

    return result;
  }
}

module.exports = {
  FormulaParser,
  FormulaHelpers,
};
