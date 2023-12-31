const FormulaError = require('../error');
const { FormulaHelpers, WildCard, Address } = require('../helpers');
const { Types } = require('../types');
const Collection = require('../../grammar/type/collection');

const H = FormulaHelpers;
const Utils = require('../../grammar/utils');

const grammarUtils = new Utils();

const ReferenceFunctions = {
  ADDRESS: (rowNumber, columnNumber, absNum, a1, sheetText) => {
    rowNumber = H.accept(rowNumber, Types.NUMBER);
    columnNumber = H.accept(columnNumber, Types.NUMBER);
    absNum = H.accept(absNum, Types.NUMBER, 1);
    a1 = H.accept(a1, Types.BOOLEAN, true);
    sheetText = H.accept(sheetText, Types.STRING, '');

    if (rowNumber < 1 || columnNumber < 1 || absNum < 1 || absNum > 4) {
      throw FormulaError.VALUE;
    }

    let result = '';
    if (sheetText.length > 0) {
      if (/[^A-Za-z_.\d\u007F-\uFFFF]/.test(sheetText)) {
        result += `'${sheetText}'!`;
      } else {
        result += `${sheetText}!`;
      }
    }
    if (a1) {
      // A1 style
      result += absNum === 1 || absNum === 3 ? '$' : '';
      result += Address.columnNumberToName(columnNumber);
      result += absNum === 1 || absNum === 2 ? '$' : '';
      result += rowNumber;
    } else {
      // R1C1 style
      result += 'R';
      result += absNum === 4 || absNum === 3 ? `[${rowNumber}]` : rowNumber;
      result += 'C';
      result +=
        absNum === 4 || absNum === 2 ? `[${columnNumber}]` : columnNumber;
    }
    return result;
  },

  AREAS: (refs) => {
    refs = H.accept(refs);
    if (refs instanceof Collection) {
      return refs.length;
    }
    return 1;
  },
  // The parser should automatically handle all of the calculations.
  // Thus, this is just a formality for the Google sheets users
  ARRAYFORMULA: (value) => value.value,

  CHOOSE: (context, indexNum, ...values) => {
    const idx = H.accept(indexNum);
    const idxArr = H.flattenDeep([idx]).map(
      (index) => H.accept(index, Types.NUMBER) - 1
    );
    const arr = H.accept(values, Types.ARRAY);
    idxArr.forEach((index) => {
      if (index < 0 || index > arr.length - 1) {
        throw FormulaError.VALUE;
      }
    });
    const rv = [];
    idxArr.forEach((index) => rv.push(H.accept(arr[index])));
    if (rv.length > 1) {
      throw FormulaError.CUSTOM(
        'Can not test for multiple rows yet, need to get JSON testing PR merged'
      );
    }
    return [rv];
  },

  // Special
  COLUMN: (context, obj) => {
    if (obj == null) {
      if (context.position.col != null) return context.position.col;
      throw Error('FormulaParser.parse is called without position parameter.');
    } else {
      if (typeof obj !== 'object' || Array.isArray(obj)) {
        throw FormulaError.VALUE;
      }
      if (H.isCellRef(obj)) {
        return obj.ref.col;
      }
      if (H.isRangeRef(obj)) {
        return obj.ref.from.col;
      }
      throw Error('ReferenceFunctions.COLUMN should not reach here.');
    }
  },

  // Special
  COLUMNS: (context, obj) => {
    if (obj == null) {
      throw Error('COLUMNS requires one argument');
    }
    if (typeof obj !== 'object' || Array.isArray(obj)) throw FormulaError.VALUE;
    if (H.isCellRef(obj)) {
      return 1;
    }
    if (H.isRangeRef(obj)) {
      return Math.abs(obj.ref.from.col - obj.ref.to.col) + 1;
    }
    throw Error('ReferenceFunctions.COLUMNS should not reach here.');
  },

  FILTER: (returnArray, ...boolArrays) => {
    const bools = boolArrays.map((boolArray) =>
      H.accept(boolArray, Types.ARRAY)
    );
    acceptedMatrix = H.accept(returnArray, Types.ARRAY, undefined, false);
    acceptedArray = acceptedMatrix.map((row) => H.accept(row, Types.ARRAY));

    if (
      bools.some((bool) => bool.length !== acceptedArray.length) ||
      acceptedMatrix === undefined
    ) {
      throw FormulaError.VALUE;
    }
    return acceptedArray.filter((row, index) =>
      bools.every((bool) => bool[index])
    );
  },

  HLOOKUP: (lookupValue, tableArray, rowIndexNum, rangeLookup) => {
    // preserve type of lookupValue
    lookupValue = H.accept(lookupValue);
    try {
      tableArray = H.accept(tableArray, Types.ARRAY, undefined, false);
    } catch (e) {
      // catch #VALUE! and throw #N/A
      if (e instanceof FormulaError) throw FormulaError.NA;
      throw e;
    }
    rowIndexNum = H.accept(rowIndexNum, Types.NUMBER);
    rangeLookup = H.accept(rangeLookup, Types.BOOLEAN, true);

    // check if rowIndexNum out of bound
    if (rowIndexNum < 1) throw FormulaError.VALUE;
    if (tableArray[rowIndexNum - 1] === undefined) throw FormulaError.REF;

    const lookupType = typeof lookupValue; // 'number', 'string', 'boolean'
    // approximate lookup (assume the array is sorted)
    if (rangeLookup) {
      let prevValue =
        lookupType === typeof tableArray[0][0] ? tableArray[0][0] : null;
      for (let i = 1; i < tableArray[0].length; i++) {
        const currValue = tableArray[0][i];
        const type = typeof currValue;
        // skip the value if type does not match
        if (type !== lookupType) continue;
        // if the previous two values are greater than lookup value, throw #N/A
        if (prevValue > lookupValue && currValue > lookupValue) {
          throw FormulaError.NA;
        }
        if (currValue === lookupValue) return tableArray[rowIndexNum - 1][i];
        // if previous value <= lookup value and current value > lookup value
        if (
          prevValue != null &&
          currValue > lookupValue &&
          prevValue <= lookupValue
        ) {
          return tableArray[rowIndexNum - 1][i - 1];
        }
        prevValue = currValue;
      }
      if (prevValue == null) throw FormulaError.NA;
      if (tableArray[0].length === 1) {
        return tableArray[rowIndexNum - 1][0];
      }
      return prevValue;
    }
    // exact lookup with wildcard support

    let index = -1;
    if (WildCard.isWildCard(lookupValue)) {
      index = tableArray[0].findIndex((item) =>
        WildCard.toRegex(lookupValue, 'i').test(item)
      );
    } else {
      index = tableArray[0].findIndex((item) => item === lookupValue);
    }
    // the exact match is not found
    if (index === -1) throw FormulaError.NA;
    return tableArray[rowIndexNum - 1][index];
  },

  // Special
  INDEX: (context, ranges, rowNum, colNum, areaNum) => {
    // retrieve values
    rowNum = context.utils.extractRefValue(rowNum);
    rowNum = { value: rowNum.val, isArray: rowNum.isArray };
    rowNum = H.accept(rowNum, Types.NUMBER);
    rowNum = Math.trunc(rowNum);

    if (colNum == null) {
      colNum = 1;
    } else {
      colNum = context.utils.extractRefValue(colNum);
      colNum = { value: colNum.val, isArray: colNum.isArray };
      colNum = H.accept(colNum, Types.NUMBER, 1);
      colNum = Math.trunc(colNum);
    }

    if (areaNum == null) {
      areaNum = 1;
    } else {
      areaNum = context.utils.extractRefValue(areaNum);
      areaNum = { value: areaNum.val, isArray: areaNum.isArray };
      areaNum = H.accept(areaNum, Types.NUMBER, 1);
      areaNum = Math.trunc(areaNum);
    }

    // get the range area that we want to index
    // ranges can be cell ref, range ref or array constant
    let range = ranges;
    // many ranges (Reference form)
    if (ranges instanceof Collection) {
      range = ranges.refs[areaNum - 1];
    } else if (areaNum > 1) {
      throw FormulaError.REF;
    }

    if (rowNum === 0 && colNum === 0) {
      return range;
    }

    // query the whole column
    if (rowNum === 0) {
      if (H.isRangeRef(range)) {
        if (range.ref.to.col - range.ref.from.col < colNum - 1) {
          throw FormulaError.REF;
        }
        range.ref.from.col += colNum - 1;
        range.ref.to.col = range.ref.from.col;
        return range;
      }
      if (Array.isArray(range)) {
        const res = [];
        range.forEach((row) => res.push([row[colNum - 1]]));
        return res;
      }
    }
    // query the whole row
    if (colNum === 0) {
      if (H.isRangeRef(range)) {
        if (range.ref.to.row - range.ref.from.row < rowNum - 1) {
          throw FormulaError.REF;
        }
        range.ref.from.row += rowNum - 1;
        range.ref.to.row = range.ref.from.row;
        return range;
      }
      if (Array.isArray(range)) {
        return range[colNum - 1];
      }
    }
    // query single cell
    if (rowNum !== 0 && colNum !== 0) {
      // range reference
      if (H.isRangeRef(range)) {
        range = range.ref;
        if (
          range.to.row - range.from.row < rowNum - 1 ||
          range.to.col - range.from.col < colNum - 1
        ) {
          throw FormulaError.REF;
        }
        return {
          ref: {
            row: range.from.row + rowNum - 1,
            col: range.from.col + colNum - 1,
          },
        };
      }
      // cell reference
      if (H.isCellRef(range)) {
        range = range.ref;
        if (rowNum > 1 || colNum > 1) throw FormulaError.REF;
        return {
          ref: { row: range.row + rowNum - 1, col: range.col + colNum - 1 },
        };
      }
      // array constant
      if (Array.isArray(range)) {
        if (range.length < rowNum || range[0].length < colNum) {
          throw FormulaError.REF;
        }
        return range[rowNum - 1][colNum - 1];
      }
    }
  },

  INDIRECT: (context, refText, A1 = true) => {
    const refTextAccepted = H.accept(refText, Types.STRING);
    const A1Bool = H.accept(A1, Types.BOOLEAN);
    let returnValue = null;
    if (A1Bool) {
      returnValue = grammarUtils.parseAddress(refTextAccepted);
      if (returnValue.ref.to != null) {
        return FormulaError.CUSTOM(
          'Can not test for multiple rows yet, need to get JSON testing PR merged'
        );
      }
    } else {
      returnValue = grammarUtils.parseR1C1(context.position, refTextAccepted);
      if (returnValue.ref.to != null) {
        return FormulaError.CUSTOM(
          'Can not test for multiple rows yet, need to get JSON testing PR merged'
        );
      }
    }

    if (!H.checkValidAddress(returnValue)) {
      return 0;
    }
    return returnValue;
  },

  ISELEMENTVAL: (lookupArrayArg, valArg) => {
    const lookupArray = H.accept(lookupArrayArg, Types.ARRAY);
    const val = H.accept(valArg);

    return lookupArray.map((e) => H.accept(e) === val).map((e) => [e]);
  },

  MATCH: (lookupValue, lookupArray, matchType) => {
    const lookupValueAccepted = H.accept(lookupValue);
    const lookupArrayAccepted = H.accept(lookupArray, Types.COLLECTIONS);
    const matchTypeAccepted = H.accept(matchType, Types.Number, 1);
    if (lookupArrayAccepted.length > 1 && lookupArrayAccepted[0].length > 1) {
      throw FormulaError.NA;
    }
    const flattenedLookUpArray = H.flattenDeep(lookupArrayAccepted);
    let result;
    if (matchTypeAccepted > 0) {
      if (flattenedLookUpArray[0] > lookupValueAccepted) throw FormulaError.NA;
      result = flattenedLookUpArray.findIndex(
        (element) => lookupValueAccepted >= element
      );
    } else if (matchTypeAccepted === 0) {
      result = flattenedLookUpArray.findIndex(
        (element) => lookupValueAccepted === element
      );
      if (result === -1) throw FormulaError.NA;
    } else if (matchTypeAccepted < 0) {
      if (flattenedLookUpArray[0] < lookupValueAccepted) throw FormulaError.NA;
      result = flattenedLookUpArray.findIndex(
        (element) => lookupValueAccepted <= element
      );
    }
    return result + 1;
  },

  /**
   *
   * @param {*} Origin : Required. The reference from which you want to base the offset.
   *                        Reference must refer to a cell or range of adjacent cells; otherwise,
   *                        OFFSET returns the #VALUE! error value.
   * @param {*} Rows : Required. The number of rows, up or down, that you want the upper-left cell to refer to.
   *                   Using 5 as the rows argument specifies that the upper-left cell in the reference is five rows below reference.
   *                   Rows can be positive (which means below the starting reference) or negative (which means above the starting reference).
   * @param {*} Cols : Required. The number of columns, to the left or right, that you want the upper-left cell of the result to refer to.
   *                   Using 5 as the cols argument specifies that the upper-left cell in the reference is five columns to the right of reference.
   *                   Cols can be positive (which means to the right of the starting reference) or negative (which means to the left of the starting reference).
   * @param {*} Height : Optional. The height, in number of rows, that you want the returned reference to be. Height must be a positive number.
   * @param {*} Width : Optional. The width, in number of columns, that you want the returned reference to be. Width must be a positive number.
   */
  OFFSET: (context, origin, rows, cols, height = null, width = null) => {
    let newRow = null;
    let newCol = null;
    if (!H.isRangeRef(origin) && !H.isCellRef(origin)) {
      throw FormulaError.VALUE;
    }
    if (H.isRangeRef(origin)) {
      newRow = origin.ref.from.row + H.accept(rows, Types.NUMBER);
      newCol = origin.ref.from.col + H.accept(cols, Types.NUMBER);
    } else if (H.isCellRef(origin)) {
      newRow = origin.ref.row + H.accept(rows, Types.NUMBER);
      newCol = origin.ref.col + H.accept(cols, Types.NUMBER);
    } else {
      throw 'Unreachable Code Error';
    }

    if (newRow <= 0 || newCol <= 0) {
      throw FormulaError.CUSTOM('#REF', 'Out of Range Error');
    }

    const isRangeRef = H.isRangeRef(origin) != null;

    if (height == null && !isRangeRef) {
      height = 0;
    } else if (height == null && isRangeRef) {
      height = origin.ref.to.row - origin.ref.from.row + 1;
    } else if (height != null) {
      height = H.accept(height, Types.NUMBER);
    } else {
      throw 'Unreachable Code Error';
    }

    if (width == null && !isRangeRef) {
      width = 0;
    } else if (width == null && isRangeRef) {
      width = origin.ref.to.col - origin.ref.from.col + 1;
    } else if (width != null) {
      width = H.accept(width, H.NUMBER);
    } else {
      throw 'Unreachable Code Error';
    }
    const finalRow = newRow + height - Math.sign(height);
    const finalCol = newCol + width - Math.sign(width);
    if (finalRow < 0 || finalCol < 0) {
      throw FormulaError.CUSTOM('#REF', 'Out of Range Error');
    }
    const top = Math.min(newRow, finalRow);
    const left = Math.min(newCol, finalCol);
    const bottom = Math.max(newRow, finalRow);
    const right = Math.max(newCol, finalCol);
    ref = {
      ref: { from: { row: top, col: left }, to: { row: bottom, col: right } },
    };
    return context.utils.extractRefValue(ref).val;
  },

  // Special
  ROW: (context, obj) => {
    if (obj == null) {
      if (context.position.row != null) return context.position.row;
      throw Error('FormulaParser.parse is called without position parameter.');
    } else {
      if (typeof obj !== 'object' || Array.isArray(obj)) {
        throw FormulaError.VALUE;
      }
      if (H.isCellRef(obj)) {
        return obj.ref.row;
      }
      if (H.isRangeRef(obj)) {
        return obj.ref.from.row;
      }
      throw Error('ReferenceFunctions.ROW should not reach here.');
    }
  },

  // Special
  ROWS: (context, obj) => {
    if (obj == null) {
      throw Error('ROWS requires one argument');
    }
    if (typeof obj !== 'object' || Array.isArray(obj)) throw FormulaError.VALUE;
    if (H.isCellRef(obj)) {
      return 1;
    }
    if (H.isRangeRef(obj)) {
      return Math.abs(obj.ref.from.row - obj.ref.to.row) + 1;
    }
    throw Error('ReferenceFunctions.ROWS should not reach here.');
  },

  /**
   *
   * @param {*} array : REQUIRED The range, or array to sort
   * @param {*} sort_index : OPTIONAL A number indicating the row or column to sort by
   * @param {*} sort_order : OPTIONAL A number indicating the desired sort order; 1 for
   *                         ascending order (default), -1 for descending order
   * @param {*} by_col : OPTIONAL A logical value indicating the desired sort direction;
   *                     FALSE to sort by row (default), TRUE to sort by column
   */
  SORT: (
    array,
    sort_index = 1,
    sort_order = 1,
    by_col = false,
    ...sortArgs
  ) => {
    array = H.accept(array, Types.ARRAY, null, false);
    sort_index = H.accept(sort_index);
    sort_order = H.accept(sort_order);
    const sortValues = [{ sort_index: sort_index - 1, sort_order }];
    for (let i = 0; i < sortArgs.length - 1; i += 2) {
      const sort_index = H.accept(sortArgs[i], null, null);
      const sort_order = H.accept(sortArgs[i + 1], null, null);
      if (!sort_index || !sort_order) throw FormulaError.VALUE;
      sortValues.push({ sort_index: sort_index - 1, sort_order });
    }
    by_col = H.accept(by_col);

    const checkValidValues = ({ sort_index, sort_order }) => {
      if (sort_index < 0 || sort_index > array.length - 1) {
        throw FormulaError.VALUE;
      }
      if (![1, -1].includes(sort_order)) {
        throw FormulaError.VALUE;
      }
    };
    sortValues.map((val) => checkValidValues(val));
    if (typeof by_col !== 'boolean') {
      throw FormulaError.VALUE;
    }

    if (by_col) {
      array = ReferenceFunctions.TRANSPOSE(array);
    }

    const map = new Map();
    map.set('number', 1);
    map.set('string', 2);
    map.set('boolean', 3);

    const getOrderValue = (aVal, bVal) => {
      const aPriority = map.get(typeof aVal);
      const bPriority = map.get(typeof bVal);
      if (aPriority != bPriority) {
        return aPriority - bPriority;
      }

      if (typeof aVal === 'boolean') {
        if (aVal === bVal) {
          return 0;
        }
        if (aVal) {
          return 1;
        }
        if (bVal) {
          return -1;
        }
        throw `Unreachable Code ERROR: aVal type: ${typeof aVal}; bVal type: ${typeof bVal}; aPriority: ${aPriority}; bPriority: ${bPriority}`;
      }
      if (typeof aVal === 'number') {
        return aVal - bVal;
      }
      if (typeof aVal === 'string') {
        return aVal.localeCompare(bVal);
      }
    };

    const getValue = (x, sort_index = 0) => {
      val = H.accept(x[sort_index]);
      return typeof val === 'object' ? val.result : val;
    };

    let sortedArr = array.sort((a, b) => {
      a = H.accept(a, Types.ARRAY);
      b = H.accept(b, Types.ARRAY);

      return sortValues.reduce(
        (prevOrderValue, { sort_index, sort_order }) =>
          prevOrderValue === 0
            ? getOrderValue(getValue(a, sort_index), getValue(b, sort_index)) *
              sort_order
            : prevOrderValue,
        0
      );

      throw `2 Unreachable Code ERROR: aVal type: ${typeof aVal}; bVal type: ${typeof bVal}; aPriority: ${aPriority}; bPriority: ${bPriority}`;
    });

    if (by_col) {
      sortedArr = ReferenceFunctions.TRANSPOSE(array);
    }
    return sortedArr;
  },

  TRANSPOSE: (array) => {
    array = H.accept(array, Types.ARRAY, undefined, false);
    // https://github.com/numbers/numbers.js/blob/master/lib/numbers/matrix.js#L171
    const result = [];

    for (let i = 0; i < array[0].length; i++) {
      result[i] = [];

      for (let j = 0; j < array.length; j++) {
        result[i][j] = array[j][i];
      }
    }

    return result;
  },

  /**
   * @param {*} range - The data to filter by unique entries.
   * @returns unique rows in the provided source range, discarding duplicates.
   *          Rows are returned in the order in which they first appear in the source range.
   * */
  UNIQUE: (range) => {
    try {
      range = H.accept(range, Types.ARRAY, null, false);
    } catch (e) {
      range = H.accept(range);
      return range;
    }
    // Checks to see if range is an array of arrays. If it is not,
    // then the first row is Unique to its self and thus can be returned
    if (typeof range[0] !== 'object') {
      return range;
    }
    const seen = new Set();
    const rv = [];
    for (let index = 0; index < range.length; index++) {
      let currValue = range[index];
      let returnValue = currValue;
      currValue = H.accept(currValue);
      returnValue = H.accept(returnValue);

      // Since JS compares based upon pointers, converting our Arrays to strings allows for element wise comparisons.
      if (typeof currValue.join === 'function') {
        currValue = currValue.join();
      }
      if (!seen.has(currValue)) {
        seen.add(currValue);
        rv.push(returnValue);
      }
    }
    return rv;
  },

  VLOOKUP: (lookupValue, tableArray, colIndexNum, rangeLookup) => {
    // preserve type of lookupValue
    lookupValue = H.accept(lookupValue);
    try {
      tableArray = H.accept(tableArray, Types.ARRAY, undefined, false);
    } catch (e) {
      // catch #VALUE! and throw #N/A
      if (e instanceof FormulaError) throw FormulaError.NA;
      throw e;
    }
    colIndexNum = H.accept(colIndexNum, Types.NUMBER);
    rangeLookup = H.accept(rangeLookup, Types.BOOLEAN, true);

    // check if colIndexNum out of bound
    if (colIndexNum < 1) throw FormulaError.VALUE;
    if (tableArray[0][colIndexNum - 1] === undefined) throw FormulaError.REF;

    const lookupType = typeof lookupValue; // 'number', 'string', 'boolean'

    // approximate lookup (assume the array is sorted)
    if (rangeLookup) {
      let prevValue =
        lookupType === typeof tableArray[0][0] ? tableArray[0][0] : null;
      for (let i = 1; i < tableArray.length; i++) {
        const currRow = tableArray[i];
        const currValue = tableArray[i][0];
        const type = typeof currValue;
        // skip the value if type does not match
        if (type !== lookupType) continue;
        // if the previous two values are greater than lookup value, throw #N/A
        if (prevValue > lookupValue && currValue > lookupValue) {
          throw FormulaError.NA;
        }
        if (currValue === lookupValue) return currRow[colIndexNum - 1];
        // if previous value <= lookup value and current value > lookup value
        if (
          prevValue != null &&
          currValue > lookupValue &&
          prevValue <= lookupValue
        ) {
          return tableArray[i - 1][colIndexNum - 1];
        }
        prevValue = currValue;
      }
      if (prevValue == null) throw FormulaError.NA;
      if (tableArray.length === 1) {
        return tableArray[0][colIndexNum - 1];
      }
      return prevValue;
    }
    // exact lookup with wildcard support

    let index = -1;
    if (WildCard.isWildCard(lookupValue)) {
      index = tableArray.findIndex((currRow) =>
        WildCard.toRegex(lookupValue, 'i').test(currRow[0])
      );
    } else {
      index = tableArray.findIndex((currRow) => currRow[0] === lookupValue);
    }
    // the exact match is not found
    if (index === -1) throw FormulaError.NA;
    return tableArray[index][colIndexNum - 1];
  },
  /** *
   * @param lookup_value: the value we are looking for in our lookup_array
   * @param lookup_array: the array were we search for lookup_value
   * @param return_array: the array where our return value is
   * @param if_not_found: OPTIONAL default value if lookup_value is not found
   * @param match_mode: OPTIONAL:
   *                         0: If none found, return #N/A. This is the default
   *                         -1 - Exact match. If none found, return the next smaller item.
   *                          1 - Exact match. If none found, return the next larger item.
   *                          2 - A wildcard match where *, ?, and ~ have special meaning.
   * @param search_mode: OPTIONAL:
   *                          1 - Perform a search starting at the first item. This is the default.
   *                         -1 - Perform a reverse search starting at the last item.
   *                          2 - Perform a binary search that relies on lookup_array being sorted in ascending order. If not sorted, invalid results will be returned.
   *                         -2 - Perform a binary search that relies on lookup_array being sorted in descending order. If not sorted, invalid results will be returned.
   * Microsoft Link: https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929
   ** */
  XLOOKUP: (
    lookup_value,
    lookup_array,
    return_array,
    if_not_found = null,
    match_mode = 0,
    search_mode = 1
  ) => {
    try {
      lookup_array = H.accept(lookup_array, Types.ARRAY);
      return_array = H.accept(return_array, Types.ARRAY);
    } catch (e) {
      if (e instanceof FormulaError) {
        throw FormulaError.NA;
      }
      throw e;
    }
    if (lookup_array.length != return_array.length) {
      throw FormulaError.NA;
    }
    if (Array.isArray(lookup_array[0])) {
      throw FormulaError.VALUE;
    }
    if (Array.isArray(return_array[0])) {
      throw FormulaError.VALUE;
    }

    lookup_value = H.accept(lookup_value);
    search_mode = H.accept(search_mode, Types.NUMBER);
    match_mode = H.accept(match_mode, Types.NUMBER);
    if (![1, -1, 2, -2].includes(search_mode)) {
      throw FormulaError.VALUE;
    }
    if (![0, 1, -1, 2].includes(match_mode)) {
      throw FormulaError.VALUE;
    }
    if (match_mode === 2 && [2, -2].includes(search_mode)) {
      throw FormulaError.VALUE;
    }
    if (if_not_found != null) {
      if (if_not_found.omitted) {
        if_not_found = null;
      } else {
        if_not_found = H.accept(
          if_not_found,
          null,
          null,
          true,
          false,
          (test = true)
        );
      }
    }
    // If search mode is 1 or -1, then we run a linear search on the input arrays
    if ([1, -1].includes(search_mode)) {
      // Transform is 0 if search mode is 1 (we want to go through the array in order)
      // Transform is the last index if search_mode is -1 (we go through in reverse order)
      const transform = search_mode === 1 ? 0 : lookup_array.length - 1;
      const minDiff = { index: -1, value: Number.MAX_VALUE };
      const maxDiff = { index: -1, value: Number.MAX_VALUE };

      for (let i = 0; i < lookup_array.length; i++) {
        const currIndex = Math.abs(transform - i);
        const currValue = H.accept(lookup_array[currIndex]);
        const comparison = H.XLOOKUP_HELPER(
          lookup_value,
          currValue,
          match_mode != 2
        );

        if (comparison === 0) {
          return return_array[currIndex];
        }
        if ([1, -1].includes(match_mode)) {
          if (comparison < 0 && Math.abs(comparison) < minDiff.value) {
            minDiff.index = currIndex;
            minDiff.value = Math.abs(comparison);
          }
          if (comparison > 0 && comparison < maxDiff.value) {
            maxDiff.index = currIndex;
            maxDiff.value = comparison;
          }
        }
      }
      if (minDiff.index >= 0 && match_mode === -1) {
        return return_array[minDiff.index];
      }
      if (maxDiff.index >= 0 && match_mode === 1) {
        return return_array[maxDiff.index];
      }
      if (if_not_found != null) {
        return if_not_found;
      }
      throw FormulaError.NA;

      // In the case where search mode is 2 or -2, we run a binary search
    } else if ([2, -2].includes(search_mode)) {
      let front = 0;
      let back = lookup_array.length - 1;
      while (front < back - 1) {
        const middle = Math.floor((front + back) / 2);
        const currValue = H.accept(lookup_array[middle]);
        const comparison = H.XLOOKUP_HELPER(
          lookup_value,
          currValue,
          match_mode != 2,
          [2, -2].includes(search_mode)
        );
        if (comparison === 0) {
          return return_array[middle];
        }
        if (comparison < 0 && search_mode === 2) {
          front = middle;
        } else if (comparison < 0 && search_mode === -2) {
          back = middle;
        } else if (comparison > 0 && search_mode === 2) {
          back = middle;
        } else if (comparison > 0 && search_mode === -2) {
          front = middle;
        } else {
          throw 'Unreachable Code Error';
        }
      }
      const comparisonFront = H.XLOOKUP_HELPER(
        lookup_value,
        lookup_array[front],
        match_mode != 2
      );
      const comparisonBack = H.XLOOKUP_HELPER(
        lookup_value,
        lookup_array[back],
        match_mode != 2
      );

      if (comparisonFront === 0) {
        return return_array[front];
      }
      if (comparisonBack === 0) {
        return return_array[back];
      }
      // If search mode === 2 search_mode === 1, we have to check front first,
      // b/c its in ascending order and we want the next largest item
      if (comparisonFront > 0 && match_mode === 1 && search_mode === 2) {
        return return_array[front];
      }
      if (comparisonBack > 0 && match_mode === 1 && search_mode === 2) {
        return return_array[back];
      }
      // If search mode === 2 search_mode === -1, we have to check back first,
      // b/c its in ascending order and we want the next smaller item
      if (comparisonBack < 0 && match_mode === -1 && search_mode === 2) {
        return return_array[back];
      }
      if (comparisonFront < 0 && match_mode === -1 && search_mode === 2) {
        return return_array[front];
      }
      // If search mode === -2 search_mode === 1, we have to check back first,
      // b/c its in descending order and we want the next largest item
      if (comparisonBack > 0 && match_mode === 1 && search_mode === -2) {
        return return_array[back];
      }
      if (comparisonFront > 0 && match_mode === 1 && search_mode === -2) {
        return return_array[front];
      }
      // If search mode === -2 search_mode === -1, we have to check front first,
      // b/c its in descending order and we want the next smaller item
      if (comparisonFront < 0 && match_mode === -1 && search_mode === -2) {
        return return_array[front];
      }
      if (comparisonBack < 0 && match_mode === -1 && search_mode === -2) {
        return return_array[back];
      }
      if (if_not_found) {
        return if_not_found;
      }
    }
    throw FormulaError.NA;
  },
};

module.exports = ReferenceFunctions;
