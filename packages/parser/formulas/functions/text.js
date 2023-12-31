const FormulaError = require('../error');
const { FormulaHelpers, WildCard } = require('../helpers');
const { Types } = require('../types');

const H = FormulaHelpers;

// Spreadsheet number format
const ssf = require('../../ssf/ssf');

// Change number to Thai pronunciation string
const bahttext = require('bahttext');

// full-width and half-width converter
const charsets = {
  latin: { halfRE: /[!-~]/g, fullRE: /[！-～]/g, delta: 0xfee0 },
  hangul1: { halfRE: /[ﾡ-ﾾ]/g, fullRE: /[ᆨ-ᇂ]/g, delta: -0xedf9 },
  hangul2: { halfRE: /[ￂ-ￜ]/g, fullRE: /[ᅡ-ᅵ]/g, delta: -0xee61 },
  kana: {
    delta: 0,
    half: '｡｢｣､･ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ',
    full:
      '。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシ' +
      'スセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン゛゜',
  },
  extras: {
    delta: 0,
    half: '¢£¬¯¦¥₩\u0020|←↑→↓■°',
    full: '￠￡￢￣￤￥￦\u3000￨￩￪￫￬￭￮',
  },
};
const toFull = (set) => (c) =>
  set.delta
    ? String.fromCharCode(c.charCodeAt(0) + set.delta)
    : [...set.full][[...set.half].indexOf(c)];
const toHalf = (set) => (c) =>
  set.delta
    ? String.fromCharCode(c.charCodeAt(0) - set.delta)
    : [...set.half][[...set.full].indexOf(c)];
const re = (set, way) => set[`${way}RE`] || new RegExp(`[${set[way]}]`, 'g');
const sets = Object.keys(charsets).map((i) => charsets[i]);
const toFullWidth = (str0) =>
  sets.reduce((str, set) => str.replace(re(set, 'half'), toFull(set)), str0);
const toHalfWidth = (str0) =>
  sets.reduce((str, set) => str.replace(re(set, 'full'), toHalf(set)), str0);

const searchOnce = (findText, withinText, startNum, isCaseSensitive) => {
  if (typeof withinText !== 'string') {
    return FormulaError.VALUE;
  }
  if (startNum < 1 || startNum > withinText.length) throw FormulaError.VALUE;
  const findTextRegex =
    WildCard.isWildCard(findText) || !isCaseSensitive
      ? WildCard.toRegex(findText, 'i')
      : findText;
  const res = withinText.slice(startNum - 1).search(findTextRegex);
  if (res === -1) return FormulaError.VALUE;
  return res + startNum;
};

const prepostFix = (text, numChars, sliceFunc) => {
  text = H.accept(text, null, undefined, false);
  numChars = H.accept(numChars, Types.NUMBER, 1);
  if (numChars < 0) throw FormulaError.VALUE;

  const textArr = Array.isArray(text) ? text : [[text]];
  return textArr.map((row) =>
    row.map((textVal) => {
      const cellText = String(textVal);
      if (numChars > cellText.length) return cellText;
      return sliceFunc(cellText, numChars);
    })
  );
};

const TextFunctions = {
  ASC: (text) => {
    text = H.accept(text, Types.STRING);
    return toHalfWidth(text);
  },

  BAHTTEXT: (number) => {
    number = H.accept(number, Types.NUMBER);
    try {
      return bahttext(number);
    } catch (e) {
      throw Error(
        `Error in https://github.com/jojoee/bahttext \n${e.toString()}`
      );
    }
  },

  CHAR: (number) => {
    number = H.accept(number, Types.NUMBER);
    if (number > 255 || number < 1) throw FormulaError.VALUE;
    return String.fromCharCode(number);
  },

  CLEAN: (text) => {
    text = H.accept(text, Types.STRING);
    return text.replace(/[\x00-\x1F]/g, '');
  },

  CODE: (text) => {
    text = H.accept(text, Types.STRING);
    if (text.length === 0) throw FormulaError.VALUE;
    return text.charCodeAt(0);
  },

  CONCAT: (...params) => {
    let text = '';
    // does not allow union
    H.flattenParams(params, Types.STRING, false, (item) => {
      item = H.accept(item, Types.STRING);
      text += item;
    });
    return text;
  },

  CONCATENATE: (...params) => {
    let text = '';
    if (params.length === 0) {
      throw Error('CONCATENATE need at least one argument.');
    }
    params.forEach((param) => {
      // does not allow range reference, array, union
      param = H.accept(param, Types.STRING);
      text += param;
    });

    return text;
  },

  DBCS: (text) => {
    text = H.accept(text, Types.STRING);
    return toFullWidth(text);
  },

  DOLLAR: (number, decimals) => {
    number = H.accept(number, Types.NUMBER);
    decimals = H.accept(decimals, Types.NUMBER, 2);
    const decimalString = Array(decimals).fill('0').join('');
    // Note: does not support locales
    // TODO: change currency based on user locale or settings from this library
    return ssf
      .format(`$#,##0.${decimalString}_);($#,##0.${decimalString})`, number)
      .trim();
  },

  EXACT: (searchText, searchedInText) => {
    const text1 = H.accept(searchText);
    const text2 = H.accept(searchedInText);
    let withinTextArr;
    if (Array.isArray(text2)) {
      withinTextArr = text2;
    } else {
      withinTextArr = [[text2]];
    }
    return withinTextArr.map((row) => row.map((t) => t === text1));
  },

  FIND: (findText, withinText, startNum) => {
    findText = H.accept(findText, Types.STRING);
    withinText = H.accept(withinText, null, undefined, false);
    startNum = H.accept(startNum, Types.NUMBER, 1);
    let withinTextArr;
    if (Array.isArray(withinText)) {
      withinTextArr = withinText;
    } else {
      withinTextArr = [[withinText]];
    }
    return withinTextArr.map((row) =>
      row.map((val) =>
        searchOnce(findText, H.accept(val, null, null), startNum, true)
      )
    );
  },

  FINDB: (...params) => TextFunctions.FIND(...params),

  FIXED: (number, decimals, noCommas) => {
    number = H.accept(number, Types.NUMBER);
    decimals = H.accept(decimals, Types.NUMBER, 2);
    noCommas = H.accept(noCommas, Types.BOOLEAN, false);

    const decimalString = Array(decimals).fill('0').join('');
    const comma = noCommas ? '' : '#,';
    return ssf
      .format(
        `${comma}##0.${decimalString}_);(${comma}##0.${decimalString})`,
        number
      )
      .trim();
  },

  LEFT: (text, numChars) => prepostFix(text, numChars, (s, n) => s.slice(0, n)),

  LEFTB: (...params) => TextFunctions.LEFT(...params),

  LEN: (text) => {
    text = H.accept(text, Types.STRING);
    return text.length;
  },

  LENB: (...params) => TextFunctions.LEN(...params),

  LOWER: (text) => {
    text = H.accept(text);
    let withinTextArr;
    if (Array.isArray(text)) {
      withinTextArr = text;
    } else {
      withinTextArr = [[text]];
    }
    return withinTextArr.map((row) => row.map((t) => t.toLowerCase()));
  },

  MID: (text, startNum, numChars) => {
    text = H.accept(text, Types.STRING);
    startNum = H.accept(startNum, Types.NUMBER);
    numChars = H.accept(numChars, Types.NUMBER);
    if (startNum > text.length) return '';
    if (startNum < 1 || numChars < 1) throw FormulaError.VALUE;
    return text.slice(startNum - 1, startNum + numChars - 1);
  },

  MIDB: (...params) => TextFunctions.MID(...params),

  NUMBERVALUE: (text, decimalSeparator, groupSeparator) => {
    text = H.accept(text, Types.STRING);
    // TODO: support reading system locale and set separators
    decimalSeparator = H.accept(decimalSeparator, Types.STRING, '.');
    groupSeparator = H.accept(groupSeparator, Types.STRING, ',');

    if (text.length === 0) return 0;
    if (decimalSeparator.length === 0 || groupSeparator.length === 0) {
      throw FormulaError.VALUE;
    }
    decimalSeparator = decimalSeparator[0];
    groupSeparator = groupSeparator[0];
    if (
      decimalSeparator === groupSeparator ||
      text.indexOf(decimalSeparator) < text.lastIndexOf(groupSeparator)
    ) {
      throw FormulaError.VALUE;
    }

    const res = text
      .replace(groupSeparator, '')
      .replace(decimalSeparator, '.')
      // remove chars that not related to number
      .replace(/[^\-0-9.%()]/g, '')
      .match(/([(-]*)([0-9]*[.]*[0-9]+)([)]?)([%]*)/);
    if (!res) throw FormulaError.VALUE;
    // ["-123456.78%%", "(-", "123456.78", ")", "%%"]
    const leftParenOrMinus = res[1].length;
    const rightParen = res[3].length;
    const percent = res[4].length;
    let number = Number(res[2]);
    if (
      leftParenOrMinus > 1 ||
      (leftParenOrMinus && !rightParen) ||
      (!leftParenOrMinus && rightParen) ||
      isNaN(number)
    ) {
      throw FormulaError.VALUE;
    }
    number /= 100 ** percent;
    return leftParenOrMinus ? -number : number;
  },

  PHONETIC: () => {},

  PROPER: (text) => {
    text = H.accept(text, [Types.STRING]);
    text = text.toLowerCase();
    text = text.charAt(0).toUpperCase() + text.slice(1);
    return text.replace(/(?:[^a-zA-Z])([a-zA-Z])/g, (letter) =>
      letter.toUpperCase()
    );
  },

  REPLACE: (old_text, start_num, num_chars, new_text) => {
    old_text = H.accept(old_text, [Types.STRING]);
    start_num = H.accept(start_num, [Types.NUMBER]);
    num_chars = H.accept(num_chars, [Types.NUMBER]);
    new_text = H.accept(new_text, [Types.STRING]);

    const arr = old_text.split('');
    arr.splice(start_num - 1, num_chars, new_text);

    return arr.join('');
  },

  REPLACEB: (...params) => TextFunctions.REPLACE(...params),

  REPT: (text, number_times) => {
    text = H.accept(text, Types.STRING);
    number_times = H.accept(number_times, Types.NUMBER);
    let str = '';

    for (let i = 0; i < number_times; i++) {
      str += text;
    }
    return str;
  },

  RIGHT: (text, numChars) =>
    prepostFix(text, numChars, (s, n) => s.slice(s.length - n)),

  RIGHTB: (...params) => TextFunctions.RIGHT(...params),

  SEARCH: (searchText, searchedInText, startingNum) => {
    const findText = H.accept(searchText, Types.STRING);
    const withinText = H.accept(searchedInText, null, undefined, false);
    const startNum = H.accept(startingNum, Types.NUMBER, 1);
    let withinTextArr;
    if (Array.isArray(withinText)) {
      withinTextArr = withinText;
    } else {
      withinTextArr = [[withinText]];
    }
    return withinTextArr.map((row) =>
      row.map((val) =>
        searchOnce(findText, H.accept(val, null, null), startNum, false)
      )
    );
  },

  SEARCHB: (...params) => TextFunctions.SEARCH(...params),

  /** *
   * @param text: The text to divide.
   * @param delimiter: The character or characters to use to split text.
   * @param split_by_each: OPTIONAL: Whether or not to divide text around each character contained in delimiter.
   * @param remove_empty_text: OPTIONAL:
   *                           Whether or not to remove empty text messages from the split results. The default behavior is to treat
   *                           consecutive delimiters as one (if TRUE). If FALSE, empty cells values are added between consecutive delimiters.
   * Google link: https://support.google.com/docs/answer/3094136?hl=en
   *
   ** */
  SPLIT: (text, delimiter, split_by_each = true, remove_empty_text = true) => {
    text = H.accept(text, Types.STRING);
    delimiter = H.accept(delimiter, Types.STRING);
    split_by_each = H.accept(split_by_each, Types.BOOLEAN);
    remove_empty_text = H.accept(remove_empty_text, Types.BOOLEAN);

    // Exit out of any ReGex special characters in both delimiter and text in order to preserve
    // the user's input sequence
    delimiter.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

    if (split_by_each) {
      delimiter = `[${delimiter}]`;
      delimiter = new RegExp(delimiter);
    }
    const rv = text.split(delimiter);

    if (!remove_empty_text) {
      return rv;
    }
    filteredRV = rv.filter((s) => s);
    return filteredRV;
  },

  SUBSTITUTE: (text, old_text, new_text, instance = null) => {
    const accepted_text = H.accept(text);
    const accepted_old = H.accept(old_text);
    const accepted_new = H.accept(new_text);
    let rv = null;
    original_text = null;
    changed_text = null;

    if (
      ['boolean', 'number', 'string'].includes(typeof accepted_text) &&
      ['boolean', 'number', 'string'].includes(typeof accepted_old) &&
      ['boolean', 'number', 'string'].includes(typeof accepted_new)
    ) {
      rv = H.changeEscapeCharacters(accepted_text.toString());
      original_text = H.changeEscapeCharacters(accepted_old.toString());
      changed_text = H.changeEscapeCharacters(accepted_new.toString());
    } else throw FormulaError.VALUE;
    const regexReplaceAll = WildCard.toRegex(original_text, 'g');

    if (instance === null) return rv.replace(regexReplaceAll, changed_text);

    let counter = H.accept(instance, Types.NUMBER);
    return rv.replace(regexReplaceAll, (t) => {
      counter -= 1;
      if (counter === 0) return changed_text;
      return t;
    });
  },

  T: (value) => {
    // extract the real parameter
    value = H.accept(value);
    if (typeof value === 'string') return value;
    return '';
  },

  TEXT: (value, formatText) => {
    value = H.accept(value, Types.NUMBER);
    formatText = H.accept(formatText, Types.STRING);
    // I know ssf contains bugs...
    try {
      return ssf.format(formatText, value);
    } catch (e) {
      console.error(e);
      throw FormulaError.VALUE;
    }
  },

  TEXTJOIN: (...params) => {},

  TRIM: (text) => {
    text = H.accept(text, [Types.STRING]);
    return text.replace(/^\s+|\s+$/g, '');
  },

  UNICHAR: (number) => {
    number = H.accept(number, [Types.NUMBER]);
    if (number <= 0) throw FormulaError.VALUE;
    return String.fromCharCode(number);
  },

  UNICODE: (text) => TextFunctions.CODE(text),
};

module.exports = TextFunctions;
