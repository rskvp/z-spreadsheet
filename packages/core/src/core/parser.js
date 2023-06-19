import FastFormulaParser, { FormulaHelpers } from 'fast-formula-parser';

export default class FormulaParser extends FastFormulaParser {
  constructor(table) {
    const parserConfig = {
      functions: {
        TIMEOUT: (text = 'message', seconds = 1) =>
          new Promise((resolve) => {
            setTimeout(() => {
              resolve(FormulaHelpers.accept(text));
            }, FormulaHelpers.accept(seconds) * 1000);
          }),
      },
      onCell: ({ sheet, row, col }) => {
        // Find index of sheet
        const sheetIndex = table
          .getData()
          .findIndex((item) => item.name === sheet);

        const cellText = table.cell(row - 1, col - 1, sheetIndex).text;
        if (/^\d+$/.test(cellText.replaceAll(' ', ''))) {
          const number = Number(cellText.replaceAll(' ', ''));
          if (Number.isNaN(number) === false) return cellText;
          return 1;
        }
        return 2;
      },
    };

    super(parserConfig);
  }
}
