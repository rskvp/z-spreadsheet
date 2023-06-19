const { expect } = require('chai');
const { FormulaParser } = require('../../grammar/hooks');
const FormulaError = require('../../formulas/error');

const parser = new FormulaParser({
  onCell: (ref) => 1,
  onRange: () => [[1, 2]],
  onStructuredReference: (
    tableName,
    columnName,
    thisRow,
    specialItem,
    sheet,
    position
  ) => {
    if (thisRow || specialItem) {
      // Single cell
      return { row: 2, col: 2 };
    }
    // Full column
    return {
      sheet: 'Sheet 1',
      from: {
        row: 1,
        col: 1,
      },
      to: {
        row: 10,
        col: 1,
      },
    };
  },
});

const position = { row: 1, col: 1, sheet: 'Sheet1' };

describe('Structured References', () => {
  it('should parse table and column reference', async () => {
    const actual = await parser.parseAsync(
      '=Table Name[@COLUMN_NAME]',
      position,
      true
    );
    expect(actual).to.eq(1);
  });

  it('thisRow will be false', async () => {
    const actual = await parser.parseAsync(
      '=TABLE[COLUMN_NAME]',
      position,
      true
    );
    expect(actual).to.deep.eq([[1, 2]]);
  });

  it('can detect columns without table', async () => {
    const actual = await parser.parseAsync('=[@COLUMN_NAME]', position, true);
    expect(actual).to.deep.eq(1);
  });

  it('can detect columns with special characters', async () => {
    const actual = await parser.parseAsync('=[@%COLUMN_NAME]', position, true);
    expect(actual).to.deep.eq(1);
  });

  it('can detect columns enclosed in []', async () => {
    const actual = await parser.parseAsync('=[@[Commission]]', position, true);
    expect(actual).to.deep.eq(1);
  });

  it('can parse single @', async () => {
    const actual = await parser.parseAsync('=TableName[@]', position, true);
    expect(actual).to.eq(1);
  });

  it('can parse headers', async () => {
    const actual = await parser.parseAsync(
      '=DeptSales[[#Headers],[Region]:[Commission Amount]]',
      position,
      true
    );
    expect(actual).to.eq(1);
  });

  it('can parse empty headers', async () => {
    const actual = await parser.parseAsync(
      '=DeptSales[[#Headers]]',
      position,
      true
    );
    expect(actual).to.eq(1);
  });
});
