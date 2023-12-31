const { expect } = require('chai');
const { DepParser } = require('../../grammar/dependency/hooks');

const depParser = new DepParser({
  onVariable: (variable) =>
    variable === 'aaaa'
      ? { from: { row: 1, col: 1 }, to: { row: 2, col: 2 } }
      : { row: 1, col: 1 },
});
const position = { row: 1, col: 1, sheet: 'Sheet1' };

describe('Dependency parser', () => {
  it('parse SUM(1,)', () => {
    const actual = depParser.parse('=SUM(1,)', position);
    expect(actual).to.deep.eq([]);
  });

  it('should parse single cell', () => {
    let actual = depParser.parse('=A1', position);
    expect(actual).to.deep.eq([position]);
    actual = depParser.parse('=A1+1', position);
    expect(actual).to.deep.eq([position]);
  });

  it('should parse the same cell/range once', () => {
    let actual = depParser.parse('=A1+A1+A1', position);
    expect(actual).to.deep.eq([position]);
    actual = depParser.parse('=A1:C3+A1:C3+A1:C3', position);
    expect(actual).to.deep.eq([
      {
        from: { row: 1, col: 1 },
        to: { row: 3, col: 3 },
        sheet: position.sheet,
      },
    ]);

    actual = depParser.parse('=A1:C3+A1:C3+A1:C3+A1+B1', position);
    expect(actual).to.deep.eq([
      {
        from: { row: 1, col: 1 },
        to: { row: 3, col: 3 },
        sheet: position.sheet,
      },
    ]);
  });

  it('should parse ranges', () => {
    let actual = depParser.parse('=A1:C3', position);
    expect(actual).to.deep.eq([
      { sheet: 'Sheet1', from: { row: 1, col: 1 }, to: { row: 3, col: 3 } },
    ]);
    actual = depParser.parse('=A:C', position);
    expect(actual).to.deep.eq([
      {
        sheet: 'Sheet1',
        from: { row: 1, col: 1 },
        to: { row: 1048576, col: 3 },
      },
    ]);
    actual = depParser.parse('=1:3', position);
    expect(actual).to.deep.eq([
      { sheet: 'Sheet1', from: { row: 1, col: 1 }, to: { row: 3, col: 16384 } },
    ]);
  });

  it('should parse variable', () => {
    const actual = depParser.parse('=aaaa', position);
    expect(actual).to.deep.eq([
      { sheet: 'Sheet1', from: { row: 1, col: 1 }, to: { row: 2, col: 2 } },
    ]);
  });

  it('should parse basic formulas', () => {
    // data types
    let actual = depParser.parse('=TRUE+A1+#VALUE!+{1}', position);
    expect(actual).to.deep.eq([{ sheet: 'Sheet1', row: 1, col: 1 }]);

    // function without args
    actual = depParser.parse('=A1+FUN()', position);
    expect(actual).to.deep.eq([{ sheet: 'Sheet1', row: 1, col: 1 }]);

    // prefix
    actual = depParser.parse('=++A1', position);
    expect(actual).to.deep.eq([{ sheet: 'Sheet1', row: 1, col: 1 }]);

    // postfix
    actual = depParser.parse('=A1%', position);
    expect(actual).to.deep.eq([{ sheet: 'Sheet1', row: 1, col: 1 }]);

    // intersect
    actual = depParser.parse('=A1:A3 A3:B3', position);
    expect(actual).to.deep.eq([{ sheet: 'Sheet1', row: 3, col: 1 }]);

    // union
    actual = depParser.parse('=(A1:C1, A2:E9)', position);
    expect(actual).to.deep.eq([
      { sheet: 'Sheet1', from: { row: 1, col: 1 }, to: { row: 1, col: 3 } },
      { sheet: 'Sheet1', from: { row: 2, col: 1 }, to: { row: 9, col: 5 } },
    ]);
  });

  it('should parse complex formula', () => {
    const actual = depParser.parse(
      '=IF(MONTH($K$1)<>MONTH($K$1-(WEEKDAY($K$1,1)-(start_day-1))-IF((WEEKDAY($K$1,1)-(start_day-1))<=0,7,0)+(ROW(O5)-ROW($K$3))*7+(COLUMN(O5)-COLUMN($K$3)+1)),"",$K$1-(WEEKDAY($K$1,1)-(start_day-1))-IF((WEEKDAY($K$1,1)-(start_day-1))<=0,7,0)+(ROW(O5)-ROW($K$3))*7+(COLUMN(O5)-COLUMN($K$3)+1))',
      position
    );
    expect(actual).to.deep.eq([
      {
        col: 11,
        row: 1,
        sheet: 'Sheet1',
      },
      {
        col: 1,
        row: 1,
        sheet: 'Sheet1',
      },
      {
        col: 15,
        row: 5,
        sheet: 'Sheet1',
      },
      {
        col: 11,
        row: 3,
        sheet: 'Sheet1',
      },
    ]);
  });

  it('should not throw error', () => {
    expect(() => depParser.parse('=SUM(1))', position, true)).to.not.throw();

    expect(() => depParser.parse('=SUM(1+)', position, true)).to.not.throw();

    expect(() => depParser.parse('=SUM(1+)', position, true)).to.not.throw();
  });
});
