const { expect } = require('chai');
const Collection = require('../../grammar/type/collection');

describe('Collection', () => {
  it('should throw error', () => {
    expect(
      () => new Collection([], [{ row: 1, col: 1, sheet: 'Sheet1' }])
    ).to.throw('Collection: data length should match references length.');
  });

  it('should not throw error', () => {
    expect(
      () => new Collection([1], [{ row: 1, col: 1, sheet: 'Sheet1' }])
    ).to.not.throw();
  });
});
