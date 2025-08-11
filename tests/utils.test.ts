import { toHex4 } from "../src/utils";

describe('toHex4', () => {
  it("0", () => { expect(toHex4(0)).toBe("00000000"); });
  it("1", () => { expect(toHex4(1)).toBe("00000001"); });
  it("255", () => { expect(toHex4(255)).toBe("000000ff"); });
  it("4294967295", () => { expect(toHex4(4294967295)).toBe("ffffffff"); });
});
