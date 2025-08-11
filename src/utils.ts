const hexLo = "0123456789abcdef";
const hexHi = "0123456789ABCDEF";

/**
 * little uint32 to hex string
 * 
 * @internal
 */
export function toHex4(value: number, upper?: boolean): string {
  const tbl = upper ? hexHi : hexLo;
  return tbl[(value >> 28) & 15]
    + tbl[(value >> 24) & 15]
    + tbl[(value >> 20) & 15]
    + tbl[(value >> 16) & 15]
    + tbl[(value >> 12) & 15]
    + tbl[(value >> 8) & 15]
    + tbl[(value >> 4) & 15]
    + tbl[(value) & 15];
}

/**
 * little uint16 to hex string
 * 
 * @internal
 */
export function toHex2(value: number, upper?: boolean): string {
  const tbl = upper ? hexHi : hexLo;
  return tbl[(value >> 12) & 15]
    + tbl[(value >> 8) & 15]
    + tbl[(value >> 4) & 15]
    + tbl[(value) & 15];
}
