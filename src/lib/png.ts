/*
 * Minimal PNG tEXt reader — enough to pull the embedded character card out of
 * a SillyTavern PNG (keyword `chara` = V2 base64 JSON, `ccv3` = V3 base64 JSON).
 */

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

export function readPngTextChunks(bytes: Uint8Array): Record<string, string> {
  for (let i = 0; i < PNG_SIGNATURE.length; i++) {
    if (bytes[i] !== PNG_SIGNATURE[i]) throw new Error('Not a PNG file');
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const out: Record<string, string> = {};
  let pos = 8;
  while (pos + 8 <= bytes.length) {
    const length = view.getUint32(pos);
    const type = String.fromCharCode(bytes[pos + 4], bytes[pos + 5], bytes[pos + 6], bytes[pos + 7]);
    const dataStart = pos + 8;
    if (type === 'tEXt' && dataStart + length <= bytes.length) {
      const data = bytes.subarray(dataStart, dataStart + length);
      const nul = data.indexOf(0);
      if (nul > 0) {
        const keyword = latin1(data.subarray(0, nul));
        out[keyword] = latin1(data.subarray(nul + 1));
      }
    }
    if (type === 'IEND') break;
    pos = dataStart + length + 4; // skip data + CRC
  }
  return out;
}

function latin1(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
  return s;
}

/** base64 → UTF-8 string (card payloads are base64-encoded UTF-8 JSON). */
export function base64ToUtf8(b64: string): string {
  const bin = atob(b64.replace(/\s/g, ''));
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}
