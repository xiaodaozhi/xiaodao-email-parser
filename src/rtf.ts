// Initial LZFu dictionary defined by [MS-OXRTFCP].
const RTF_DICTIONARY = Buffer.from(
  '{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}' +
  '{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript ' +
  '\\fdecor MS Sans SerifSymbolArialTimes New RomanCourier}' +
  '{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain' +
  '\\f0\\fs20\\b\\i\\u\\tab\\tx',
  'ascii',
);

const COMPRESSED_MAGIC = 0x75465a4c;
const UNCOMPRESSED_MAGIC = 0x414c454d;

/** Decode PidTagRtfCompressed. Returns null for malformed or unknown data. */
export function decompressRtf(input: Buffer): string | null {
  if (input.length < 16) return null;

  const rawSize = input.readUInt32LE(4);
  const magic = input.readUInt32LE(8);
  if (rawSize === 0 || rawSize > 256 * 1024 * 1024) return null;

  if (magic === UNCOMPRESSED_MAGIC) {
    return input.subarray(16, Math.min(16 + rawSize, input.length)).toString('latin1').replace(/\0+$/, '');
  }
  if (magic !== COMPRESSED_MAGIC) return null;

  const dictionary = Buffer.alloc(4096);
  RTF_DICTIONARY.copy(dictionary, 0, 0, Math.min(RTF_DICTIONARY.length, dictionary.length));
  let dictionaryOffset = RTF_DICTIONARY.length & 0x0fff;
  let inputOffset = 16;
  const output = Buffer.alloc(rawSize);
  let outputOffset = 0;

  while (inputOffset < input.length && outputOffset < rawSize) {
    const flags = input[inputOffset++];
    for (let bit = 0; bit < 8 && outputOffset < rawSize && inputOffset < input.length; bit++) {
      if ((flags & (1 << bit)) === 0) {
        const value = input[inputOffset++];
        output[outputOffset++] = value;
        dictionary[dictionaryOffset] = value;
        dictionaryOffset = (dictionaryOffset + 1) & 0x0fff;
        continue;
      }

      if (inputOffset + 1 >= input.length) break;
      const first = input[inputOffset++];
      const second = input[inputOffset++];
      let referenceOffset = ((first << 4) | (second >> 4)) & 0x0fff;
      const length = (second & 0x0f) + 2;

      for (let index = 0; index < length && outputOffset < rawSize; index++) {
        const value = dictionary[referenceOffset];
        referenceOffset = (referenceOffset + 1) & 0x0fff;
        output[outputOffset++] = value;
        dictionary[dictionaryOffset] = value;
        dictionaryOffset = (dictionaryOffset + 1) & 0x0fff;
      }
    }
  }

  return outputOffset === rawSize ? output.toString('latin1').replace(/\0+$/, '') : null;
}
