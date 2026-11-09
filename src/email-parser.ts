import fs from 'node:fs/promises';
import CFB from 'cfb';
import EmlParser from './eml-parser.js';
import MsgParser from './msg-parser.js';
import type {
  EmailFormat,
  EmailParserOptions,
  ParseOptions,
  ParsedEmail,
} from './types.js';

const CFB_SIGNATURE = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
const KNOWN_EML_HEADERS = new Set([
  'bcc', 'cc', 'content-transfer-encoding', 'content-type', 'date', 'from',
  'message-id', 'mime-version', 'received', 'reply-to', 'sender', 'subject', 'to',
]);

export function detectEmailFormat(buffer: Buffer): EmailFormat | null {
  if (!Buffer.isBuffer(buffer)) {
    throw new TypeError('Input must be a Buffer.');
  }
  if (buffer.length >= CFB_SIGNATURE.length && buffer.subarray(0, 8).equals(CFB_SIGNATURE)) {
    // A CFB signature is shared by Outlook MSG, Word DOC and several other
    // compound documents. Confirm the MAPI property stream when the container
    // is readable, while preserving a useful MSG error for truncated inputs.
    if (buffer.length < 512) return 'msg';
    try {
      const cfb = CFB.read(buffer, { type: 'buffer' });
      const rootIndex = cfb.FileIndex.findIndex(entry => entry.type === 5);
      const root = cfb.FullPaths[rootIndex]?.replace(/\/?$/, '/');
      const hasMapiProperties = root && cfb.FullPaths.some(path =>
        path.toLowerCase() === `${root}__properties_version1.0`.toLowerCase(),
      );
      return hasMapiProperties ? 'msg' : null;
    } catch {
      return 'msg';
    }
  }

  const sampleSize = Math.min(buffer.length, 256 * 1024);
  const headerSample = buffer.subarray(0, sampleSize).toString('latin1');
  const boundary = /\r?\n\r?\n/.exec(headerSample);
  // A header-only RFC 5322 message may end immediately after the last header
  // line, without the empty line that would normally separate the body.
  if (!boundary && sampleSize !== buffer.length) return null;

  const headerEnd = boundary?.index ?? headerSample.length;

  let validHeaders = 0;
  let knownHeaders = 0;
  let hasHeader = false;
  let canUseMboxSeparator = true;
  for (const line of headerSample.slice(0, headerEnd).split(/\r?\n/)) {
    if (!line) continue;
    if (/^[ \t]/.test(line)) {
      if (!hasHeader) return null;
      continue;
    }
    const match = /^([!-9;-~]+)\s*:/.exec(line);
    if (!match) {
      // Accept the optional Unix mbox envelope line before RFC 5322 headers.
      if (canUseMboxSeparator && /^From\s+\S/.test(line)) {
        canUseMboxSeparator = false;
        continue;
      }
      return null;
    }
    hasHeader = true;
    canUseMboxSeparator = false;
    validHeaders++;
    if (KNOWN_EML_HEADERS.has(match[1].toLowerCase())) knownHeaders++;
  }
  return validHeaders > 0 && knownHeaders > 0 ? 'eml' : null;
}

/** Auto-detecting parser for both Outlook MSG and RFC/MIME EML files. */
class EmailParser {
  private readonly msgParser: MsgParser;
  private readonly emlParser: EmlParser;

  public constructor(options: EmailParserOptions = {}) {
    this.msgParser = new MsgParser();
    this.emlParser = new EmlParser(options.eml);
  }

  public async parse(buffer: Buffer, options: ParseOptions = {}): Promise<ParsedEmail> {
    if (!Buffer.isBuffer(buffer)) {
      throw new TypeError('Input must be a Buffer. Use parseFile() for file paths.');
    }

    const requestedFormat = options.format ?? 'auto';
    if (requestedFormat !== 'auto' && requestedFormat !== 'msg' && requestedFormat !== 'eml') {
      throw new RangeError(`Unsupported format option: ${String(requestedFormat)}.`);
    }
    const format = requestedFormat === 'auto' ? detectEmailFormat(buffer) : requestedFormat;
    if (!format) {
      throw new Error('Unsupported email format: expected an Outlook .msg or RFC/MIME .eml file.');
    }
    return format === 'msg' ? this.msgParser.parse(buffer) : this.emlParser.parse(buffer);
  }

  public async parseFile(filePath: string, options: ParseOptions = {}): Promise<ParsedEmail> {
    return this.parse(await fs.readFile(filePath), options);
  }
}

export default EmailParser;
