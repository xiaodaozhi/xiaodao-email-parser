import fs from 'node:fs/promises';
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
    return 'msg';
  }

  const headerSample = buffer.subarray(0, Math.min(buffer.length, 256 * 1024)).toString('latin1');
  const boundary = /\r?\n\r?\n/.exec(headerSample);
  if (!boundary) return null;

  let validHeaders = 0;
  let knownHeaders = 0;
  for (const line of headerSample.slice(0, boundary.index).split(/\r?\n/)) {
    if (/^[ \t]/.test(line)) continue;
    const match = /^([!-9;-~]+)\s*:/.exec(line);
    if (!match) continue;
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
