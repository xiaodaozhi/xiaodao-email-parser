import fs from 'node:fs/promises';
import PostalMime from 'postal-mime';
import type { PostalMimeOptions } from 'postal-mime';
import {
  addressToRecipients,
  createPreview,
  extractRawHeaders,
  headersToRecord,
  normalizeContentId,
  normalizeSubject,
  parseImportance,
  receivedDateFromHeaders,
  toValidDate,
} from './email-utils.js';
import type { EmlParserOptions, ParsedEmail } from './types.js';

const DEFAULT_MAX_INPUT_SIZE = 100 * 1024 * 1024;

/** Parser for RFC 5322 / MIME `.eml` messages. */
class EmlParser {
  private readonly maxInputSize: number;
  private readonly postalOptions: PostalMimeOptions;

  public constructor(options: EmlParserOptions = {}) {
    const { maxInputSize = DEFAULT_MAX_INPUT_SIZE, ...postalOptions } = options;
    if (!Number.isSafeInteger(maxInputSize) || maxInputSize <= 0) {
      throw new RangeError('maxInputSize must be a positive safe integer.');
    }
    this.maxInputSize = maxInputSize;
    this.postalOptions = { attachmentEncoding: 'arraybuffer', ...postalOptions };
  }

  public async parse(buffer: Buffer): Promise<ParsedEmail> {
    if (!Buffer.isBuffer(buffer)) {
      throw new TypeError('Input must be a Buffer. Use parseFile() for file paths.');
    }
    if (buffer.length > this.maxInputSize) {
      throw new RangeError(`EML input exceeds the configured ${this.maxInputSize}-byte limit.`);
    }

    let email;
    try {
      email = await PostalMime.parse(buffer, this.postalOptions);
    } catch (error: unknown) {
      const detail = error instanceof Error ? error.message : String(error);
      throw new Error(`Invalid .eml file: ${detail}`, { cause: error });
    }

    const parsedHeaders = headersToRecord(email.headers);
    const from = addressToRecipients(email.from)[0] ?? null;
    const body = email.text || null;
    const bodyHtml = email.html || null;

    return {
      format: 'eml',
      subject: email.subject || null,
      from,
      replyTo: addressToRecipients(email.replyTo),
      to: addressToRecipients(email.to),
      cc: addressToRecipients(email.cc),
      bcc: addressToRecipients(email.bcc),
      body,
      bodyHtml,
      bodyRtf: null,
      attachments: email.attachments.map(attachment => {
        const content = typeof attachment.content === 'string'
          ? Buffer.from(attachment.content, attachment.encoding === 'base64' ? 'base64' : 'utf8')
          : attachment.content instanceof ArrayBuffer
            ? Buffer.from(new Uint8Array(attachment.content))
            : Buffer.from(attachment.content);
        return {
          filename: attachment.filename || null,
          content,
          contentType: attachment.mimeType || 'application/octet-stream',
          contentId: normalizeContentId(attachment.contentId),
          contentLocation: null,
          size: content.length,
        };
      }),
      sentDate: toValidDate(email.date),
      receivedDate: receivedDateFromHeaders(email.headers),
      createdDate: null,
      modifiedDate: null,
      messageClass: null,
      importance: parseImportance(parsedHeaders),
      messageSize: buffer.length,
      conversationTopic: parsedHeaders?.['thread-topic'] || null,
      normalizedSubject: normalizeSubject(email.subject),
      messageId: email.messageId || null,
      headers: extractRawHeaders(buffer),
      parsedHeaders,
      preview: createPreview(body, bodyHtml),
      _rawProperties: {
        sender: email.sender ?? null,
        deliveredTo: email.deliveredTo ?? null,
        returnPath: email.returnPath ?? null,
        inReplyTo: email.inReplyTo ?? null,
        references: email.references ?? null,
        headers: email.headers,
      },
    };
  }

  public async parseFile(filePath: string): Promise<ParsedEmail> {
    return this.parse(await fs.readFile(filePath));
  }
}

export default EmlParser;
