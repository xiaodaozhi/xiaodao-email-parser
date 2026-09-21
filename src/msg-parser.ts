import CFB from 'cfb';
import fs from 'node:fs';
import {
  PROPERTY_IDS,
  PROPERTY_TYPES,
  RECIPIENT_TYPES,
  IMPORTANCE,
  FILETIME_EPOCH_OFFSET,
} from './constants.js';
import {
  createPreview,
  decodeBuffer,
  normalizeContentId,
  normalizeSubject,
  parseAddressHeader,
  parseHeaderBlock,
  toValidDate,
} from './email-utils.js';
import { decompressRtf } from './rtf.js';
import type {
  ParsedEmail,
  EmailRecipient,
  RecipientInfo,
  EmailAttachmentInfo,
  RawEmailData,
  EntryStore,
  SubStorageBucket,
} from './types.js';

/**
 * Parser for Outlook .msg files (Compound File Binary / OLE2 format).
 *
 * Usage:
 *   const parser = new MsgParser();
 *   const email = parser.parseFile('message.msg');
 *   console.log(email.subject, email.from);
 */
class MsgParser {
  private cfb!: ReturnType<typeof CFB.read>;
  private sourceSize = 0;
  private _entries: EntryStore = {
    topPropsStream: null,
    topSubstg: {},
    recipients: {},
    attachments: {},
  };

  // ------------------------------------------------------------------
  //  Public API
  // ------------------------------------------------------------------

  /**
   * Parse a .msg file from a Buffer.
   */
  public parse(buffer: Buffer): ParsedEmail {
    if (!Buffer.isBuffer(buffer)) {
      throw new TypeError('Input must be a Buffer. Use parseFile() for file paths.');
    }
    this.sourceSize = buffer.length;
    try {
      this.cfb = CFB.read(buffer, { type: 'buffer' });
    } catch (err: unknown) {
      const detail = err instanceof Error ? err.message : String(err);
      throw new Error(`Invalid .msg file: ${detail}`, { cause: err });
    }
    this._classifyEntries();
    if (!this._entries.topPropsStream || this._entries.topPropsStream.length < 32) {
      throw new Error('Invalid .msg file: the root MAPI property stream is missing or truncated.');
    }
    return this._assembleMessage();
  }

  /**
   * Parse a .msg file from a file path (synchronous).
   */
  public parseFile(filePath: string): ParsedEmail {
    return this.parse(fs.readFileSync(filePath));
  }

  // ------------------------------------------------------------------
  //  CFB entry classification
  // ------------------------------------------------------------------

  private _classifyEntries(): void {
    this._entries = {
      topPropsStream: null,
      topSubstg: {},
      recipients: {},
      attachments: {},
    };

    // Find the root entry's path — CFB may use '/' or 'Root Entry/' etc.
    const rootPath = this._findRootPath();

    for (let i = 0; i < this.cfb.FullPaths.length; i++) {
      const path = this.cfb.FullPaths[i];
      const entry = this.cfb.FileIndex[i];

      if (!entry || !path) continue;
      if (!entry.content || entry.type === 5) continue;

      const name = this._entryName(path);
      const content = Buffer.from(entry.content);

      // Determine parent directory path
      const parentPath = this._parentPath(path);

      if (this._isRootChild(parentPath, rootPath)) {
        // Direct child of root → top-level property or substg stream
        this._classifyTopLevel(name, content);
      } else {
        // Child of a sub-storage → recipient or attachment property
        const storageName = this._entryName(parentPath);
        if (storageName.startsWith('__recip_version1.0_#')) {
          const idx = storageName.slice('__recip_version1.0_#'.length);
          this._addToBucket(this._entries.recipients, idx, name, content);
        } else if (storageName.startsWith('__attach_version1.0_#')) {
          const idx = storageName.slice('__attach_version1.0_#'.length);
          this._addToBucket(this._entries.attachments, idx, name, content);
        }
      }
    }
  }

  /**
   * Find the root entry's FullPath from the CFB structure.
   * Root entry always has type === 5.
   */
  private _findRootPath(): string {
    for (let i = 0; i < this.cfb.FileIndex.length; i++) {
      if (this?.cfb?.FileIndex?.[i]?.type === 5) {
        let rootPath = this.cfb.FullPaths[i];
        if (!rootPath) continue;
        if (!rootPath.endsWith('/')) rootPath += '/';
        return rootPath;
      }
    }
    return '/'; // fallback
  }

  /**
   * Check whether `parentPath` is the root (i.e. the entry is a top-level child).
   */
  private _isRootChild(parentPath: string, rootPath: string): boolean {
    if (parentPath === '/') return true;
    const normalized = parentPath.endsWith('/') ? parentPath : parentPath + '/';
    return normalized === rootPath;
  }

  private _classifyTopLevel(name: string, content: Buffer): void {
    if (name === '__properties_version1.0') {
      this._entries.topPropsStream = content;
    } else if (name.startsWith('__substg1.0_')) {
      const tag = name.slice('__substg1.0_'.length);
      this._entries.topSubstg[tag] = content;
    }
  }

  private _addToBucket(
    bucket: Record<string, SubStorageBucket>,
    idx: string,
    name: string,
    content: Buffer,
  ): void {
    if (!bucket[idx]) bucket[idx] = { props: null, substg: {} };
    if (name === '__properties_version1.0') {
      bucket[idx].props = content;
    } else if (name.startsWith('__substg1.0_')) {
      const tag = name.slice('__substg1.0_'.length);
      bucket[idx].substg[tag] = content;
    }
  }

  // ------------------------------------------------------------------
  //  Message assembly
  // ------------------------------------------------------------------

  private _assembleMessage(): ParsedEmail {
    const raw = this._collectAllProperties();
    return this._buildResult(raw);
  }

  private _collectAllProperties(): RawEmailData {
    const tags: Record<string, unknown> = {};

    if (this._entries.topPropsStream) {
      const streamProps = this._readPropertyStream(this._entries.topPropsStream, 32);
      Object.assign(tags, streamProps);
    }

    for (const [tag, content] of Object.entries(this._entries.topSubstg)) {
      const value = this._readSubstgValue(content, tag);
      if (value !== undefined) tags[tag] = value;
    }

    return {
      properties: this._resolveProperties(tags),
      recipients: this._parseRecipients(),
      attachments: this._parseAttachments(),
      size: this.sourceSize,
    };
  }

  // ------------------------------------------------------------------
  //  Property stream reader (__properties_version1.0)
  // ------------------------------------------------------------------

  private _readPropertyStream(buf: Buffer, headerSize = 8): Record<string, unknown> {
    const result: Record<string, unknown> = {};
    let offset = headerSize;

    // [MS-OXMSG] defines every property entry as exactly 16 bytes. Variable
    // values live in sibling __substg1.0_* streams; only fixed values are
    // stored in the final eight bytes of an entry.
    while (offset + 16 <= buf.length) {
      const propertyTag = buf.readUInt32LE(offset);
      if (propertyTag === 0) break;

      const propId = (propertyTag >>> 16) & 0xffff;
      const propType = propertyTag & 0xffff;
      const tagStr = this._makeTagStr(propId, propType);
      const info = PROPERTY_TYPES[propType];

      if (info?.fixedSize && info.fixedSize > 0 && info.fixedSize <= 8) {
        const value = this._readFixedValue(buf, offset + 8, propType, info.fixedSize);
        if (value !== undefined) result[tagStr] = value;
      }
      offset += 16;
    }

    return result;
  }

  private _readFixedValue(buf: Buffer, offset: number, type: number, size: number): unknown | undefined {
    if (offset < 0 || offset + size > buf.length) return undefined;
    let value: unknown;
    switch (type) {
      case 0x0002:
        value = buf.readInt16LE(offset);
        break;
      case 0x0003:
        value = buf.readInt32LE(offset);
        break;
      case 0x0004:
        value = buf.readFloatLE(offset);
        break;
      case 0x0005:
        value = buf.readDoubleLE(offset);
        break;
      case 0x0006:
        value = Number(buf.readBigInt64LE(offset));
        break;
      case 0x000b:
        value = buf.readUInt16LE(offset) !== 0;
        break;
      case 0x0014:
        value = Number(buf.readBigInt64LE(offset));
        break;
      case 0x0040:
        value = Number(buf.readBigUInt64LE(offset));
        break;
      default:
        value = buf.slice(offset, offset + size);
    }
    return value;
  }

  // ------------------------------------------------------------------
  //  Individual property stream reader (__substg1.0_XXXXXXXX)
  // ------------------------------------------------------------------

  private _readSubstgValue(content: Buffer, tag: string): unknown {
    const type = parseInt(tag.slice(4), 16);

    switch (type) {
      case 0x001e:
        return content.toString('latin1').replace(/\0+$/, '');
      case 0x001f:
        return content.toString('ucs2').replace(/\0+$/, '');
      case 0x0102:
        return Buffer.from(content);
      case 0x0003:
        return content.length >= 4 ? content.readInt32LE(0) : undefined;
      case 0x0002:
        return content.length >= 2 ? content.readInt16LE(0) : undefined;
      case 0x000b:
        return content.length >= 2 ? content.readUInt16LE(0) !== 0 : undefined;
      case 0x0014:
        return content.length >= 8 ? Number(content.readBigInt64LE(0)) : undefined;
      case 0x0040:
        return content.length >= 8 ? Number(content.readBigUInt64LE(0)) : undefined;
      case 0x0004:
        return content.length >= 4 ? content.readFloatLE(0) : undefined;
      case 0x0005:
        return content.length >= 8 ? content.readDoubleLE(0) : undefined;
      default:
        return content;
    }
  }

  // ------------------------------------------------------------------
  //  Resolution: raw tags → named properties
  // ------------------------------------------------------------------

  private _resolveProperties(tags: Record<string, unknown>): Record<string, unknown> {
    const named: Record<string, { value: unknown; type: number }> = {};

    for (const [tag, value] of Object.entries(tags)) {
      if (tag.length < 8) continue;
      const propId = tag.slice(0, 4).toLowerCase();
      const propType = parseInt(tag.slice(4), 16);
      const name = PROPERTY_IDS[propId];
      if (!name) continue;

      const existing = named[name];
      if (!existing || this._preferNew(propType, existing.type)) {
        named[name] = { value, type: propType };
      }
    }

    const result: Record<string, unknown> = {};
    const codePageEntry = named.internetCodePage?.value;
    const codePage = typeof codePageEntry === 'number' ? codePageEntry : undefined;
    for (const [name, { value, type }] of Object.entries(named)) {
      result[name] = type === 0x001e && typeof value === 'string' && codePage
        ? decodeBuffer(Buffer.from(value, 'latin1'), codePage)
        : value;

      if (name.endsWith('Time') && typeof result[name] === 'number' && result[name] > 0) {
        result[name] = this._fileTimeToDate(result[name] as number);
      }
    }

    if (result.importance !== undefined && typeof result.importance === 'number') {
      const mapped = IMPORTANCE[result.importance];
      if (mapped !== undefined) {
        result.importance = mapped;
      }
    }

    const senderName = result.sentRepresentingName || result.senderName;
    const senderEmail = result.sentRepresentingSmtpEmail || result.sentRepresentingEmail ||
      result.senderSmtpEmail || result.senderEmail;
    if (senderName || senderEmail) {
      result.from = {
        name: (senderName as string) || null,
        email: (senderEmail as string) || null,
      };
    }

    return result;
  }

  private _preferNew(newType: number, existingType: number): boolean {
    const unicode = 0x001f;
    const ansi = 0x001e;
    if (newType === unicode && existingType !== unicode) return true;
    if (existingType === unicode && newType !== unicode) return false;
    return true;
  }

  // ------------------------------------------------------------------
  //  Recipients
  // ------------------------------------------------------------------

  _parseRecipients(): RecipientInfo[] {
    const result: RecipientInfo[] = [];

    for (const [, bucket] of Object.entries(this._entries.recipients)) {
      const tags: Record<string, unknown> = {};

      if (bucket.props) {
        Object.assign(tags, this._readPropertyStream(bucket.props));
      }
      for (const [tag, content] of Object.entries(bucket.substg)) {
        const value = this._readSubstgValue(content, tag);
        if (value !== undefined) tags[tag] = value;
      }

      const props = this._resolveProperties(tags);

      const recipient: RecipientInfo = {
        name: (props.recipientDisplayName as string) || (props.displayName as string) || null,
        email: (props.smtpEmailAddress as string) || (props.emailAddress as string) || null,
        type: RECIPIENT_TYPES[props.recipientType as number] || null,
      };

      result.push(recipient);
    }

    const order: Record<string, number> = { to: 0, cc: 1, bcc: 2 };
    result.sort((a, b) => (order[a.type ?? ''] ?? 9) - (order[b.type ?? ''] ?? 9));

    return result;
  }

  // ------------------------------------------------------------------
  //  Attachments
  // ------------------------------------------------------------------

  _parseAttachments(): EmailAttachmentInfo[] {
    const result: EmailAttachmentInfo[] = [];

    for (const [, bucket] of Object.entries(this._entries.attachments)) {
      const tags: Record<string, unknown> = {};

      if (bucket.props) {
        Object.assign(tags, this._readPropertyStream(bucket.props));
      }
      for (const [tag, content] of Object.entries(bucket.substg)) {
        const value = this._readSubstgValue(content, tag);
        if (value !== undefined) tags[tag] = value;
      }

      const props = this._resolveProperties(tags);
      const data = Buffer.isBuffer(props.attachmentData) ? props.attachmentData : null;

      const attachment: EmailAttachmentInfo = {
        filename: (props.attachmentLongFilename as string) || (props.attachmentFilename as string) || null,
        content: data,
        contentType: (props.attachmentMimeType as string) || 'application/octet-stream',
        contentId: normalizeContentId(props.attachmentContentId as string | undefined),
        contentLocation: (props.attachmentContentLocation as string) || null,
        size: data ? data.length : (props.attachmentSize as number) || 0,
      };

      result.push(attachment);
    }

    return result;
  }

  // ------------------------------------------------------------------
  //  Result builder
  // ------------------------------------------------------------------

  _buildResult(raw: RawEmailData): ParsedEmail {
    const p = raw.properties;
    const recipients = raw.recipients;
    const attachments = raw.attachments;

    let to = recipients.filter(r => r.type === 'to');
    let cc = recipients.filter(r => r.type === 'cc');
    let bcc = recipients.filter(r => r.type === 'bcc');

    let rawHeaders: string | null = null;
    let parsedHeaders: Record<string, string> | null = null;
    if (p.internetHeaders) {
      if (typeof p.internetHeaders === 'string') {
        rawHeaders = p.internetHeaders;
      } else if (Buffer.isBuffer(p.internetHeaders)) {
        rawHeaders = this._decodeBinaryHeaders(p.internetHeaders);
      }
      parsedHeaders = rawHeaders ? this._parseHeaders(rawHeaders) : null;
    }

    let headerFrom: EmailRecipient | null = null;
    let replyTo: EmailRecipient[] = [];
    if (parsedHeaders && Object.keys(parsedHeaders).length > 0) {
      const fromHeaders = this._parseRecipientsFromHeaders(parsedHeaders);
      to = this._mergeRecipients(to, fromHeaders.to);
      cc = this._mergeRecipients(cc, fromHeaders.cc);
      bcc = this._mergeRecipients(bcc, fromHeaders.bcc);
      headerFrom = parseAddressHeader(parsedHeaders.from)[0] ?? null;
      replyTo = parseAddressHeader(parsedHeaders['reply-to']);
    }

    const codePage = typeof p.internetCodePage === 'number' ? p.internetCodePage : undefined;
    const body = typeof p.body === 'string'
      ? p.body
      : Buffer.isBuffer(p.body) ? decodeBuffer(p.body, codePage) : null;
    const bodyHtml = typeof p.bodyHtml === 'string'
      ? p.bodyHtml
      : Buffer.isBuffer(p.bodyHtml) ? decodeBuffer(p.bodyHtml, codePage) : null;
    const bodyRtf = Buffer.isBuffer(p.bodyRtfCompressed)
      ? decompressRtf(p.bodyRtfCompressed)
      : typeof p.bodyRtfCompressed === 'string' ? p.bodyRtfCompressed : null;

    if (to.length === 0 && body) {
      const bodyHeaders = this._parseHeadersFromBody(body);
      if (bodyHeaders) {
        const fromBody = this._parseRecipientsFromHeaders(bodyHeaders);
        to = this._mergeRecipients(to, fromBody.to);
        cc = this._mergeRecipients(cc, fromBody.cc);
        bcc = this._mergeRecipients(bcc, fromBody.bcc);
      }
    }

    const subject = (p.subject as string) || parsedHeaders?.subject || null;
    const importance = p.importance === 'low' || p.importance === 'high' || p.importance === 'normal'
      ? p.importance
      : 'normal';

    return {
      format: 'msg',
      subject,
      from: this._mergeSender((p.from as EmailRecipient) || null, headerFrom),
      replyTo,
      to: to.length > 0 ? to.map(r => ({ name: r.name, email: r.email })) : [],
      cc: cc.length > 0 ? cc.map(r => ({ name: r.name, email: r.email })) : [],
      bcc: bcc.length > 0 ? bcc.map(r => ({ name: r.name, email: r.email })) : [],
      body,
      bodyHtml,
      bodyRtf,
      attachments,
      sentDate: (p.clientSubmitTime as Date) || toValidDate(parsedHeaders?.date),
      receivedDate: (p.messageDeliveryTime as Date) || null,
      createdDate: (p.creationTime as Date) || null,
      modifiedDate: (p.lastModificationTime as Date) || null,
      messageClass: (p.messageClass as string) || null,
      importance,
      messageSize: typeof p.messageSize === 'number' ? p.messageSize : (raw.size ?? null),
      conversationTopic: (p.conversationTopic as string) || null,
      normalizedSubject: (p.normalizedSubject as string) || normalizeSubject(subject),
      messageId: parsedHeaders?.['message-id'] || null,
      headers: rawHeaders,
      parsedHeaders,
      preview: (p.preview as string) || createPreview(body, bodyHtml),
      _rawProperties: p,
    };
  }

  private _mergeRecipients(primary: RecipientInfo[], fallback: RecipientInfo[]): RecipientInfo[] {
    if (primary.length === 0) return fallback;
    if (fallback.length !== primary.length) return primary;
    return primary.map((recipient, index) => ({
      name: recipient.name || fallback[index]?.name || null,
      email: this._isSmtpAddress(recipient.email)
        ? recipient.email
        : fallback[index]?.email || recipient.email || null,
      type: recipient.type,
    }));
  }

  private _mergeSender(primary: EmailRecipient | null, fallback: EmailRecipient | null): EmailRecipient | null {
    if (!primary) return fallback;
    if (!fallback) return primary;
    return {
      name: primary.name || fallback.name,
      email: this._isSmtpAddress(primary.email) ? primary.email : fallback.email || primary.email,
    };
  }

  private _isSmtpAddress(value: string | null): boolean {
    return Boolean(value && /^[^\s@]+@[^\s@]+$/.test(value));
  }

  // ------------------------------------------------------------------
  //  Utilities
  // ------------------------------------------------------------------

  private _entryName(path: string): string {
    const idx = path.lastIndexOf('/');
    return idx >= 0 ? path.slice(idx + 1) : path;
  }

  private _parentPath(path: string): string {
    const idx = path.lastIndexOf('/');
    return idx > 0 ? path.slice(0, idx) : '/';
  }

  private _makeTagStr(propId: number, propType: number): string {
    const id = propId.toString(16).padStart(4, '0').toLowerCase();
    const type = propType.toString(16).padStart(4, '0').toLowerCase();
    return id + type;
  }

  _fileTimeToDate(fileTime: number | bigint): Date | null {
    try {
      const msSince1601 = Number(fileTime) / 10000;
      const date = new Date(msSince1601 - FILETIME_EPOCH_OFFSET);
      return Number.isNaN(date.getTime()) ? null : date;
    } catch {
      return null;
    }
  }

  private _parseHeaders(raw: string): Record<string, string> | null {
    return parseHeaderBlock(raw);
  }

  private _parseHeadersFromBody(body: string): Record<string, string> | null {
    const separator = /\r?\n\r?\n/.exec(body);
    let headerEnd = separator?.index ?? -1;

    if (headerEnd === -1) {
      const lines = body.split(/\r?\n/);
      let lastHeaderLine = -1;

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (!line) continue;

        const trimmed = line.trim();

        if (trimmed.length === 0) {
          lastHeaderLine = i;
          break;
        }

        const isContinuation = /^\s/.test(line);

        if (i === 0) {
          if (trimmed.includes(':') && !trimmed.startsWith(' ') && !trimmed.startsWith('\t')) {
            lastHeaderLine = i;
          }
        } else if (isContinuation) {
          lastHeaderLine = i;
        } else if (trimmed.includes(':') && !trimmed.startsWith(' ') && !trimmed.startsWith('\t')) {
          lastHeaderLine = i;
        } else {
          const trimmedLower = trimmed.toLowerCase();
          if (trimmedLower.startsWith('dear ') ||
              trimmedLower.startsWith('hi ') ||
              trimmedLower.startsWith('hello ') ||
              trimmedLower.startsWith('caution:') ||
              trimmedLower.startsWith('dikkat:')) {
            break;
          }
          if (trimmed.includes(':') || isContinuation) {
            lastHeaderLine = i;
          } else {
            break;
          }
        }
      }

      if (lastHeaderLine >= 0) {
        headerEnd = 0;
        for (let i = 0; i <= lastHeaderLine; i++) {
          headerEnd += (lines?.[i]?.length || 0) + (body.includes('\r\n') ? 2 : 1);
        }
      } else {
        return null;
      }
    }

    const headerSection = body.slice(0, headerEnd);
    return this._parseHeaders(headerSection);
  }

  private _decodeBinaryHeaders(buf: Buffer): string | null {
    if (!buf || buf.length === 0) return null;

    let headersText = '';

    if (buf.length >= 4 && buf.readUInt32LE(0) === 0) {
      const remaining = buf.slice(4);
      if (remaining.length > 0) {
        if (remaining.length > 1 && remaining[1] === 0) {
          headersText = remaining.toString('ucs2');
        } else {
          headersText = remaining.toString('latin1');
        }
      }
    } else {
      headersText = buf.length > 1 && buf[1] === 0 ? buf.toString('ucs2') : buf.toString('latin1');
    }

    headersText = headersText.replace(/^\x00+/, '').replace(/\x00+$/, '');
    return headersText || null;
  }

  private _parseRecipientsFromHeaders(headers: Record<string, string>): { to: RecipientInfo[]; cc: RecipientInfo[]; bcc: RecipientInfo[] } {
    const parseAddressList = (headerValue: string | undefined): RecipientInfo[] => {
      return parseAddressHeader(headerValue).map(recipient => ({ ...recipient, type: null }));
    };

    return {
      to: parseAddressList(headers.to),
      cc: parseAddressList(headers.cc),
      bcc: parseAddressList(headers.bcc),
    };
  }
}

export default MsgParser;
