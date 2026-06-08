// server/utils/tools/msgparser/msg-parser.ts

import CFB from 'cfb';
import fs from 'node:fs';
import {
  PROPERTY_IDS,
  PROPERTY_TYPES,
  RECIPIENT_TYPES,
  IMPORTANCE,
  FILETIME_EPOCH_OFFSET,
} from './constants';
import type {
  ParsedEmail,
  EmailRecipient,
  RecipientInfo,
  EmailAttachmentInfo,
  RawEmailData,
  EntryStore,
  SubStorageBucket,
  PropertyTypeInfo,
  TypedReadResult,
} from './types';

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
    try {
      this.cfb = CFB.read(buffer, { type: 'buffer' });
    } catch (err: unknown) {
      if (err instanceof Error && err.message && err.message.includes('Header Signature')) {
        throw new Error('Invalid .msg file: header signature mismatch. The file may be corrupted or not an Outlook .msg file.');
      }
      throw err;
    }
    this._classifyEntries();
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

  private _classifySubLevel(path: string, name: string, content: Buffer): void {
    const parent = this._parentPath(path);
    let storageName = parent;
    if (storageName.startsWith('/')) storageName = storageName.slice(1);

    if (storageName.startsWith('__recip_version1.0_#')) {
      const idx = storageName.slice('__recip_version1.0_#'.length);
      this._addToBucket(this._entries.recipients, idx, name, content);
    } else if (storageName.startsWith('__attach_version1.0_#')) {
      const idx = storageName.slice('__attach_version1.0_#'.length);
      this._addToBucket(this._entries.attachments, idx, name, content);
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
      const streamProps = this._readPropertyStream(this._entries.topPropsStream);
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
    };
  }

  // ------------------------------------------------------------------
  //  Property stream reader (__properties_version1.0)
  // ------------------------------------------------------------------

  private _readPropertyStream(buf: Buffer): Record<string, unknown> {
    const result: Record<string, unknown> = {};
    let offset = 8;

    while (offset + 8 <= buf.length) {
      const propertyTag = buf.readUInt32LE(offset);
      if (propertyTag === 0) break;

      offset += 8;

      const propId = (propertyTag >>> 16) & 0xffff;
      const propType = propertyTag & 0xffff;
      const tagStr = this._makeTagStr(propId, propType);

      try {
        const { value, bytesRead } = this._readTypedValue(buf, offset, propType);
        if (value !== undefined) result[tagStr] = value;
        offset += bytesRead;
      } catch {
        const skip = this._guessValueSize(propType);
        if (skip > 0) offset += skip;
        else break;
      }
    }

    return result;
  }

  private _readTypedValue(buf: Buffer, offset: number, type: number): TypedReadResult {
    const info: PropertyTypeInfo | undefined = PROPERTY_TYPES[type];

    if (!info) {
      const guessed = this._guessValueSize(type);
      return { value: undefined, bytesRead: guessed > 0 ? guessed : 0 };
    }

    if (info.fixedSize > 0) {
      return this._readFixedValue(buf, offset, type, info.fixedSize);
    }
    if (info.fixedSize === -1) {
      return this._readVariableValue(buf, offset, type);
    }
    if (info.fixedSize === -2) {
      return this._readMultiValue(buf, offset, type);
    }

    return { value: undefined, bytesRead: 0 };
  }

  private _readFixedValue(buf: Buffer, offset: number, type: number, size: number): TypedReadResult {
    let value: unknown;
    switch (type) {
      case 0x0002:
        value = buf.readInt16LE(offset);
        break;
      case 0x0003:
        value = buf.readUInt32LE(offset);
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
        value = buf.readUInt32LE(offset) + buf.readUInt32LE(offset + 4) * 0x100000000;
        break;
      default:
        value = buf.slice(offset, offset + size);
    }
    return { value, bytesRead: size };
  }

  private _readVariableValue(buf: Buffer, offset: number, type: number): TypedReadResult {
    if (offset + 4 > buf.length) return { value: undefined, bytesRead: 0 };

    const byteCount = buf.readUInt32LE(offset);
    if (byteCount === 0 || offset + 4 + byteCount > buf.length) {
      return { value: undefined, bytesRead: 4 };
    }

    const dataStart = offset + 4;
    let value: unknown;

    switch (type) {
      case 0x001e:
        value = buf.toString('latin1', dataStart, dataStart + byteCount).replace(/\0+$/, '');
        break;
      case 0x001f:
        value = buf.toString('ucs2', dataStart, dataStart + byteCount).replace(/\0+$/, '');
        break;
      case 0x0102:
        value = Buffer.from(buf.slice(dataStart, dataStart + byteCount));
        break;
      default:
        value = Buffer.from(buf.slice(dataStart, dataStart + byteCount));
    }

    return { value, bytesRead: 4 + byteCount };
  }

  private _readMultiValue(buf: Buffer, offset: number, type: number): TypedReadResult {
    if (offset + 4 > buf.length) return { value: undefined, bytesRead: 0 };
    const count = buf.readUInt32LE(offset);
    offset += 4;
    const values: unknown[] = [];
    let totalBytes = 4;

    for (let i = 0; i < count; i++) {
      if (offset >= buf.length) break;

      const elemType = type & 0x0fff;
      const elemInfo: PropertyTypeInfo | undefined = PROPERTY_TYPES[elemType];

      if (elemInfo && elemInfo.fixedSize > 0) {
        const { value, bytesRead } = this._readFixedValue(buf, offset, elemType, elemInfo.fixedSize);
        values.push(value);
        offset += bytesRead;
        totalBytes += bytesRead;
      } else if (elemInfo && elemInfo.fixedSize === -1) {
        if (offset + 4 > buf.length) break;
        const elemLen = buf.readUInt32LE(offset);
        offset += 4;
        totalBytes += 4;
        const elemEnd = Math.min(offset + elemLen, buf.length);
        const raw = buf.slice(offset, elemEnd);
        if (elemType === 0x001f) {
          values.push(raw.toString('ucs2').replace(/\0+$/, ''));
        } else if (elemType === 0x001e) {
          values.push(raw.toString('latin1').replace(/\0+$/, ''));
        } else {
          values.push(Buffer.from(raw));
        }
        totalBytes += (elemEnd - offset);
        offset = elemEnd;
      } else {
        break;
      }
    }

    return { value: values, bytesRead: totalBytes };
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
        return content.readUInt32LE(0);
      case 0x0002:
        return content.readInt16LE(0);
      case 0x000b:
        return content.readUInt16LE(0) !== 0;
      case 0x0014:
        return Number(content.readBigInt64LE(0));
      case 0x0040:
        return Number(content.readBigUInt64LE(0));
      case 0x0004:
        return content.readFloatLE(0);
      case 0x0005:
        return content.readDoubleLE(0);
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
    for (const [name, { value }] of Object.entries(named)) {
      result[name] = value;

      if (name.endsWith('Time') && typeof value === 'number' && value > 0) {
        result[name] = this._fileTimeToDate(value);
      }
    }

    if (result.importance !== undefined && typeof result.importance === 'number') {
      const mapped = IMPORTANCE[result.importance];
      if (mapped !== undefined) {
        result.importance = mapped;
      }
    }

    if (result.senderName || result.senderEmail) {
      result.from = {
        name: (result.senderName as string) || null,
        email: (result.senderEmail as string) || null,
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
        name: (props.recipientDisplayName as string) || (props.displayName as string) || (props.recipientName as string) || null,
        email: (props.smtpEmailAddress as string) || (props.recipientEmailAddress as string) || null,
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
      const data = (props.attachmentData as Buffer) || null;

      const attachment: EmailAttachmentInfo = {
        filename: (props.attachmentLongFilename as string) || (props.attachmentFilename as string) || null,
        content: data,
        contentType: (props.attachmentContentType as string) || (props.attachmentMimeType as string) || 'application/octet-stream',
        contentId: (props.attachmentContentId as string) || null,
        contentLocation: (props.attachmentContentLocation as string) || null,
        size: data ? data.length : 0,
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

    let parsedHeaders: Record<string, string> | null = null;
    if (p.internetHeaders) {
      if (typeof p.internetHeaders === 'string') {
        parsedHeaders = this._parseHeaders(p.internetHeaders);
      } else if (Buffer.isBuffer(p.internetHeaders)) {
        parsedHeaders = this._parseBinaryHeaders(p.internetHeaders);
      }
    }

    if (parsedHeaders && Object.keys(parsedHeaders).length > 0) {
      const fromHeaders = this._parseRecipientsFromHeaders(parsedHeaders);
      if (to.length === 0) to = fromHeaders.to;
      if (cc.length === 0) cc = fromHeaders.cc;
      if (bcc.length === 0) bcc = fromHeaders.bcc;
    }

    if (to.length === 0 && p.body && typeof p.body === 'string') {
      const bodyHeaders = this._parseHeadersFromBody(p.body);
      if (bodyHeaders) {
        const fromBody = this._parseRecipientsFromHeaders(bodyHeaders);
        if (to.length === 0) to = fromBody.to;
        if (cc.length === 0) cc = fromBody.cc;
        if (bcc.length === 0) bcc = fromBody.bcc;
      }
    }

    return {
      subject: (p.subject as string) || null,
      from: (p.from as EmailRecipient) || null,
      to: to.length > 0 ? to.map(r => ({ name: r.name, email: r.email })) : [],
      cc: cc.length > 0 ? cc.map(r => ({ name: r.name, email: r.email })) : [],
      bcc: bcc.length > 0 ? bcc.map(r => ({ name: r.name, email: r.email })) : [],
      body: (p.body as string) || null,
      bodyHtml: (p.bodyHtml as string) || null,
      bodyRtf: (p.bodyRtf as string) || null,
      attachments,
      sentDate: (p.clientSubmitTime as Date) || null,
      receivedDate: (p.messageDeliveryTime as Date) || null,
      createdDate: (p.creationTime as Date) || null,
      modifiedDate: (p.lastModificationTime as Date) || null,
      messageClass: (p.messageClass as string) || null,
      importance: (p.importance as 'low' | 'normal' | 'high') || 'normal',
      messageSize: (p.messageSize as number) || null,
      conversationTopic: (p.conversationTopic as string) || null,
      normalizedSubject: (p.normalizedSubject as string) || null,
      headers: (p.internetHeaders as string) || null,
      parsedHeaders,
      preview: (p.preview as string) || null,
      _rawProperties: p,
    };
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

  _fileTimeToDate(fileTime: number): Date | null {
    try {
      const msSince1601 = Number(fileTime) / 10000;
      return new Date(msSince1601 - FILETIME_EPOCH_OFFSET);
    } catch {
      return null;
    }
  }

  private _guessValueSize(type: number): number {
    if (type > 0x1000) return 4;
    if (type <= 0x0014) return 8;
    if (type === 0x001e || type === 0x001f || type === 0x0102) return 4;
    return 0;
  }

  private _parseHeaders(raw: string): Record<string, string> | null {
    if (!raw) return null;
    const headers: Record<string, string> = {};
    let currentKey: string | null = null;
    const lines = raw.split(/\r?\n/);

    for (const line of lines) {
      if (/^\s/.test(line) && currentKey) {
        headers[currentKey] += ' ' + line.trim();
      } else {
        const idx = line.indexOf(':');
        if (idx > 0) {
          currentKey = line.slice(0, idx).trim().toLowerCase();
          const value = line.slice(idx + 1).trim();
          headers[currentKey] = value;
        } else {
          currentKey = null;
        }
      }
    }

    return headers;
  }

  private _parseHeadersFromBody(body: string): Record<string, string> | null {
    let headerEnd = body.indexOf('\r\n\r\n');

    if (headerEnd === -1) {
      const lines = body.split('\r\n');
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
          headerEnd += (lines?.[i]?.length || 0) + 2;
        }
      } else {
        return null;
      }
    }

    const headerSection = body.slice(0, headerEnd);
    return this._parseHeaders(headerSection);
  }

  private _parseBinaryHeaders(buf: Buffer): Record<string, string> | null {
    if (!buf || buf.length === 0) return null;

    let headersText = '';

    if (buf.length >= 4 && buf.readUInt32LE(0) === 0) {
      const remaining = buf.slice(4);
      if (remaining.length > 0) {
        if (remaining[0] === 0 && remaining[1] !== 0) {
          headersText = remaining.toString('ucs2');
        } else {
          headersText = remaining.toString('latin1');
        }
      }
    } else {
      headersText = buf.toString('latin1');
    }

    headersText = headersText.replace(/^\x00+/, '').replace(/\x00+$/, '');
    return headersText ? this._parseHeaders(headersText) : null;
  }

  private _parseRecipientsFromHeaders(headers: Record<string, string>): { to: RecipientInfo[]; cc: RecipientInfo[]; bcc: RecipientInfo[] } {
    const parseAddressList = (headerValue: string | undefined): RecipientInfo[] => {
      if (!headerValue) return [];
      const result: RecipientInfo[] = [];
      const addresses = headerValue.split(/,/).map(s => s.trim()).filter(Boolean);

      for (const addr of addresses) {
        const emailMatch = addr.match(/<([^>]+)>/);
        if (emailMatch) {
          const email = emailMatch[1];
          const name = addr.replace(/<[^>]+>/, '').trim().replace(/^["']|["']$/g, '');
          result.push({
            name: name || null,
            email: email || null,
            type: null,
          });
        } else if (addr.includes('@')) {
          result.push({
            name: null,
            email: addr.trim(),
            type: null,
          });
        }
      }
      return result;
    };

    return {
      to: parseAddressList(headers.to),
      cc: parseAddressList(headers.cc),
      bcc: parseAddressList(headers.bcc),
    };
  }
}

export default MsgParser;
