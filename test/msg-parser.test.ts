import { describe, it } from 'node:test';
import assert from 'node:assert';
import CFB from 'cfb';
import EmailParser, { MsgParser } from '../index.js';

describe('MsgParser', () => {
  it('should construct a new instance', () => {
    const parser = new MsgParser();
    assert.ok(parser instanceof MsgParser);
  });

  it('should throw on non-Buffer input to parse()', () => {
    const parser = new MsgParser();
    assert.throws(() => parser.parse(null as unknown as Buffer), /Buffer/);
    assert.throws(() => parser.parse('string' as unknown as Buffer), /Buffer/);
    assert.throws(() => parser.parse(123 as unknown as Buffer), /Buffer/);
  });

  it('should throw on non-existent file', () => {
    const parser = new MsgParser();
    assert.throws(() => parser.parseFile('/nonexistent/file.msg'));
  });

  it('should throw a descriptive error for an invalid buffer', () => {
    const parser = new MsgParser();
    const buf = Buffer.alloc(1024);
    buf.write('not an msg file');
    assert.throws(() => parser.parse(buf), /Invalid .msg file/);
  });

  it('should internally resolve property preferences (Unicode > ANSI)', () => {
    const parser = new MsgParser();
    const preferNew = (parser as unknown as { _preferNew(newType: number, existingType: number): boolean })._preferNew.bind(parser);

    assert.strictEqual(preferNew(0x001f, 0x001e), true);
    assert.strictEqual(preferNew(0x001e, 0x001f), false);
    assert.strictEqual(preferNew(0x0003, 0x0003), true);
  });

  it('should handle FILETIME to Date conversion', () => {
    const parser = new MsgParser();
    const date = parser._fileTimeToDate(0);
    assert.ok(date instanceof Date);
    assert.strictEqual(isNaN(date.getTime()), false);
  });

  it('should parse internet headers', () => {
    const parser = new MsgParser();
    const raw = 'From: test@example.com\r\nSubject: Test\r\nDate: Mon, 01 Jan 2024 00:00:00 +0000\r\n';
    const headers = (parser as unknown as { _parseHeaders(raw: string): Record<string, string> | null })._parseHeaders(raw);
    assert.ok(headers);
    assert.strictEqual(headers.from, 'test@example.com');
    assert.strictEqual(headers.subject, 'Test');
    assert.ok(headers.date);
  });

  it('should handle attachment parsing with empty attachments bucket', () => {
    const parser = new MsgParser();
    const attachments = parser._parseAttachments();
    assert.deepStrictEqual(attachments, []);
  });

  it('should handle recipient parsing with empty recipients bucket', () => {
    const parser = new MsgParser();
    const recipients = parser._parseRecipients();
    assert.deepStrictEqual(recipients, []);
  });

  it('should build a sane result from empty properties', () => {
    const parser = new MsgParser();
    const result = parser._buildResult({
      properties: {},
      recipients: [],
      attachments: [],
    });
    assert.strictEqual(result.format, 'msg');
    assert.strictEqual(result.subject, null);
    assert.strictEqual(result.importance, 'normal');
    assert.deepStrictEqual(result.to, []);
    assert.deepStrictEqual(result.cc, []);
    assert.deepStrictEqual(result.bcc, []);
    assert.deepStrictEqual(result.attachments, []);
    assert.strictEqual(result.body, null);
    assert.strictEqual(result.bodyHtml, null);
    assert.deepStrictEqual(result.replyTo, []);
  });

  it('should parse recipients from internet headers when recipients bucket is empty', () => {
    const parser = new MsgParser();
    const result = parser._buildResult({
      properties: {
        internetHeaders: 'To: John Doe <john@example.com>, jane@example.com\r\nCc: Bob Smith <bob@example.com>\r\nBcc: hidden@example.com\r\n',
      },
      recipients: [],
      attachments: [],
    });
    assert.strictEqual(result.to.length, 2);
    assert.strictEqual(result.to[0].name, 'John Doe');
    assert.strictEqual(result.to[0].email, 'john@example.com');
    assert.strictEqual(result.to[1].name, null);
    assert.strictEqual(result.to[1].email, 'jane@example.com');
    assert.strictEqual(result.cc.length, 1);
    assert.strictEqual(result.cc[0].name, 'Bob Smith');
    assert.strictEqual(result.cc[0].email, 'bob@example.com');
    assert.strictEqual(result.bcc.length, 1);
    assert.strictEqual(result.bcc[0].name, null);
    assert.strictEqual(result.bcc[0].email, 'hidden@example.com');
  });

  it('should prefer recipients from bucket over headers', () => {
    const parser = new MsgParser();
    const result = parser._buildResult({
      properties: {
        internetHeaders: 'To: header@example.com\r\n',
      },
      recipients: [{ name: 'Bucket Name', email: 'bucket@example.com', type: 'to' }],
      attachments: [],
    });
    assert.strictEqual(result.to.length, 1);
    assert.strictEqual(result.to[0].email, 'bucket@example.com');
  });

  it('should parse recipients from binary internetHeaders', () => {
    const parser = new MsgParser();
    const prefix = Buffer.from([0x00, 0x00, 0x00, 0x00]);
    const headersText = 'To: <ozdagdeviren@enka.com>\r\n';
    const binaryHeaders = Buffer.concat([prefix, Buffer.from(headersText)]);

    const result = parser._buildResult({
      properties: {
        internetHeaders: binaryHeaders,
      },
      recipients: [],
      attachments: [],
    });
    assert.strictEqual(result.to.length, 1);
    assert.strictEqual(result.to[0].email, 'ozdagdeviren@enka.com');
  });

  it('should parse recipients from binary internetHeaders with leading zeros', () => {
    const parser = new MsgParser();
    const binaryHeaders = Buffer.from([0x00, 0x00, 0x00, 0x00, 0x54, 0x6f, 0x3a, 0x20, 0x3c, 0x6f, 0x7a, 0x64, 0x61, 0x67, 0x64, 0x65, 0x76, 0x69, 0x72, 0x65, 0x6e, 0x40, 0x65, 0x6e, 0x6b, 0x61, 0x2e, 0x63, 0x6f, 0x6d, 0x3e, 0x0d, 0x0a]);

    const result = parser._buildResult({
      properties: {
        internetHeaders: binaryHeaders,
      },
      recipients: [],
      attachments: [],
    });
    assert.strictEqual(result.to.length, 1);
    assert.strictEqual(result.to[0].email, 'ozdagdeviren@enka.com');
  });

  it('should parse UTF-16 binary internet headers', () => {
    const parser = new MsgParser();
    const binaryHeaders = Buffer.concat([
      Buffer.alloc(4),
      Buffer.from('To: unicode@example.com\r\n\0', 'utf16le'),
    ]);
    const result = parser._buildResult({
      properties: { internetHeaders: binaryHeaders },
      recipients: [],
      attachments: [],
    });
    assert.strictEqual(result.to[0].email, 'unicode@example.com');
  });

  it('should parse recipients from body when internetHeaders is empty', () => {
    const parser = new MsgParser();
    const body = 'Received: from smtp.example.com\r\nTo: <ozdagdeviren@enka.com>\r\nSubject: Test\r\n\r\nDear Ozdagdeviren,\r\nThis is a test email.';

    const result = parser._buildResult({
      properties: {
        body,
        internetHeaders: Buffer.from([0x00, 0x00, 0x00, 0x00]),
      },
      recipients: [],
      attachments: [],
    });
    assert.strictEqual(result.to.length, 1);
    assert.strictEqual(result.to[0].email, 'ozdagdeviren@enka.com');
  });

  it('should read fixed values from 16-byte MSG property entries', () => {
    const parser = new MsgParser();
    const stream = Buffer.alloc(32 + 32);
    stream.writeUInt32LE((0x0017 << 16) | 0x0003, 32);
    stream.writeInt32LE(2, 40);
    stream.writeUInt32LE((0x0e08 << 16) | 0x0003, 48);
    stream.writeInt32LE(1234, 56);

    const read = (parser as unknown as {
      _readPropertyStream(buffer: Buffer, headerSize: number): Record<string, unknown>;
    })._readPropertyStream(stream, 32);

    assert.strictEqual(read['00170003'], 2);
    assert.strictEqual(read['0e080003'], 1234);
  });

  it('should parse quoted commas and groups in recipient headers', () => {
    const parser = new MsgParser();
    const result = parser._buildResult({
      properties: {
        internetHeaders: 'To: "Doe, Jane" <jane@example.com>, Team: Bob <bob@example.com>;\r\n',
      },
      recipients: [],
      attachments: [],
    });

    assert.deepStrictEqual(result.to, [
      { name: 'Doe, Jane', email: 'jane@example.com' },
      { name: 'Bob', email: 'bob@example.com' },
    ]);
  });

  it('should decode literal LZFu compressed RTF data', () => {
    const parser = new MsgParser();
    const rtf = Buffer.from('{\\rtf1}', 'ascii');
    const compressed = Buffer.alloc(16 + 1 + rtf.length);
    compressed.writeUInt32LE(compressed.length - 4, 0);
    compressed.writeUInt32LE(rtf.length, 4);
    compressed.writeUInt32LE(0x75465a4c, 8);
    compressed[16] = 0;
    rtf.copy(compressed, 17);

    const result = parser._buildResult({
      properties: { bodyRtfCompressed: compressed },
      recipients: [],
      attachments: [],
    });
    assert.strictEqual(result.bodyRtf, '{\\rtf1}');
  });

  it('should parse a minimal CFB-backed MSG end to end', async () => {
    type WritableCfb = ReturnType<typeof CFB.read>;
    const writableCfb = CFB as unknown as {
      utils: {
        cfb_new(): WritableCfb;
        cfb_add(cfb: WritableCfb, path: string, content: Buffer): void;
      };
      write(cfb: WritableCfb, options: { type: 'buffer' }): Buffer;
    };
    const container = writableCfb.utils.cfb_new();
    const properties = Buffer.alloc(48);
    properties.writeUInt32LE((0x0017 << 16) | 0x0003, 32);
    properties.writeInt32LE(2, 40);

    const unicode = (value: string): Buffer => Buffer.from(`${value}\0`, 'utf16le');
    writableCfb.utils.cfb_add(container, '/__properties_version1.0', properties);
    writableCfb.utils.cfb_add(container, '/__substg1.0_0037001F', unicode('Synthetic message'));
    writableCfb.utils.cfb_add(container, '/__substg1.0_1000001F', unicode('Synthetic body'));
    writableCfb.utils.cfb_add(
      container,
      '/__substg1.0_007D001F',
      unicode('From: Sender <sender@example.com>\r\nTo: receiver@example.com\r\nMessage-ID: <synthetic@example.com>\r\n'),
    );

    const source = writableCfb.write(container, { type: 'buffer' });
    const result = new MsgParser().parse(source);
    assert.strictEqual(result.format, 'msg');
    assert.strictEqual(result.subject, 'Synthetic message');
    assert.strictEqual(result.body, 'Synthetic body');
    assert.strictEqual(result.importance, 'high');
    assert.deepStrictEqual(result.from, { name: 'Sender', email: 'sender@example.com' });
    assert.deepStrictEqual(result.to, [{ name: null, email: 'receiver@example.com' }]);
    assert.strictEqual(result.messageId, '<synthetic@example.com>');
    assert.strictEqual(result.messageSize, source.length);

    const autoDetected = await new EmailParser().parse(source);
    assert.strictEqual(autoDetected.format, 'msg');
    assert.strictEqual(autoDetected.subject, 'Synthetic message');
  });
});
