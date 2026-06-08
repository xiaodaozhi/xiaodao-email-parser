import { describe, it } from 'node:test';
import assert from 'node:assert';
import MsgParser from '../index.js';

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
    assert.strictEqual(result.subject, null);
    assert.strictEqual(result.importance, 'normal');
    assert.deepStrictEqual(result.to, []);
    assert.deepStrictEqual(result.cc, []);
    assert.deepStrictEqual(result.bcc, []);
    assert.deepStrictEqual(result.attachments, []);
    assert.strictEqual(result.body, null);
    assert.strictEqual(result.bodyHtml, null);
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
});
