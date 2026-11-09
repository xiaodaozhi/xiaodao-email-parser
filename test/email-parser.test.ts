import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import CFB from 'cfb';
import EmailParser, { EmlParser, detectEmailFormat } from '../index.js';

function eml(lines: string[]): Buffer {
  return Buffer.from(lines.join('\r\n'), 'utf8');
}

describe('EmailParser and EmlParser', () => {
  it('detects EML, MSG, and unsupported input', () => {
    assert.strictEqual(detectEmailFormat(eml(['From: a@example.com', '', 'hello'])), 'eml');
    assert.strictEqual(
      detectEmailFormat(Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])),
      'msg',
    );
    assert.strictEqual(detectEmailFormat(Buffer.from('plain text only')), null);
  });

  it('does not classify other CFB documents as MSG', () => {
    type WritableCfb = ReturnType<typeof CFB.read>;
    const writableCfb = CFB as unknown as {
      utils: {
        cfb_new(): WritableCfb;
        cfb_add(cfb: WritableCfb, path: string, content: Buffer): void;
      };
      write(cfb: WritableCfb, options: { type: 'buffer' }): Buffer;
    };
    const container = writableCfb.utils.cfb_new();
    writableCfb.utils.cfb_add(container, '/WordDocument', Buffer.from('not an MSG'));
    writableCfb.utils.cfb_add(container, '/nested/__properties_version1.0', Buffer.alloc(32));

    assert.strictEqual(detectEmailFormat(writableCfb.write(container, { type: 'buffer' })), null);
  });

  it('parses encoded headers, addresses, metadata, and quoted-printable text', async () => {
    const source = eml([
      'From: =?UTF-8?B?5byg5LiJ?= <zhang@example.com>',
      'To: "Doe, Jane" <jane@example.com>, Team: Bob <bob@example.com>;',
      'Reply-To: support@example.com',
      'Subject: =?UTF-8?B?5rWL6K+V6YKu5Lu2?=',
      'Date: Mon, 1 Jan 2024 08:00:00 +0800',
      'Message-ID: <message-1@example.com>',
      'Importance: high',
      'MIME-Version: 1.0',
      'Content-Type: text/plain; charset=utf-8',
      'Content-Transfer-Encoding: quoted-printable',
      '',
      'Hello, =E4=B8=96=E7=95=8C!',
    ]);

    const result = await new EmailParser().parse(source);
    assert.strictEqual(result.format, 'eml');
    assert.strictEqual(result.subject, '测试邮件');
    assert.deepStrictEqual(result.from, { name: '张三', email: 'zhang@example.com' });
    assert.deepStrictEqual(result.to, [
      { name: 'Doe, Jane', email: 'jane@example.com' },
      { name: 'Bob', email: 'bob@example.com' },
    ]);
    assert.deepStrictEqual(result.replyTo, [{ name: null, email: 'support@example.com' }]);
    assert.match(result.body ?? '', /Hello, 世界!/);
    assert.strictEqual(result.importance, 'high');
    assert.strictEqual(result.messageId, '<message-1@example.com>');
    assert.strictEqual(result.messageSize, source.length);
    assert.strictEqual(result.sentDate?.toISOString(), '2024-01-01T00:00:00.000Z');
    assert.match(result.headers ?? '', /^From:/);
    assert.strictEqual(result.parsedHeaders?.subject, '=?UTF-8?B?5rWL6K+V6YKu5Lu2?=');
  });

  it('parses nested multipart bodies, inline content, and attachments', async () => {
    const source = eml([
      'From: sender@example.com',
      'To: receiver@example.com',
      'Subject: Multipart',
      'MIME-Version: 1.0',
      'Content-Type: multipart/mixed; boundary="outer"',
      '',
      '--outer',
      'Content-Type: multipart/alternative; boundary="inner"',
      '',
      '--inner',
      'Content-Type: text/plain; charset=utf-8',
      '',
      'Plain body',
      '--inner',
      'Content-Type: text/html; charset=utf-8',
      'Content-Transfer-Encoding: base64',
      '',
      Buffer.from('<p>HTML body</p>').toString('base64'),
      '--inner--',
      '--outer',
      "Content-Type: application/octet-stream; name*=utf-8''%E6%8A%A5%E5%91%8A.txt",
      "Content-Disposition: attachment; filename*=utf-8''%E6%8A%A5%E5%91%8A.txt",
      'Content-Transfer-Encoding: base64',
      'Content-ID: <attachment-1>',
      '',
      'SGVsbG8=',
      '--outer--',
      '',
    ]);

    const result = await new EmlParser().parse(source);
    assert.match(result.body ?? '', /Plain body/);
    assert.match(result.bodyHtml ?? '', /HTML body/);
    assert.strictEqual(result.attachments.length, 1);
    assert.strictEqual(result.attachments[0].filename, '报告.txt');
    assert.strictEqual(result.attachments[0].content?.toString(), 'Hello');
    assert.strictEqual(result.attachments[0].contentId, 'attachment-1');
    assert.strictEqual(result.attachments[0].size, 5);
  });

  it('rejects unsupported auto-detected data and oversized EML input', async () => {
    await assert.rejects(
      new EmailParser().parse(Buffer.from('not an email')),
      /Unsupported email format/,
    );
    await assert.rejects(
      new EmlParser({ maxInputSize: 4 }).parse(Buffer.from('12345')),
      /exceeds the configured/,
    );
    await assert.rejects(
      new EmailParser().parse(Buffer.from('Subject: test\r\n\r\n'), { format: 'pdf' as 'eml' }),
      /Unsupported format option/,
    );
  });

  it('supports explicitly selecting EML format for header-only messages', async () => {
    const result = await new EmailParser().parse(
      Buffer.from('Subject: Header only\r\n'),
      { format: 'eml' },
    );
    assert.strictEqual(result.subject, 'Header only');
    assert.strictEqual(result.format, 'eml');
  });

  it('auto-detects header-only EML and preserves its raw headers', async () => {
    const source = Buffer.from('Subject: Header only\r\n', 'utf8');

    assert.strictEqual(detectEmailFormat(source), 'eml');
    const result = await new EmailParser().parse(source);
    assert.strictEqual(result.subject, 'Header only');
    assert.strictEqual(result.headers, 'Subject: Header only\r\n');
  });

  it('does not treat a header followed by an unseparated body as EML', () => {
    const source = Buffer.from('Subject: Looks like a header\r\nplain text body', 'utf8');

    assert.strictEqual(detectEmailFormat(source), null);
  });

  it('does not treat HTTP traffic as EML', () => {
    const source = Buffer.from([
      'HTTP/1.1 200 OK',
      'Content-Type: text/html',
      '',
      '<html></html>',
    ].join('\r\n'), 'latin1');

    assert.strictEqual(detectEmailFormat(source), null);
  });

  it('uses HTML for previews when the plain-text part is blank', async () => {
    const source = eml([
      'From: sender@example.com',
      'MIME-Version: 1.0',
      'Content-Type: multipart/alternative; boundary="preview"',
      '',
      '--preview',
      'Content-Type: text/plain; charset=utf-8',
      '',
      '   ',
      '--preview',
      'Content-Type: text/html; charset=utf-8',
      '',
      '<p>HTML fallback</p>',
      '--preview--',
      '',
    ]);

    const result = await new EmlParser().parse(source);
    assert.strictEqual(result.preview, 'HTML fallback');
  });

  it('decodes encoded conversation topics', async () => {
    const source = eml([
      'Thread-Topic: =?UTF-8?B?5rWL6K+V6YKu5Lu2?=',
      '',
    ]);

    const result = await new EmlParser().parse(source);
    assert.strictEqual(result.conversationTopic, '测试邮件');
  });

  it('ignores an mbox envelope line as a parsed mail header', async () => {
    const source = eml([
      'From 13292f7ca4914b8fad350d448253472c234034a2 Mon Sep 17 00:00:00 2001',
      'From: sender@example.com',
      'Subject: Patch mail',
      '',
      'Body',
    ]);

    assert.strictEqual(detectEmailFormat(source), 'eml');
    const result = await new EmailParser().parse(source);
    assert.deepStrictEqual(Object.keys(result.parsedHeaders ?? {}), ['from', 'subject']);
    assert.strictEqual(result.from?.email, 'sender@example.com');
  });
});
