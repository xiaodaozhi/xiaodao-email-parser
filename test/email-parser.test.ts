import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
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
});
