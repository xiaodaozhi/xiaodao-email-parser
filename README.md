# xiaodao-email-parser

[![npm version](https://img.shields.io/npm/v/xiaodao-email-parser)](https://www.npmjs.com/package/xiaodao-email-parser)
[![npm downloads](https://img.shields.io/npm/dm/xiaodao-email-parser)](https://www.npmjs.com/package/xiaodao-email-parser)
[![CI](https://github.com/xiaodaozhi/xiaodao-email-parser/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/xiaodaozhi/xiaodao-email-parser/actions/workflows/ci.yml)
[![TypeScript](https://img.shields.io/badge/TypeScript-3178C6?logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
[![Node.js version](https://img.shields.io/node/v/xiaodao-email-parser)](https://www.npmjs.com/package/xiaodao-email-parser)
[![License](https://img.shields.io/npm/l/xiaodao-email-parser)](LICENSE)

English | [简体中文](README.ZH.md)

A Node.js and TypeScript email file parser that automatically detects and parses:

- Outlook `.msg` files (Compound File Binary / OLE2 + MAPI)
- RFC 5322 / MIME `.eml` files

The default `EmailParser` export detects the format from the file content and returns a consistent `ParsedEmail` object. Dedicated `MsgParser` and `EmlParser` exports are also available for format-specific use cases and for migrating existing synchronous MSG integrations.

## Features

- Detect MSG and EML from their content instead of relying on file extensions
- Parse senders, reply-to addresses, recipients, subjects, dates, importance, and Message-ID values
- Extract plain text, HTML, and MSG compressed RTF bodies
- Handle multipart messages, base64, quoted-printable, RFC 2047 encoded words, and RFC 2231 filenames
- Extract regular and inline attachments with MIME type and Content-ID metadata
- Preserve raw headers, parsed headers, and format-specific low-level properties
- Enforce MIME nesting, header-size, and input-size limits
- Ship complete TypeScript declarations

## Installation

```bash
npm install xiaodao-email-parser
```

Node.js 20 or later is required. The package is published as an ES module. To run the JavaScript-compatible quick-start examples, use an `.mjs` file or set `"type": "module"` in your project's `package.json`.

### Upgrading from 1.0.7

Version 2 requires Node.js 20 or later, up from Node.js 18 in 1.0.7. `EmailParser` remains asynchronous, so continue to use `await`; `MsgParser` remains synchronous. If you rely on MIME parsing details or nested messages, test those cases when upgrading: the underlying `postal-mime` dependency was updated from 3.x to 4.x.

## Quick start: auto-detect MSG or EML

```typescript
import EmailParser from 'xiaodao-email-parser';

const parser = new EmailParser();
const email = await parser.parseFile('message.eml'); // message.msg also works

console.log(email.format); // 'eml' or 'msg'
console.log(email.subject);
console.log(email.from);
console.log(email.to);
console.log(email.attachments);
```

For example, an EML message might produce `format: 'eml'`, subject `'Meeting'`, sender `{ name: 'Alice', email: 'alice@example.com' }`, and an empty attachments array. Values depend on the message; missing fields can be `null` or empty arrays.

You can also parse a `Buffer`:

```typescript
import fs from 'node:fs/promises';
import EmailParser from 'xiaodao-email-parser';

const source = await fs.readFile('message.msg');
const email = await new EmailParser().parse(source);
```

The unified parser always uses an asynchronous API, including when the detected input is MSG.

File reads and parsing can fail. Handle the rejected promise, for example:

```typescript
try {
  const email = await new EmailParser().parseFile('message.eml');
  console.log(email.subject);
} catch (error) {
  console.error('Could not read or parse the email:', error);
}
```

Failures can include an unreadable file, an unrecognized or invalid message, or an EML size or MIME limit being exceeded.

## Dedicated parsers

### Parse MSG synchronously

```typescript
import { MsgParser } from 'xiaodao-email-parser';

const email = new MsgParser().parseFile('message.msg');
console.log(email.body, email.bodyHtml, email.bodyRtf);
```

`MsgParser.parse()` and `MsgParser.parseFile()` remain synchronous for existing MSG integrations.

### Parse EML

```typescript
import { EmlParser } from 'xiaodao-email-parser';

const parser = new EmlParser({
  maxInputSize: 50 * 1024 * 1024,
  maxNestingDepth: 50,
  maxHeadersSize: 2 * 1024 * 1024,
  maxRfc822NestingDepth: 5,
});

const email = await parser.parseFile('message.eml');
```

`EmlParser` accepts inputs up to 100 MiB by default. This limit applies to EML only; `MsgParser` has no equivalent input-size option. Both `parseFile()` implementations read the entire file into memory before parsing, so the EML limit is checked after the read. For untrusted files, enforce a size limit before loading their contents and account for possible file changes between checking and reading.

## API

Package exports:

```typescript
import EmailParser, {
  EmailParser as NamedEmailParser,
  MsgParser,
  EmlParser,
  detectEmailFormat,
} from 'xiaodao-email-parser';
```

### `EmailParser`

```typescript
new EmailParser(options?: EmailParserOptions)
parser.parse(buffer: Buffer, options?: ParseOptions): Promise<ParsedEmail>
parser.parseFile(filePath: string, options?: ParseOptions): Promise<ParsedEmail>
```

Content-based auto-detection is used by default. You can also force a format:

```typescript
await parser.parse(buffer, { format: 'eml' });
await parser.parseFile(filePath, { format: 'msg' });
```

### `MsgParser`

```typescript
new MsgParser()
parser.parse(buffer: Buffer): ParsedEmail
parser.parseFile(filePath: string): ParsedEmail
```

Parses Outlook MSG only. Both methods are synchronous and throw a descriptive error when the input is not a valid CFB/MAPI message.

### `EmlParser`

```typescript
new EmlParser(options?: EmlParserOptions)
parser.parse(buffer: Buffer): Promise<ParsedEmail>
parser.parseFile(filePath: string): Promise<ParsedEmail>
```

Parses RFC 5322/MIME EML only. The following options can be configured:

```typescript
interface EmlParserOptions {
  maxInputSize?: number;
  rfc822Attachments?: boolean;
  forceRfc822Attachments?: boolean;
  maxNestingDepth?: number;
  maxHeadersSize?: number;
  maxRfc822NestingDepth?: number;
}
```

| Option | Default | Behavior |
| --- | --- | --- |
| `maxInputSize` | 100 MiB | Maximum EML input size in bytes; must be a positive safe integer. |
| `rfc822Attachments` | `false` | Treat `message/rfc822` parts without a Content-Disposition as attachments. |
| `forceRfc822Attachments` | `false` | Treat all `message/rfc822` parts as attachments. |
| `maxNestingDepth` | `256` | Maximum MIME part nesting depth. |
| `maxHeadersSize` | `2097152` | Maximum total header size in bytes across MIME parts and inline nested messages. |
| `maxRfc822NestingDepth` | `10` | Maximum inline `message/rfc822` recursion depth; deeper messages are returned as attachments. Set to `0` to disable inline parsing. |

The MIME option defaults above come from the current `postal-mime` 4.2.x dependency. To configure EML parsing through the unified parser, pass the same options under `eml`:

```typescript
const parser = new EmailParser({
  eml: { maxInputSize: 20 * 1024 * 1024, maxNestingDepth: 50 },
});
const email = await parser.parseFile('message.eml');
```

### `detectEmailFormat`

```typescript
import { detectEmailFormat } from 'xiaodao-email-parser';

const format = detectEmailFormat(buffer); // 'msg' | 'eml' | null
```

MSG detection checks the CFB signature and, for readable containers, the root MAPI property stream. A truncated or unreadable CFB container is still treated as a possible MSG so parsing can report the error. EML detection checks the header structure and known email headers. The function returns `null` when it cannot determine the format reliably; the unified parser throws a descriptive error in that case.

### `ParsedEmail`

```typescript
interface ParsedEmail {
  format: 'msg' | 'eml';
  subject: string | null;
  from: EmailRecipient | null;
  replyTo: EmailRecipient[];
  to: EmailRecipient[];
  cc: EmailRecipient[];
  bcc: EmailRecipient[];
  body: string | null;
  bodyHtml: string | null;
  bodyRtf: string | null;
  attachments: EmailAttachment[];
  sentDate: Date | null;
  receivedDate: Date | null;
  createdDate: Date | null;
  modifiedDate: Date | null;
  messageClass: string | null;
  importance: 'low' | 'normal' | 'high' | null;
  messageSize: number | null;
  conversationTopic: string | null;
  normalizedSubject: string | null;
  messageId: string | null;
  headers: string | null;
  parsedHeaders: Record<string, string> | null;
  preview: string | null;
  _rawProperties: Record<string, unknown>;
}
```

For EML, `_rawProperties` contains sender, return-path, references, and the unfolded header array. For MSG, it contains the resolved low-level MAPI properties.

Each `EmailRecipient` has a nullable `name` and `email`. The `from` field is either one such address object or `null`; `replyTo`, `to`, `cc`, and `bcc` are arrays and may be empty.

### Attachments

```typescript
interface EmailAttachment {
  filename: string | null;
  content: Buffer | null;
  contentType: string;
  contentId: string | null;
  contentLocation: string | null;
  size: number;
}
```

Example: save available attachment contents to a newly created temporary directory. The generated names avoid trusting filenames from the message; `wx` also prevents overwriting an existing file.

```typescript
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';

const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), 'email-attachments-'));
for (const [index, attachment] of email.attachments.entries()) {
  if (attachment.content !== null) {
    const outputPath = path.join(outputDir, `attachment-${index}`);
    await fs.writeFile(outputPath, attachment.content, { flag: 'wx' });
    console.log('Saved:', outputPath, 'Original name:', attachment.filename);
  }
}
```

Attachment names and contents come from the message and should be treated as untrusted. If you need to preserve original names or file extensions, validate them according to your application's storage and security requirements.

## Architecture

```text
EmailParser (asynchronous, auto-detecting)
├── MsgParser (synchronous, CFB/MAPI)
└── EmlParser (asynchronous, RFC 5322/MIME)
```

Each dedicated parser owns its format-specific decoding. Shared utilities normalize addresses, headers, subjects, previews, dates, and Content-ID values before both implementations map their output to the same result type.

## Development

```bash
npm ci
npm run build
npm test
node --import tsx examples/parse.ts path/to/message.msg
node --import tsx examples/parse.ts path/to/message.eml
```

## Limitations

- Encrypted and password-protected messages are not decrypted.
- MSG custom named properties and embedded MSG attachments are not fully expanded yet.
- Nested EML messages beyond `maxRfc822NestingDepth` are returned as attachments rather than expanded inline. The underlying parser's `rfc822DepthExceeded` flag is not exposed in `EmailAttachment`; do not assume the attachment list contains every deeply nested payload when scanning mail.
- For backward compatibility, `parsedHeaders` is a key-value object and repeated headers are merged. The original EML header array is available at `_rawProperties.headers`.
- The parser only extracts content; it never executes HTML, scripts, or attachments. Treat `bodyHtml` as untrusted and sanitize it before rendering in a web page.

## License

MIT
