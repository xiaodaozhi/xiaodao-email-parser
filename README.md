# xiaodao-msg-parser

A Node.js parser for Microsoft Outlook `.msg` files (Compound File Binary / OLE2 format). Extract email metadata, body content, recipients, and attachments from saved Outlook messages.

## Features

- Parse Outlook `.msg` files synchronously
- Extract email metadata (subject, sender, recipients, dates)
- Support for plain text, HTML, and RTF body content
- Attachment extraction with content and metadata
- Internet headers parsing
- TypeScript support with full type definitions
- Zero external runtime dependencies (only `cfb` for OLE2 parsing)

## Installation

```bash
npm install xiaodao-msg-parser
```

## Quick Start

```typescript
import MsgParser from 'xiaodao-msg-parser';

const parser = new MsgParser();
const email = parser.parseFile('message.msg');

console.log('Subject:', email.subject);
console.log('From:', email.from?.name, '<' + email.from?.email + '>');
console.log('To:', email.to.map(r => r.email).join(', '));
console.log('Date:', email.sentDate);
console.log('Body:', email.body?.slice(0, 200));
```

Or parse from a Buffer:

```typescript
import MsgParser from 'xiaodao-msg-parser';
import fs from 'node:fs';

const parser = new MsgParser();
const buffer = fs.readFileSync('message.msg');
const email = parser.parse(buffer);
```

## API Reference

### `MsgParser`

#### `parseFile(filePath: string): ParsedEmail`

Parse a `.msg` file from a file path (synchronous).

```typescript
const email = parser.parseFile('/path/to/message.msg');
```

#### `parse(buffer: Buffer): ParsedEmail`

Parse a `.msg` file from a Buffer.

```typescript
const email = parser.parse(buffer);
```

### `ParsedEmail`

The parsed email object contains the following fields:

| Field | Type | Description |
|-------|------|-------------|
| `subject` | `string \| null` | Email subject line |
| `from` | `EmailRecipient \| null` | Sender information |
| `to` | `EmailRecipient[]` | To recipients |
| `cc` | `EmailRecipient[]` | CC recipients |
| `bcc` | `EmailRecipient[]` | BCC recipients |
| `body` | `string \| null` | Plain text body |
| `bodyHtml` | `string \| null` | HTML body |
| `bodyRtf` | `string \| null` | RTF body |
| `attachments` | `EmailAttachment[]` | Array of attachments |
| `sentDate` | `Date \| null` | Date the email was sent |
| `receivedDate` | `Date \| null` | Date the email was received |
| `createdDate` | `Date \| null` | Creation date |
| `modifiedDate` | `Date \| null` | Last modification date |
| `messageClass` | `string \| null` | Outlook message class (e.g., `IPM.Note`) |
| `importance` | `'low' \| 'normal' \| 'high' \| null` | Importance level |
| `messageSize` | `number \| null` | Message size in bytes |
| `conversationTopic` | `string \| null` | Conversation topic |
| `headers` | `string \| null` | Raw internet headers |
| `parsedHeaders` | `Record<string, string> \| null` | Parsed headers as key-value pairs |
| `preview` | `string \| null` | Preview text |
| `_rawProperties` | `Record<string, unknown>` | Raw MAPI properties |

### `EmailRecipient`

```typescript
interface EmailRecipient {
  name: string | null;   // Display name
  email: string | null;  // Email address
}
```

### `EmailAttachment`

```typescript
interface EmailAttachment {
  filename: string | null;           // File name
  content: Buffer | null;            // File content as Buffer
  contentType: string;               // MIME type
  contentId: string | null;          // Content-ID (for inline images)
  contentLocation: string | null;    // Content location
  size: number;                      // Size in bytes
}
```

## Examples

### Extract and save attachments

```typescript
import MsgParser from 'xiaodao-msg-parser';
import fs from 'node:fs';

const parser = new MsgParser();
const email = parser.parseFile('message.msg');

for (const attachment of email.attachments) {
  if (attachment.content && attachment.filename) {
    fs.writeFileSync(attachment.filename, attachment.content);
    console.log(`Saved: ${attachment.filename} (${attachment.size} bytes)`);
  }
}
```

### Get inline images from HTML body

```typescript
import MsgParser from 'xiaodao-msg-parser';

const parser = new MsgParser();
const email = parser.parseFile('message.msg');

// Find inline images by Content-ID
for (const attachment of email.attachments) {
  if (attachment.contentId) {
    console.log(`Inline image: cid:${attachment.contentId}`);
    // Use in HTML: <img src="cid:contentId">
  }
}
```

### Parse headers for specific information

```typescript
import MsgParser from 'xiaodao-msg-parser';

const parser = new MsgParser();
const email = parser.parseFile('message.msg');

if (email.parsedHeaders) {
  console.log('Message-ID:', email.parsedHeaders['message-id']);
  console.log('Return-Path:', email.parsedHeaders['return-path']);
  console.log('X-Priority:', email.parsedHeaders['x-priority']);
}
```

### Handle different body formats

```typescript
import MsgParser from 'xiaodao-msg-parser';

const parser = new MsgParser();
const email = parser.parseFile('message.msg');

let bodyContent: string;

if (email.bodyHtml) {
  // Use HTML body
  bodyContent = email.bodyHtml;
} else if (email.body) {
  // Use plain text body
  bodyContent = email.body;
} else if (email.bodyRtf) {
  // Use RTF body (may need additional processing)
  bodyContent = email.bodyRtf;
} else {
  bodyContent = '(No body content)';
}
```

## Development

### Build

```bash
npm run build
```

### Test

```bash
npm test
```

### Run example

```bash
node --import tsx examples/parse.ts path/to/message.msg
```

## How It Works

This parser reads the Compound File Binary (CFB/OLE2) structure of `.msg` files and extracts MAPI properties from the following locations:

1. **Top-level properties** - Subject, body, dates, etc.
2. **Recipient streams** (`__recip_version1.0_#*`) - To, CC, BCC recipients
3. **Attachment streams** (`__attach_version1.0_#*`) - File attachments

When recipient information is not available in the standard MAPI streams, the parser falls back to extracting recipient information from:
1. Internet headers (`PidTagInternetHeaders`)
2. Email body content (when headers contain `To:`, `Cc:`, `Bcc:` fields)

## Limitations

- Only supports synchronous parsing
- Does not decrypt encrypted or password-protected messages
- RTF body is returned as-is (no RTF to HTML conversion)
- Named properties with custom GUIDs are not fully supported

## License

MIT

## Related

- [cfb](https://github.com/SheetJS/js-cfb) - Compound File Binary format parser
- [MAPI Property Tags](https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/mapi-property-tags) - Microsoft MAPI documentation
