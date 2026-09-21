# xiaodao-email-parser

[English](README.md) | 简体中文

面向 Node.js 和 TypeScript 的电子邮件文件解析器，可自动识别并解析：

- Outlook `.msg`（Compound File Binary / OLE2 + MAPI）
- RFC 5322 / MIME `.eml`

默认导出的 `EmailParser` 会根据文件内容自动识别格式，并返回一致的 `ParsedEmail` 结构。同时提供专用的 `MsgParser` 和 `EmlParser`，便于只处理单一格式，或将既有同步 MSG 调用迁移到 `xiaodao-email-parser`。

## 功能

- 自动识别 MSG 与 EML，不依赖文件扩展名
- 解析发件人、回复地址、收件人、主题、日期、重要性和 Message-ID
- 支持纯文本、HTML 和 MSG 压缩 RTF 正文
- 支持 multipart、base64、quoted-printable、RFC 2047 编码字和 RFC 2231 文件名
- 提取普通附件与内嵌附件的内容、MIME 类型和 Content-ID
- 保留原始邮件头、规范化邮件头和底层格式属性
- 提供 MIME 嵌套深度、邮件头大小和输入体积限制
- 完整 TypeScript 类型声明

## 安装

```bash
npm install xiaodao-email-parser
```

需要 Node.js 18 或更高版本。

## 快速开始：自动识别 MSG / EML

```typescript
import EmailParser from 'xiaodao-email-parser';

const parser = new EmailParser();
const email = await parser.parseFile('message.eml'); // 也可以是 message.msg

console.log(email.format);      // 'eml' 或 'msg'
console.log(email.subject);
console.log(email.from);
console.log(email.to);
console.log(email.attachments);
```

也可以解析 `Buffer`：

```typescript
import fs from 'node:fs/promises';
import EmailParser from 'xiaodao-email-parser';

const source = await fs.readFile('message.msg');
const email = await new EmailParser().parse(source);
```

统一解析器是异步 API，即使输入是 MSG 也应始终使用 `await`。

## 专用解析器

### 同步解析 MSG

```typescript
import { MsgParser } from 'xiaodao-email-parser';

const email = new MsgParser().parseFile('message.msg');
console.log(email.body, email.bodyHtml, email.bodyRtf);
```

`MsgParser.parse()` 和 `MsgParser.parseFile()` 保持同步，适合从原项目迁移的代码。

### 解析 EML

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

`EmlParser` 默认最多接受 100 MiB 输入。MIME 解析还带有邮件头总大小和嵌套深度保护。

## API

包导出关系：

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

默认使用内容自动识别，也可以明确指定格式：

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

仅处理 Outlook MSG，所有方法都是同步方法。输入不是有效 CFB/MAPI 消息时会抛出描述性错误。

### `EmlParser`

```typescript
new EmlParser(options?: EmlParserOptions)
parser.parse(buffer: Buffer): Promise<ParsedEmail>
parser.parseFile(filePath: string): Promise<ParsedEmail>
```

仅处理 RFC 5322/MIME EML，支持以下限制选项：

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

### `detectEmailFormat`

```typescript
import { detectEmailFormat } from 'xiaodao-email-parser';

const format = detectEmailFormat(buffer); // 'msg' | 'eml' | null
```

MSG 使用 CFB 文件签名识别；EML 使用头部结构与标准邮件头识别。无法可靠判断时返回 `null`，统一解析器则抛出描述性错误。

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

对于 EML，`_rawProperties` 保存 sender、return-path、references 以及未折叠的邮件头数组；对于 MSG，它保存解析后的底层 MAPI 属性。

### 附件

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

保存附件示例：

```typescript
import fs from 'node:fs/promises';

for (const attachment of email.attachments) {
  if (attachment.filename && attachment.content) {
    await fs.writeFile(attachment.filename, attachment.content);
  }
}
```

## 架构

```text
EmailParser (异步、自动识别)
├── MsgParser (同步、CFB/MAPI)
└── EmlParser (异步、RFC 5322/MIME)
```

两个专用解析器负责各自格式，公共工具统一处理地址、邮件头、主题、预览、日期和 Content-ID，最终都映射到同一结果类型。

## 开发

```bash
npm run build
npm test
node --import tsx examples/parse.ts path/to/message.msg
node --import tsx examples/parse.ts path/to/message.eml
```

## 限制

- 不解密加密或受密码保护的邮件。
- MSG 自定义命名属性和嵌入式 MSG 附件尚未完整展开。
- `parsedHeaders` 为兼容旧 API 使用键值对象；重复邮件头会合并。EML 的原始头部数组保存在 `_rawProperties.headers`。
- 解析器只解析内容，不执行 HTML、脚本或附件。

## License

MIT
