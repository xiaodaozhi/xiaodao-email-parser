# xiaodao-email-parser

[![npm version](https://img.shields.io/npm/v/xiaodao-email-parser)](https://www.npmjs.com/package/xiaodao-email-parser)
[![npm downloads](https://img.shields.io/npm/dm/xiaodao-email-parser)](https://www.npmjs.com/package/xiaodao-email-parser)
[![CI](https://github.com/xiaodaozhi/xiaodao-email-parser/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/xiaodaozhi/xiaodao-email-parser/actions/workflows/ci.yml)
[![TypeScript](https://img.shields.io/badge/TypeScript-3178C6?logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
[![Node.js version](https://img.shields.io/node/v/xiaodao-email-parser)](https://www.npmjs.com/package/xiaodao-email-parser)
[![License](https://img.shields.io/npm/l/xiaodao-email-parser)](LICENSE)

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

需要 Node.js 20 或更高版本。本包以 ES 模块形式发布。如需将快速开始中的示例作为 JavaScript 运行，请使用 `.mjs` 文件，或在项目的 `package.json` 中设置 `"type": "module"`。

### 从 1.0.7 升级

2.x 要求 Node.js 20 或更高版本，而 1.0.7 的最低要求为 Node.js 18。`EmailParser` 仍为异步 API，请继续使用 `await`；`MsgParser` 仍为同步 API。如果项目依赖 MIME 解析细节或嵌套邮件，请在升级时测试相关场景：底层 `postal-mime` 依赖已从 3.x 升级到 4.x。

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

例如，一封 EML 邮件可能返回 `format: 'eml'`、主题 `'Meeting'`、发件人 `{ name: 'Alice', email: 'alice@example.com' }` 和空附件数组。实际值取决于邮件内容；缺失字段可能是 `null` 或空数组。

也可以解析 `Buffer`：

```typescript
import fs from 'node:fs/promises';
import EmailParser from 'xiaodao-email-parser';

const source = await fs.readFile('message.msg');
const email = await new EmailParser().parse(source);
```

统一解析器是异步 API，即使输入是 MSG 也应始终使用 `await`。

读取文件或解析邮件可能失败，请处理被拒绝的 Promise，例如：

```typescript
try {
  const email = await new EmailParser().parseFile('message.eml');
  console.log(email.subject);
} catch (error) {
  console.error('无法读取或解析邮件：', error);
}
```

常见失败原因包括文件无法读取、邮件格式无法识别或内容无效，以及 EML 超过大小或 MIME 解析限制。

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

`EmlParser` 默认最多接受 100 MiB 输入。该限制仅适用于 EML，`MsgParser` 没有对应的输入大小选项。两种 `parseFile()` 都会先将整个文件读入内存，再开始解析，因此 EML 的大小检查发生在读取之后。处理不可信文件时，应在加载内容前限制文件大小，并考虑检查与读取之间文件可能发生变化。

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

仅处理 RFC 5322/MIME EML，支持以下选项：

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

| 选项 | 默认值 | 作用 |
| --- | --- | --- |
| `maxInputSize` | 100 MiB | EML 输入的最大字节数，必须是正的安全整数。 |
| `rfc822Attachments` | `false` | 将没有 Content-Disposition 的 `message/rfc822` 部分视为附件。 |
| `forceRfc822Attachments` | `false` | 将所有 `message/rfc822` 部分视为附件。 |
| `maxNestingDepth` | `256` | MIME 部分的最大嵌套深度。 |
| `maxHeadersSize` | `2097152` | MIME 部分和内联嵌套邮件的邮件头总大小上限，单位为字节。 |
| `maxRfc822NestingDepth` | `10` | 内联 `message/rfc822` 的最大递归深度；更深层的邮件会作为附件返回。设为 `0` 可禁用内联解析。 |

上述 MIME 选项的默认值来自当前的 `postal-mime` 4.2.x 依赖。使用统一解析器时，可在 `eml` 下传入相同选项：

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

MSG 使用 CFB 文件签名识别，并对可读取的容器检查根目录下的 MAPI 属性流；截断或无法读取的 CFB 容器仍会被视为可能的 MSG，以便解析时报告错误。EML 使用头部结构与标准邮件头识别。无法可靠判断时返回 `null`，统一解析器则抛出描述性错误。

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

每个 `EmailRecipient` 包含可为 `null` 的 `name` 和 `email`。`from` 是单个地址对象或 `null`；`replyTo`、`to`、`cc` 和 `bcc` 是数组，可能为空。

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

下面的示例将可用的附件内容保存到新建的临时目录。文件名由程序生成，避免直接使用邮件提供的文件名；`wx` 可以防止覆盖已有文件。

```typescript
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';

const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), 'email-attachments-'));
for (const [index, attachment] of email.attachments.entries()) {
  if (attachment.content !== null) {
    const outputPath = path.join(outputDir, `attachment-${index}`);
    await fs.writeFile(outputPath, attachment.content, { flag: 'wx' });
    console.log('已保存：', outputPath, '原始文件名：', attachment.filename);
  }
}
```

附件文件名和内容均来自邮件，应视为不可信数据。如果需要保留原始文件名或扩展名，请按照应用的存储与安全要求进行验证。

## 架构

```text
EmailParser (异步、自动识别)
├── MsgParser (同步、CFB/MAPI)
└── EmlParser (异步、RFC 5322/MIME)
```

两个专用解析器负责各自格式，公共工具统一处理地址、邮件头、主题、预览、日期和 Content-ID，最终都映射到同一结果类型。

## 开发

```bash
npm ci
npm run build
npm test
node --import tsx examples/parse.ts path/to/message.msg
node --import tsx examples/parse.ts path/to/message.eml
```

## 限制

- 不解密加密或受密码保护的邮件。
- MSG 自定义命名属性和嵌入式 MSG 附件尚未完整展开。
- 超过 `maxRfc822NestingDepth` 的嵌套 EML 会作为附件返回，而不会继续内联展开。底层解析器的 `rfc822DepthExceeded` 标记未在 `EmailAttachment` 中暴露；扫描邮件时，不应假定附件列表已经包含所有深层嵌套的内容。
- `parsedHeaders` 为兼容旧 API 使用键值对象；重复邮件头会合并。EML 的原始头部数组保存在 `_rawProperties.headers`。
- 解析器只解析内容，不执行 HTML、脚本或附件。`bodyHtml` 是不可信内容，在网页中展示前应进行适当的 HTML 清理。

## License

MIT
