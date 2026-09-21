import EmailParser from './src/email-parser.js';

export { default as MsgParser } from './src/msg-parser.js';
export { default as EmlParser } from './src/eml-parser.js';
export { EmailParser };
export { detectEmailFormat } from './src/email-parser.js';
export default EmailParser;
export type {
  EmailFormat,
  EmailFormatHint,
  ParseOptions,
  EmailParserOptions,
  EmlParserOptions,
  ParsedEmail,
  EmailRecipient,
  EmailAttachment,
  RecipientInfo,
  EmailAttachmentInfo,
} from './src/types.js';
