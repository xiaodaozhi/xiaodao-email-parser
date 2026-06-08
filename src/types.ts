// server/utils/tools/msgparser/types.ts

// ---------------------------------------------------------------------------
//  Public API types
// ---------------------------------------------------------------------------

export interface EmailRecipient {
  name: string | null;
  email: string | null;
}

export interface EmailAttachment {
  filename: string | null;
  content: Buffer | null;
  contentType: string;
  contentId: string | null;
  contentLocation: string | null;
  size: number;
}

export interface ParsedEmail {
  subject: string | null;
  from: EmailRecipient | null;
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
  headers: string | null;
  parsedHeaders: Record<string, string> | null;
  preview: string | null;
  _rawProperties: Record<string, unknown>;
}

// ---------------------------------------------------------------------------
//  Internal types
// ---------------------------------------------------------------------------

export interface RecipientInfo {
  name: string | null;
  email: string | null;
  type: string | null;
}

export interface EmailAttachmentInfo {
  filename: string | null;
  content: Buffer | null;
  contentType: string;
  contentId: string | null;
  contentLocation: string | null;
  size: number;
}

export interface RawEmailData {
  properties: Record<string, unknown>;
  recipients: RecipientInfo[];
  attachments: EmailAttachmentInfo[];
}

export interface TaggedValue {
  value: unknown;
  type: number;
}

export interface SubStorageBucket {
  props: Buffer | null;
  substg: Record<string, Buffer>;
}

export interface EntryStore {
  topPropsStream: Buffer | null;
  topSubstg: Record<string, Buffer>;
  recipients: Record<string, SubStorageBucket>;
  attachments: Record<string, SubStorageBucket>;
}

export interface PropertyTypeInfo {
  name: string;
  fixedSize: number;
}

export interface TypedReadResult<T = unknown> {
  value: T | undefined;
  bytesRead: number;
}
