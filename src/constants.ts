import type { PropertyTypeInfo } from './types.js';

// ---------------------------------------------------------------------------
//  MAPI Property ID → canonical name map
//  (hex string key, lowercase)
// ---------------------------------------------------------------------------
// Using fromEntries to handle duplicate keys (multiple hex IDs map to the same
// canonical name — e.g. ANSI + Unicode variants of the same property).
const PROPERTY_ENTRIES: [string, string][] = [
  // Identification
  ['0037', 'subject'],
  ['0e1d', 'normalizedSubject'],

  // Sender
  ['0c1a', 'senderName'],
  ['0c1f', 'senderEmail'],
  ['5d01', 'senderSmtpEmail'],
  ['0042', 'sentRepresentingName'],
  ['0065', 'sentRepresentingEmail'],
  ['5d02', 'sentRepresentingSmtpEmail'],

  // Body
  ['1000', 'body'],
  ['1013', 'bodyHtml'],
  ['1009', 'bodyRtfCompressed'],
  ['3fd9', 'preview'],

  // Dates
  ['0039', 'clientSubmitTime'],
  ['0e06', 'messageDeliveryTime'],
  ['3007', 'creationTime'],
  ['3008', 'lastModificationTime'],

  // Message info
  ['0017', 'importance'],
  ['001a', 'messageClass'],
  ['0e08', 'messageSize'],
  ['0e07', 'messageFlags'],
  ['3fde', 'internetCodePage'],
  ['007d', 'internetHeaders'],

  // Conversation
  ['0070', 'conversationTopic'],
  ['0071', 'conversationIndex'],

  // Recipient properties
  ['3001', 'displayName'],
  ['3002', 'emailAddressType'],
  ['3003', 'emailAddress'],
  ['39fe', 'smtpEmailAddress'],
  ['5ff6', 'recipientDisplayName'],
  ['0c15', 'recipientType'],

  // Attachment properties
  ['3701', 'attachmentData'],
  ['3704', 'attachmentFilename'],
  ['3705', 'attachmentMethod'],
  ['3707', 'attachmentLongFilename'],
  ['370e', 'attachmentMimeType'],
  ['3712', 'attachmentContentId'],
  ['3713', 'attachmentContentLocation'],
  ['0e20', 'attachmentSize'],
];

export const PROPERTY_IDS: Record<string, string> = Object.fromEntries(PROPERTY_ENTRIES);

// ---------------------------------------------------------------------------
//  MAPI Property Types
//  fixedSize: >0 = fixed size, -1 = variable-length, -2 = multi-valued
// ---------------------------------------------------------------------------
export const PROPERTY_TYPES: Record<number, PropertyTypeInfo> = {
  0x0001: { name: 'PT_NULL',      fixedSize: 0 },
  0x0002: { name: 'PT_SHORT',     fixedSize: 2 },
  0x0003: { name: 'PT_LONG',      fixedSize: 4 },
  0x0004: { name: 'PT_FLOAT',     fixedSize: 4 },
  0x0005: { name: 'PT_DOUBLE',    fixedSize: 8 },
  0x0006: { name: 'PT_CURRENCY',  fixedSize: 8 },
  0x0007: { name: 'PT_APPTIME',   fixedSize: 8 },
  0x000a: { name: 'PT_ERROR',     fixedSize: 4 },
  0x000b: { name: 'PT_BOOLEAN',   fixedSize: 2 },
  0x0014: { name: 'PT_I8',        fixedSize: 8 },
  0x001e: { name: 'PT_STRING8',   fixedSize: -1 },
  0x001f: { name: 'PT_UNICODE',   fixedSize: -1 },
  0x0040: { name: 'PT_SYSTIME',   fixedSize: 8 },
  0x0048: { name: 'PT_CLSID',     fixedSize: 16 },
  0x0102: { name: 'PT_BINARY',    fixedSize: -1 },

  // Multi-valued
  0x1002: { name: 'PT_MV_SHORT',    fixedSize: -2 },
  0x1003: { name: 'PT_MV_LONG',     fixedSize: -2 },
  0x1004: { name: 'PT_MV_FLOAT',    fixedSize: -2 },
  0x1005: { name: 'PT_MV_DOUBLE',   fixedSize: -2 },
  0x1014: { name: 'PT_MV_I8',       fixedSize: -2 },
  0x101e: { name: 'PT_MV_STRING8',  fixedSize: -2 },
  0x101f: { name: 'PT_MV_UNICODE',  fixedSize: -2 },
  0x1040: { name: 'PT_MV_SYSTIME',  fixedSize: -2 },
  0x1102: { name: 'PT_MV_BINARY',   fixedSize: -2 },
};

// ---------------------------------------------------------------------------
//  Recipient type codes
// ---------------------------------------------------------------------------
export const RECIPIENT_TYPES: Record<number, string> = {
  1: 'to',
  2: 'cc',
  3: 'bcc',
};

// ---------------------------------------------------------------------------
//  Importance levels
// ---------------------------------------------------------------------------
export const IMPORTANCE: Record<number, string> = {
  0: 'low',
  1: 'normal',
  2: 'high',
};

// ---------------------------------------------------------------------------
//  FILETIME epoch offset (100-ns intervals between 1601-01-01 and 1970-01-01)
// ---------------------------------------------------------------------------
export const FILETIME_EPOCH_OFFSET = 11644473600000; // ms
