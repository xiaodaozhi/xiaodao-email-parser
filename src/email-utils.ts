import { addressParser, decodeWords } from 'postal-mime';
import type { Address, Header } from 'postal-mime';
import type { EmailRecipient } from './types.js';

const CODE_PAGE_LABELS: Record<number, string> = {
  65001: 'utf-8',
  1200: 'utf-16le',
  20127: 'us-ascii',
  28591: 'iso-8859-1',
  936: 'gbk',
  950: 'big5',
  1250: 'windows-1250',
  1251: 'windows-1251',
  1252: 'windows-1252',
  1253: 'windows-1253',
  1254: 'windows-1254',
  1255: 'windows-1255',
  1256: 'windows-1256',
  1257: 'windows-1257',
  1258: 'windows-1258',
};

export function parseHeaderBlock(raw: string): Record<string, string> | null {
  if (!raw.trim()) return null;

  const result: Record<string, string> = {};
  let currentKey: string | null = null;

  for (const line of raw.split(/\r?\n/)) {
    if (/^[ \t]/.test(line) && currentKey) {
      result[currentKey] += ` ${line.trim()}`;
      continue;
    }

    const separator = line.indexOf(':');
    if (separator <= 0) {
      currentKey = null;
      continue;
    }

    const key = line.slice(0, separator).trim().toLowerCase();
    if (!/^[!-9;-~]+$/.test(key)) {
      currentKey = null;
      continue;
    }

    currentKey = key;
    const value = line.slice(separator + 1).trim();
    result[key] = result[key] ? `${result[key]}, ${value}` : value;
  }

  return Object.keys(result).length > 0 ? result : null;
}

export function headersToRecord(headers: Header[]): Record<string, string> | null {
  if (headers.length === 0) return null;

  const result: Record<string, string> = {};
  for (const header of headers) {
    result[header.key] = result[header.key]
      ? `${result[header.key]}, ${header.value}`
      : header.value;
  }
  return result;
}

export function extractRawHeaders(buffer: Buffer): string | null {
  const source = buffer.toString('latin1');
  const match = /\r?\n\r?\n/.exec(source);
  if (!match || match.index === 0) return null;
  return source.slice(0, match.index);
}

function flattenAddresses(addresses: Address[]): EmailRecipient[] {
  const result: EmailRecipient[] = [];
  for (const item of addresses) {
    if (item.group) {
      result.push(...flattenAddresses(item.group));
    } else {
      result.push({
        name: item.name ? decodeWords(item.name) : null,
        email: item.address || null,
      });
    }
  }
  return result;
}

export function parseAddressHeader(value: string | undefined): EmailRecipient[] {
  if (!value) return [];
  try {
    return flattenAddresses(addressParser(value));
  } catch {
    return [];
  }
}

export function addressToRecipients(address: Address | Address[] | undefined): EmailRecipient[] {
  if (!address) return [];
  return flattenAddresses(Array.isArray(address) ? address : [address]);
}

export function toValidDate(value: string | undefined): Date | null {
  if (!value) return null;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

export function receivedDateFromHeaders(headers: Header[]): Date | null {
  const deliveryDate = headers.find(header => header.key === 'delivery-date')?.value;
  const parsedDeliveryDate = toValidDate(deliveryDate);
  if (parsedDeliveryDate) return parsedDeliveryDate;

  const received = headers.find(header => header.key === 'received')?.value;
  if (!received) return null;
  const separator = received.lastIndexOf(';');
  return separator >= 0 ? toValidDate(received.slice(separator + 1).trim()) : null;
}

export function normalizeSubject(subject: string | null | undefined): string | null {
  if (!subject) return null;
  const normalized = subject.replace(/^\s*((re|fw|fwd)\s*:\s*)+/i, '').trim();
  return normalized || subject.trim() || null;
}

export function createPreview(text: string | null | undefined, html?: string | null): string | null {
  const source = text || html?.replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>');
  if (!source) return null;
  const compact = source.replace(/\s+/g, ' ').trim();
  return compact ? compact.slice(0, 255) : null;
}

export function parseImportance(headers: Record<string, string> | null): 'low' | 'normal' | 'high' {
  if (!headers) return 'normal';
  const importance = headers.importance?.trim().toLowerCase();
  if (importance === 'high') return 'high';
  if (importance === 'low') return 'low';

  const priority = headers['x-priority'] || headers['x-msmail-priority'] || headers.priority;
  if (/high|urgent/i.test(priority || '')) return 'high';
  if (/low|non-urgent/i.test(priority || '')) return 'low';
  const numeric = priority ? Number.parseInt(priority, 10) : Number.NaN;
  if (numeric === 1 || numeric === 2) return 'high';
  if (numeric === 4 || numeric === 5) return 'low';
  return 'normal';
}

export function decodeBuffer(buffer: Buffer, codePage?: number): string {
  const label = codePage ? CODE_PAGE_LABELS[codePage] : undefined;
  try {
    return new TextDecoder(label || 'utf-8').decode(buffer).replace(/\0+$/, '');
  } catch {
    return buffer.toString('utf8').replace(/\0+$/, '');
  }
}

export function normalizeContentId(value: string | null | undefined): string | null {
  if (!value) return null;
  const normalized = value.trim().replace(/^<|>$/g, '');
  return normalized || null;
}
