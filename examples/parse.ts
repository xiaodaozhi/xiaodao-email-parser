import EmailParser from '../index.js';
import path from 'node:path';

const filePath = process.argv[2];
if (!filePath) {
  console.error('Usage: node --import tsx examples/parse.ts <path-to-msg-or-eml-file>');
  process.exit(1);
}

const absolutePath = path.resolve(filePath);
console.log('Parsing:', absolutePath);
console.log('');

const parser = new EmailParser();

try {
  const email = await parser.parseFile(absolutePath);

  console.log('=== Email Summary ===');
  console.log('Format:         ', email.format.toUpperCase());
  console.log('Subject:        ', email.subject);
  console.log('From:           ', email.from ? `${email.from.name} <${email.from.email}>` : 'N/A');
  console.log('To:             ', email.to.map(r => `${r.name} <${r.email}>`).join('; '));
  console.log('CC:             ', email.cc.map(r => `${r.name} <${r.email}>`).join('; '));
  console.log('Date:           ', email.sentDate ? email.sentDate.toISOString() : 'N/A');
  console.log('Message Class:  ', email.messageClass);
  console.log('Importance:     ', email.importance);
  console.log('');

  console.log('=== Body (first 500 chars) ===');
  if (email.body) {
    console.log(email.body.slice(0, 500));
  } else if (email.bodyHtml) {
    console.log('(HTML body, showing plain text extract...)');
    const stripped = email.bodyHtml.replace(/<[^>]+>/g, '').slice(0, 500);
    console.log(stripped);
  } else {
    console.log('(no body content)');
  }
  console.log('');

  console.log('=== Attachments ===');
  if (email.attachments.length === 0) {
    console.log('(none)');
  } else {
    for (const att of email.attachments) {
      console.log(`  - ${att.filename || '(unnamed)'} (${att.contentType}, ${att.size} bytes)`);
    }
  }
  console.log('');

  console.log('=== Internet Headers (first 20) ===');
  if (email.parsedHeaders) {
    const keys = Object.keys(email.parsedHeaders).slice(0, 20);
    for (const key of keys) {
      const val = String(email.parsedHeaders[key]).slice(0, 120);
      console.log(`  ${key}: ${val}`);
    }
  } else {
    console.log('(none)');
  }
} catch (err) {
  console.error('Failed to parse email file:', err instanceof Error ? err.message : String(err));
  process.exit(1);
}
