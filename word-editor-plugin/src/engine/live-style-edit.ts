/** Apply named style definitions through the public template API, without replacing content or history. */
import type { SuperDoc } from 'superdoc';
import { crc32, writeZip, type ZipEntry } from './docx-parts';
import { changedStyleDefinitions, editStyles, liveStyleTogglesSupported, type StyleEdit } from './docx-style-edit';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type MaybePromise, type DocReceipt } from './document-api';

interface StyleHost {
  activeEditor?: { doc?: {
    info?: (input: object) => MaybePromise<{ revision?: string }>;
    templates?: { apply?: (input: {
      source: { kind: 'base64'; data: string; filename: string }; bodyPolicy: 'preserve';
    }, options: { expectedRevision: string }) => MaybePromise<DocReceipt> };
  } | null } | null;
}

/** A styles-only OPC source: no body, settings, defaults, theme, headers or media from the snapshot. */
export function styleTemplateBytes(styles: string): Uint8Array<ArrayBuffer> {
  const rel = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const parts: Record<string, string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${office}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/document.xml': '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>',
    'word/_rels/document.xml.rels': `<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="${office}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': styles,
  };
  const encoder = new TextEncoder();
  const entries: ZipEntry[] = Object.entries(parts).map(([name, xml]) => {
    const data = encoder.encode(xml);
    return { name, nameBytes: encoder.encode(name), data, crc: crc32(data), uncompressedSize: data.length,
      versionMadeBy: 20, versionNeeded: 20, flags: 0, method: 0, modTime: 0, modDate: 0,
      internalAttrs: 0, externalAttrs: 0 };
  });
  return writeZip(entries);
}

function base64(bytes: Uint8Array): string {
  const chunks: string[] = [];
  for (let start = 0; start < bytes.length; start += 8192) {
    chunks.push(String.fromCharCode(...bytes.subarray(start, start + 8192)));
  }
  return btoa(chunks.join(''));
}

export async function styleRevision(host: SuperDoc): Promise<string | null> {
  const info = (host as StyleHost).activeEditor?.doc?.info;
  const revision = info ? (await info({})).revision : undefined;
  return typeof revision === 'string' ? revision : null;
}

export async function applyLiveStyleEdit(
  host: SuperDoc, xml: string, edit: StyleEdit, expectedRevision: string,
): Promise<{ outcome: CommandOutcome; id?: string }> {
  if (edit.kind === 'delete') return { outcome: { ok: false, message: 'מחיקת סגנון אינה נתמכת בלי לפגוע בהיסטוריית המסמך', reason: 'unsupported' } };
  const edited = editStyles(xml, edit);
  if (!edited.ok) return { outcome: { ok: false, message: edited.message, reason: 'style-edit-refused' } };
  if (!liveStyleTogglesSupported(xml, edited.xml, edit, edited.id)) {
    return { outcome: { ok: false, message: 'השינוי הזה בירושת מודגש או נטוי אינו נתמך כרגע. בחרו סגנון בסיס אחר; המסמך לא שונה.', reason: 'unsupported-style-inheritance' } };
  }
  const definitions = changedStyleDefinitions(xml, edited.xml);
  if (!definitions) return { outcome: { ok: false, message: 'גיליון הסגנונות אינו קריא', reason: 'style-edit-refused' } };
  if (edited.xml === xml) return { outcome: { ok: true }, id: edited.id };
  const apply = (host as StyleHost).activeEditor?.doc?.templates?.apply;
  if (!apply) return { outcome: { ok: false, message: 'ניהול סגנונות אינו נתמך בגרסת המנוע הזאת', reason: 'unsupported' } };
  try {
    const receipt = await apply({ source: { kind: 'base64', data: base64(styleTemplateBytes(definitions)), filename: 'style-edit.docx' }, bodyPolicy: 'preserve' }, { expectedRevision });
    if (receipt?.success !== true) return { outcome: { ok: false, message: receiptFailureText('שינוי הסגנון נכשל', receipt), reason: receipt?.failure?.code ?? 'engine-refused' } };
    return { outcome: { ok: true }, id: edited.id };
  } catch (error) {
    return { outcome: { ok: false, message: thrownText('שינוי הסגנון נכשל', error), reason: 'threw' } };
  }
}
