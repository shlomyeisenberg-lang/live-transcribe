// @vitest-environment node
// ‏Blob של jsdom אינו מממש `arrayBuffer`, והמודול קורא דרכו.
import { describe, expect, it } from 'vitest';
import { crc32, readEntryText, readZip, writeZip, type Bytes, type ZipEntry } from '../../src/engine/docx-parts';
import { editDocxStyles, readDocxStyles } from '../../src/engine/docx-style-package';
import { listStyles } from '../../src/engine/docx-style-edit';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const enc = new TextEncoder();

function entry(name: string, text: string): ZipEntry {
  const data = enc.encode(text) as Bytes;
  return {
    name,
    nameBytes: enc.encode(name) as Bytes,
    versionMadeBy: 20,
    versionNeeded: 10,
    flags: 0,
    method: 0,
    modTime: 0,
    modDate: 0x21,
    crc: crc32(data),
    internalAttrs: 0,
    externalAttrs: 0,
    data,
    uncompressedSize: data.byteLength,
  };
}

const STYLES =
  `<w:styles ${W}>` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Mine"><w:name w:val="שלי"/><w:basedOn w:val="Normal"/></w:style>' +
  '</w:styles>';
const BODY =
  `<w:document ${W}><w:body><w:p><w:pPr><w:pStyle w:val="Mine"/></w:pPr><w:r><w:t>א</w:t></w:r></w:p></w:body></w:document>`;
const HEADER = `<w:hdr ${W}><w:p><w:pPr><w:pStyle w:val="Mine"/></w:pPr></w:p></w:hdr>`;

function docx(): Blob {
  const bytes = writeZip([
    entry('word/document.xml', BODY),
    entry('word/styles.xml', STYLES),
    entry('word/header1.xml', HEADER),
  ]);
  return new Blob([bytes]);
}

async function part(blob: Blob, name: string): Promise<string> {
  const found = readZip(new Uint8Array(await blob.arrayBuffer()))!.find((e) => e.name === name)!;
  return (await readEntryText(found))!;
}

describe('docx style package', () => {
  it('reads styles.xml', async () => {
    expect(await readDocxStyles(docx())).toBe(STYLES);
  });

  it('modifies a style inside the package and leaves the body alone', async () => {
    const result = await editDocxStyles(docx(), { kind: 'modify', id: 'Mine', format: { bold: true } });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(await part(result.docx, 'word/styles.xml')).toContain('<w:b/><w:bCs/>');
    expect(await part(result.docx, 'word/document.xml')).toBe(BODY);
  });

  it('deletes a style and its references in every paragraph part', async () => {
    const result = await editDocxStyles(docx(), { kind: 'delete', id: 'Mine' });
    if (!result.ok) throw new Error(result.message);
    expect(listStyles(await part(result.docx, 'word/styles.xml')).map((s) => s.id)).toEqual(['Normal']);
    expect(await part(result.docx, 'word/document.xml')).not.toContain('pStyle');
    expect(await part(result.docx, 'word/header1.xml')).not.toContain('pStyle');
  });

  it('reports a refusal without touching the package', async () => {
    const result = await editDocxStyles(docx(), { kind: 'delete', id: 'Normal' });
    expect(result).toMatchObject({ ok: false });
  });

  it('reports an unreadable package', async () => {
    const result = await editDocxStyles(new Blob(['not a zip']), { kind: 'delete', id: 'Mine' });
    expect(result).toMatchObject({ ok: false });
  });
});
