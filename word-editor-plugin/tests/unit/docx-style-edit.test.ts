import { describe, expect, it } from 'vitest';
import {
  editStyles,
  isStyleNameTaken,
  listStyles,
  newStyleId,
  readOwnFormat,
  resolveFormat,
  stripStyleReferences,
} from '../../src/engine/docx-style-edit';

const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function sheet(styles: string, defaults = '<w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:styles xmlns:w="${NS}"><w:docDefaults><w:rPrDefault>${defaults}</w:rPrDefault>` +
    `<w:pPrDefault><w:pPr><w:bidi/></w:pPr></w:pPrDefault></w:docDefaults>${styles}</w:styles>`
  );
}

const NORMAL = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>';
const HEADING1 =
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
  '<w:next w:val="Normal"/><w:link w:val="Heading1Char"/><w:qFormat/>' +
  '<w:pPr><w:keepNext/><w:spacing w:before="240"/><w:outlineLvl w:val="0"/></w:pPr>' +
  '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:b/><w:bCs/><w:sz w:val="32"/></w:rPr></w:style>';
const HEADING1_CHAR =
  '<w:style w:type="character" w:customStyle="1" w:styleId="Heading1Char"><w:name w:val="Heading 1 Char"/>' +
  '<w:link w:val="Heading1"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>';

const BASE = sheet(NORMAL + HEADING1 + HEADING1_CHAR);

/** סדר הילדים הישירים של אלמנט, לפי שמותיהם. */
function childOrder(xml: string, parent: string): string[] {
  const match = new RegExp(`<w:${parent}>([\\s\\S]*?)</w:${parent}>`).exec(xml);
  return [...(match?.[1] ?? '').matchAll(/<w:([A-Za-z]+)[\s/>]/g)]
    .map((m) => m[1])
    .filter((name, index, all) => all.indexOf(name) === index);
}

function styleXml(xml: string, id: string): string {
  return new RegExp(`<w:style [^>]*w:styleId="${id}"[\\s\\S]*?</w:style>`).exec(xml)?.[0] ?? '';
}

describe('listStyles', () => {
  it('reads ids, names, chain and flags', () => {
    const styles = listStyles(BASE);
    expect(styles.map((s) => s.id)).toEqual(['Normal', 'Heading1', 'Heading1Char']);
    expect(styles[0]).toMatchObject({ isDefault: true, custom: false });
    expect(styles[1]).toMatchObject({ name: 'heading 1', basedOn: 'Normal', next: 'Normal', link: 'Heading1Char' });
    expect(styles[2]).toMatchObject({ type: 'character', custom: true });
  });

  it('returns nothing for unreadable input', () => {
    expect(listStyles('<not-word/>')).toEqual([]);
    expect(listStyles(null)).toEqual([]);
  });
});

describe('resolveFormat', () => {
  it('layers docDefaults, then the basedOn chain', () => {
    expect(resolveFormat(BASE, 'Normal')).toMatchObject({ fontSizePt: 12 });
    expect(resolveFormat(BASE, 'Heading1')).toMatchObject({ fontSizePt: 16, bold: true, spaceBeforePt: 12 });
  });

  it('prefers the complex-script side, which is what Hebrew text reads', () => {
    const xml = sheet(
      '<w:style w:type="paragraph" w:styleId="X"><w:name w:val="X"/>' +
        '<w:rPr><w:rFonts w:ascii="Arial" w:cs="David"/><w:b w:val="0"/><w:bCs/><w:sz w:val="20"/><w:szCs w:val="28"/></w:rPr></w:style>',
    );
    expect(readOwnFormat(xml, 'X')).toMatchObject({ fontFamily: 'David', bold: true, fontSizePt: 14 });
  });
});

describe('editStyles — modify', () => {
  it('writes both script stacks and keeps unrelated declarations', () => {
    const result = editStyles(BASE, {
      kind: 'modify',
      id: 'Heading1',
      format: { fontFamily: 'David', fontSizePt: 18, italic: true, color: 'ff0000' },
    });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const style = styleXml(result.xml, 'Heading1');
    expect(style).toContain('<w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David"/>');
    expect(style).not.toContain('majorHAnsi');
    expect(style).toContain('<w:sz w:val="36"/><w:szCs w:val="36"/>');
    expect(style).toContain('<w:i/><w:iCs/>');
    expect(style).toContain('<w:color w:val="FF0000"/>');
    // מה שהדיאלוג לא נגע בו נשאר.
    expect(style).toContain('<w:keepNext/>');
    expect(style).toContain('<w:outlineLvl w:val="0"/>');
    expect(style).toContain('<w:b/>');
    expect(resolveFormat(result.xml, 'Heading1')).toMatchObject({ fontFamily: 'David', fontSizePt: 18, italic: true });
  });

  it('keeps rPr and pPr children in schema order', () => {
    const result = editStyles(BASE, {
      kind: 'modify',
      id: 'Heading1',
      format: { bold: false, underline: true, color: '112233', alignment: 'center', indentStartCm: 1, spaceAfterPt: 6 },
    });
    if (!result.ok) throw new Error(result.message);
    const style = styleXml(result.xml, 'Heading1');
    expect(childOrder(style, 'rPr')).toEqual(['rFonts', 'b', 'bCs', 'color', 'sz', 'u']);
    expect(childOrder(style, 'pPr')).toEqual(['keepNext', 'spacing', 'ind', 'jc', 'outlineLvl']);
    expect(style).toContain('<w:b w:val="0"/><w:bCs w:val="0"/>');
    expect(style).toContain('<w:spacing w:before="240" w:after="120"/>');
    expect(style).toContain('<w:ind w:start="567"/>');
  });

  it('null removes the declaration so the value is inherited again', () => {
    const result = editStyles(BASE, { kind: 'modify', id: 'Heading1', format: { fontSizePt: null, spaceBeforePt: null } });
    if (!result.ok) throw new Error(result.message);
    const style = styleXml(result.xml, 'Heading1');
    expect(style).not.toContain('w:sz ');
    expect(style).not.toContain('<w:spacing');
    expect(resolveFormat(result.xml, 'Heading1').fontSizePt).toBe(12);
  });

  it('mirrors run changes to the linked character style', () => {
    const result = editStyles(BASE, { kind: 'modify', id: 'Heading1', format: { fontSizePt: 20, alignment: 'center' } });
    if (!result.ok) throw new Error(result.message);
    const linked = styleXml(result.xml, 'Heading1Char');
    expect(linked).toContain('<w:sz w:val="40"/><w:szCs w:val="40"/>');
    expect(linked).not.toContain('<w:pPr');
  });

  it('writes a hanging indent as hanging', () => {
    const result = editStyles(BASE, { kind: 'modify', id: 'Normal', format: { firstLineCm: -0.5 } });
    if (!result.ok) throw new Error(result.message);
    expect(styleXml(result.xml, 'Normal')).toContain('<w:ind w:hanging="283"/>');
    expect(resolveFormat(result.xml, 'Normal').firstLineCm).toBe(-0.5);
  });

  it('writes the outline level where Word reads headings, in schema order', () => {
    const result = editStyles(BASE, { kind: 'modify', id: 'Normal', format: { outlineLevel: 1, alignment: 'center' } });
    if (!result.ok) throw new Error(result.message);
    const style = styleXml(result.xml, 'Normal');
    expect(style).toContain('<w:jc w:val="center"/><w:outlineLvl w:val="1"/>');
    expect(resolveFormat(result.xml, 'Normal').outlineLevel).toBe(1);
    // ‏`Heading1` יורש מ„רגיל” אבל מצהיר רמה משלו.
    expect(resolveFormat(result.xml, 'Heading1').outlineLevel).toBe(0);
    const body = editStyles(result.xml, { kind: 'modify', id: 'Normal', format: { outlineLevel: null } });
    if (!body.ok) throw new Error(body.message);
    expect(styleXml(body.xml, 'Normal')).not.toContain('outlineLvl');
  });

  it('refuses a basedOn cycle', () => {
    const result = editStyles(BASE, { kind: 'modify', id: 'Normal', basedOn: 'Heading1', format: {} });
    expect(result).toMatchObject({ ok: false });
  });

  it('is a no-op on an empty patch', () => {
    const result = editStyles(BASE, { kind: 'modify', id: 'Heading1', format: {} });
    expect(result).toMatchObject({ ok: true, xml: BASE });
  });
});

describe('editStyles — create', () => {
  it('appends a custom quick style with a readable unique id', () => {
    const result = editStyles(BASE, {
      kind: 'create',
      name: 'ציטוט מקור',
      type: 'paragraph',
      basedOn: 'Normal',
      next: null,
      format: { fontFamily: 'Frank Ruehl', alignment: 'justify' },
    });
    if (!result.ok) throw new Error(result.message);
    expect(result.id).toBe('Style');
    const style = styleXml(result.xml, 'Style');
    expect(style).toContain('w:customStyle="1"');
    const order = ['<w:name', '<w:basedOn', '<w:next', '<w:qFormat', '<w:pPr', '<w:rPr'].map((tag) => style.indexOf(tag));
    expect(order.every((at) => at >= 0)).toBe(true);
    expect([...order].sort((a, b) => a - b)).toEqual(order);
    expect(style).toContain('<w:next w:val="Style"/>');
    expect(style).toContain('<w:jc w:val="both"/>');
    expect(result.xml.endsWith('</w:style></w:styles>')).toBe(true);
    expect(listStyles(result.xml).find((s) => s.id === 'Style')).toMatchObject({ name: 'ציטוט מקור', custom: true });
  });

  it('refuses a duplicate name, case-insensitively', () => {
    expect(isStyleNameTaken(BASE, 'Heading 1')).toBe(true);
    expect(editStyles(BASE, { kind: 'create', name: 'HEADING 1', type: 'paragraph', basedOn: null, next: null, format: {} }))
      .toMatchObject({ ok: false });
  });

  it('derives ids that do not collide', () => {
    expect(newStyleId(BASE, 'heading1')).toBe('heading11');
    expect(newStyleId(BASE, 'My Style')).toBe('MyStyle');
  });

  it('escapes the name', () => {
    const result = editStyles(BASE, { kind: 'create', name: 'A & "B"', type: 'paragraph', basedOn: null, next: null, format: {} });
    if (!result.ok) throw new Error(result.message);
    expect(result.xml).toContain('<w:name w:val="A &amp; &quot;B&quot;"/>');
    expect(listStyles(result.xml).slice(-1)[0]?.name).toBe('A & "B"');
  });
});

describe('editStyles — rename and delete', () => {
  const custom =
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Mine"><w:name w:val="שלי"/><w:basedOn w:val="Heading1"/>' +
    '<w:link w:val="MineChar"/></w:style>' +
    '<w:style w:type="character" w:customStyle="1" w:styleId="MineChar"><w:name w:val="שלי תו"/><w:link w:val="Mine"/></w:style>' +
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Child"><w:name w:val="ילד"/><w:basedOn w:val="Mine"/><w:next w:val="Mine"/></w:style>';
  const xml = sheet(NORMAL + HEADING1 + custom);

  it('renames custom styles only', () => {
    const renamed = editStyles(xml, { kind: 'rename', id: 'Mine', name: 'חדש' });
    if (!renamed.ok) throw new Error(renamed.message);
    expect(listStyles(renamed.xml).find((s) => s.id === 'Mine')?.name).toBe('חדש');
    expect(editStyles(xml, { kind: 'rename', id: 'Heading1', name: 'x' })).toMatchObject({ ok: false });
  });

  it('deletes a custom style with its linked style, and rebases dependents', () => {
    const result = editStyles(xml, { kind: 'delete', id: 'Mine' });
    if (!result.ok) throw new Error(result.message);
    expect(result.removed).toEqual(['Mine', 'MineChar']);
    const ids = listStyles(result.xml).map((s) => s.id);
    expect(ids).toEqual(['Normal', 'Heading1', 'Child']);
    const child = listStyles(result.xml).find((s) => s.id === 'Child');
    expect(child).toMatchObject({ basedOn: 'Heading1', next: null });
  });

  it('refuses built-in styles', () => {
    expect(editStyles(xml, { kind: 'delete', id: 'Heading1' })).toMatchObject({ ok: false });
    expect(editStyles(xml, { kind: 'delete', id: 'Normal' })).toMatchObject({ ok: false });
  });

  it('strips references from body parts', () => {
    const body =
      `<w:document xmlns:w="${NS}"><w:body>` +
      '<w:p><w:pPr><w:pStyle w:val="Mine"/><w:bidi/></w:pPr><w:r><w:rPr><w:rStyle w:val="MineChar"/></w:rPr><w:t>א</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Mine2"/></w:pPr></w:p></w:body></w:document>';
    const out = stripStyleReferences(body, ['Mine', 'MineChar']);
    expect(out).not.toBeNull();
    expect(out).not.toContain('"Mine"');
    expect(out).not.toContain('MineChar');
    expect(out).toContain('<w:pStyle w:val="Mine2"/>');
    expect(stripStyleReferences(body, ['Other'])).toBeNull();
  });
});
