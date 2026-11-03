import { describe, expect, it } from 'vitest';
import { formatPatch, selectionStyleFormat, withoutThemedFont } from '../../src/engine/style-from-selection';
import { resolveFormat } from '../../src/engine/docx-style-edit';
import { emptyParagraphFormat, type ParagraphDeclarations } from '../../src/engine/paragraph-format';

const NONE: ParagraphDeclarations = {
  spaceBefore: false,
  spaceAfter: false,
  line: false,
  indentStart: false,
  indentEnd: false,
  firstLine: false,
};

/** מתאם שמדווח כמו המנוע על ריצה רגילה: שחור „אוטומטי” ויישור לימין. */
function adapter(values: Record<string, unknown>, active: Record<string, boolean> = {}) {
  return {
    has: (id: string) => id in values || id in active,
    getState: (id: string) => ({ supported: true, enabled: true, active: active[id] ?? false, value: values[id] }) as never,
  };
}

describe('selectionStyleFormat', () => {
  it('לוקח מהפסקה רק את מה שהיא מצהירה — 0 שאינו מוצהר הוא של הסגנון', () => {
    const snapshot = emptyParagraphFormat();
    expect(selectionStyleFormat(null, { snapshot, declared: NONE })).toEqual({});

    const declared = { ...NONE, spaceAfter: true };
    const withAfter = { ...snapshot, spacing: { ...snapshot.spacing, afterTwips: 120 } };
    expect(selectionStyleFormat(null, { snapshot: withAfter, declared })).toEqual({ spaceAfterPt: 6 });
  });

  it('קורא את מצב הריצה מהפקודות', () => {
    const format = selectionStyleFormat(
      adapter({ 'font-family': 'David', 'font-size': 18, 'text-color': '#c0504d', 'text-align': 'center' }, { bold: true }),
      null,
    );
    expect(format).toMatchObject({ fontFamily: 'David', fontSizePt: 18, bold: true, color: 'C0504D', alignment: 'center' });
  });
});

describe('formatPatch', () => {
  it('שחור וימין שהמנוע מדווח על סגנון שאינו מצהיר אותם אינם שינוי', () => {
    expect(formatPatch({ color: '000000', alignment: 'right', bold: false }, {})).toEqual({});
  });

  it('שינוי אמיתי נכתב', () => {
    expect(formatPatch({ fontSizePt: 22, spaceAfterPt: 4 }, { fontSizePt: 16, spaceAfterPt: 4 })).toEqual({ fontSizePt: 22 });
  });
});

describe('גופן ערכת נושא', () => {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const xml =
    `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:cs="Arial"/></w:rPr></w:rPrDefault></w:docDefaults>` +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
    '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:cstheme="majorBidi"/></w:rPr></w:style></w:styles>';

  it('אינו מוצג כגופן של רמה נמוכה יותר', () => {
    expect(resolveFormat(xml, 'Normal').fontFamily).toBe('Arial');
    expect(resolveFormat(xml, 'Heading1').fontFamily).toBe('');
  });

  it('„מהבחירה” אינו מחליף אותו בשם קבוע', () => {
    const baseline = resolveFormat(xml, 'Heading1');
    const wanted = withoutThemedFont({ fontFamily: 'Aptos Display', fontSizePt: 28 }, baseline);
    expect(formatPatch(wanted, baseline)).toEqual({ fontSizePt: 28 });
    // סגנון עם שם גופן רגיל — כן.
    expect(withoutThemedFont({ fontFamily: 'David' }, resolveFormat(xml, 'Normal'))).toEqual({ fontFamily: 'David' });
  });
});
