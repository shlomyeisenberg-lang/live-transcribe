// @vitest-environment node
// ‏Blob של jsdom אינו מממש `arrayBuffer`, והמודול קורא דרכו.
/**
 * ברירות המחדל של מסמך חדש. הבדיקות הראשיות רצות על התבנית העברית שנגזרת
 * מהמסמך הריק של המנוע **המותקן** — אותה תבנית שמסמך חדש נפתח ממנה.
 */
import { describe, expect, it } from 'vitest';
import { patchBlankDocumentXml, patchBlankStylesXml } from '../../src/engine/blank-document';
import { deriveHebrewBlankDocx, readZip as readNodeZip } from '../../scripts/blank-docx';
import { editRunDefaults, resolveFormat } from '../../src/engine/docx-style-edit';
import {
  applyNewDocumentDefaults,
  customizeBlankDocx,
  hasNewDocumentDefaults,
  normalizeNewDocumentDefaults,
  newDocumentDefaultsSupported,
  NO_NEW_DOCUMENT_DEFAULTS,
  readTemplateStyles,
  rememberStyle,
  withRunDefaults,
  withStyleDefault,
  type NewDocumentDefaults,
} from '../../src/engine/new-document-defaults';

const DOCX = deriveHebrewBlankDocx({ patchDocument: patchBlankDocumentXml, patchStyles: patchBlankStylesXml });
const TEMPLATE = readNodeZip(DOCX).find((entry) => entry.name === 'word/styles.xml')!.data.toString('utf8');
const blank = (): Blob => new Blob([new Uint8Array(DOCX)]);

function styleXml(xml: string, id: string): string {
  return new RegExp(`<w:style [^>]*w:styleId="${id}">[\\s\\S]*?</w:style>`).exec(xml)![0];
}

function rPrDefault(xml: string): string {
  return /<w:rPrDefault>[\s\S]*?<\/w:rPrDefault>/.exec(xml)![0];
}

describe('editRunDefaults', () => {
  it('גופן וגודל בשני צדי הכתב, וערכת הנושא מוסרת; שאר ההצהרות נשארות', () => {
    const xml = editRunDefaults(TEMPLATE, { fontFamily: 'David', fontSizePt: 14 })!;
    const run = rPrDefault(xml);
    expect(run).toMatch(/<w:rFonts [^>]*w:ascii="David"/);
    expect(run).toMatch(/<w:rFonts [^>]*w:hAnsi="David"/);
    expect(run).toMatch(/<w:rFonts [^>]*w:cs="David"/);
    expect(run).not.toMatch(/Theme=|cstheme=/);
    expect(run).toContain('<w:sz w:val="28"/><w:szCs w:val="28"/>');
    expect(run).toContain('w:bidi="he-IL"');
    // רק docDefaults השתנה.
    expect(xml.replace(rPrDefault(xml), '')).toBe(TEMPLATE.replace(rPrDefault(TEMPLATE), ''));
  });

  it('גיליון בלי docDefaults מקבל אותו ראשון', () => {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    expect(editRunDefaults(`<w:styles ${W}><w:style w:styleId="Normal"/></w:styles>`, { fontSizePt: 13 })).toBe(
      `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:styleId="Normal"/></w:styles>`,
    );
  });
});

describe('applyNewDocumentDefaults', () => {
  it('בלי ברירות מחדל — הגיליון כמות שהוא', () => {
    expect(applyNewDocumentDefaults(TEMPLATE, NO_NEW_DOCUMENT_DEFAULTS)).toBe(TEMPLATE);
  });

  it('„רגיל” מיושר לימין במקום לשני הצדדים, וכותרת 1 בגופן, גודל, צבע וריווח אחרים', () => {
    const defaults: NewDocumentDefaults = {
      run: {},
      styles: {
        Normal: { alignment: 'right' },
        Heading1: { fontFamily: 'Frank Ruehl CLM', fontSizePt: 18, bold: true, color: '1F3864', spaceBeforePt: 12 },
      },
    };
    const xml = applyNewDocumentDefaults(TEMPLATE, defaults);
    expect(resolveFormat(TEMPLATE, 'Normal').alignment).toBe('justify');
    expect(resolveFormat(xml, 'Normal').alignment).toBe('right');
    // בפסקת bidi ‏`left` הוא לוגי — תחילת השורה, כלומר ימין.
    expect(styleXml(xml, 'Normal')).toContain('<w:jc w:val="left"/>');
    const heading = resolveFormat(xml, 'Heading1');
    expect(heading).toMatchObject({ fontFamily: 'Frank Ruehl CLM', fontSizePt: 18, bold: true, color: '1F3864', spaceBeforePt: 12 });
    // הצד המורכב נכתב יחד עם הלטיני — אחרת Word מציג עברית בגופן ובגודל הישנים.
    expect(styleXml(xml, 'Heading1')).toMatch(/<w:rFonts [^>]*w:cs="Frank Ruehl CLM"/);
    expect(styleXml(xml, 'Heading1')).toContain('<w:szCs w:val="36"/>');
    expect(styleXml(xml, 'Heading1')).toContain('<w:bCs/>');
    // הכותרת עדיין יורשת את היישור מ„רגיל”.
    expect(styleXml(xml, 'Heading1')).not.toContain('<w:jc');
    expect(heading.alignment).toBe('right');
  });

  it('סגנון שאינו בתבנית מדולג, והשאר חל', () => {
    const xml = applyNewDocumentDefaults(TEMPLATE, { run: { fontSizePt: 13 }, styles: { NoSuchStyle: { bold: true } } });
    expect(rPrDefault(xml)).toContain('<w:sz w:val="26"/>');
  });
});

describe('rememberStyle — רק ההבדל מהתבנית נשמר', () => {
  it('העיצוב השלם של „רגיל” עם יישור לימין נשמר כיישור בלבד', () => {
    const shown = { ...resolveFormat(TEMPLATE, 'Normal'), alignment: 'right' as const };
    const next = rememberStyle(NO_NEW_DOCUMENT_DEFAULTS, TEMPLATE, 'Normal', shown);
    expect(next!.styles).toEqual({ Normal: { alignment: 'right' } });
  });

  it('חזרה לערך של התבנית מוחקת את הטלאי, והטלאים האחרים נשארים', () => {
    const before: NewDocumentDefaults = { run: { fontSizePt: 13 }, styles: { Normal: { alignment: 'right' }, Heading2: { bold: true } } };
    // מה שהדיאלוג מציג במסמך חדש, עם „רגיל” שהוחזר ליישור לשני הצדדים.
    const current = applyNewDocumentDefaults(TEMPLATE, before);
    const shown = { ...resolveFormat(current, 'Normal'), alignment: 'justify' as const };
    const next = rememberStyle(before, TEMPLATE, 'Normal', shown);
    expect(next).toEqual({ run: { fontSizePt: 13 }, styles: { Heading2: { bold: true } } });
  });

  it('כותרת ביחס לתבנית שכבר נושאת את „רגיל” של המשתמש', () => {
    const before: NewDocumentDefaults = { run: {}, styles: { Normal: { alignment: 'right' } } };
    const current = applyNewDocumentDefaults(TEMPLATE, before);
    const shown = { ...resolveFormat(current, 'Heading1'), fontSizePt: 22 };
    expect(rememberStyle(before, TEMPLATE, 'Heading1', shown)!.styles).toEqual({
      Normal: { alignment: 'right' },
      Heading1: { fontSizePt: 22 },
    });
  });
});

describe('customizeBlankDocx — על התבנית המותקנת', () => {
  it('בלי ברירות מחדל מוחזרת אותה תבנית', async () => {
    const source = blank();
    expect(await customizeBlankDocx(source, NO_NEW_DOCUMENT_DEFAULTS)).toBe(source);
  });

  it('styles.xml נושא את ההגדרות, ושאר החלקים זהים', async () => {
    const defaults: NewDocumentDefaults = { run: { fontFamily: 'David', fontSizePt: 14 }, styles: { Heading1: { fontSizePt: 18 } } };
    const out = await customizeBlankDocx(blank(), defaults);
    expect(out.type).toBe('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    const styles = (await readTemplateStyles(out))!;
    expect(styles).toBe(applyNewDocumentDefaults(TEMPLATE, defaults));
    const before = readNodeZip(DOCX);
    const after = readNodeZip(Buffer.from(await out.arrayBuffer()));
    expect(after.map((entry) => entry.name)).toEqual(before.map((entry) => entry.name));
    for (const entry of before) {
      if (entry.name === 'word/styles.xml') continue;
      expect(after.find((other) => other.name === entry.name)!.data.equals(entry.data)).toBe(true);
    }
  });
});

describe('normalizeNewDocumentDefaults', () => {
  it('ערך חסר או פגום — אין ברירות מחדל', () => {
    for (const raw of [null, undefined, 'x', 3, []]) {
      expect(hasNewDocumentDefaults(normalizeNewDocumentDefaults(raw))).toBe(false);
    }
  });

  it('שדה פגום נזרק לבדו, ומזהה סגנון פגום נזרק', () => {
    expect(
      normalizeNewDocumentDefaults({
        run: { fontFamily: '  David ', fontSizePt: -3 },
        styles: {
          Heading1: { fontSizePt: 18, color: 'red', alignment: 'diagonal', bold: 'yes', italic: false },
          '<bad id>': { bold: true },
          Heading2: { color: 42 },
        },
      }),
    ).toEqual({ run: { fontFamily: 'David' }, styles: { Heading1: { fontSizePt: 18, italic: false } } });
  });

  it('„ברירות מחדל” אחרי „שנה סגנון” על „רגיל” שמצהיר גופן — הבחירה האחרונה חלה', () => {
    // מסמך מהדיסק שבו „רגיל” מצהיר Times New Roman 11; „שנה סגנון” ← מרכז,
    // עם „גם לכל מסמך חדש” — הטלאי נושא גם את הגופן והגודל של המסמך ההוא.
    const disk = TEMPLATE.replace(
      /(<w:style [^>]*w:styleId="Normal">[\s\S]*?)(<\/w:style>)/,
      '$1<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>$2',
    );
    const styled = rememberStyle(NO_NEW_DOCUMENT_DEFAULTS, TEMPLATE, 'Normal', {
      ...resolveFormat(disk, 'Normal'),
      alignment: 'center',
    });
    expect(styled!.styles.Normal?.fontFamily).toBe('Times New Roman');

    // ואז „ברירות מחדל” ← David 15, „גם לכל מסמך חדש”.
    const next = withRunDefaults(styled!, { fontFamily: 'David', fontSizePt: 15 });
    const normal = resolveFormat(applyNewDocumentDefaults(TEMPLATE, next), 'Normal');
    expect(normal.fontFamily).toBe('David');
    expect(normal.fontSizePt).toBe(15);
    // והיישור שנבחר בסגנון נשאר.
    expect(normal.alignment).toBe('center');
    // רק השדה שנשלח יורד: גודל לבד משאיר את הגופן של „רגיל”.
    expect(withRunDefaults(styled!, { fontSizePt: 15 }).styles.Normal?.fontFamily).toBe('Times New Roman');
  });

  it('withRunDefaults ממזג, ו-withStyleDefault עם טלאי ריק מוחק', () => {
    const merged = withRunDefaults({ run: { fontFamily: 'David' }, styles: {} }, { fontSizePt: 14 });
    expect(merged.run).toEqual({ fontFamily: 'David', fontSizePt: 14 });
    const withHeading = withStyleDefault(merged, 'Heading1', { bold: true });
    expect(withStyleDefault(withHeading, 'Heading1', {}).styles).toEqual({});
  });
});

const TOGGLE_TEMPLATE = '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Child"><w:name w:val="Child"/><w:basedOn w:val="Normal"/><w:rPr><w:i/><w:iCs/></w:rPr></w:style></w:styles>';

describe('stored style graph safety', () => {
  it('refuses an imported no-op Normal italic preference against the installed template', () => {
    const disk = '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:i/><w:iCs/></w:rPr></w:style></w:styles>';
    const previous: NewDocumentDefaults = { run: { fontSizePt: 14 }, styles: { Heading1: { color: '123456' } } };
    const snapshot = JSON.stringify(previous);
    expect(rememberStyle(previous, TEMPLATE, 'Normal', resolveFormat(disk, 'Normal'))).toBeNull();
    expect(JSON.stringify(previous)).toBe(snapshot);
  });

  it('rejects unsafe stored Normal italic, retaining run defaults and the original style graph', () => {
    const preferences = { run: { fontFamily: 'David', fontSizePt: 14 }, styles: { Normal: { italic: true }, Heading1: { color: '123456' } } };
    expect(newDocumentDefaultsSupported(TEMPLATE, preferences)).toBe(false);
    const xml = applyNewDocumentDefaults(TEMPLATE, preferences);
    expect(xml).toBe(editRunDefaults(TEMPLATE, preferences.run));
    expect(resolveFormat(xml, 'Heading4').italic).toBe(true);
  });

  it('validates the complete final graph and resolves ancestors before children for either storage order', () => {
    const parent = { italic: true };
    const child = { italic: true, color: '123456' };
    const first: NewDocumentDefaults = { run: {}, styles: { Child: child, Normal: parent } };
    const second: NewDocumentDefaults = { run: {}, styles: { Normal: parent, Child: child } };
    expect(newDocumentDefaultsSupported(TOGGLE_TEMPLATE, { run: {}, styles: { Normal: parent } })).toBe(false);
    expect(newDocumentDefaultsSupported(TOGGLE_TEMPLATE, first)).toBe(true);
    expect(newDocumentDefaultsSupported(TOGGLE_TEMPLATE, second)).toBe(true);
    const xml = applyNewDocumentDefaults(TOGGLE_TEMPLATE, first);
    expect(xml).toBe(applyNewDocumentDefaults(TOGGLE_TEMPLATE, second));
    expect(resolveFormat(xml, 'Normal').italic).toBe(true);
    expect(resolveFormat(xml, 'Child')).toMatchObject({ italic: true, color: '123456' });
    expect(styleXml(xml, 'Child')).not.toMatch(/<w:i(?:Cs)?\b/);
  });

  it('checks inherited bold in both script stacks as well as italic', () => {
    const template = TOGGLE_TEMPLATE.replace(/<w:i\/>/g, '<w:b/>').replace(/<w:iCs\/>/g, '<w:bCs/>');
    expect(newDocumentDefaultsSupported(template, { run: {}, styles: { Normal: { bold: true } } })).toBe(false);
    for (const styles of [{ Child: { bold: true }, Normal: { bold: true } }, { Normal: { bold: true }, Child: { bold: true } }]) {
      const preferences = { run: {}, styles };
      expect(newDocumentDefaultsSupported(template, preferences)).toBe(true);
      const xml = applyNewDocumentDefaults(template, preferences);
      expect(resolveFormat(xml, 'Child').bold).toBe(true);
      expect(styleXml(xml, 'Child')).not.toMatch(/<w:b(?:Cs)?\b/);
    }
  });

  it('rejects an incompatible complete combination in both storage orders', () => {
    for (const styles of [{ Child: { italic: false }, Normal: { italic: true } }, { Normal: { italic: true }, Child: { italic: false } }]) {
      const preferences = { run: {}, styles };
      expect(newDocumentDefaultsSupported(TOGGLE_TEMPLATE, preferences)).toBe(false);
      expect(applyNewDocumentDefaults(TOGGLE_TEMPLATE, preferences)).toBe(TOGGLE_TEMPLATE);
    }
    expect(rememberStyle(NO_NEW_DOCUMENT_DEFAULTS, null, 'Normal', { italic: true })).toBeNull();
  });
});
