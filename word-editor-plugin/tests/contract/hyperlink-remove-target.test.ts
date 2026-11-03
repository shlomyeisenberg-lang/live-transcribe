/**
 * חוזה ה-target של `hyperlinks.remove`, מול ה-`.d.ts` **האמיתי** של המנוע.
 *
 * ## למה זה נבדק כך
 *
 * `hyperlinks-manage.ts` שלח בעבר `remove({ within: TextAddress })` — טווח
 * טקסט, בדיוק כפי ש-`wrap` מצפה. אבל `HyperlinksRemoveInput` (הקובץ שקבוע
 * למטה) דורש `target: HyperlinkTarget` — כתובת **צומת הקישור עצמו**
 * (`InlineNodeAddress & {nodeType:'hyperlink'}`), לא טווח. התוצאה שנמדדה
 * בדפדפן: `TypeError: Cannot read properties of undefined (reading
 * 'anchor')`, בלי שה-`<w:hyperlink>` הוסר מהמסמך.
 *
 * הבדיקה כאן קוראת את ה-`.d.ts` הארוז בפועל ולא מניחה את צורתו — שינוי
 * עתידי בחוזה (הוספת `within`, שינוי שם השדה) יפיל את הבדיקה הזאת, לא רק
 * את ההתנהגות בדפדפן.
 *
 * מה שלא נבדק כאן: ההתנהגות בדפדפן החי (זו באימות QA) ולוגיקת ההתאמה בין
 * הבחירה לכתובת הקישור — זו נבדקת ב-`tests/unit/hyperlinks-manage.test.ts`.
 */
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

const HYPERLINKS_TYPES_DTS = join(
  process.cwd(),
  'node_modules/superdoc/dist/document-api/src/hyperlinks/hyperlinks.types.d.ts',
);

function read(path: string): string {
  return readFileSync(path, 'utf8');
}

function interfaceBody(text: string, name: string): string {
  const start = text.indexOf(`export interface ${name} {`);
  if (start === -1) throw new Error(`הממשק ${name} לא נמצא — חוזה הקישורים של superdoc השתנה`);
  return text.slice(start, text.indexOf('}', start));
}

describe('HyperlinksRemoveInput — target הוא HyperlinkTarget, לא within', () => {
  const dts = read(HYPERLINKS_TYPES_DTS);

  it('`HyperlinksRemoveInput.target` הוא `HyperlinkTarget`', () => {
    const body = interfaceBody(dts, 'HyperlinksRemoveInput');
    expect(body).toContain('target: HyperlinkTarget');
    expect(body).not.toContain('within');
  });

  it('`HyperlinkTarget` הוא כתובת inline מסוג hyperlink (לא TextAddress)', () => {
    expect(dts).toContain(
      "export type HyperlinkTarget = InlineNodeAddress & {\n    nodeType: 'hyperlink';\n};",
    );
  });

  it('`HyperlinksWrapInput.target` דווקא `TextAddress` — כדי שהאבחנה מ-remove לא תיטשטש', () => {
    const body = interfaceBody(dts, 'HyperlinksWrapInput');
    expect(body).toContain('target: TextAddress');
  });

  it('`HyperlinksListQuery` מדגם ל-`items[]` — לא `stories[].hyperlinks[]`', () => {
    // HyperlinksListResult = DiscoveryOutput<HyperlinkDomain>, ו-DiscoveryOutput
    // הוא DiscoveryResult<DiscoveryItem<TDomain>> — עוגן: HyperlinkDomain נושא
    // `address`, שממנו נגזר ה-target הנכון ל-remove.
    expect(dts).toContain('export interface HyperlinkDomain {\n    address: HyperlinkTarget;');
  });
});
