/**
 * שער נגד הכפתור המת.
 *
 * זו לא בדיקת סגנון. שלוש לשוניות נבנו מראש עם תשעה כפתורים בלי `@click`
 * בכלל — לא disabled, לא שום סימן, נראים בדיוק כמו כפתור עובד. שניים מהם אף
 * הציגו קיצור מקלדת (`Alt+Ctrl+F`, `F7`) שלא נרשם בשום מקום. אף בדיקה לא
 * נכשלה, ואף typecheck לא התלונן: פקד שאינו עושה כלום הוא HTML תקין לחלוטין.
 *
 * הכלל „לכל פקד יש מטפל, או שהוא מנוטרל במפורש” **אינו כאן יותר**: הוא נמדד
 * מעכשיו בהרכבה (tests/component/ribbon-tabs.test.ts), שלוחצת על כל כפתור בכל
 * שמונה הלשוניות ודורשת שמשהו נצפה יקרה — ולכן היא תופסת גם את מה שסריקה אינה
 * יכולה לתפוס: `@click` שקיים ומצביע על `doCut(){}` ריקה.
 *
 * מה שנשאר כאן הוא בדיוק מה שרינדור **אינו** יכול לענות עליו:
 *   1. לכל פקד יש חיווט של `disabled` — כלומר מישהו החליט מתי הוא זמין. פקד
 *      שאין לו חיווט כזה נראה בהרכבה בדיוק כמו פקד זמין שהיכולת שלו קיימת.
 *   2. אין `shortcut` בלשוניות האלה, כל עוד אין קיצור רשום בפועל. „קיצור
 *      שמוצג ואינו רשום” היא טענה על המעטפת, לא על מה שנרנדר.
 *   3. שלושת פקדי „אוצריא” אמנם פולטים event — וזה נמדד בהרכבה — אבל שמישהו
 *      **עונה** לו ב-App.vue אינו נראה מהרכבה של הלשונית לבדה.
 */
import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

/** vitest רץ משורש המאגר, ולכן src/ נמצא ביחס ל-cwd. */
const TABS = join(process.cwd(), 'src/ui/ribbon/tabs');

const FILES = [
  'ShulchanTab.vue',
  'LayoutTab.vue',
  'ReferencesTab.vue',
  'ReviewTab.vue',
  'OtzariaTab.vue',
  'DeveloperTab.vue',
] as const;

const SOURCES = new Map(FILES.map((file) => [file, readFileSync(join(TABS, file), 'utf8')]));

/** כל תג פקד בקובץ, כולל התכונות שלו. כל הפקדים בלשוניות האלה נסגרים ב-`/>`. */
function controls(source: string): string[] {
  return [...source.matchAll(/<Ribbon(?:Button|MenuButton)\b[\s\S]*?\/>/g)].map((m) => m[0]);
}

/** שורת הזיהוי של פקד בכשל: התווית שלו, כדי שאפשר יהיה למצוא אותו. */
function labelOf(control: string): string {
  return control.match(/label="([^"]*)"/)?.[1] ?? control.slice(0, 60);
}

describe('פקדי הלשוניות פריסה, הפניות, סקירה, אוצריא ומפתחים', () => {
  it('נמצאו פקדים לבדוק בכל הלשוניות', () => {
    for (const file of FILES) {
      expect(controls(SOURCES.get(file)!).length, file).toBeGreaterThan(0);
    }
  });

  it('לכל פקד יש חיווט של disabled', () => {
    const unguarded: string[] = [];
    for (const file of FILES) {
      for (const control of controls(SOURCES.get(file)!)) {
        if (!/:disabled=/.test(control)) unguarded.push(`${file}: ${labelOf(control)}`);
      }
    }
    expect(unguarded).toEqual([]);
  });

  it('אין קיצור מקלדת מוצג, כי אין קיצור רשום', () => {
    // `Alt+Ctrl+F`, `Alt+Ctrl+D` ו-`F7` הוצגו כאן ולא היו קיימים.
    const fake: string[] = [];
    for (const file of FILES) {
      for (const control of controls(SOURCES.get(file)!)) {
        if (/shortcut=/.test(control)) fake.push(`${file}: ${labelOf(control)}`);
      }
    }
    expect(fake).toEqual([]);
  });

  it('שלושת פקדי „אוצריא” קוראים ל-SDK ולא מסתפקים בהודעת סטטוס', () => {
    // ההודעות שהיו כאן („פותח את ספריית אוצריא...”) תיארו פעולה שלא קרתה.
    const tab = SOURCES.get('OtzariaTab.vue')!;
    for (const event of ['insert-citation', 'search-otzaria', 'open-library']) {
      expect(tab, event).toContain(`$emit('${event}')`);
    }
    const app = readFileSync(join(process.cwd(), 'src/App.vue'), 'utf8');
    expect(app).toContain('insertCitation(');
    expect(app).toContain('openSearchTab(');
    expect(app).toContain('openLibrary()');
  });

  it('„חידוש”/„קושיא”/„תירוץ” אינם חוזרים כפקדים מנוטרלים', () => {
    // הם היו כאן שלושה כפתורים מנוטרלים עם הסבר, כי אין למנוע פעולה ציבורית
    // שיוצרת סגנון פסקה בשם (`doc.styles.apply` מקבל `scope:'docDefaults'`
    // בלבד). הם ירדו מהלשונית כפקד שאינו נצרך; הבקשה למנוע נשארה פתוחה
    // (superdoc/docx-editor#3975), ואם תתקבל — החזרה שלהם היא חיווט לפעולה
    // אמיתית, לא `:disabled="true"`.
    const tab = SOURCES.get('OtzariaTab.vue')!;
    for (const label of ['חידוש', 'קושיא', 'תירוץ']) {
      expect(controls(tab).map(labelOf), label).not.toContain(label);
    }
  });

  it('„סקירה” מציעה גם דחייה של כל השינויים, ולא רק קבלה', () => {
    // `rejectAllChanges` היה ב-registry מהיום הראשון, ולא היה לו פקד — כלומר
    // אפשר היה לקבל את כל השינויים ולא לדחות אותם.
    const review = SOURCES.get('ReviewTab.vue')!;
    expect(review).toContain("useCommand('rejectAllChanges')");
    expect(review).toContain('label="דחה את כל השינויים"');
  });

});

/**
 * „קובץ” נבדקת בנפרד ולא נוספה ל-FILES שמעל, ובכוונה: היא **כן** מציגה קיצורי
 * מקלדת (Ctrl+S, Ctrl+Shift+S, Ctrl+P), ושלושתם רשומים בפועל במטפל המקלדת של
 * App.vue — כלומר הכלל „אין קיצור מוצג” אינו חל עליה, בעוד הכלל על חיווט
 * ה-`disabled` דווקא כן.
 *
 * מה שנמדד כאן ואינו נראה בהרכבה: שלכל פקד **יש** תנאי, ושהיחיד שאין לו הוא
 * היחיד שאין לו תנאי אמיתי. המצבים עצמם — אין מסמך, שמירה שרצה, פתיחה שרצה —
 * נמדדים בהרכבה ב-tests/component/ribbon-tabs.test.ts.
 */
describe('פקדי לשונית „קובץ”', () => {
  const FILE_TAB = readFileSync(join(TABS, 'FileTab.vue'), 'utf8');

  /**
   * ארבעת הפקדים בלי `:disabled`, ובכוונה: „אודות”, „דיווח על באג”, „קיצורים” ו„ניהול
   * קיצורים” הם דיאלוגים של התוסף, אינם נוגעים במסמך ואינם נוגעים במנוע.
   * `:disabled="false"` קבוע היה נראה כמו תנאי ואינו תנאי, וזה גרוע מהיעדרו.
   *
   * „ניהול קיצורים” אינו יוצא מן הכלל חדש אלא אותו כלל: הדיאלוג עצמו אינו
   * דורש מסמך. **הצמדת** העיצוב כן דורשת אחד, וזה נבדק בזמן הלחיצה על
   * הצירוף — לא כאן.
   *
   * הסדר הוא סדר ה-DOM: המחסנית („קיצורים”, „ניהול קיצורים”) לפני „אודות”,
   * שיורד לקצה השמאלי. ראו את ההנמקה ב-FileTab.vue.
   */
  const UNCONDITIONAL = ['קיצורים', 'ניהול קיצורים', 'דיווח על באג', 'אודות'];

  it('נמצאו אחד-עשר פקדים', () => {
    // תשעה מאז ש„ייצוא ל-Word” הוסר — הורדה לתיקיית ההורדות שנראתה כמו
    // שמירה ולא הייתה שמירה (ראו FileTab.vue). „ייצוא לאוצריא” יושב בלשונית
    // „אוצריא” — הוא פעולה מול הספרייה, לא פעולת קובץ. העשירי הוא „ניהול
    // קיצורים”, והאחד-עשר „דיווח על באג”.
    expect(controls(FILE_TAB)).toHaveLength(11);
  });

  it('לכל פקד יש חיווט של disabled — חוץ מארבעת הדיאלוגים', () => {
    const unguarded = controls(FILE_TAB)
      .filter((control) => !/:disabled=/.test(control))
      .map((control) => labelOf(control));

    expect(unguarded).toEqual(UNCONDITIONAL);
  });

  it('התנאי הוא מצב המעטפת, ולא `true` קשיח', () => {
    // `:disabled="true"` כאן היה אומר „הפקד לא ממומש”, וכל האחד-עשר ממומשים.
    for (const control of controls(FILE_TAB)) {
      expect(control, labelOf(control)).not.toContain(':disabled="true"');
    }
    for (const state of ['hasDocument', 'hasPdfExport', 'isSaving', 'isOpening']) {
      expect(FILE_TAB, state).toContain(state);
    }
  });

  it('המעטפת מעבירה את שלושת המצבים דרך הרצועה', () => {
    // שלוש שכבות, וכל אחת מהן יכולה להישמט בלי שה-typecheck יתלונן: prop
    // שאינו נמסר פשוט נופל לברירת המחדל, כלומר פקד שמנוטרל לנצח.
    const ribbon = readFileSync(join(process.cwd(), 'src/ui/ribbon/Ribbon.vue'), 'utf8');
    const app = readFileSync(join(process.cwd(), 'src/App.vue'), 'utf8');

    for (const attribute of [':has-document', ':is-saving', ':is-opening']) {
      expect(ribbon, `Ribbon → FileTab: ${attribute}`).toContain(attribute);
      expect(app, `App → Ribbon: ${attribute}`).toContain(attribute);
    }
    expect(app).toContain('const hasDocument = computed(() => activeSuperdoc.value !== null)');
  });
});
