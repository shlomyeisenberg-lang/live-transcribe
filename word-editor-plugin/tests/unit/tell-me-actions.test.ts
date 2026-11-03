import { describe, it, expect } from 'vitest';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  TELL_ME_ACTIONS,
  DEFAULT_SUGGESTED_IDS,
  searchTellMeActions,
  normalizeSearchTerm,
  type TellMeAction,
} from '../../src/ui/shell/tell-me-actions';
import { ICONS } from '../../src/ui/icons/icons';
import { SHORTCUTS } from '../../src/ui/shortcuts/registry';
import { RIBBON_TABS } from '../../src/ui/ribbon/tabs';

describe('קטלוג הפקודות Tell Me (tell-me-actions)', () => {
  it('יש פקודות בקטלוג', () => {
    expect(TELL_ME_ACTIONS.length).toBeGreaterThan(30);
  });

  it('כל המזהים ייחודיים', () => {
    const ids = TELL_ME_ACTIONS.map((a) => a.id);
    const uniqueIds = new Set(ids);
    expect(ids.length).toBe(uniqueIds.size);
  });

  it('כל פקודה כוללת כותרת, קטגוריה ואייקון חוקי שקיים ב-ICONS', () => {
    for (const action of TELL_ME_ACTIONS) {
      expect(action.title.trim().length, `action ${action.id} has empty title`).toBeGreaterThan(0);
      expect(action.category.trim().length, `action ${action.id} has empty category`).toBeGreaterThan(0);
      expect(ICONS[action.icon], `action ${action.id} references missing icon '${action.icon}'`).toBeDefined();
    }
  });

  it('לכל פקודה יש יעד תקף (command / shellAction / customAction / ribbonTab)', () => {
    for (const action of TELL_ME_ACTIONS) {
      const hasTarget = Boolean(
        action.command || action.shellAction || action.customAction || action.ribbonTab,
      );
      expect(hasTarget, `action ${action.id} has no execution target`).toBe(true);
    }
  });

  /**
   * בדיקת חוזה: כל `command.id` בקטלוג הוא פקודה שהמנוע מכיר.
   *
   * למה זה נדרש: `command.id` הוא `string` חופשי, ו-`commandAdapter.run` מיירט
   * מזהה לא מוכר ומחזיר „הפעולה אינה מוכרת למנוע” — הפריט נראה תקין בתפריט
   * ולחיצה עליו נכשלת תמיד. 24 מ-38 הפריטים נכשלו כך לפני הבדיקה הזאת
   * (`align` במקום `text-align`, `style` במקום `linked-style`, `paste` שאינו
   * פקודת מנוע כלל). `shellAction`, לעומת זאת, מוקלד ונופל בבנייה.
   *
   * מקור האמת: רג׳יסטרי הפקודות של SuperDoc 2.11.0 (`create-super-doc-ui`),
   * 50 מזהים. הרשימה מועתקת לכאן ולא נקראת מהמנוע: הבדיקה רצה בלי מנוע,
   * ושדרוג שמוסיף/מסיר פקודה צריך להיראות כשינוי מפורש בקובץ הזה.
   */
  const ENGINE_COMMAND_IDS = new Set([
    'acceptAllChanges', 'acceptChange', 'bold', 'bullet-list', 'clear-formatting', 'copy-format',
    'direction-ltr', 'direction-rtl', 'document-mode', 'font-family', 'font-size', 'formatting-marks',
    'highlight-color', 'image', 'indent-decrease', 'indent-increase', 'italic', 'line-height', 'link',
    'linked-style', 'measurement-unit', 'numbered-list', 'redo', 'rejectAllChanges', 'rejectChange',
    'ruler', 'setFontFamily', 'setFontSize', 'strikethrough', 'table-add-column-after',
    'table-add-column-before', 'table-add-row-after', 'table-add-row-before', 'table-delete',
    'table-delete-column', 'table-delete-row', 'table-fix', 'table-insert',
    'table-of-contents-insert', 'table-remove-borders', 'table-split-cell', 'text-align',
    'text-color', 'track-changes-accept-selection', 'track-changes-reject-selection',
    'underline', 'undo', 'zoom', 'zoom-fit-width',
  ]);

  it('כל command.id בקטלוג הוא פקודת מנוע קיימת', () => {
    const unknown = TELL_ME_ACTIONS.filter((a) => a.command && !ENGINE_COMMAND_IDS.has(a.command.id)).map(
      (a) => `${a.id} → ${a.command!.id}`,
    );
    expect(unknown).toEqual([]);
  });

  it('רג׳יסטרי הקיצורים מסכים עם רשימת המנוע — כדי שהרשימה כאן לא תתיישן בשקט', () => {
    const commands = (SHORTCUTS as readonly { id: string; command?: string }[]).filter((s) => s.command);
    const unknown = commands.filter((s) => !ENGINE_COMMAND_IDS.has(s.command!)).map((s) => `${s.id} → ${s.command}`);
    expect(unknown).toEqual([]);
  });

  /**
   * ה-payload של `linked-style` הוא `w:styleId` של OOXML — CamelCase בלי מקף
   * (`Heading1`, לא `heading-1`). מזהה לא מוכר נכשל *סגור*: המנוע מדווח הצלחה
   * ואינו משנה כלום, ולכן זה נבדק כאן ולא נראה בריצה.
   */
  it('סגנונות ב-payload הם מזהי OOXML מהרג׳יסטרי', () => {
    const known = new Set(
      (SHORTCUTS as readonly { command?: string; payload?: unknown }[])
        .filter((s) => s.command === 'linked-style')
        .map((s) => (s.payload as { style: string }).style),
    );
    for (const action of TELL_ME_ACTIONS) {
      if (action.command?.id !== 'linked-style') continue;
      const style = (action.command.payload as { style?: string } | null)?.style;
      expect(style && known.has(style), `action ${action.id}: style '${style}' אינו ברג׳יסטרי`).toBe(true);
    }
  });

  it('כל ribbonTab בקטלוג הוא לשונית קיימת ברצועה', () => {
    const tabs = new Set(RIBBON_TABS.map((t) => t.id));
    const unknown = TELL_ME_ACTIONS.filter((a) => a.ribbonTab && !tabs.has(a.ribbonTab)).map(
      (a) => `${a.id} → ${a.ribbonTab}`,
    );
    expect(unknown).toEqual([]);
  });

  it('פקדים שהוסרו מהרצועה אינם נשארים בתוצאות החיפוש', () => {
    // „סגנון תורני” ירד מלשונית „אוצריא”: להשאיר אותו כאן היה מוליך
    // משתמש ללשונית שבה אין עוד פקד כזה.
    const ids = searchTellMeActions('סגנון תורני').map((action) => action.id);
    expect(ids).not.toContain('otzaria-torah-styles');
  });

  it('מזהי הפעולות המוצעות קיימים כולם בקטלוג', () => {
    const actionIds = new Set(TELL_ME_ACTIONS.map((a) => a.id));
    for (const id of DEFAULT_SUGGESTED_IDS) {
      expect(actionIds.has(id), `suggested action id '${id}' not in TELL_ME_ACTIONS`).toBe(true);
    }
  });
});

describe('אלגוריתם חיפוש Tell Me (searchTellMeActions)', () => {
  it('שאילתה ריקה מחזירה את הפעולות המוצעות כברירת מחדל', () => {
    const results = searchTellMeActions('');
    expect(results.length).toBe(DEFAULT_SUGGESTED_IDS.length);
    expect(results.map((r) => r.id)).toEqual(DEFAULT_SUGGESTED_IDS);
  });

  it('חיפוש "הדפסה" מוצא את פעולת ההדפסה ראשונה', () => {
    const results = searchTellMeActions('הדפסה');
    expect(results.length).toBeGreaterThan(0);
    expect(results[0].id).toBe('file-print');
  });

  it('חיפוש "טבלה" מוצא את פעולת הוספת טבלה ראשונה', () => {
    const results = searchTellMeActions('טבלה');
    expect(results.length).toBeGreaterThan(0);
    expect(results[0].id).toBe('insert-table');
  });

  it('חיפוש "מרכז" מוצא את יישור למרכז', () => {
    const results = searchTellMeActions('מרכז');
    expect(results.length).toBeGreaterThan(0);
    expect(results[0].id).toBe('para-align-center');
  });

  it('חיפוש מונח באנגלית (כמו bold) מוצא מודגש דרך מילות מפתח', () => {
    const results = searchTellMeActions('bold');
    expect(results.length).toBeGreaterThan(0);
    expect(results[0].id).toBe('font-bold');
  });

  it('חיפוש "שמור" מוצא גם שמירה וגם שמירה בשם', () => {
    const results = searchTellMeActions('שמור');
    const ids = results.map((r) => r.id);
    expect(ids).toContain('file-save');
    expect(ids).toContain('file-save-as');
  });

  it('נרמול טקסט עברי מסיר ניקוד ומסיר רווחים', () => {
    // שָׁמוֹר -> שמור
    const withVowels = '\u05E9\u05B8\u05C1\u05DE\u05D5\u05B9\u05E8';
    expect(normalizeSearchTerm(withVowels)).toBe('שמור');
    expect(normalizeSearchTerm('  Hello World  ')).toBe('hello world');
  });

  it('חיפוש עם ניקוד עברי מוצא את הפקודה המתאימה', () => {
    const withVowels = '\u05D4\u05B7\u05D3\u05B0\u05E4\u05B8\u05BC\u05E1\u05B8\u05D4'; // הַדְפָּסָה
    const results = searchTellMeActions(withVowels);
    expect(results.length).toBeGreaterThan(0);
    expect(results[0].id).toBe('file-print');
  });

  /** הבאג שדווח: „השלמה מהספר” החזירה „לא נמצאו פקודות מתאימות”. */
  it('חיפוש "השלמה מהספר" מוצא את הפקד של אוצריא ראשון', () => {
    const results = searchTellMeActions('השלמה מהספר');
    expect(results[0]?.id).toBe('otzaria-book-completion');
  });

  /** צירוף חלקי של מילים מתוך השם — מה שמשתמש באמת מקליד. */
  it('שתי מילים לא רצופות מתוך שם הפקד מוצאות אותו', () => {
    expect(searchTellMeActions('השלמה ספר').map((r) => r.id)).toContain('otzaria-book-completion');
    expect(searchTellMeActions('גבולות עמוד').map((r) => r.id)).toContain('layout-page-borders');
  });

  /** אות סופית אינה אמורה לשבור התאמת תחילית: „עמוד” מול „עמודים”. */
  it('נרמול משווה אותיות סופיות לרגילות', () => {
    expect(normalizeSearchTerm('מסמך')).toBe('מסמכ');
    expect(searchTellMeActions('מספור עמודים').map((r) => r.id)).toContain('layout-page-numbering');
  });

  it('שאילתה שאינה קיימת מחזירה רשימה ריקה', () => {
    expect(searchTellMeActions('xyznonexistentquery123')).toEqual([]);
    // מילה נפוצה אחת מתוך שלוש אינה התאמה: „לא” הוא תחילית של מילה במילות
    // המפתח של „סוגריים לא סגורים”, והשאילתה הזאת החזירה בגללה תוצאות.
    expect(searchTellMeActions('זזזזז לא קיים')).toEqual([]);
  });

  /** האינדקס המנורמל נשמר לכל קטלוג בנפרד — קטלוג משלו אינו מקבל את של ברירת המחדל. */
  it('קטלוג שנמסר במפורש נחפש לעצמו ולא מהאינדקס של ברירת המחדל', () => {
    const own: TellMeAction[] = [
      { id: 'own-1', title: 'צבע גופן', category: 'בדיקה', keywords: ['צבע'], icon: 'bold' },
    ];
    expect(searchTellMeActions('צבע', own).map((a) => a.id)).toEqual(['own-1']);
    expect(searchTellMeActions('הדפסה', own)).toEqual([]);
    // ברירת המחדל לא נפגעה
    expect(searchTellMeActions('הדפסה')[0]?.id).toBe('file-print');
  });
});

/**
 * שער הכיסוי: כל פקד ברצועה נמצא בחיפוש.
 *
 * הבאג שהוליד את הבדיקה: „השלמה מהספר” — כפתור גדול בלשונית „אוצריא” — פשוט
 * לא היה בקטלוג, ולכן חיפוש שמו החזיר „לא נמצאו פקודות מתאימות”. הוא לא היה
 * לבד: 55 מ-96 תוויות הרצועה לא היו ניתנות למציאה. הקטלוג מתוחזק ביד, ולכן
 * פקד שנוסף לרצועה אינו מגיע לחיפוש מעצמו — וזה בדיוק מה שנספר כאן.
 *
 * הבדיקה קוראת את קובצי הלשוניות ולא רשימה מועתקת: רשימה מועתקת הייתה מזדקנת
 * בדיוק כמו הקטלוג שהיא באה לשמור.
 */
describe('כיסוי הרצועה בחיפוש Tell Me', () => {
  const TABS_DIR = join(process.cwd(), 'src/ui/ribbon/tabs');
  /** `label="..."` בלבד — `:label="expr"` הוא ביטוי, לא תווית לחיפוש. */
  const STATIC_LABEL = /(?<![:\w-])label="([^"]+)"/g;

  const RIBBON_LABELS = readdirSync(TABS_DIR)
    .filter((file) => file.endsWith('.vue'))
    .flatMap((file) =>
      [...readFileSync(join(TABS_DIR, file), 'utf8').matchAll(STATIC_LABEL)].map((m) => ({
        file,
        label: m[1],
      })),
    );

  it('נאספו תוויות מכל לשוניות הרצועה', () => {
    expect(RIBBON_LABELS.length).toBeGreaterThan(80);
  });

  it('כל תווית של פקד ברצועה מוחזרת בחיפוש שלה', () => {
    const missing: string[] = [];

    for (const { file, label } of RIBBON_LABELS) {
      // „שמור בשם...” — שלוש הנקודות הן מוסכמה של „נפתח דיאלוג”, לא חלק מהשם.
      const query = label.replace(/[.…]+$/u, '').trim();
      const normalized = normalizeSearchTerm(query);
      // „מוחזרת” אינו „מדורגת ראשונה”: „הסר” מחזיר בצדק גם „הסר קישור”. מה
      // שנדרש הוא שהפריט של הפקד עצמו יהיה שם — לפי כותרת או מילת מפתח זהה.
      const found = searchTellMeActions(query).some(
        (action) =>
          normalizeSearchTerm(action.title) === normalized ||
          action.keywords.some((keyword) => normalizeSearchTerm(keyword) === normalized),
      );
      if (!found) missing.push(`${file}: ${label}`);
    }

    expect(missing).toEqual([]);
  });
});
