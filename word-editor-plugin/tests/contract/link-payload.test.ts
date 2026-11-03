/**
 * חוזה ה-payload של פקודת `link`, מול הפונקציות **האמיתיות** של superdoc@2.8.0.
 *
 * ## למה זה נבדק כך
 *
 * `linkCmd.run()` נקרא בלי payload בכלל, ו-`executeLinkCommand` מחזיר `false`
 * בלי לגעת במסמך כשאין `href`. כלומר הכפתור נכשל סגור בכל לחיצה, בלי שאף
 * בדיקה תיפול ובלי שהמשתמש יראה דבר. הצורה הנכונה אינה נגזרת מהטיפוסים: החוזה
 * הוא `readLinkPayloadHref` (שמכיר `href` **או** `value`), `readLinkPayloadText`
 * (שנצרך רק במסלול אחד מארבעה) ו-`readLinkPayloadTarget` (שנבדק לפני הבחירה
 * החיה). שלושתם פנימיים ל-chunk, ואף typecheck לא יגן עליהם.
 *
 * לכן הם **נחלצים מהמקור הארוז ומורצים כמו שהם**, בדיוק כמו ב-
 * tests/contract/command-payloads.test.ts. זה לא שחזור הלוגיקה — זה הקוד
 * שהחבילה שולחת.
 *
 * ## מה נבדק
 *
 * 1. `linkPayload` מייצר `href` שהמנוע מחלץ, ולא צורה שתיבלע.
 * 2. ה-`target` שנתפס מהבחירה עובר את `linkPayloadHasExplicitTarget` — זה מה
 *    שמאפשר לפקודה לרוץ אחרי שהדיאלוג גזל את המיקוד מהעורך.
 * 3. אותו `target` מיתרגם לכתובות טקסט שהמנוע יעטוף (`textAddressesFromTarget`)
 *    או לכתובת סמן שיכניס אליה (`collapsedTextAddressFromTarget`) — ולא
 *    לרשימה ריקה, שהיא הכשל הסגור.
 * 4. כתובת פסולה אינה מגיעה למנוע בכלל.
 *
 * מה שלא נבדק כאן: ה-mutation עצמו (`hyperlinks.wrap` / `insert` / `patch`)
 * והמצב במסמך. אלה דורשים מנוע חי ונבדקים באימות בדפדפן. מה שכן נבדק הוא
 * השלב שנכשל בשקט.
 */
import { describe, it, expect } from 'vitest';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { linkPayload, normalizeLinkHref } from '../../src/engine/payloads';

const CHUNKS_DIR = join(process.cwd(), 'node_modules/superdoc/dist/chunks');

/** שם הקובץ נושא hash שמשתנה בכל build של החבילה, ולכן הוא נמצא ולא נכתב. */
function readControllerChunk(): string {
  const file = readdirSync(CHUNKS_DIR).find((name) =>
    /^create-super-doc-ui-.*\.es\.js$/.test(name),
  );
  if (!file) throw new Error('לא נמצא ה-chunk של controller ה-UI ב-superdoc');
  return readFileSync(join(CHUNKS_DIR, file), 'utf8');
}

const CHUNK = readControllerChunk();

/**
 * פונקציה ברמת המודול. `}` בתחילת שורה הוא סוף הפונקציה — מה שהופך את החילוץ
 * לחד-משמעי בלי מנתח JS.
 */
function extractModuleFunction(name: string): string {
  const match = CHUNK.match(new RegExp(String.raw`^function ${name}\([\s\S]*?\n\}`, 'm'));
  if (!match) {
    throw new Error(`הפונקציה ${name} לא נמצאה ב-chunk — חוזה הקישור של superdoc השתנה`);
  }
  return match[0];
}

/**
 * פונקציה בתוך ה-factory של ה-controller: מוזחת ב-tab אחד, ולכן `\t}` הוא
 * סופה. ההזחה מוסרת כדי שהגוף יוכל לרוץ ברמת המודול.
 */
function extractControllerFunction(name: string): string {
  const match = CHUNK.match(new RegExp(String.raw`^\tfunction ${name}\([\s\S]*?\n\t\}`, 'm'));
  if (!match) {
    throw new Error(`הפונקציה ${name} לא נמצאה ב-chunk — חוזה הקישור של superdoc השתנה`);
  }
  return match[0].replace(/^\t/gm, '');
}

/** `linkPayloadHasExplicitTarget` הוא closure בתוך ה-controller ולא פונקציה. */
function extractExplicitTargetPredicate(): string {
  const match = CHUNK.match(/^\tconst linkPayloadHasExplicitTarget = \([\s\S]*?\n\t\};/m);
  if (!match) {
    throw new Error('linkPayloadHasExplicitTarget לא נמצא ב-chunk — חוזה הקישור של superdoc השתנה');
  }
  return match[0].replace(/^\t/gm, '').replace(/^const /, 'var ');
}

type TextAddress = { kind: 'text'; blockId: string; range: { start: number; end: number } };

interface LinkValidators {
  readLinkPayloadHref(payload: unknown): unknown;
  readLinkPayloadText(payload: unknown, href: string): string;
  readLinkPayloadTarget(payload: unknown): unknown;
  linkPayloadHasExplicitTarget(payload: unknown): boolean;
  textAddressesFromTarget(target: unknown): TextAddress[];
  collapsedTextAddressFromTarget(target: unknown): TextAddress | null;
}

const CONTROLLER_FUNCTIONS = [
  'readLinkPayloadRecord',
  'readLinkPayloadHref',
  'readLinkPayloadText',
  'readLinkPayloadTarget',
] as const;

const MODULE_FUNCTIONS = [
  'isLooseObject',
  'textAddressesFromTarget',
  'collapsedTextAddressFromTarget',
] as const;

const EXPORTED = [
  'readLinkPayloadHref',
  'readLinkPayloadText',
  'readLinkPayloadTarget',
  'linkPayloadHasExplicitTarget',
  'textAddressesFromTarget',
  'collapsedTextAddressFromTarget',
] as const;

const engine: LinkValidators = new Function(
  [
    ...MODULE_FUNCTIONS.map(extractModuleFunction),
    ...CONTROLLER_FUNCTIONS.map(extractControllerFunction),
    extractExplicitTargetPredicate(),
    `return { ${EXPORTED.join(', ')} };`,
  ].join('\n'),
)() as LinkValidators;

/* ------------------------------------------------------------------ */
/* יעדים בצורה שהמנוע מחזיר מ-`doc.selection.current()`                */
/* ------------------------------------------------------------------ */

/** בחירה עם טווח — המסלול של `hyperlinks.wrap`. */
const RANGE_TARGET = {
  kind: 'text',
  segments: [{ blockId: 'p1', range: { start: 3, end: 9 } }],
};

/** סמן בלבד — המסלול של `hyperlinks.insert`. */
const CARET_TARGET = {
  kind: 'text',
  segments: [{ blockId: 'p1', range: { start: 4, end: 4 } }],
};

describe('href — מה שהמנוע מחלץ מה-payload', () => {
  it('המפתח `href` הוא זה שנקרא, והכתובת מגיעה שלמה', () => {
    const payload = linkPayload({ href: 'https://otzaria.org/a' });

    expect(engine.readLinkPayloadHref(payload)).toBe('https://otzaria.org/a');
  });

  it('mailto עובר כמו שהוא', () => {
    const payload = linkPayload({ href: 'mailto:info@otzaria.org' });

    expect(engine.readLinkPayloadHref(payload)).toBe('mailto:info@otzaria.org');
  });

  it('בלי payload — זה המצב שהיה בפקד, והוא כשל סגור', () => {
    // `executeLinkCommand` ממשיך לענף האחרון ומחזיר false בלי לגעת במסמך.
    expect(engine.readLinkPayloadHref(undefined)).toBeUndefined();
  });

  it('המפתח האינטואיטיבי `url` נבלע בשקט — ולכן `href` אינו בחירת סגנון', () => {
    // בדיוק סוג הכשל שהפיל את `{ fontFamily }` ו-`{ zoom }`: payload שנראה
    // סביר לחלוטין בקוד, שהמנוע אינו מחלץ ממנו כלום, והפקודה מחזירה false.
    expect(engine.readLinkPayloadHref({ url: 'https://otzaria.org' })).toBeUndefined();
    expect(engine.readLinkPayloadHref({ href: 'https://otzaria.org' })).toBe('https://otzaria.org');
  });

  it('כתובת פסולה אינה מגיעה למנוע בכלל', () => {
    for (const href of ['javascript:alert(1)', 'www.otzaria.org', '', '   ']) {
      expect(linkPayload({ href }), href).toBeNull();
    }
  });
});

describe('text — נצרך רק כשאין טווח מסומן', () => {
  it('הטקסט שהמשתמש הזין הוא זה שיוכנס', () => {
    const payload = linkPayload({ href: 'https://otzaria.org', text: 'אוצריא' });

    expect(engine.readLinkPayloadText(payload, 'https://otzaria.org/')).toBe('אוצריא');
  });

  it('בלי טקסט המנוע נופל לכתובת עצמה, ולכן ההכנסה אינה נכשלת', () => {
    // `executeLinkCommand` דוחה `text.trim() === ''` במסלול ההכנסה. הנפילה
    // לכתובת היא מה שמאפשר „הוסף קישור” בלי למלא טקסט להצגה.
    const payload = linkPayload({ href: 'https://otzaria.org' });

    expect(payload?.text).toBeUndefined();
    expect(engine.readLinkPayloadText(payload, 'https://otzaria.org/')).toBe('https://otzaria.org/');
  });
});

describe('target — מה שמאפשר לפקודה לרוץ אחרי שהדיאלוג גזל את המיקוד', () => {
  it('היעד שנתפס נקרא ראשון, לפני הבחירה החיה', () => {
    const payload = linkPayload({ href: 'https://otzaria.org', target: RANGE_TARGET });

    expect(engine.readLinkPayloadTarget(payload)).toBe(RANGE_TARGET);
    expect(engine.linkPayloadHasExplicitTarget(payload)).toBe(true);
  });

  it('בלי יעד הפקודה תלויה בבחירה החיה — וזה מה שנשבר עם דיאלוג', () => {
    const payload = linkPayload({ href: 'https://otzaria.org' });

    expect(engine.linkPayloadHasExplicitTarget(payload)).toBe(false);
  });

  it('טווח מסומן מיתרגם לכתובת שהמנוע יעטוף', () => {
    const payload = linkPayload({ href: 'https://otzaria.org', target: RANGE_TARGET });
    const addresses = engine.textAddressesFromTarget(engine.readLinkPayloadTarget(payload));

    // רשימה ריקה כאן פירושה שהמנוע נופל למסלול ההכנסה במקום לעטיפה, כלומר
    // הטקסט המסומן לא היה מקבל את הקישור.
    expect(addresses).toEqual([{ kind: 'text', blockId: 'p1', range: { start: 3, end: 9 } }]);
  });

  it('סמן בלבד מיתרגם לכתובת הכנסה, ולא לעטיפה', () => {
    const payload = linkPayload({ href: 'https://otzaria.org', text: 'אוצריא', target: CARET_TARGET });
    const target = engine.readLinkPayloadTarget(payload);

    expect(engine.textAddressesFromTarget(target)).toEqual([]);
    expect(engine.collapsedTextAddressFromTarget(target)).toEqual({
      kind: 'text',
      blockId: 'p1',
      range: { start: 4, end: 4 },
    });
  });

  it('יעד שנשמר בכותרת עליונה שומר על ה-story', () => {
    // בלי ה-story הכתובת מתפרשת כגוף המסמך, והקישור נכתב במקום הלא נכון.
    const story = { kind: 'story', storyType: 'headerFooterSlot' };
    const payload = linkPayload({
      href: 'https://otzaria.org',
      target: { ...RANGE_TARGET, story },
    });

    expect(engine.textAddressesFromTarget(engine.readLinkPayloadTarget(payload))).toEqual([
      { kind: 'text', blockId: 'p1', range: { start: 3, end: 9 }, story },
    ]);
  });
});

describe('normalizeLinkHref מול הפרסור של המנוע', () => {
  it('הצורה הקנונית היא זו שנשלחת, ואין בה תווי בקרה', () => {
    // `new URL` משמיט תווי בקרה בפרסור, ולכן בדיקת תחילית על המחרוזת הגולמית
    // הייתה מפספסת `java\nscript:`.
    expect(normalizeLinkHref('java\nscript:alert(1)')).toBeNull();
    expect(normalizeLinkHref('  https://otzaria.org  ')).toBe('https://otzaria.org/');
  });

  it('כל כתובת שאושרה נקראת בחזרה זהה על ידי המנוע', () => {
    for (const raw of [
      'https://otzaria.org',
      'https://otzaria.org/ספר?q=1#הערה',
      'http://127.0.0.1:8080/a',
      'mailto:info@otzaria.org',
    ]) {
      const payload = linkPayload({ href: raw });
      expect(payload, raw).not.toBeNull();
      expect(engine.readLinkPayloadHref(payload), raw).toBe(payload!.href);
    }
  });
});
