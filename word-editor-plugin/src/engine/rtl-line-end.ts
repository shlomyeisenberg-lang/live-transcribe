/**
 * מקש `End` בשורה עברית — הבאג שדווח: „בקובץ הוורד שהמסמך נשמר הנקודה עוברת
 * לתחילת המסמך”.
 *
 * ## מה נמדד
 *
 * ב-superdoc 2.12.0, Chrome אמיתי, על ה-dist הארוז: בפסקה שיש בה `<w:bidi/>`
 * מקש `End` מעביר את הסמן ל**תחילת** השורה במקום לסופה. שתי פסקאות באותו
 * מסמך, סמן שהונח באמצע בלחיצה אמיתית:
 *
 *   | הפסקה                | אמצע | `Home` | `End`    | מה שנכון |
 *   |----------------------|------|--------|----------|----------|
 *   | עברית, שורה גולשת 1  |  48  |   0    | **0**    |    98    |
 *   | עברית, שורה גולשת 2  | 144  |  98    | **98**   |   191    |
 *   | אנגלית, שורה 1       |  45  |   0    |   92     |    92    |
 *   | אנגלית, שורה 2       | 141  |  92    |  157     |   157    |
 *
 * כלומר בעברית `End` הוא בדיוק `Home`, ובאנגלית הוא תקין. `Shift+End` נגרר
 * אחריו ובוחר אחורה (0..48 במקום 48..98); `Ctrl+Home` ו-`Ctrl+End` נמדדו
 * **תקינים** ואינם נוגעים כאן, וכך גם `Home` עצמו.
 *
 * ## למה זה נראה למשתמש כמו „הנקודה קפצה לתחילת המסמך”
 *
 * מי שסיים להקליד משפט, לחץ `End` כדי לחזור לסוף השורה והקליד נקודה — קיבל
 * את הנקודה בהיסט 0, כלומר כתו הראשון בפסקה. נמדד: `<w:p><w:pPr><w:bidi/>
 * </w:pPr><w:r><w:t>.</w:t></w:r><w:r><w:t>ברוך הבא לעולם</w:t></w:r></w:p>`.
 * זה אינו עניין של ציור אלא של סדר לוגי — Word מצייר את הפסקה הזאת בדיוק
 * כמו שהעורך מצייר אותה, ובפסקה הראשונה „תחילת הפסקה” היא תחילת המסמך.
 *
 * ## למה התיקון כאן ולא במנוע
 *
 * המשטח שמקבל את המקש הוא של המנוע (אין במסמך אף `contenteditable`; ההקלדה
 * נכנסת ל-`textarea` נסתרת שלו), והקוד שמזיז את הסמן אינו בחלקים הפתוחים של
 * SuperDoc. לכן אין כאן PR אלא יירוט: המאזין יושב על ה-div **שלנו**
 * (`paintedHost` מחזיר את `editor-stack__host`), ורץ לפני שהאירוע מגיע
 * למנוע.
 *
 * דווח למעלה: superdoc/docx-editor#3986. נמדד שוב על superdoc 2.13.0 (מנוע
 * 0.12.0) על עץ נקי — אותה התנהגות. כשה-issue ייסגר, למחוק את המודול הזה ואת
 * השער שלו במקום להשאיר עקיפה מיותרת.
 *
 * ## מה היירוט קורא, ולמה דווקא כך
 *
 * שני דברים חייבים להיות **סינכרוניים**, מפני שההחלטה אם לבטל את האירוע
 * חייבת ליפול בתוך המאזין עצמו:
 *
 * 1. **הסמן** — `activeEditor.host.readLiveSelectionSyncSnapshot()`. נמדד
 *    שהוא מחזיר ערך ולא הבטחה, ובתוכו `selectionTarget.start/end` עם
 *    `blockId` ו-`offset`. `doc.selection.current()`, לעומתו, מחזיר הבטחה —
 *    ואילו חיכינו לה, המנוע כבר היה מזיז את הסמן.
 * 2. **השורה המצוירת** — הפסקה מצוירת כ-fragment שנושא `data-source-node-id`
 *    ו-`data-pm-start`, וכל שורה בתוכה היא **ילד ישיר** שנושא `data-pm-start`
 *    /`data-pm-end`. ההמרה למספר שהמנוע מדבר בו היא חיסור אחד:
 *    `offset = pm − fragmentPmStart`. זה לא נוחש אלא אומת מול המנוע עצמו —
 *    באנגלית, שבה `End` תקין, הוא נותן בדיוק את המספרים שהמנוע נותן (92 ו-
 *    157 בטבלה שלמעלה).
 *
 * `dir="rtl"` על אלמנט השורה הוא ההצהרה של המנוע עצמו על כיוון השורה, והוא
 * נכתב **רק** על שורות RTL (נמדד: במסמך בלי `w:bidi` התכונה אינה קיימת
 * והכיוון המחושב הוא `ltr`). קריאת תכונה ולא `getComputedStyle` היא גם מה
 * שמונע חישוב סגנון מחדש בכל הקשה — ראו „הקלדה שאינה מחשבת סגנון” בהיסטוריה.
 *
 * ## מה קורה כשמשהו מזה חסר
 *
 * שום דבר. אין fragment, אין תכונות, אין `authoring` — המאזין חוזר בלי לבטל
 * את האירוע, והמנוע עושה את מה שהוא עושה היום. שדרוג מנוע שישנה את הצורה
 * הזאת יחזיר את הבאג ולא ישבור הקלדה, ושער ה-QA
 * (`scripts/qa/rtl-line-end-qa.mjs`) הוא זה שיתפוס אותו.
 *
 * ## גבול גלישה: היסט אחד לשתי שורות
 *
 * בשורה שגולשת, ההיסט שבתפר הוא גם סופה של השורה הראשונה וגם תחילתה של
 * השנייה, והתצלום הסינכרוני מחזיר היסט, לא שורה. כברירת מחדל התפר נשאר שייך
 * לשורה הקודמת, ולכן `End` חוזר יציב ואינו מטייל שורה-שורה. החריג המדויק הוא
 * `Home` ואחריו `End`: מאזין ה-Home יודע שהמנוע העביר את הסמן לתחילת השורה
 * השנייה, ולכן ההקשה הבאה בוחרת את השורה הבאה. כל הקשה אחרת או לחיצת עכבר
 * מאפסות את הרמז הזה.
 *
 * ומה ש**כן** נמדד ואינו סתירה לזה: הקלדה בתפר מעבירה את הסמן לשורה הבאה
 * בפועל — התו נכנס לתיבת השורה השנייה והסמן מצויר שם (נמדד: פסקה שגלשה
 * 0..98 / 98..146, נקודה שהוקלדה בהיסט 98 נכנסה לשורה השנייה, שהפכה
 * ל-98..147) — ולכן `End` שאחריה מגיע לסוף השורה השנייה, וזה הדבר הנכון.
 */

/** שורה מצוירת אחת: טווח ה-pm שלה, והאם המנוע הכריז עליה מימין לשמאל. */
export interface PaintedLine {
  pmStart: number;
  pmEnd: number;
  rtl: boolean;
}

/** נקודת קצה של בחירה, כפי שהמנוע מחזיר אותה. */
interface TextPoint {
  kind?: string;
  blockId?: string;
  offset?: number;
  story?: unknown;
}

/** התצלום הסינכרוני של הבחירה. */
interface LiveSelectionSnapshot {
  selectionTarget?: {
    kind?: string;
    start?: TextPoint | null;
    end?: TextPoint | null;
    story?: unknown;
    coordinateSpace?: string;
  } | null;
}

/** מה שנצרך מ-`activeEditor`. מוגדר כאן ולא מיובא — ההסבר ב-document-api.ts. */
export interface LineEndEditor {
  host?: {
    readLiveSelectionSyncSnapshot?: () => LiveSelectionSnapshot | null | undefined;
  } | null;
  authoring?: {
    setSelectionTarget?: (input: {
      target: {
        kind: 'selection';
        start: TextPoint;
        end: TextPoint;
        story: unknown;
        coordinateSpace?: string;
      };
      collapse?: 'start' | 'end';
      focus?: boolean;
    }) => unknown;
  } | null;
}

export interface LineEndHost {
  activeEditor?: LineEndEditor | null;
}

export interface RtlLineEndHandle {
  dispose(): void;
}

export interface RtlLineEndOptions {
  /** ה-div שהמנוע מצייר בתוכו — `paintedHost(editor.ui)`. */
  host: HTMLElement | null;
  /** ה-SuperDoc של אותו מסמך. */
  superdoc: LineEndHost | null | undefined;
}

/** מה שנקרא מהאירוע. `isComposing`/`keyCode` חסרים בדמויות של הבדיקות. */
type LineEndKeyEvent = Pick<KeyboardEvent, 'key' | 'ctrlKey' | 'metaKey' | 'altKey'> &
  Partial<Pick<KeyboardEvent, 'isComposing' | 'keyCode'>>;

/**
 * הרכבה במקלדת (IME). התקן של הריפו ב-`src/ui/shortcuts/match.ts`: דפדפן
 * שאינו מציב `isComposing` מדווח `keyCode === 229`.
 *
 * ‏`Home` ו-`End` הם מקשי חלונית הרכבה בדיוק כמו החצים — הם מזיזים בתוך
 * המועמד ולא בתוך המסמך. כאן זה חוסם, ולא רק לא-רצוי: המסלול שלמטה בולע
 * את האירוע ב-`stopImmediatePropagation`, כלומר חלונית ההרכבה לא הייתה
 * רואה אותו כלל. אותו שיקול בדיוק ב-`isHorizontalArrow` שב-`rtl-caret.ts`,
 * ואותו מתקין מרכיב את שני המודולים זה לצד זה.
 */
function isComposingEvent(event: LineEndKeyEvent): boolean {
  return event.isComposing === true || event.keyCode === 229;
}

/** האם ההקשה היא `End` „נקי” — עם Shift או בלעדיו, ובלי Ctrl/Meta/Alt. */
export function isLineEndKey(event: LineEndKeyEvent): boolean {
  if (isComposingEvent(event)) return false;
  return event.key === 'End' && !event.ctrlKey && !event.metaKey && !event.altKey;
}

/** `Home` רגיל הוא הרמז היחיד שמסיר את העמימות של תפר שורות. */
function isPlainHomeKey(event: LineEndKeyEvent): boolean {
  if (isComposingEvent(event)) return false;
  return event.key === 'Home' && !event.ctrlKey && !event.metaKey && !event.altKey;
}

/**
 * ההיסט שהסמן צריך להגיע אליו, או `null` כשהשורה אינה עברית (ואז המנוע
 * מטפל בעצמו, ונמדד שהוא עושה זאת נכון).
 *
 * היסט שיושב בתפר בין שתי שורות נקרא כסופה של הקודמת, אלא אם הוא הגיע מיד
 * אחרי `Home` — ראו „גבול גלישה” בהערת הפתיחה. כך `End` חוזר אינו מטייל,
 * ואילו `Home` ואז `End` מגיעים לסוף השורה שה-Home בחר.
 */
export function lineEndOffset(
  lines: readonly PaintedLine[],
  fragmentPmStart: number,
  caretOffset: number,
  preferNextAtBoundary = false,
): number | null {
  const caretPm = fragmentPmStart + caretOffset;
  let chosen: PaintedLine | null = null;

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index]!;
    if (caretPm < line.pmStart) break;
    if (caretPm <= line.pmEnd) {
      const next = lines[index + 1];
      if (preferNextAtBoundary && caretPm === line.pmEnd && next?.pmStart === caretPm) continue;
      chosen = line;
      break;
    }
  }

  if (!chosen || !chosen.rtl) return null;
  return chosen.pmEnd - fragmentPmStart;
}

/** ה-fragment של הפסקה שהסמן בה, ואיתו טווח ה-pm שלו. */
function readFragment(
  host: HTMLElement,
  blockId: string,
  caretOffset: number,
): { base: number; lines: PaintedLine[] } | null {
  const fragments = host.querySelectorAll<HTMLElement>('[data-source-node-id][data-pm-start]');

  for (const fragment of Array.from(fragments)) {
    if (fragment.getAttribute('data-source-node-id') !== blockId) continue;
    const base = Number(fragment.getAttribute('data-pm-start'));
    if (!Number.isFinite(base)) continue;
    // פסקה שנחצית בין עמודים מצוירת כשני fragment עם אותו מזהה; הסמן שייך
    // לזה שטווחו מכיל אותו.
    const end = Number(fragment.getAttribute('data-pm-end'));
    const caretPm = base + caretOffset;
    if (Number.isFinite(end) && (caretPm < base || caretPm > end)) continue;

    const lines: PaintedLine[] = [];
    for (const child of Array.from(fragment.children)) {
      const pmStart = Number(child.getAttribute('data-pm-start'));
      const pmEnd = Number(child.getAttribute('data-pm-end'));
      if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) continue;
      lines.push({ pmStart, pmEnd, rtl: child.getAttribute('dir') === 'rtl' });
    }
    if (lines.length) return { base, lines };
  }

  return null;
}

/** מתקין את היירוט. אין לו מצב, והוא נפרק עם המסמך. */
export function installRtlLineEnd({ host, superdoc }: RtlLineEndOptions): RtlLineEndHandle {
  if (!host) return { dispose() {} };

  let preferNextAtBoundary = false;

  const onKeyDown = (event: KeyboardEvent): void => {
    if (isPlainHomeKey(event)) {
      preferNextAtBoundary = true;
      return;
    }
    if (!isLineEndKey(event)) {
      preferNextAtBoundary = false;
      return;
    }
    const homePrecededEnd = preferNextAtBoundary;
    preferNextAtBoundary = false;

    const editor = superdoc?.activeEditor;
    const setSelectionTarget = editor?.authoring?.setSelectionTarget;
    const readSnapshot = editor?.host?.readLiveSelectionSyncSnapshot;
    if (typeof setSelectionTarget !== 'function' || typeof readSnapshot !== 'function') return;

    let snapshot: LiveSelectionSnapshot | null | undefined;
    try {
      snapshot = readSnapshot.call(editor?.host);
    } catch {
      return;
    }

    const selection = snapshot?.selectionTarget;
    const head = selection?.end;
    const anchor = selection?.start;
    if (!head || head.kind !== 'text' || typeof head.blockId !== 'string') return;
    if (typeof head.offset !== 'number') return;

    const fragment = readFragment(host, head.blockId, head.offset);
    if (!fragment) return;

    const target = lineEndOffset(fragment.lines, fragment.base, head.offset, homePrecededEnd);
    if (target === null) return;

    // מכאן ואילך זה שלנו: המנוע לא יראה את ההקשה.
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();

    // ה-`story` מגיע מהתצלום. ברירת המחדל שהייתה כאן — `{kind:'story',
    // storyType:'body'}` — מעולם לא נמדדה כנכונה, והייתה נכתבת למנוע דווקא
    // במצב שבו הוא עצמו לא דיווח story (הערת שוליים? כותרת?). נמדד בשער
    // (`check:line-end`, „ה-story בתצלום”) שהתצלום נושא אותו גם על העוגן וגם
    // על הקצה, ולכן אין מקרה שבו הברירה הזאת מגיעה. אותה הכרעה ב-`rtl-caret.ts`.
    const story = selection?.story ?? head.story;
    const coordinateSpace = selection?.coordinateSpace;
    const point: TextPoint = { kind: 'text', blockId: head.blockId, offset: target, story };
    const start = event.shiftKey && anchor && anchor.kind === 'text' ? anchor : point;

    // הסמן כבר שם ואין Shift — אין מה לכתוב, ורק ביטול האירוע נדרש (בלעדיו
    // המנוע היה מקפיץ אותו לתחילת השורה).
    if (!event.shiftKey && target === head.offset) return;

    try {
      const written = setSelectionTarget.call(editor?.authoring, {
        target: {
          kind: 'selection',
          start,
          end: point,
          story,
          ...(typeof coordinateSpace === 'string' ? { coordinateSpace } : {}),
        },
        focus: true,
      });
      // הפעולה א-סינכרונית במנוע. כשל מאוחר (למשל אם המסמך התפרק בין
      // התצלום לכתיבה) אינו צריך לייצר rejection לא מטופלת.
      void Promise.resolve(written).catch(() => {});
    } catch {
      /* בחירה שלא נכתבה אינה סיבה להפיל הקלדה */
    }
  };

  const clearHomeHint = (): void => {
    preferNextAtBoundary = false;
  };

  host.addEventListener('keydown', onKeyDown, true);
  host.addEventListener('pointerdown', clearHomeHint, true);

  return {
    dispose() {
      host.removeEventListener('keydown', onKeyDown, true);
      host.removeEventListener('pointerdown', clearHomeHint, true);
    },
  };
}
