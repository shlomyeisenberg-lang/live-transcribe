/**
 * ההחלה של התיקון האוטומטי על המנוע. ההחלטה **מה** לתקן ב-`autocorrect.ts`
 * ונבדקת ביחידה; כאן רק מה שהמנוע מכתיב: מתי לקרוא, מה לקרוא, ואיך להחליף.
 *
 * ## הדפוס הוא של ההשלמה מהספר, ולא של זיהוי הרשימות
 *
 * `list-autoformat-install.ts` נבנה סביב הקשה בודדת: הוא חייב להמיר ברגע ה-
 * רווח, לפני שהמנוע התיישב, ולכן הוא מחזיק רצף מקשים, עוגנים ונתיב מהיר. תיקון
 * אוטומטי אינו זקוק לזה ואינו יכול להרשות לעצמו את הסיבוך: טעות בזיהוי רשימה
 * היא פסקה שגויה, וטעות בתיקון היא **טקסט של המשתמש שהוחלף**. לכן כאן הדפוס
 * הפשוט והמדוד של `book-completion-overlay.ts`:
 *
 * - **ממתינים לשקט** (`DEBOUNCE_MS`) אחרי ההקלדה. קריאות מהמנוע עונות רק כשהוא
 *   נרגע (list-autoformat-install.ts, סעיף 1), וקריאה באמצע הקלדה רציפה אינה
 *   חוזרת עד שהיא נעצרת.
 * - **קוראים** `doc.selection.current` ואז `doc.ranges.resolve` על חלון שמסתיים
 *   **בסמן עצמו**, ומחליפים דרך `doc.insert({ target })` — אותה הכנסה שהשלמה
 *   מהספר מקבלת בה את ההצעה, ושנמדדה שם בשער.
 * - **טוקן ריצה** (`evalToken`) פוסל כל קריאה שבאוויר ברגע שהמשתמש חזר להקליד.
 *
 * ## מה נסרק: רק מה שהוקלד, ולא כל מה שלפני הסמן
 *
 * סריקה של החלון כולו הייתה מתקנת טקסט שהמשתמש לא הקליד: הדבקה שמכילה `(c)`,
 * מסמך שנפתח עם `teh` שנשאר בכוונה, תיקון שהמשתמש ביטל. לכן המתקין סופר **כמה
 * תווים הוקלדו** מאז האיפוס האחרון (`typed`), ומתקן רק תיקון שנוגע באזור הזה או
 * צמוד אליו. מה שמאפס: חץ, עכבר, Tab, Enter, קיצור, הדבקה, גרירה, חיתוך.
 *
 * ‏„צמוד אליו” ולא „בתוכו”: מי שלחץ בסוף `teh` והקליד רווח ביקש לתקן אותה, וכך
 * גם ב-Word. ההבדל מ„בתוכו” הוא תו אחד, והוא מה שמבדיל בין שני המצבים.
 *
 * ## ביטול: Ctrl+Z מחזיר את המקור, ואינו מתקן אותו שוב
 *
 * ההחלפה היא צעד היסטוריה משלה, ולכן Ctrl+Z הראשון מחזיר את מה שהוקלד — כמו
 * ב-Word. אבל ההקשה הבאה סורקת שוב ומוצאת את המקור. לכן כל תיקון שהוחל נרשם
 * (בלוק, היסט, מקור), וסריקה שמוצאת בדיוק אותו שוב **מדלגת עליו**. ההיסטים
 * אבסולוטיים ולכן הרשומה מתיישנת כשמקלידים לפני הטוקן — במקרה כזה התיקון יחזור,
 * וזו הטעות הנסבלת: חוסר בביטול זכור, לא תיקון שגוי.
 *
 * ## מה שאינו כאן
 *
 * - **Enter אינו מתקן את המילה האחרונה.** אחריו הסמן כבר בבלוק החדש, והמילה
 *   בבלוק הקודם. קריאה ממנו דורשת ללמוד את הבלוק הקודם מהמנוע, וזו הרחבה.
 *   חריג אחד: דפוס סמל שהמתין לתו הבא כתחילית של דפוס ארוך ממנו (`--` מול
 *   `-->`). ההערכה שבשקט שלפני ה-Enter כבר ידעה איפה הוא (`pending`), ולכן
 *   Enter או Tab מחליפים אותו שם. בלי קריאה שהושלמה לפני Enter אין עוגן בטוח,
 *   ולכן הקלדה מהירה ומיד Enter משאירה את המקור. כל עריכה נוספת מבטלת גם
 *   החלפה שכבר ממתינה לקריאת טווח — עדיף להחמיץ תיקון מאשר לדרוס טקסט חדש.
 * ביטול בלחיצה על המחוון נעשה ב-`changeBack`, אחרי בדיקת הטקסט העדכני.
 */
import type { CompletionDoc, CompletionTarget } from './book-completion-overlay';
import { planAutoCorrectWithPending, type AutoCorrectFix, type AutoCorrectOptions } from './autocorrect';
import { applyFixesToText, createAutoCorrectIndicator, locateWord, type IndicatorEntry } from './autocorrect-indicator';
import type { DocReceipt } from './document-api';
import type { ResolvedRangeLike, SelectionPointLike, SelectionTargetLike } from './word-selection';

/** כמה זמן שקט לפני שקוראים. כמו ב-`INPUT_DEBOUNCE_MS` של ההשלמה מהספר, ומעט יותר. */
export const AUTOCORRECT_DEBOUNCE_MS = 200;

/**
 * הקשר שנקרא **לפני** האזור שהוקלד, כדי שמילה שהוקלדה בסופה תיבדק בשלמותה ולא
 * כשבר שמתחיל באמצע. התיקון עצמו מסונן לאזור (ראו `touchesTyped`).
 */
export const AUTOCORRECT_CONTEXT_CHARS = 32;

/** תקרת התווים שנספרים. מעבר לה הספירה אינה מחמירה דבר, ורק מרחיבה את הקריאה. */
export const AUTOCORRECT_TYPED_CAP = 200;
/** הגבול של `preview.text` ב-`ranges.resolve`: מעליו התוצאה קטומה ונדחית (נמדד, ראו `word-selection.ts`). */
export const AUTOCORRECT_WINDOW_MAX = 200;

const SUPPRESS_TTL_MS = 10 * 60 * 1000;
const SUPPRESS_MAX = 60;

export interface AutoCorrectInstallOptions {
  container: HTMLElement;
  host: CompletionTarget;
  /** ההגדרות החיות. נקראות בכל הערכה — שינוי בהן אינו דורש התקנה מחדש. */
  getOptions: () => AutoCorrectOptions;
  /** האם משהו אחר מחזיק את ההקלדה כרגע (למשל דיאלוג מודאלי). */
  isSuppressed?: () => boolean;
  /** „הפסק לתקן אוטומטית” בתפריט הסימון: ההגדרות החדשות נשמרות אצל הקורא. */
  onStopCorrecting?: (fix: { rule: AutoCorrectFix['rule']; original: string; replacement: string }) => void;
  /** „שליטה באפשרויות התיקון האוטומטי…” בתפריט הסימון. */
  onOpenOptions?: () => void;
  /** האם לדווח לקונסולה על כל שלב. ברירת המחדל: `isAutoCorrectDebug`. */
  debug?: () => boolean;
  now?: () => number;
}

/**
 * מצב אבחון: `localStorage.setItem('otzaria-autocorrect-debug', '1')` ורענון,
 * או `window.otzariaAutocorrectDebug = true`.
 *
 * קיים כי כל כישלון כאן **שקט בכוונה** — טקסט שלא הוחלף אינו שגיאה למשתמש — ומי
 * שמדווח „זה לא מחליף” אינו יכול לדעת אם ההקשה לא נקלטה, הקריאה נדחתה או
 * ההכנסה נכשלה. המצב הזה אומר איזה.
 */
export function isAutoCorrectDebug(): boolean {
  try {
    const flag = (globalThis as { otzariaAutocorrectDebug?: unknown }).otzariaAutocorrectDebug;
    if (flag === true) return true;
    return globalThis.localStorage?.getItem('otzaria-autocorrect-debug') === '1';
  } catch {
    return false;
  }
}

export interface AutoCorrectHandle {
  /** מנקה את הספירה ואת זיכרון הביטולים — מסמך שהוחלף. */
  reset(): void;
  dispose(): void;
}

/**
 * דפוס סמל שממתין לתו הבא (`mayStillGrow` ב-autocorrect.ts), כפי שההערכה
 * האחרונה ראתה אותו. Enter או Tab מכריעים שלא יגדל, ואז הוא מוחלף כאן.
 */
interface PendingFix {
  readonly blockId: string;
  readonly story: unknown;
  readonly start: number;
  readonly original: string;
  readonly replacement: string;
}

/** תיקון שהוחל, כדי שביטול שלו לא יתוקן שוב. */
interface AppliedFix {
  readonly blockId: string;
  readonly start: number;
  readonly original: string;
  readonly at: number;
}

interface CaretWindow {
  readonly text: string;
  readonly base: number;
  readonly blockId: string;
  readonly story: unknown;
  readonly startsInsideToken: boolean;
}

function pointAt(blockId: string, offset: number, story: unknown): SelectionPointLike {
  const point: SelectionPointLike = { kind: 'text', blockId, offset };
  if (story !== undefined && story !== null) point.story = story;
  return point;
}

/** סמן מכווץ בלבד. טווח נבחר אינו „הקלדה”, ואין מה לתקן לפניו. */
function readCaretSeed(
  target: SelectionTargetLike | null | undefined,
): { blockId: string; offset: number; story: unknown } | null {
  if (!target || target.kind !== 'selection') return null;
  if (target.coordinateSpace !== undefined && target.coordinateSpace !== 'visible') return null;

  const { start, end } = target;
  if (start?.kind !== 'text' || end?.kind !== 'text') return null;
  if (typeof start.blockId !== 'string' || start.blockId !== end.blockId) return null;
  if (typeof start.offset !== 'number' || start.offset !== end.offset) return null;

  return { blockId: start.blockId, offset: start.offset, story: target.story ?? start.story ?? null };
}

/**
 * הטקסט שלפני הסמן, מסתיים **בסמן** ומתחיל `length` תווים לפניו.
 *
 * `null` כשמשהו אינו מתיישב — והכלל הוא שאז **אין תיקון**. טקסט שאינו באורך
 * שההיסטים אומרים (תו שהמנוע מייצג אחרת, חלון קטום) הוא טקסט שאי אפשר לחשב עליו
 * היסטים, ותיקון על היסטים שגויים מוחק טקסט אחר.
 */
type WindowResult = { ok: true; window: CaretWindow } | { ok: false; reason: string };

async function readCaretWindow(doc: CompletionDoc, length: number): Promise<WindowResult> {
  if (typeof doc.selection?.current !== 'function') return { ok: false, reason: 'אין doc.selection.current' };
  if (typeof doc.ranges?.resolve !== 'function') return { ok: false, reason: 'אין doc.ranges.resolve' };

  const info = await doc.selection.current();
  const seed = readCaretSeed(info?.selectionTarget);
  if (!seed) {
    return { ok: false, reason: `הבחירה אינה סמן מכווץ בבלוק (${JSON.stringify(info?.selectionTarget ?? null)})` };
  }

  const from = Math.max(0, seed.offset - length);
  const request: Record<string, unknown> = {
    start: { kind: 'point', point: pointAt(seed.blockId, from, seed.story) },
    end: { kind: 'point', point: pointAt(seed.blockId, seed.offset, seed.story) },
  };
  if (seed.story) request.in = seed.story;

  const resolved: ResolvedRangeLike | undefined = await doc.ranges.resolve(request);
  const text = resolved?.preview?.text;
  if (typeof text !== 'string' || resolved?.preview?.truncated === true) {
    return { ok: false, reason: `ranges.resolve לא החזירה טקסט שלם (${JSON.stringify(resolved?.preview ?? null)})` };
  }

  const base = resolved?.target?.start?.offset ?? from;
  const end = resolved?.target?.end?.offset ?? base + text.length;
  if (end !== seed.offset || base + text.length !== end) {
    return {
      ok: false,
      reason: `ההיסטים אינם מתיישבים: סמן=${seed.offset}, בסיס=${base}, סוף=${end}, אורך הטקסט=${text.length}`,
    };
  }

  // כשהחלון אינו מתחיל בבלוק, התו הראשון הוא שומר גבול. הוא נכלל באותה
  // קריאה של עד 200 תווים, אך אינו חלק מהטקסט שמתכננים עליו תיקונים.
  const guarded = base > 0 && text.length > 0;
  return {
    ok: true,
    window: {
      text: guarded ? text.slice(1) : text,
      base: base + (guarded ? 1 : 0),
      blockId: seed.blockId,
      story: seed.story,
      startsInsideToken: guarded && !/\s/u.test(text[0]!),
    },
  };
}

/** האם התיקון נוגע באזור שהוקלד או צמוד אליו. הסף הוא תו אחד לפני תחילת האזור. */
export function touchesTyped(fix: AutoCorrectFix, textLength: number, typed: number): boolean {
  return fix.end >= textLength - typed;
}

export function installAutoCorrect(options: AutoCorrectInstallOptions): AutoCorrectHandle {
  const { container, host } = options;
  /**
   * ה-`doc` נקרא **בכל שימוש** ולא פעם אחת בהתקנה, כמו `list-autoformat-install`:
   * ההתקנה רצה כש-container וה-superdoc כבר קיימים, אבל `activeEditor` עשוי
   * להיות עדיין `null` באותו רגע, ומתקין שתפס `null` נשאר עיוור עד ההתקנה הבאה.
   */
  const getDoc = (): CompletionDoc | null =>
    (host as { activeEditor?: { doc?: CompletionDoc | null } | null } | null | undefined)?.activeEditor?.doc ?? null;
  const now = options.now ?? (() => Date.now());
  const isDebug = options.debug ?? isAutoCorrectDebug;
  const log = (stage: string, detail?: unknown): void => {
    if (isDebug()) console.info('[otzaria-word][תיקון אוטומטי]', stage, detail ?? '');
  };

  const indicator = createAutoCorrectIndicator(container, {
    changeBack: (entry) => changeBack(entry),
    stop: options.onStopCorrecting ? (entry) => options.onStopCorrecting?.(entry) : undefined,
    openOptions: options.onOpenOptions,
    invalidate: (resetTyping, event) => {
      if (!resetTyping && event?.type === 'input' && isSettlementInput(event)) return;
      if (resetTyping) resetTyped();
      else cancelPending();
    },
  });

  let disposed = false;
  let composing = false;
  let typed = 0;
  let evalToken = 0;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let applied: AppliedFix[] = [];
  let pending: PendingFix | null = null;
  let settlement: { entry: PendingFix; token: number; doc: CompletionDoc; expectedInput: boolean } | null = null;

  function cancelPending(): void {
    evalToken += 1;
    if (timer !== undefined) clearTimeout(timer);
    timer = undefined;
    settlement = null;
  }

  /** איפוס הספירה. כל מה שאינו הקלדה של תו. */
  function resetTyped(): void {
    typed = 0;
    pending = null;
    cancelPending();
  }

  function remember(blockId: string, start: number, original: string): void {
    const at = now();
    applied = applied.filter((entry) => at - entry.at < SUPPRESS_TTL_MS);
    applied.push({ blockId, start, original, at });
    if (applied.length > SUPPRESS_MAX) applied = applied.slice(applied.length - SUPPRESS_MAX);
  }

  function wasUndone(blockId: string, start: number, original: string): boolean {
    const at = now();
    return applied.some(
      (entry) =>
        entry.blockId === blockId &&
        entry.start === start &&
        entry.original === original &&
        at - entry.at < SUPPRESS_TTL_MS,
    );
  }

  /**
   * „חזור ל…”: מחליפה את מה שהוחלף בחזרה למקור. הטקסט שבהיסט נקרא ונבדק קודם —
   * אם אינו מה שהוחלף (המסמך זז), לא נוגעים בו.
   */
  async function changeBack(entry: IndicatorEntry): Promise<void> {
    const doc = getDoc();
    if (!doc || typeof doc.insert !== 'function' || typeof doc.ranges?.resolve !== 'function') return;
    if (disposed || options.isSuppressed?.()) return;
    const found = locateWord(entry, container);
    if (!found) return;
    resetTyped();
    const token = evalToken;
    const start = found.start;
    const story = entry.story;
    // replacement יכול להגיע ל-255 תווים; preview יחיד מוגבל ל-200.
    let current = '';
    for (let offset = 0; offset < entry.replacement.length; offset += AUTOCORRECT_WINDOW_MAX) {
      const end = Math.min(entry.replacement.length, offset + AUTOCORRECT_WINDOW_MAX);
      const request: Record<string, unknown> = {
        start: { kind: 'point', point: pointAt(entry.blockId, start + offset, story) },
        end: { kind: 'point', point: pointAt(entry.blockId, start + end, story) },
      };
      if (story) request.in = story;
      const resolved = await doc.ranges.resolve(request);
      if (disposed || token !== evalToken || resolved?.preview?.truncated) return;
      const text = resolved?.preview?.text;
      if (typeof text !== 'string' || text.length !== end - offset) return;
      current += text;
    }
    if (current !== entry.replacement) {
      log('חזרה: הטקסט שבהיסט אינו מה שהוחלף', current);
      return;
    }
    const target: SelectionTargetLike = {
      kind: 'selection',
      start: pointAt(entry.blockId, start, story),
      end: pointAt(entry.blockId, start + entry.replacement.length, story),
      ...(story ? { story } : {}),
    };
    const receipt = await doc.insert({ value: entry.original, type: 'text', target });
    if (receipt?.success === false) {
      console.warn('[otzaria-word] תיקון אוטומטי: החזרה למקור נדחתה', receipt.failure ?? receipt);
      return;
    }
    // המקור נרשם כבר ב-`remember`, ולכן ההקשה הבאה לא תתקן אותו שוב.
    remember(entry.blockId, start, entry.original);
    log('חזרה למקור', { from: entry.replacement, to: entry.original });
  }

  /**
   * רושמת את המילים שהוחלפו אצל הסימון. כאן רק **היכן הן בבלוק** (לפי הטקסט
   * שאחרי ההחלפה); המיקום על המסך נקרא מה-DOM בכל ריחוף. ללא מלבן סמן, ללא
   * המתנה למנוע — ולכן הסימון אינו תלוי בתזמון של ציור הסמן.
   */
  function registerIndicators(
    window_: { text: string; base: number; blockId: string; story: unknown },
    fixes: readonly AutoCorrectFix[],
  ): void {
    const { text, base, blockId, story } = window_;
    const result = applyFixesToText(text, fixes);
    const entries: IndicatorEntry[] = fixes.map((fix, index) => ({
      rule: fix.rule,
      original: fix.original,
      replacement: fix.replacement,
      blockId,
      story,
      start: base + result.starts[index]!,
      context: {
        before: result.text.slice(Math.max(0, result.starts[index]! - 3), result.starts[index]!),
        after: result.text.slice(
          result.starts[index]! + fix.replacement.length,
          result.starts[index]! + fix.replacement.length + 3,
        ),
      },
    }));
    log('סימון', { count: entries.length });
    indicator.show(entries);
  }

  async function evaluate(token: number): Promise<void> {
    const settings = options.getOptions();
    const doc = getDoc();
    log('הערכה', { typed, enabled: settings.enabled, composing, hasDoc: doc !== null });
    if (!settings.enabled || composing || disposed || typed === 0 || !doc) return;
    if (options.isSuppressed?.()) {
      log('מושתק (דיאלוג פתוח)');
      return;
    }

    // `ranges.resolve` חותך את `preview.text` ב-200 תווים ומסמן `truncated`, ותוצאה קטומה נדחית.
    // לכן החלון כולו (מה שהוקלד + ההקשר) לא עובר את הגבול — אחרת אחרי ~170 תווים רצופים התיקון נעצר.
    const length = Math.min(Math.min(typed, AUTOCORRECT_TYPED_CAP) + AUTOCORRECT_CONTEXT_CHARS, AUTOCORRECT_WINDOW_MAX);
    let result: WindowResult;
    try {
      result = await readCaretWindow(doc, length);
    } catch (error) {
      console.warn('[otzaria-word] תיקון אוטומטי: קריאת הסמן נכשלה', error);
      return;
    }
    if (token !== evalToken || disposed) {
      log('נפסלה: המשתמש חזר להקליד');
      return;
    }
    if (!result.ok) {
      log('לא נקרא טקסט', result.reason);
      return;
    }

    const { text, base, blockId, story, startsInsideToken } = result.window;
    const { fixes: planned, pending: deferred } = planAutoCorrectWithPending(
      text, settings, { atBlockStart: base === 0, startsInsideToken },
    );
    const fixes = planned.filter(
      (fix) => touchesTyped(fix, text.length, typed) && !wasUndone(blockId, base + fix.start, fix.original),
    );
    log('תכנון', { text, planned: planned.length, afterFilter: fixes.length });

    // The same planning pass records a suffix deferred by a longer symbol pattern.
    const waiting = deferred && touchesTyped(deferred, text.length, typed) &&
      !wasUndone(blockId, base + deferred.start, deferred.original) ? deferred : null;
    // התיקונים שלפניו מזיזים אותו, כולם לפניו ואינם חופפים.
    const shift = waiting
      ? fixes.reduce((sum, fix) => (fix.end <= waiting.start ? sum + fix.replacement.length - (fix.end - fix.start) : sum), 0)
      : 0;
    const nextPending: PendingFix | null = waiting
      ? { blockId, story, start: base + waiting.start + shift, original: waiting.original, replacement: waiting.replacement }
      : null;
    if (fixes.length === 0) {
      pending = nextPending;
      return;
    }

    // מההיסט הגבוה לנמוך (כך `planAutoCorrect` מחזירה): החלפה לא מזיזה היסט
    // של מה שלפניה. כל הקשה בדרך פוסלת את הנותרים — הם יחושבו מחדש בשקט הבא.
    let completed = 0;
    let currentLength = text.length;
    for (const fix of fixes) {
      if (token !== evalToken || disposed) return;

      const target: SelectionTargetLike = {
        kind: 'selection',
        start: pointAt(blockId, base + fix.start, story),
        end: pointAt(blockId, base + fix.end, story),
        ...(story ? { story } : {}),
      };

      if (typeof doc.insert !== 'function') {
        log('אין doc.insert');
        return;
      }

      let receipt: DocReceipt;
      try {
        receipt = await doc.insert({ value: fix.replacement, type: 'text', target });
      } catch (error) {
        console.warn('[otzaria-word] תיקון אוטומטי: ההחלפה נכשלה', error);
        return;
      }
      if (receipt?.success === false) {
        // לא שקט: הכנסה שהמנוע דחה היא הדבר היחיד כאן שהמשתמש לא יכול לנחש.
        console.warn('[otzaria-word] תיקון אוטומטי: המנוע דחה את ההחלפה', receipt.failure ?? receipt);
        return;
      }
      log('הוחלף', { from: fix.original, to: fix.replacement, start: base + fix.start });

      remember(blockId, base + fix.start, fix.original);
      if (token === evalToken) {
        const delta = fix.replacement.length - (fix.end - fix.start);
        const typedStart = currentLength - typed;
        if (fix.start >= typedStart) typed = Math.min(AUTOCORRECT_TYPED_CAP, Math.max(0, typed + delta));
        else if (fix.end > typedStart) typed = Math.max(0, currentLength - fix.end);
        // תיקון שכולו לפני אזור ההקלדה מזיז את שני קצות האזור באותה מידה.
        currentLength += delta;
      }
      completed += 1;
    }

    // רק אם כל התיקונים הוחלו: הסימון מחשב היסטים מהטקסט הסופי, וחלקי אינו מתיישב.
    if (completed === fixes.length && !disposed && token === evalToken) {
      registerIndicators(result.window, fixes);
      pending = nextPending;
    }
  }

  /**
   * Enter או Tab אחרי דפוס שממתין: מחליפה אותו במקום שההערכה ראתה. הטקסט שם
   * נקרא ונבדק קודם, כמו ב-`changeBack` — אם אינו הדפוס, לא נוגעים בו.
   */
  async function settlePending(job: NonNullable<typeof settlement>): Promise<void> {
    const { entry, doc, token } = job;
    const current = (): boolean => {
      const settings = options.getOptions();
      return !disposed && token === evalToken && settlement === job && getDoc() === doc &&
        !composing && !options.isSuppressed?.() && settings.enabled && settings.replaceAsYouType &&
        settings.replacements.some((item) => item.from === entry.original && item.to === entry.replacement);
    };
    if (typeof doc.insert !== 'function' || typeof doc.ranges?.resolve !== 'function' || !current()) return;
    const end = entry.start + entry.original.length;
    const request: Record<string, unknown> = {
      start: { kind: 'point', point: pointAt(entry.blockId, entry.start, entry.story) },
      end: { kind: 'point', point: pointAt(entry.blockId, end, entry.story) },
    };
    if (entry.story) request.in = entry.story;
    const resolved = await doc.ranges.resolve(request);
    if (!current() || resolved?.preview?.truncated || resolved?.preview?.text !== entry.original ||
      resolved?.target?.start?.offset !== entry.start || resolved?.target?.end?.offset !== end) {
      log('המתנה: הטקסט שבהיסט אינו הדפוס', resolved?.preview?.text);
      return;
    }
    const target: SelectionTargetLike = {
      kind: 'selection',
      start: pointAt(entry.blockId, entry.start, entry.story),
      end: pointAt(entry.blockId, end, entry.story),
      ...(entry.story ? { story: entry.story } : {}),
    };
    const receipt = await doc.insert({ value: entry.replacement, type: 'text', target });
    if (receipt?.success === false) {
      console.warn('[otzaria-word] תיקון אוטומטי: המנוע דחה את ההחלפה', receipt.failure ?? receipt);
      return;
    }
    remember(entry.blockId, entry.start, entry.original);
    log('הוחלף אחרי סוף השורה', { from: entry.original, to: entry.replacement, start: entry.start });
  }

  function scheduleEvaluate(): void {
    cancelPending();
    const token = evalToken;
    timer = setTimeout(() => {
      timer = undefined;
      void evaluate(token);
    }, AUTOCORRECT_DEBOUNCE_MS);
  }

  /** מקש שאינו מוסיף תו: חץ, Tab, Enter, Escape, קיצור. כולם מעבירים את הסמן או את ההקשר. */
  function onKeyDown(event: KeyboardEvent): void {
    if (composing) return;

    const isShortcut = (event.ctrlKey || event.metaKey) && !(event.ctrlKey && event.altKey);
    if (isShortcut) {
      resetTyped();
      return;
    }

    if ((event.key === 'Enter' || event.key === 'Tab') && !event.shiftKey && !event.altKey && pending) {
      const entry = pending;
      resetTyped();
      const doc = getDoc();
      if (!doc) return;
      const job = { entry, token: evalToken, doc, expectedInput: true };
      settlement = job;
      // This uses the same timer and generation as normal evaluation: any subsequent
      // edit, undo, navigation, reset or disposal cancels it, including an in-flight read.
      timer = setTimeout(() => {
        timer = undefined;
        void settlePending(job).catch((error) => {
          console.warn('[otzaria-word] תיקון אוטומטי: ההחלפה אחרי סוף השורה נכשלה', error);
        }).finally(() => { if (settlement === job) settlement = null; });
      }, AUTOCORRECT_DEBOUNCE_MS);
      return;
    }

    if (event.key === 'Backspace') {
      pending = null;
      typed = Math.max(0, typed - 1);
      // בדיקה שכבר מתוזמנת או רצה חושבה על הטקסט שלפני המחיקה.
      cancelPending();
      return;
    }

    if (event.key.length === 1) {
      pending = null;
      cancelPending();
      typed = Math.min(AUTOCORRECT_TYPED_CAP, typed + 1);
      log('הקשה', { key: event.key, typed });
      return;
    }

    // Shift/Alt/CapsLock לבדם אינם שינוי: מי שלוחץ Shift באמצע מילה ממשיך להקליד.
    if (event.key === 'Shift' || event.key === 'Alt' || event.key === 'CapsLock' || event.key === 'AltGraph') {
      return;
    }

    resetTyped();
  }

  /**
   * `input` מתזמן הערכה; הדבקה, מחיקה והיסטוריה אינם הקלדה. `inputType` אינו
   * מובטח במשטח ההקלדה של המנוע, ולכן היעדרו נחשב הקלדה — והספירה (`typed`)
   * היא ההגנה האמיתית, לא הסוג.
   */
  function isSettlementInput(event: Event): boolean {
    const input = event as InputEvent;
    return !!settlement?.expectedInput && (input.inputType === 'insertParagraph' || input.inputType === 'insertLineBreak' ||
      (input.inputType === 'insertText' && input.data === '\t'));
  }

  function onInput(event: Event): void {
    const type = (event as InputEvent).inputType;
    // Enter's own structural input is expected; every other input invalidates the job.
    if (settlement && isSettlementInput(event)) {
      settlement.expectedInput = false;
      return;
    }
    if (type && /^(history|insertFromPaste|insertFromDrop|insertFromYank)/.test(type)) {
      resetTyped();
      return;
    }
    if (type && type.startsWith('delete')) {
      cancelPending();
      return;
    }
    if (!composing) scheduleEvaluate();
  }

  /**
   * `keyup` מתזמן גם הוא, כמו ב-`book-completion-overlay.ts` שמאזינה לשניהם:
   * `input` אינו מובטח במשטח ההקלדה של המנוע, ומתקין שהסתמך עליו לבדו לא היה
   * נקרא כלל. הכפילות בטוחה — `scheduleEvaluate` מאחדת קריאות לאחת, והספירה
   * (`typed`) היא מי שמחליטה אם יש מה לבדוק.
   */
  function onKeyUp(event: KeyboardEvent): void {
    if (composing || event.key.length !== 1) return;
    if (event.ctrlKey || event.metaKey) return;
    scheduleEvaluate();
  }

  const onReset = (): void => resetTyped();
  const onCompositionStart = (): void => {
    composing = true;
    cancelPending();
  };
  const onCompositionEnd = (): void => {
    composing = false;
    // תו שנוצר ב-IME נספר כאן, כי `keydown` שלו הוא `Process` ואינו מוסיף ספירה.
    typed = Math.min(AUTOCORRECT_TYPED_CAP, typed + 1);
    scheduleEvaluate();
  };

  container.addEventListener('keydown', onKeyDown, true);
  container.addEventListener('input', onInput);
  container.addEventListener('keyup', onKeyUp);
  container.addEventListener('compositionstart', onCompositionStart);
  container.addEventListener('compositionend', onCompositionEnd);
  for (const name of ['paste', 'cut', 'drop', 'pointerdown']) {
    container.addEventListener(name, onReset, true);
  }

  return {
    reset() {
      indicator.clear();
      resetTyped();
      applied = [];
    },
    dispose() {
      disposed = true;
      indicator.dispose();
      cancelPending();
      container.removeEventListener('keydown', onKeyDown, true);
      container.removeEventListener('input', onInput);
      container.removeEventListener('keyup', onKeyUp);
      container.removeEventListener('compositionstart', onCompositionStart);
      container.removeEventListener('compositionend', onCompositionEnd);
      for (const name of ['paste', 'cut', 'drop', 'pointerdown']) {
        container.removeEventListener(name, onReset, true);
      }
    },
  };
}
