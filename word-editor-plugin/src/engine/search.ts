/**
 * שכבת החיפוש-והחלפה. „חיפוש והחלפה” עובר דרך כאן ולא קורא ל-`ui.search`
 * בכלל — וזו הנקודה המרכזית של המודול הזה, לא רק פרט מימוש.
 *
 * ## למה לא `ui.search`
 *
 * הגרסה הקודמת של הקובץ הזה עטפה את `ui.search.search()`/`.find()`. זה נמדד
 * שבור על מסמך רב-פסקאות: 8 מופעים על פני כמה פסקאות, והמנוע דיווח 4 בלבד
 * (המונה בדיאלוג הראה „1 מתוך 4”). `replaceAll()` היה נאמן למה שנמצא — החליף
 * בדיוק את ארבעת אלה, לא פחות — כך שהתקלה האמיתית היא בכיסוי של החיפוש, לא
 * בהחלפה עצמה. המנוע עצמו רשם את האבחנה: `projection-incomplete:
 * exact-complete projection did not cover the full document (contiguous 1 of
 * 7 ordinals)`. זו תקלה במנגנון „projection” הפנימי שהחיפוש-לפי-שאילתה של
 * המנוע נשען עליו, ולא נפתרה גם ב-superdoc 2.10.0 (נבדק). פירוט מלא ומדוד
 * ב-docs/superdoc-2.10-review.md (”'החלף הכל' — האבחנה הקודמת הייתה שגויה”)
 * וב-docs/button-audit.md (טבלת „קשה לתקן”, שורה א').
 *
 * הפתרון: לעקוף את ה-projection לגמרי, ולא לנסות "לתקן" אותו — אין לנו גישה
 * למנגנון הפנימי הזה, והרישיון אוסר לגעת בו. `doc.blocks.list
 * ({includeText:true})` הוא ה-API הציבורי שכבר משמש בביטחון ב-
 * `engine/formatting-marks-layer.ts` וב-`engine/line-number-layer.ts` לקריאת
 * הטקסט הקנוני **המלא** של המסמך — טקסט שאומת שם, פסקה-פסקה, מול DOM אמיתי.
 * המודול הזה קורא את אותה רשימת בלוקים (עם דפדוף על `offset`/`limit`, כדי
 * שמסמך גדול לא ייחתך — אותה טכניקה בדיוק כמו `engine/caret-anchor.ts`),
 * ומעביר אותה ל-`engine/text-search.ts` — לוגיקה טהורה שמוצאת בעצמה את כל
 * המופעים, בלי שום תלות ב-projection של המנוע.
 *
 * ההחלפה עוברת דרך `doc.replace({target, text})` — פעולת ה-Document API
 * הציבורית שמקבלת `SelectionTarget` **מדויק** (בלוק+היסט, לא "המופע הבא
 * שהמנוע מוצא"), ולכן גם היא אינה תלויה ב-projection. מדידה ישירה מול
 * ה-dist הארוז (headless, docx מיוצא נבדק zip-by-zip) אישרה: `doc.blocks
 * .list()` על מסמך עם 8 פסקאות, מופע אחד בכל אחת, מחזיר את כל השמונה; קריאת
 * `doc.replace()` עם `SelectionTarget` בודד מחליפה בדיוק את המופע המבוקש
 * ומחזירה `textRangeShifts` שמאשר את ה-delta הצפוי באורך. „החלף הכל” כאן
 * מריץ קריאת `replace` אחת לכל מופע, **מהאחרון לראשון בתוך כל בלוק**
 * (`matchesForReplacement` ב-text-search.ts) — כי החלפה משנה את אורך הבלוק,
 * וכל מופע אחרי נקודת ההחלפה זז בהתאם; החלפה מהסוף ואחורה היא היחידה
 * ששומרת את ההיסטים המקוריים תקפים למופעים הבאים.
 *
 * ## מה שנשאר זהה למשתמש
 *
 * `SearchAdapter`, `SearchState`, `SearchOutcome`, `idleSearchState`,
 * `searchCounterText`, `replaceControlsVisible` — כל החוזה הציבורי שהדיאלוג
 * (`ui/panels/FindReplaceDialog.vue`) ו-`App.vue` צורכים נשאר כפי שהוא. רק
 * המימוש הפנימי של `createSearchAdapter` השתנה, וגם ההזרקה: הוא מקבל את
 * המופע (`SuperDoc`, כדי לגשת גם ל-`activeEditor.doc` וגם ל-`ui`) ולא רק את
 * ה-controller.
 *
 * הדגשת המופע הפעיל (מה ש-`ui.search`'s highlighting עשה בעבר) מתבצעת כאן
 * דרך `ui.selection.apply` (בחירת הטווח המדויק במסמך — נמדד: משנה בפועל את
 * ה-selection הלוגי של המנוע, `quotedText` בסנפשוט מאשר) ו-
 * `ui.viewport.scrollIntoView` (גלילה אליו) — שני API-ים ציבוריים, לא DOM
 * פנימי. כשל בהם הוא ויזואלי-בלבד ואינו הופך תוצאת חיפוש/החלפה תקינה לכשל.
 */
import type { SuperDoc } from 'superdoc';
import type { SelectionTarget, TextAddress } from 'superdoc/ui';
import { REASON_TEXT } from './command-adapter';
import { receiptFailureText, thrownText, type MaybePromise } from './document-api';
import {
  findAllMatches,
  matchToTarget,
  advanceActiveIndex,
  activeIndexAfterReplace,
  matchesForReplacement,
  groupMatchesByBlock,
  type SearchableBlock,
  type TextMatch,
  type BlockMatchCount,
} from './text-search';

/** השקטה לחיפוש-בזמן-הקלדה. ראו התיעוד המקורי: 250ms הוא ההפרש שבו אדם מפסיק להקליד. */
export const SEARCH_DEBOUNCE_MS = 250;

/** ההסבר שהדיאלוג מציג במקום פקדי ההחלפה. ראו `replaceControlsVisible`. */
export const REPLACE_UNAVAILABLE_TEXT = 'החלפה אינה זמינה במסמך הזה כרגע';

/** מה שמוצג כשיש שאילתה ואין לה התאמות. תשובה, לא שגיאה. */
export const NO_MATCHES_TEXT = 'אין התאמות להחלפה';

/** מה שמוצג כשמבקשים החלפה בלי שאילתת חיפוש. */
export const NO_QUERY_TEXT = 'יש להזין טקסט לחיפוש לפני החלפה';

/** כמה בלוקים לבקש בכל קריאת `blocks.list`, וכמה קריאות לכל היותר. כמו caret-anchor.ts. */
const BLOCKS_PAGE_SIZE = 500;
const BLOCKS_MAX_PAGES = 50;

/**
 * תקרות `doc.mutations.apply` (mutation plan). מוצהרות כאן ולא מיובאות:
 * הן מוגדרות ב-`document-api/src/types/mutation-plan.types.ts` הפנימי,
 * ו-import מנתיב פנימי אסור (tests/unit/engine-boundaries.test.ts). הערכים
 * (200 צעדים, 500 יעדים) הם מה שנמדד באותו קובץ ב-node_modules בזמן הכתיבה
 * — אם המנוע ישנה אותם, החריגה מהם כאן פשוט מפעילה את הנתיב הרגיל
 * (`callReplace` לכל מופע) מוקדם יותר, ולא שוברת דבר.
 */
const MUTATIONS_MAX_STEPS = 200;
const MUTATIONS_MAX_TARGETS = 500;

/** בלוק אחד כפי ש-`doc.blocks.list({includeText:true})` מחזיר אותו, בחלק שנצרך כאן. */
interface SearchBlockEntry {
  nodeId?: unknown;
  nodeType?: unknown;
  text?: unknown;
}

interface SearchBlocksListResult {
  blocks?: readonly SearchBlockEntry[];
}

interface SearchReplaceReceipt {
  success?: boolean;
  failure?: { code?: string; message?: string };
}

/**
 * קבלת `doc.mutations.apply({atomic:true, steps:[...]})`. `steps[i].matchCount`
 * הוא מה שמאפשר לאמת שהמנוע באמת החליף בדיוק את מה שציפינו בכל בלוק —
 * ראו `applyAtomicRewriteChunks`.
 */
interface SearchMutationsApplyResult {
  success?: boolean;
  steps?: readonly { matchCount?: unknown }[];
  failure?: { code?: string; message?: string };
}

/**
 * צעד `text.rewrite` יחיד ב-mutation plan: מחליף את **כל** המופעים המילוליים
 * (`mode:'contains'`, לא regex) של `pattern` בתוך בלוק אחד ידוע-מראש
 * (`within`) — לא בכל המסמך, ולכן אינו נשען על אותו "projection" שמקיף את
 * כל המסמך ונמדד חלקי. נמדד ישירות מול ה-dist הארוז: תת-המחרוזת בלבד
 * מוחלפת (לא כל טקסט הבלוק), עיצוב מעורב באותה פסקה נשמר, וספירת
 * ההתאמות (`matchCount`) תואמת בדיוק את מה ש-`findAllMatches` הטהור מחשב
 * — כולל במקרים גובלים (מופעים חופפים כמו "aa" בתוך "aaaa").
 */
interface SearchRewriteStep {
  id: string;
  op: 'text.rewrite';
  where: {
    by: 'select';
    select: { type: 'text'; pattern: string; mode: 'contains'; caseSensitive: boolean };
    within: { kind: 'block'; nodeType: string; nodeId: string };
    require: 'all';
  };
  args: { replacement: { text: string } };
}

/** מה שנדרש מ-`superdoc.activeEditor.doc`: רק מה שנעשה בו שימוש כאן. */
export interface SearchDocumentApi {
  blocks?: {
    list?: (input?: {
      includeText?: boolean;
      offset?: number;
      limit?: number;
    }) => MaybePromise<SearchBlocksListResult | null | undefined>;
  } | null;
  replace?: (input: {
    target: SelectionTarget;
    text: string;
  }) => MaybePromise<SearchReplaceReceipt | null | undefined>;
  /**
   * מנוע-פלנים אטומי: כמה צעדי מוטציה בקריאה אחת, `undo`-step יחיד. „החלף
   * הכל" משתמש בו כדי לא להרכיב N צעדי ביטול נפרדים ל-N מופעים — ראו
   * `applyAtomicRewriteChunks`. אופציונלי: מנוע שאינו חושף אותו פשוט מפיל
   * את „החלף הכל" לנתיב `doc.replace` הנקודתי (תקין, רק לא ב-undo אחד).
   */
  mutations?: {
    apply?: (input: {
      atomic: true;
      changeMode: 'direct' | 'tracked';
      steps: readonly SearchRewriteStep[];
    }) => MaybePromise<SearchMutationsApplyResult | null | undefined>;
  } | null;
}

interface SearchUiSelection {
  apply?: (target: SelectionTarget) => { ok?: unknown } | null | undefined;
}

interface SearchUiViewport {
  scrollIntoView?: (input: {
    target: TextAddress;
    block?: 'start' | 'center' | 'end' | 'nearest';
    behavior?: 'auto' | 'instant' | 'smooth';
  }) => MaybePromise<{ success?: unknown } | null | undefined>;
}

/** מה שנדרש מ-`superdoc.ui`: רק לצורך הדגשת המופע הפעיל (ראו הראש של הקובץ). */
export interface SearchUi {
  selection?: SearchUiSelection | null;
  viewport?: SearchUiViewport | null;
}

/** מה שנדרש מהמופע. `activeEditor.doc` לקריאה/החלפה, `ui` להדגשה בלבד. */
export interface SearchHost {
  activeEditor?: { doc?: SearchDocumentApi | null } | null;
  ui?: SearchUi | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל בבדיקות — אותה תבנית כמו caret-anchor.ts. */
export type SearchTarget = SuperDoc | SearchHost | null | undefined;

/** המצב הקריא של החיפוש. */
export interface SearchState {
  query: string;
  /** מספר ההתאמות לשאילתה הנוכחית — מכל בלוקי המסמך, לא רק מה שהמנוע "הקרין". */
  total: number;
  /** ההתאמה הפעילה, מאופס. `-1` כשאין. */
  activeIndex: number;
  /** האם session חיפוש פתוח (הדיאלוג מוצג). */
  open: boolean;
  /** האם `doc.blocks.list` חשוף במסמך הזה. */
  available: boolean;
  /**
   * האם `doc.replace` חשוף. בניגוד לגרסה הקודמת, זהו דגל-יכולת בלבד ולא
   * "האם המנוע ירשה להחליף את הקבוצה הנוכחית" — שאלה כזאת נענית בפועל, בזמן
   * הפעולה עצמה: כשל אמיתי (מסמך לקריאה בלבד וכו') מגיע כהודעה מ-`replace`/
   * `replaceAll` ולא כהסתרה מוקדמת של הכפתור.
   */
  canReplace: boolean;
  /** החלפה שנשלחה ל-Document API וטרם הסתיימה. */
  isReplacing: boolean;
}

/**
 * תוצאת פעולה. `snapshot` הוא המצב אחרי הפעולה, כדי שהקורא לא יצטרך לשאול
 * שוב; `message` תמיד בעברית ומוכן להצגה.
 */
export type SearchOutcome =
  | { ok: true; snapshot: SearchState }
  | { ok: false; message: string; reason?: string };

/** המצב לפני שיש מסמך פתוח. לא קבוע משותף — כדי שקורא לא ישנה אותו לכולם. */
export function idleSearchState(): SearchState {
  return {
    query: '',
    total: 0,
    activeIndex: -1,
    open: false,
    available: false,
    canReplace: false,
    isReplacing: false,
  };
}

/** מונה התוצאות שמוצג ליד שדה החיפוש. */
export function searchCounterText(state: SearchState): string {
  if (!state.query) return '';
  if (state.total === 0) return 'אין תוצאות';
  if (state.activeIndex < 0) return `${state.total} תוצאות`;
  return `${state.activeIndex + 1} מתוך ${state.total}`;
}

/**
 * האם להציג את פקדי ההחלפה בדיאלוג.
 *
 *   1. אין `doc.blocks.list` במסמך הזה → אין גם החלפה, והפקדים מוסתרים.
 *   2. יש חיפוש ואין התאמות → הפקדים **נשארים**. שאילתה בלי תוצאות אינה סיבה
 *      להעלים את שדה ההחלפה מתחת לאצבע של מי שמקליד בו.
 *   3. יש התאמות → `canReplace` (יכולת ה-API) מכריע.
 */
export function replaceControlsVisible(state: SearchState): boolean {
  if (!state.available) return false;
  if (state.total === 0) return true;
  return state.canReplace;
}

export interface SearchAdapter {
  getState(): SearchState;
  subscribe(listener: (state: SearchState) => void): () => void;
  open(): SearchOutcome;
  close(): void;
  clear(): void;
  find(query: string, direction: 'next' | 'prev'): Promise<SearchOutcome>;
  findDebounced(query: string, onOutcome: (outcome: SearchOutcome) => void): void;
  replace(replacement: string): Promise<SearchOutcome>;
  replaceAll(replacement: string): Promise<SearchOutcome>;
  dispose(): void;
}

function failureUnavailable(): SearchOutcome {
  return { ok: false, message: REASON_TEXT['search-unavailable'], reason: 'search-unavailable' };
}

function threw(error: unknown): SearchOutcome {
  return { ok: false, message: thrownText('פעולת החיפוש נכשלה', error), reason: 'threw' };
}

/**
 * קוראת את **כל** בלוקי המסמך עם הטקסט הקנוני המלא, בדפדוף. `null` כשאין
 * `blocks.list` בכלל, וגם כשקריאה כלשהי נכשלה — במתכוון: כיסוי חלקי הוא
 * בדיוק הבאג שהמודול הזה בא לתקן, ולכן עדיף כשל גלוי על המשך עם פחות בלוקים
 * ממה שיש במסמך.
 */
async function readAllBlocks(doc: SearchDocumentApi | undefined): Promise<SearchableBlock[] | null> {
  const list = doc?.blocks?.list;
  if (typeof list !== 'function') return null;

  const blocks: SearchableBlock[] = [];
  let offset = 0;
  for (let page = 0; page < BLOCKS_MAX_PAGES; page += 1) {
    const result = await list({ includeText: true, offset, limit: BLOCKS_PAGE_SIZE });
    const entries = result?.blocks ?? [];
    for (const entry of entries) {
      if (typeof entry?.nodeId === 'string') {
        blocks.push({
          blockId: entry.nodeId,
          text: typeof entry.text === 'string' ? entry.text : '',
          nodeType: typeof entry.nodeType === 'string' ? entry.nodeType : undefined,
        });
      }
    }
    if (entries.length < BLOCKS_PAGE_SIZE) break;
    offset += entries.length;
  }
  return blocks;
}

/** `blocks` נשמר לצד `matches`: „החלף הכל" האטומי צריך את `nodeType` לכל בלוק (ראו `applyAtomicRewriteChunks`). */
type RefreshResult =
  | { ok: true; blocks: SearchableBlock[]; matches: TextMatch[] }
  | { ok: false; outcome: SearchOutcome };

export function createSearchAdapter(host: SearchTarget): SearchAdapter {
  const listeners = new Set<(state: SearchState) => void>();
  let debounceTimer: ReturnType<typeof setTimeout> | undefined;
  let isReplacing = false;

  const state = {
    query: '',
    matches: [] as TextMatch[],
    activeIndex: -1,
    open: false,
  };

  function getDoc(): SearchDocumentApi | undefined {
    return (host as SearchHost | null | undefined)?.activeEditor?.doc ?? undefined;
  }

  function getUi(): SearchUi | undefined {
    return (host as SearchHost | null | undefined)?.ui ?? undefined;
  }

  function capabilities(): { available: boolean; canReplace: boolean } {
    const doc = getDoc();
    const available = typeof doc?.blocks?.list === 'function';
    const canReplace = available && typeof doc?.replace === 'function';
    return { available, canReplace };
  }

  function snapshot(): SearchState {
    const caps = capabilities();
    return {
      query: state.query,
      total: state.matches.length,
      activeIndex: state.activeIndex,
      open: state.open,
      available: caps.available,
      canReplace: caps.canReplace,
      isReplacing,
    };
  }

  function emit(): void {
    const current = snapshot();
    for (const listener of listeners) listener(current);
  }

  function cancelPending(): void {
    if (debounceTimer === undefined) return;
    clearTimeout(debounceTimer);
    debounceTimer = undefined;
  }

  /** מרעננת את רשימת המופעים מ-`doc.blocks.list` הטרי — לא מהמצב שאולי כבר התיישן. */
  async function refresh(query: string): Promise<RefreshResult> {
    const caps = capabilities();
    if (!caps.available) return { ok: false, outcome: failureUnavailable() };
    let blocks: SearchableBlock[] | null;
    try {
      blocks = await readAllBlocks(getDoc());
    } catch (error) {
      return { ok: false, outcome: threw(error) };
    }
    if (blocks === null) return { ok: false, outcome: failureUnavailable() };
    return { ok: true, blocks, matches: findAllMatches(blocks, query) };
  }

  /** בוחרת וגוללת אל המופע הפעיל. ויזואלי-בלבד: כשל כאן נבלע ואינו הופך תוצאה תקינה לכשל. */
  async function focusActiveMatch(): Promise<void> {
    const match = state.matches[state.activeIndex];
    if (!match) return;
    const ui = getUi();
    try {
      ui?.selection?.apply?.(matchToTarget(match) as SelectionTarget);
    } catch {
      /* ויזואלי בלבד */
    }
    try {
      const target: TextAddress = {
        kind: 'text',
        blockId: match.blockId,
        range: { start: match.start, end: match.end },
      };
      await ui?.viewport?.scrollIntoView?.({ target, block: 'center', behavior: 'smooth' });
    } catch {
      /* ויזואלי בלבד */
    }
  }

  /** קריאת `doc.replace` בודדת, עם טיפול כשל אחיד. */
  async function callReplace(
    target: TextMatch,
    replacement: string,
  ): Promise<{ ok: true } | { ok: false; message: string; reason?: string }> {
    const replaceFn = getDoc()?.replace;
    if (typeof replaceFn !== 'function') {
      return { ok: false, message: REASON_TEXT['search-unavailable'], reason: 'search-unavailable' };
    }
    try {
      const receipt = await replaceFn({ target: matchToTarget(target) as SelectionTarget, text: replacement });
      if (!receipt?.success) {
        return {
          ok: false,
          message: receiptFailureText('ההחלפה נכשלה', receipt ?? undefined),
          reason: receipt?.failure?.code,
        };
      }
      return { ok: true };
    } catch (error) {
      return { ok: false, message: thrownText('ההחלפה נכשלה', error), reason: 'threw' };
    }
  }

  /**
   * `isReplacing` נדלק **כאן**, לפני כל `await` — לא אחרי הבדיקה הראשונה
   * שדורשת קריאת מנוע. הסדר הפוך (בדוק-קיים-משהו-להחליף, ואז נעל) משאיר
   * חלון בין קריאה שנייה ל-`replace`/`replaceAll` לבין הרגע שבו `isReplacing`
   * בפועל הופך `true`: שתי קריאות שמגיעות "באותו טיק" (למשל שתי לחיצות
   * מהירות על הכפתור) עוברות את שער `if (isReplacing)` יחד, ושתיהן שולחות
   * `doc.replace` בפועל. הבדיקה שבודקת בדיוק את זה (`replace שנייה בזמן
   * שהראשונה רצה`) תפסה את הגרסה הקודמת של הקוד כתקיעה מוחלטת — לא רק כפילות.
   */
  async function mutateOne(replacement: string): Promise<SearchOutcome> {
    if (!state.query) return { ok: false, message: NO_QUERY_TEXT, reason: 'no-query' };
    if (isReplacing) {
      return { ok: false, message: 'ההחלפה הקודמת עדיין רצה', reason: 'replace-pending' };
    }

    isReplacing = true;
    emit();
    try {
      const before = await refresh(state.query);
      if (!before.ok) return before.outcome;
      state.matches = before.matches;
      if (before.matches.length === 0) {
        state.activeIndex = -1;
        return { ok: false, message: NO_MATCHES_TEXT, reason: 'no-matches' };
      }
      const index = Math.min(Math.max(state.activeIndex, 0), before.matches.length - 1);

      const receipt = await callReplace(before.matches[index], replacement);
      if (!receipt.ok) return receipt;

      const after = await refresh(state.query);
      state.matches = after.ok ? after.matches : [];
      state.activeIndex = activeIndexAfterReplace(index, state.matches.length);
    } finally {
      // `isReplacing` כבוי **כאן**, לפני שנופלים החוצה מה-try/finally.
      // כל כשל כבר `return`-ר מתוך ה-try, לפני השורה הזאת — הקוד שאחרי
      // ה-try/finally (למטה) רץ רק בהצלחה, וקורא ל-snapshot() רק אחרי
      // ש-isReplacing כבר false. בלי זה, snapshot() שמחושב בתוך ה-try
      // (לפני ה-finally) קופא עם isReplacing:true — וה-outcome המוחזר
      // (ש-App.vue דורס איתו את searchState.value בהצלחה) משאיר את
      // כפתורי ההחלפה מנוטרלים עד לפעולה לא-קשורה הבאה. נמדד בפועל
      // בדפדפן אמיתי: "החלף" מוצלח אחד השאיר את הכפתורים disabled
      // לצמיתות. ראו tests/unit/search.test.ts (`isReplacing:false
      // ב-outcome עצמו, לא רק ב-state הפנימי`) ו-
      // scripts/qa/replace-buttons-reenable-qa.mjs.
      isReplacing = false;
      emit();
    }
    void focusActiveMatch();
    return { ok: true, snapshot: snapshot() };
  }

  /**
   * מנסה להחליף את כל המופעים ב-`doc.mutations.apply` אחד (או כמה, אם
   * המסמך חוצה את התקרות של המנוע) — `undo`-step יחיד לכל קריאה, בניגוד
   * ל-`doc.replace` הנקודתי (`undo`-step לכל מופע בנפרד).
   *
   * ## למה זה נחוץ בכלל
   *
   * הגרסה הקודמת קראה ל-`handle.replaceAll(replacement)` — פעולת batch
   * יחידה במנוע הסגור, עטופה מבפנים ב-`undo`-step אחד. המעבר ל-`doc.replace`
   * הנקודתי (אחד לכל מופע) פתר את בעיית הכיסוי המקורית, אבל פתח רגרסיה
   * אחרת שנמדדה בפועל: „החלף הכל” על 3 מופעים דרש 3 Ctrl+Z נפרדים לביטול,
   * לא אחד — על מסמך גדול (60+ מופעים) זה יהיה 60+ ביטולים.
   *
   * ## למה לא `TargetWhere`/`BlockWhere`, ולמה כן `SelectWhere`+`within`
   *
   * `doc.mutations.apply({atomic:true, steps:[{op:'text.rewrite', where:{by:
   * 'target', target:SelectionTarget}, ...}]})` נראה כמו המועמד המובן
   * מאליו — הוא בדיוק היעד המדויק שכבר יש לנו מ-`matchToTarget`. **נדחה
   * בזמן ריצה**: `DocumentApiValidationError: v2 text.rewrite currently
   * requires a ref produced by query.match/find or a single text selector.`
   * (נמדד ישירות מול ה-dist הארוז). `BlockWhere` (כתיבה מחדש של **כל**
   * טקסט הבלוק) כן עובד ונותן `undo` יחיד — אבל נמדד שהוא משטח עיצוב
   * מעורב בתוך פסקה: פסקה עם "bold" מודגש ו-"zzq normal" רגיל יצאה עם
   * **כל** הפסקה מודגשת אחרי כתיבה-מחדש של טקסט שטוח (ה-`style` שמוחל
   * כברירת מחדל על טקסט מוחלף במקרה של עיצוב לא-אחיד הוא `onNonUniform:
   * 'majority'` — התיעוד של המנוע עצמו). זו הייתה רגרסיה חמורה יותר מ-N
   * ביטולים: איבוד שקט של עיצוב.
   *
   * הפתרון שכן עובד: `where:{by:'select', select:{type:'text', pattern,
   * mode:'contains', caseSensitive:false}, within:{kind:'block', nodeType,
   * nodeId}, require:'all'}` — חיפוש-טקסט של המנוע, אבל **מוגבל ל-`within`
   * בלוק אחד ידוע-מראש** (מ-`blocks.list` שלנו, לא מהשערתו של המנוע). נמדד:
   * (1) מכסה נכון מסמך עם מופע אחד בכל אחת משמונה פסקאות נפרדות — בדיוק
   * התרחיש שהיה שבור ב-`ui.search` הבלתי-מוגבל; (2) `matchCount` שמוחזר
   * תואם בדיוק את מה ש-`findAllMatches` הטהור מחשב, כולל מקרי-קצה (מופעים
   * חופפים, "a.b" מילולי ולא regex); (3) מחליף **רק** את תת-המחרוזת
   * שנמצאה — עיצוב מעורב באותה פסקה נשמר במדויק; (4) Ctrl+Z **יחיד** מבטל
   * קריאת `mutations.apply` שלמה עם כמה צעדים, ללא קשר למספר המופעים בה.
   *
   * ## שמרנות: לא לבטוח בתוצאה חלקית
   *
   * `applyAtomicRewriteChunks` מאמתת שכל צעד דיווח בדיוק את מספר המופעים
   * שציפינו לו לפני שהיא מסמנת את הבלוק שלו כ"טופל". בלוק שלא אומת (או
   * שהניסיון האטומי כולו נכשל/לא זמין) חוזר לנתיב `callReplace` הנקודתי —
   * מדויק תמיד, גם אם לא ב-`undo` אחד. הרענון הטרי לפני הנתיב הזה מבטיח
   * שהוא רואה את המצב **האמיתי** של המסמך (אחרי מה שכבר הצליח באמת), ולא
   * מנחש — כך גם סמנטיקת מעבר-יחיד (טקסט חלופי שמכיל את השאילתה עצמה,
   * כמו "cat"→"category", אינו נסרק שוב) נשמרת בשני הנתיבים גם יחד.
   */
  async function applyAtomicRewriteChunks(
    query: string,
    replacement: string,
    groups: readonly BlockMatchCount[],
    nodeTypeByBlock: ReadonlyMap<string, string>,
  ): Promise<Set<string>> {
    const handled = new Set<string>();
    const mutations = getDoc()?.mutations;
    if (typeof mutations?.apply !== 'function') return handled;

    let index = 0;
    while (index < groups.length) {
      const chunk: { group: BlockMatchCount; nodeType: string }[] = [];
      let targetSum = 0;
      while (index < groups.length && chunk.length < MUTATIONS_MAX_STEPS) {
        const group = groups[index];
        const nodeType = nodeTypeByBlock.get(group.blockId);
        if (!nodeType) break; // אין nodeType ידוע לבלוק הזה — לא לנחש, לעצור את הצ'אנק כאן
        if (targetSum + group.count > MUTATIONS_MAX_TARGETS) break;
        targetSum += group.count;
        chunk.push({ group, nodeType });
        index += 1;
      }
      if (chunk.length === 0) return handled; // בלוק בודד חוצה תקרה, או nodeType חסר — לוותר על ההמשך האטומי

      const steps: SearchRewriteStep[] = chunk.map(({ group, nodeType }, i) => ({
        id: `r${i}`,
        op: 'text.rewrite',
        where: {
          by: 'select',
          select: { type: 'text', pattern: query, mode: 'contains', caseSensitive: false },
          within: { kind: 'block', nodeType, nodeId: group.blockId },
          require: 'all',
        },
        args: { replacement: { text: replacement } },
      }));

      let receipt: SearchMutationsApplyResult | null | undefined;
      try {
        receipt = await mutations.apply({ atomic: true, changeMode: 'direct', steps });
      } catch {
        return handled;
      }
      if (!receipt?.success) return handled;

      const stepReceipts = receipt.steps ?? [];
      const allMatch = chunk.every(({ group }, i) => stepReceipts[i]?.matchCount === group.count);
      if (!allMatch) return handled; // אין ביטחון בתוצאה החלקית — הצ'אנק כולו חוזר לנתיב הנקודתי

      for (const { group } of chunk) handled.add(group.blockId);
    }
    return handled;
  }

  async function mutateAll(replacement: string): Promise<SearchOutcome> {
    if (!state.query) return { ok: false, message: NO_QUERY_TEXT, reason: 'no-query' };
    if (isReplacing) {
      return { ok: false, message: 'ההחלפה הקודמת עדיין רצה', reason: 'replace-pending' };
    }

    isReplacing = true;
    emit();
    try {
      const before = await refresh(state.query);
      if (!before.ok) return before.outcome;
      state.matches = before.matches;
      const totalBefore = before.matches.length;
      if (totalBefore === 0) {
        state.activeIndex = -1;
        return { ok: false, message: NO_MATCHES_TEXT, reason: 'no-matches' };
      }

      // שלב 1: ניסיון undo-step יחיד (או כמה, אם המסמך חוצה תקרות המנוע) —
      // ראו התיעוד המלא ב-applyAtomicRewriteChunks.
      const nodeTypeByBlock = new Map(before.blocks.map((b) => [b.blockId, b.nodeType ?? 'paragraph']));
      const groups = groupMatchesByBlock(before.matches);
      const handledBlocks = await applyAtomicRewriteChunks(state.query, replacement, groups, nodeTypeByBlock);

      // שלב 2: מה שלא טופל בשלב 1 (המנוע לא חושף mutations.apply, בלוק
      // חוצה תקרה, או matchCount לא תואם) — נתיב doc.replace המדויק-בודד,
      // על בסיס רענון טרי (לא על "before" שאולי כבר התיישן בגלל שלב 1).
      let doneInFallback = 0;
      if (handledBlocks.size < groups.length) {
        const rescan = await refresh(state.query);
        if (!rescan.ok) return rescan.outcome;
        const remainingMatches = rescan.matches.filter((m) => !handledBlocks.has(m.blockId));
        const order = matchesForReplacement(remainingMatches);
        for (const match of order) {
          const receipt = await callReplace(match, replacement);
          if (!receipt.ok) {
            const after = await refresh(state.query);
            state.matches = after.ok ? after.matches : [];
            state.activeIndex = state.matches.length ? 0 : -1;
            const doneSoFar = totalBefore - remainingMatches.length + doneInFallback;
            return {
              ok: false,
              message: `הוחלפו ${doneSoFar} מתוך ${totalBefore} לפני שההחלפה נעצרה: ${receipt.message}`,
              reason: receipt.reason,
            };
          }
          doneInFallback += 1;
        }
      }

      const after = await refresh(state.query);
      state.matches = after.ok ? after.matches : [];
      state.activeIndex = state.matches.length ? 0 : -1;
    } finally {
      // ראו ההערה המקבילה ב-mutateOne: isReplacing כבוי כאן, לפני
      // שנופלים החוצה מה-try/finally, כדי ש-snapshot() שלמטה לא יקפיא
      // isReplacing:true ב-outcome המוחזר.
      isReplacing = false;
      emit();
    }
    void focusActiveMatch();
    return { ok: true, snapshot: snapshot() };
  }

  return {
    getState: snapshot,

    subscribe(listener) {
      listeners.add(listener);
      listener(snapshot());
      return () => {
        listeners.delete(listener);
      };
    },

    open() {
      const caps = capabilities();
      if (!caps.available) return failureUnavailable();
      state.open = true;
      emit();
      return { ok: true, snapshot: snapshot() };
    },

    close() {
      cancelPending();
      state.open = false;
      state.query = '';
      state.matches = [];
      state.activeIndex = -1;
      emit();
    },

    clear() {
      cancelPending();
      state.query = '';
      state.matches = [];
      state.activeIndex = -1;
      emit();
    },

    async find(query, direction) {
      cancelPending();
      if (!query) {
        state.query = '';
        state.matches = [];
        state.activeIndex = -1;
        emit();
        return { ok: true, snapshot: snapshot() };
      }

      const sameQuery = query === state.query;
      const result = await refresh(query);
      if (!result.ok) return result.outcome;

      state.query = query;
      state.matches = result.matches;
      state.activeIndex = sameQuery
        ? advanceActiveIndex(state.activeIndex, result.matches.length, direction)
        : advanceActiveIndex(-1, result.matches.length, direction);
      emit();
      void focusActiveMatch();
      return { ok: true, snapshot: snapshot() };
    },

    findDebounced(query, onOutcome) {
      cancelPending();
      debounceTimer = setTimeout(() => {
        debounceTimer = undefined;
        void (async () => {
          if (!query) {
            state.query = '';
            state.matches = [];
            state.activeIndex = -1;
            emit();
            onOutcome({ ok: true, snapshot: snapshot() });
            return;
          }
          const result = await refresh(query);
          if (!result.ok) {
            onOutcome(result.outcome);
            return;
          }
          state.query = query;
          state.matches = result.matches;
          // חיפוש-בזמן-הקלדה תמיד "מסמן את ההתאמה הראשונה" — כמו handle.search()
          // הישן — ואינו מקדם צעד-אחר-צעד גם כשהשאילתה חוזרת על עצמה.
          state.activeIndex = result.matches.length ? 0 : -1;
          emit();
          void focusActiveMatch();
          onOutcome({ ok: true, snapshot: snapshot() });
        })();
      }, SEARCH_DEBOUNCE_MS);
    },

    replace: (replacement) => mutateOne(replacement),
    replaceAll: (replacement) => mutateAll(replacement),

    dispose() {
      cancelPending();
      listeners.clear();
    },
  };
}
