/**
 * אחרי יצירת טבלה (`table-insert` ב-InsertTab.vue) המנוע משאיר את הסמן בתא
 * הראשון שלה — נמדד. זה מקום לא ברור להמשיך ממנו, ובוורד הסמן עובר לפסקה
 * שאחרי הטבלה. המודול הזה עוקף את זה מבחוץ, בלי לגעת בפעולת ההכנסה עצמה
 * (שנשארת דרך `ui.commands`, כמו כל פקד ברצועה — §12 ב-remaining-waves.md).
 *
 * ## למה אי אפשר לקרוא את כתובת הטבלה מתוצאת הפקודה
 *
 * `ui.commands.executeAsync('table-insert', …)` מחזירה `CommandExecutionResult
 * = boolean | SuperDocUIReceipt`, ו-`SuperDocUIReceipt` נשען על `Receipt`
 * הגנרי (`ReceiptSuccess`), שאין בו שדה `table` — רק `inserted?: EntityAddress[]`,
 * וה-`EntityAddress` מוגבל להערות/מעקב שינויים. כלומר החוזה הציבורי של המסלול
 * המנותב אינו חושף את `nodeId` של הטבלה שנוצרה, גם אם `create.table` הישיר
 * (עם `CreateTableSuccessResult.table`) כן. §12: לא נשענים על שדה שאינו
 * ב-`.d.ts` הציבורי.
 *
 * ## הפתרון: הפרש לפני/אחרי, לא קריאת תוצאה
 *
 * `doc.blocks.list({ nodeTypes: ['table'] })` — ציבורי, מתועד — נקרא **לפני**
 * הרצת הפקודה וגם **אחריה**. הטבלה החדשה היא ה-`nodeId` שלא היה ברשימה
 * הראשונה. משם, `doc.blocks.list()` מלא (בלי סינון) מאתר את הבלוק הבא בתור —
 * ואם אין כזה, `doc.create.paragraph({ at: { kind: 'after', target } })` יוצר
 * אחד. הסמן עצמו זז דרך `ui.selection.apply`, עם אותו ניסיון-חוזר שכבר קיים
 * ב-`moveCaret` (clipboard.ts) לבלוק שזה עתה נוצר וטרם ניתן לאיתור.
 */
import type { SuperDoc } from 'superdoc';

import type { MaybePromise } from './document-api';

/** רשומת בלוק כפי ש-`blocks.list` מחזירה, בחלק שנצרך כאן. */
interface BlockListEntryLike {
  nodeId?: string;
}

interface BlocksListResultLike {
  blocks?: BlockListEntryLike[];
  total?: number;
}

interface BlocksListInputLike {
  offset?: number;
  limit?: number;
  nodeTypes?: string[];
}

interface CreateParagraphInsertionPointLike {
  blockId?: string;
  /** נמדד: `insertionPoint` הוא `TextAddress` — `range.start`, לא `offset`. */
  range?: { start?: number };
}

interface CreateParagraphResultLike {
  success?: boolean;
  insertionPoint?: CreateParagraphInsertionPointLike;
}

export interface TableInsertDocumentApi {
  blocks?: {
    list?: (input: BlocksListInputLike) => MaybePromise<BlocksListResultLike | undefined>;
  };
  create?: {
    paragraph?: (input: unknown) => MaybePromise<CreateParagraphResultLike | undefined>;
  };
}

/** תוצאת `apply` — נכשלת סגור עם `reason`, כמו `SelectionSurface` ב-clipboard.ts. */
interface SelectionApplyResultLike {
  ok?: boolean;
  reason?: string;
}

export interface TableInsertHost {
  activeEditor?: { doc?: TableInsertDocumentApi | null } | null;
  ui?: {
    selection?: {
      apply?: (target: unknown) => MaybePromise<SelectionApplyResultLike | undefined>;
    } | null;
  } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type TableInsertTarget = SuperDoc | TableInsertHost | null | undefined;

/** עמוד אחד הוא הרבה יותר מכל מספר טבלאות סביר במסמך אמיתי. */
const TABLE_PAGE_SIZE = 200;
/** עמוד אחד לסריקת כל הבלוקים, כשמחפשים את מה שבא אחרי הטבלה. */
const BLOCK_PAGE_SIZE = 200;
/** בלם מפני `total` שאינו יורד לעולם — אותה הכרעה כמו `readBookmarks`. */
const GUARD_LIMIT = 1000;

/** כמו `CARET_RETRIES`/`CARET_RETRY_MS` ב-clipboard.ts: בלוק שזה עתה נוצר. */
const CURSOR_RETRIES = 3;
const CURSOR_RETRY_MS = 80;

const wait = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

/**
 * כל מזהי הטבלאות במסמך, כ-`Set`. `null` כשאין `blocks.list` בכלל (כפיל חסר,
 * או גרסת מנוע ישנה) — ואז אין טעם לנסות את שאר הזרימה.
 */
export async function collectTableNodeIds(host: TableInsertTarget): Promise<Set<string> | null> {
  const list = (host as TableInsertHost | null | undefined)?.activeEditor?.doc?.blocks?.list;
  if (typeof list !== 'function') return null;

  const ids = new Set<string>();
  let offset = 0;
  let guard = 0;

  for (;;) {
    let page: BlocksListResultLike | undefined;
    try {
      page = await list({ offset, limit: TABLE_PAGE_SIZE, nodeTypes: ['table'] });
    } catch (error) {
      console.warn('[otzaria-word] קריאת רשימת הטבלאות נכשלה', error);
      return ids;
    }
    const items = page?.blocks ?? [];
    for (const item of items) {
      if (typeof item.nodeId === 'string') ids.add(item.nodeId);
    }
    if (items.length === 0) return ids;

    offset += items.length;
    const total = page?.total;
    if (!Number.isFinite(total) || offset >= (total as number)) return ids;
    if (++guard > GUARD_LIMIT) return ids;
  }
}

/**
 * מזהה הבלוק שבא מיד אחרי `nodeId`, לפי `blocks.list` המלא (בלי סינון סוג).
 *
 * `undefined` — לא נמצא כלל (או שאין `blocks.list`): מצב שאסור לפרש כ„טבלה
 * אחרונה”, כי זה עלול ליצור פסקה מיותרת. `null` — נמצא ואומת שהוא הבלוק
 * האחרון במסמך. מחרוזת — מזהה הבלוק הבא.
 *
 * לא סורקת את כל המסמך כשלא צריך: ברגע שהטבלה נמצאת בעמוד, קוראת עוד עמוד
 * קטן (`limit: 1`) רק כדי לדעת מה אחריה אם היא הייתה אחרונה בעמוד שלה.
 */
async function findBlockAfter(
  host: TableInsertTarget,
  nodeId: string,
): Promise<string | null | undefined> {
  const list = (host as TableInsertHost | null | undefined)?.activeEditor?.doc?.blocks?.list;
  if (typeof list !== 'function') return undefined;

  let offset = 0;
  let guard = 0;

  for (;;) {
    let page: BlocksListResultLike | undefined;
    try {
      page = await list({ offset, limit: BLOCK_PAGE_SIZE });
    } catch (error) {
      console.warn('[otzaria-word] קריאת רשימת הבלוקים נכשלה', error);
      return undefined;
    }
    const items = page?.blocks ?? [];
    if (items.length === 0) return undefined;

    const index = items.findIndex((entry) => entry.nodeId === nodeId);
    if (index !== -1) {
      if (index + 1 < items.length) return items[index + 1]?.nodeId ?? undefined;

      // הטבלה אחרונה בעמוד הזה. אם המסמך כולו נגמר כאן — היא גם אחרונה בו.
      const total = page?.total;
      const nextOffset = offset + items.length;
      if (Number.isFinite(total) && nextOffset >= (total as number)) return null;

      let nextPage: BlocksListResultLike | undefined;
      try {
        nextPage = await list({ offset: nextOffset, limit: 1 });
      } catch (error) {
        console.warn('[otzaria-word] קריאת הבלוק הבא אחרי הטבלה נכשלה', error);
        return undefined;
      }
      const nextId = nextPage?.blocks?.[0]?.nodeId;
      return typeof nextId === 'string' ? nextId : null;
    }

    offset += items.length;
    const total = page?.total;
    if (!Number.isFinite(total) || offset >= (total as number)) return undefined;
    if (++guard > GUARD_LIMIT) return undefined;
  }
}

/** נקודת קצה מכווצת, לשליחה ל-`ui.selection.apply`. */
function collapsedSelectionTarget(blockId: string, offset: number): unknown {
  const point = { kind: 'text', blockId, offset };
  return { kind: 'selection', start: point, end: point };
}

/**
 * קובעת את הסמן בפסקה שאחרי `nodeId` — קיימת, או חדשה אם `nodeId` הוא הבלוק
 * האחרון במסמך. לעולם אינה זורקת: כשל כאן משאיר את הסמן איפה שהמנוע הניח
 * אותו, ולא מפיל את פעולת ההכנסה שכבר הצליחה.
 */
async function placeCursorAfterBlock(host: TableInsertTarget, nodeId: string): Promise<void> {
  const target = host as TableInsertHost | null | undefined;
  const applySelection = target?.ui?.selection?.apply;
  if (typeof applySelection !== 'function') return;

  const nextBlockId = await findBlockAfter(host, nodeId);
  // `undefined`: לא הצלחנו לקבוע אם יש בלוק אחרי — לא נוגעים בסמן.
  if (nextBlockId === undefined) return;

  if (nextBlockId !== null) {
    await applyCursorWithRetry(applySelection, collapsedSelectionTarget(nextBlockId, 0));
    return;
  }

  // הטבלה אחרונה במסמך: יוצרים פסקה ריקה אחריה ונועלים את הסמן בה.
  const createParagraph = target?.activeEditor?.doc?.create?.paragraph;
  if (typeof createParagraph !== 'function') return;

  try {
    const result = await createParagraph({
      at: { kind: 'after', target: { kind: 'block', nodeType: 'table', nodeId } },
    });
    const point = result?.insertionPoint;
    if (result?.success !== true || typeof point?.blockId !== 'string') return;

    await applyCursorWithRetry(
      applySelection,
      collapsedSelectionTarget(point.blockId, point.range?.start ?? 0),
    );
  } catch (error) {
    console.warn('[otzaria-word] יצירת פסקה אחרי הטבלה נכשלה', error);
  }
}

/**
 * כמו `moveCaret` ב-clipboard.ts: הבלוק שהיעד מפנה אליו יכול היות שזה עתה
 * נוצר וטרם ניתן לאיתור — `apply` נכשל סגור עם `reason`, לא שקט.
 */
async function applyCursorWithRetry(
  applySelection: (target: unknown) => MaybePromise<SelectionApplyResultLike | undefined>,
  selectionTarget: unknown,
): Promise<void> {
  for (let attempt = 0; attempt <= CURSOR_RETRIES; attempt += 1) {
    if (attempt > 0) await wait(CURSOR_RETRY_MS);

    let result: SelectionApplyResultLike | undefined;
    try {
      result = await applySelection(selectionTarget);
    } catch (error) {
      console.warn('[otzaria-word] הזזת הסמן אחרי הטבלה נכשלה', error);
      return;
    }

    if (result?.ok !== false) return;
    if (attempt === CURSOR_RETRIES) {
      console.warn('[otzaria-word] הזזת הסמן אחרי הטבלה נדחתה', result.reason);
    }
  }
}

/**
 * מזהה את הטבלה שנוספה מאז `tableNodeIdsBefore` (תצלום מ-`collectTableNodeIds`,
 * שנקרא **לפני** הרצת `table-insert`) ומעבירה את הסמן לפסקה שאחריה.
 *
 * `tableNodeIdsBefore === null` (אין `blocks.list`) או אי-זיהוי טבלה חדשה
 * יחידה — לא נוגעים בסמן ולא מדווחים כשל: ההכנסה עצמה כבר הצליחה והמשתמש
 * קיבל טבלה; מיקום הסמן הוא שיפור, לא תנאי להצלחה.
 */
export async function placeCursorAfterInsertedTable(
  host: TableInsertTarget,
  tableNodeIdsBefore: Set<string> | null,
): Promise<void> {
  if (tableNodeIdsBefore === null) return;

  const after = await collectTableNodeIds(host);
  if (after === null) return;

  const added = [...after].filter((id) => !tableNodeIdsBefore.has(id));
  if (added.length !== 1) return;

  await placeCursorAfterBlock(host, added[0]);
}
