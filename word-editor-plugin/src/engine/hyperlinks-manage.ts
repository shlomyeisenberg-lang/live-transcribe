/**
 * ניהול היפר-קישורים (גל 22): רשימה, הסרה ועריכה — דרך
 * `hyperlinks.list/wrap/remove`. `hyperlinks.patch` אינה זמינה ב-2.8.0,
 * ולכן **עריכה = הסרה + עטיפה מחדש** — אותה תבנית של `captions.update`
 * בגל 8: שתי פעולות לא-אטומיות מעל תוכן המשתמש.
 *
 * ## מה שנמדד
 *
 * - **היעד ל-wrap הוא `TextAddress`** (`{kind:'text',blockId,range}`) —
 *   SelectionTarget נדחה ("requires a valid TextAddress target").
 * - **היעד ל-remove הוא `HyperlinkTarget`** — כתובת של **צומת הקישור עצמו**
 *   (`{kind:'inline', nodeType:'hyperlink', anchor:{start,end}}`), לא טווח
 *   טקסט. `remove({ within: TextAddress })` זורק `Cannot read properties of
 *   undefined (reading 'anchor')` — הכתובת הנכונה נלקחת מ-`list()`.
 * - **wrap דורש מפרט קישור:** `{ link: { destination: { href } } }` — href בלבד
 *   זרק "requires a link specification object".
 * - `list()` מחזיר `{ items: [{ address, properties, text }] }` (מעטפת
 *   discovery סטנדרטית) — לא `stories[].hyperlinks[]`. הסרה על כתובת שאינה
 *   קישור מחזירה `TARGET_NOT_FOUND`.
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt } from './document-api';

const UNAVAILABLE_TEXT = 'אינו זמין בגרסה זו';

export interface HyperlinkInfo {
  id?: string;
  href?: string;
  anchor?: string;
}

/** מיקום בתוך המסמך — `blockId` + היסט טקסט שטוח (נמדד: `InlineAnchor.start/end`). */
interface HyperlinkPosition {
  blockId?: string;
  offset?: number;
}

/** כתובת צומת קישור, כפי שהיא חוזרת מ-`list()` ונשלחת חזרה ל-`remove`. */
export interface HyperlinkTargetAddress {
  kind: 'inline';
  nodeType: 'hyperlink';
  anchor: { start: HyperlinkPosition; end: HyperlinkPosition };
  story?: unknown;
}

/** פריט גולמי בודד מתוך `items[]` של `hyperlinks.list()`. */
interface RawHyperlinkItem {
  id?: string;
  address?: HyperlinkTargetAddress;
  properties?: { href?: string; anchor?: string };
  text?: string;
}

interface HyperlinksApiShape {
  selection?: {
    current?: () => MaybePromise<SelectionInfoLike | undefined>;
  };
  hyperlinks?: {
    list?: (input?: Record<string, unknown>) => MaybePromise<unknown>;
    wrap?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
    remove?: (input: Record<string, unknown>) => MaybePromise<DocReceipt>;
  };
}

export interface HyperlinksHost {
  activeEditor?: { doc?: HyperlinksApiShape | null } | null;
}

export type HyperlinksTarget = SuperDoc | HyperlinksHost | null | undefined;

interface SelectionInfoLike {
  empty?: boolean;
  target?: {
    segments?: ReadonlyArray<{ blockId?: string; range?: { start?: number; end?: number } }>;
  } | null;
}

type MaybePromise<T> = T | Promise<T>;

function docOf(host: HyperlinksTarget): HyperlinksApiShape | null {
  return (host as HyperlinksHost | null | undefined)?.activeEditor?.doc ?? null;
}

/** TextAddress מהבחירה — הקטע הראשון עם blockId וטווח. `null` אם אין. */
async function readTextAddress(
  doc: HyperlinksApiShape,
): Promise<{ kind: 'text'; blockId: string; range: { start: number; end: number } } | null> {
  try {
    const info = await doc.selection?.current?.();
    const segment = info?.target?.segments?.find(
      (s) =>
        typeof s?.blockId === 'string' &&
        typeof s.range?.start === 'number' &&
        typeof s.range.end === 'number',
    );
    if (!segment) return null;
    return {
      kind: 'text',
      blockId: segment.blockId as string,
      range: { start: segment.range!.start as number, end: segment.range!.end as number },
    };
  } catch {
    return null;
  }
}

/** פריטי `items[]` הגולמיים של `hyperlinks.list()` (מעטפת discovery סטנדרטית). */
function rawHyperlinkItems(raw: unknown): RawHyperlinkItem[] {
  if (!raw || typeof raw !== 'object') return [];
  const items = (raw as { items?: unknown }).items;
  return Array.isArray(items) ? (items as RawHyperlinkItem[]) : [];
}

/** דולה את הקישורים ל-`HyperlinkInfo[]` חיצוני שטוח. */
function flattenLinks(raw: unknown): HyperlinkInfo[] {
  return rawHyperlinkItems(raw).map((item) => ({
    id: item.id,
    href: item.properties?.href,
    anchor: item.properties?.anchor,
  }));
}

/**
 * מאתרת את כתובת צומת הקישור שחופפת לטווח המסומן (`TextAddress`), כדי
 * לשלוח אותה ל-`remove`. `null` = אין קישור בטווח.
 */
async function findHyperlinkTarget(
  doc: HyperlinksApiShape,
  address: { blockId: string; range: { start: number; end: number } },
): Promise<HyperlinkTargetAddress | null> {
  const list = doc.hyperlinks?.list;
  if (typeof list !== 'function') return null;

  let raw: unknown;
  try {
    raw = await list();
  } catch {
    return null;
  }

  for (const item of rawHyperlinkItems(raw)) {
    const anchor = item.address?.anchor;
    const start = anchor?.start?.offset;
    const end = anchor?.end?.offset;
    if (
      typeof start !== 'number' ||
      typeof end !== 'number' ||
      anchor?.start?.blockId !== address.blockId ||
      anchor?.end?.blockId !== address.blockId
    ) {
      continue;
    }
    // חפיפה (כולל מגע בקצה) בין טווח הקישור לטווח המסומן.
    if (address.range.start <= end && address.range.end >= start) {
      return item.address ?? null;
    }
  }
  return null;
}

/** רשימת כל הקישורים במסמך. מחזירה `null` כשאין מנוע. */
export async function listHyperlinks(host: HyperlinksTarget): Promise<HyperlinkInfo[] | null> {
  const list = docOf(host)?.hyperlinks?.list;
  if (typeof list !== 'function') return null;
  try {
    return flattenLinks(await list());
  } catch {
    return null;
  }
}

/** „הסר קישור" על הטווח המסומן. TARGET_NOT_FOUND = כבר אין קישור = הצלחה. */
export async function removeHyperlink(host: HyperlinksTarget): Promise<CommandOutcome> {
  const failedAction = 'הסרת הקישור נכשלה';

  const doc = docOf(host);
  const remove = doc?.hyperlinks?.remove;
  if (!doc || typeof remove !== 'function') {
    return { ok: false, message: `${failedAction}: ${UNAVAILABLE_TEXT}`, reason: 'command-unsupported' };
  }

  const address = await readTextAddress(doc);
  if (!address) {
    return { ok: false, message: `${failedAction}: יש לסמן טקסט תחילה`, reason: 'selection-required' };
  }

  const target = await findHyperlinkTarget(doc, address);
  if (!target) {
    // אין קישור בטווח = המצב המבוקש כבר מתקיים. לא כשל.
    return { ok: true };
  }

  let receipt: DocReceipt;
  try {
    receipt = await remove({ target });
  } catch (error) {
    return { ok: false, message: thrownText(failedAction, error), reason: 'threw' };
  }

  // אין קישור בטווח = המצב המבוקש כבר מתקיים. לא כשל.
  if (receipt?.success === false && receipt.failure?.code === 'TARGET_NOT_FOUND') {
    return { ok: true };
  }
  if (receipt?.success === false) {
    return { ok: false, message: receiptFailureText(failedAction, receipt), reason: receipt.failure?.code };
  }

  return { ok: true };
}

/**
 * „עריכת קישור" — remove + wrap. אימות לפני מגע (`list`), ואם העטיפה
 * נכשלת אחרי הסרה — הודעה שאומרת שהקישור אבד והטקסט נשמר.
 */
export async function editHyperlink(host: HyperlinksTarget, newHref: string): Promise<CommandOutcome> {
  const failedAction = 'עריכת הקישור נכשלה';

  const href = typeof newHref === 'string' ? newHref.trim() : '';
  if (href === '' || href.length > 2048) {
    return { ok: false, message: `${failedAction}: כתובת הקישור נדרשת`, reason: 'invalid-href' };
  }

  const doc = docOf(host);
  const wrap = doc?.hyperlinks?.wrap;
  const remove = doc?.hyperlinks?.remove;
  if (!doc || typeof wrap !== 'function' || typeof remove !== 'function') {
    return { ok: false, message: `${failedAction}: ${UNAVAILABLE_TEXT}`, reason: 'command-unsupported' };
  }

  const address = await readTextAddress(doc);
  if (!address) {
    return { ok: false, message: `${failedAction}: יש לסמן טקסט תחילה`, reason: 'selection-required' };
  }

  // אימות לפני מגע: אין קישור בטווח → עטיפה ישירה, בלי remove מיותר.
  const existingTarget = await findHyperlinkTarget(doc, address);
  const hadLink = existingTarget !== null;

  if (hadLink) {
    let removed: DocReceipt;
    try {
      removed = await remove({ target: existingTarget });
    } catch (error) {
      return { ok: false, message: thrownText(failedAction, error), reason: 'threw' };
    }
    if (removed?.success === false && removed.failure?.code !== 'TARGET_NOT_FOUND') {
      return { ok: false, message: receiptFailureText(failedAction, removed), reason: removed.failure?.code };
    }
  }

  let wrappedReceipt: DocReceipt;
  try {
    wrappedReceipt = await wrap({ target: address, link: { destination: { href } } });
  } catch (error) {
    if (hadLink) {
      return {
        ok: false,
        message: `${failedAction}: הקישור הוסר אך החדש לא נוצר — הטקסט נשמר ללא קישור`,
        reason: 'link-lost',
      };
    }
    return { ok: false, message: thrownText(failedAction, error), reason: 'threw' };
  }

  if (wrappedReceipt?.success === false) {
    if (hadLink) {
      return {
        ok: false,
        message: `${failedAction}: הקישור הוסר אך החדש לא נוצר — הטקסט נשמר ללא קישור`,
        reason: wrappedReceipt.failure?.code,
      };
    }
    return { ok: false, message: receiptFailureText(failedAction, wrappedReceipt), reason: wrappedReceipt.failure?.code };
  }

  return { ok: true };
}

