/**
 * „כותרת עליונה ותחתונה” — הקבוצה הגדולה בלשונית „הוספה” של Word העברי,
 * דרך `doc.headerFooters` ושלוש פעולות של `doc.sections`.
 *
 * ## למה „הוספה” ולא „עריכה”
 *
 * ב-Word הלחיצה על „כותרת עליונה → ערוך כותרת עליונה” עושה שני דברים: היא
 * מוודאת שקיים חלק כותרת למקטע, והיא **מעבירה את הסמן לתוכו**. השני אינו
 * אפשרי כאן: `doc.selection` הוא קריאה בלבד (`current` בלבד, ראו
 * `SelectionApi`), ו-`StoryLocator` נמסר כפרמטר `in` לפעולות כתיבה — כלומר
 * המנוע יודע *לכתוב* לתוך כותרת, אבל אין פעולה ציבורית שמזיזה את המיקוד
 * לשם. לכן הפקד כאן יוצר את הכותרת ומשאיר את המשתמש להיכנס אליה בלחיצה
 * כפולה על אזור הכותרת, כמו ב-Word עצמו.
 *
 * מכאן גם ש„עריכה” על מסמך שכבר יש בו כותרת היא **הצלחה בלי מוטציה**: אין מה
 * ליצור, ויצירת חלק שני הייתה מייצרת כותרת יתומה שאיש אינו מפנה אליה.
 *
 * ## על איזה מקטע זה חל
 *
 * על כל המקטעים, בדיוק כמו page-setup.ts ומאותה סיבה: `selection.current`
 * אינו מדווח באיזה מקטע הסמן — הוא מחזיר כתובות טקסט, לא כתובת מקטע — ואין
 * דרך ציבורית למפות ביניהם. במסמך רגיל יש מקטע אחד, ו„החל על כל המסמך” הוא
 * ממילא ברירת המחדל של Word בדיאלוג „הגדרת עמוד”.
 *
 * נגזרת אחת חשובה: נוצר **חלק אחד** וכל המקטעים מפנים אליו. זו התנהגות
 * „מקושר לקודם” של Word, ולא כותרת נפרדת לכל מקטע.
 *
 * ## „קשר לקודם” על מסמך בעל מקטע אחד
 *
 * למקטע הראשון אין קודם, ולכן הפעולה חלה על המקטעים שאחריו בלבד. במסמך בעל
 * מקטע יחיד אין למה להחיל אותה, והמשתמש מקבל הודעה שאומרת את זה — ולא כשל
 * של המנוע על כתובת שאין לה משמעות.
 *
 * ## NO_OP אינה שגיאה
 *
 * כמו ב-page-setup.ts: המנוע מחזיר `NO_OP` כשהמצב המבוקש כבר קיים (מתג שכבר
 * דלוק, או ניקוי הפניה שאינה קיימת). מבחינת המשתמש זו הצלחה.
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt, type MaybePromise } from './document-api';

/** שני הסוגים שהמנוע מגדיר (`HeaderFooterKind`). אין סוג שלישי. */
export type HeaderFooterKind = 'header' | 'footer';

/** שלושת הווריאנטים שהמנוע מגדיר (`HeaderFooterVariant`). */
export type HeaderFooterVariant = 'default' | 'first' | 'even';

/**
 * הסדר קבוע: הסרה עוברת על שלושתם, וגרסה שתוסיף ווריאנט רביעי תיראה כאן
 * ולא תישאר בשקט כהפניה שלא נוקתה.
 */
export const HEADER_FOOTER_VARIANTS: readonly HeaderFooterVariant[] = ['default', 'first', 'even'];

/** ההפניות המפורשות של מקטע אחד, כפי ש-`sections.list` מדווח אותן. */
interface SectionRefs {
  default?: string;
  first?: string;
  even?: string;
}

/** מה שנקרא מ-`sections.list()`. כל שדה אופציונלי — גרסה אחרת עשויה לא לחשוף אותו. */
interface SectionItem {
  address?: unknown;
  titlePage?: boolean;
  oddEvenHeadersFooters?: boolean;
  headerRefs?: SectionRefs;
  footerRefs?: SectionRefs;
}

/** `HeaderFooterSlotAddress` של המנוע. `section` הוא `unknown` — ראו page-setup.ts. */
interface SlotAddress {
  kind: 'headerFooterSlot';
  section: unknown;
  headerFooterKind: HeaderFooterKind;
  variant: HeaderFooterVariant;
}

/** `HeaderFooterPartsMutationResult` — קבלה שנושאת גם את מזהה החלק שנוצר. */
interface PartsReceipt extends DocReceipt {
  refId?: string;
}

/** `HeaderFooterResolveResult` — מה שמעניין כאן הוא ה-status בלבד. */
interface ResolveResult {
  status?: 'explicit' | 'inherited' | 'none';
}

/** `HeaderFooterPartEntry`. `referencedBySections` ריק = חלק יתום. */
interface PartEntry {
  refId?: string;
  referencedBySections?: readonly unknown[];
}

export interface HeaderFooterDocumentApi {
  sections?: {
    // `SectionsListResult` הוא `DiscoveryOutput`: `items` הוא עמוד תחת
    // `limit`/`offset`, ו-`total` הוא המספר האמיתי — ראו `readSections`.
    list?: (query?: {
      limit?: number;
      offset?: number;
    }) => MaybePromise<{ items?: readonly SectionItem[]; total?: number } | undefined>;
    setTitlePage?: (input: { target: unknown; enabled: boolean }) => MaybePromise<DocReceipt>;
    setOddEvenHeadersFooters?: (input: { enabled: boolean }) => MaybePromise<DocReceipt>;
  };
  headerFooters?: {
    resolve?: (input: { target: SlotAddress }) => MaybePromise<ResolveResult | undefined>;
    refs?: {
      set?: (input: { target: SlotAddress; refId: string }) => MaybePromise<DocReceipt>;
      clear?: (input: { target: SlotAddress }) => MaybePromise<DocReceipt>;
      setLinkedToPrevious?: (input: {
        target: SlotAddress;
        linked: boolean;
      }) => MaybePromise<DocReceipt>;
    };
    parts?: {
      // `HeaderFootersPartsListResult` הוא `DiscoveryOutput` גם הוא — ראו
      // `deleteOrphanParts`.
      list?: (query?: {
        kind?: HeaderFooterKind;
        limit?: number;
        offset?: number;
      }) => MaybePromise<{ items?: readonly PartEntry[]; total?: number } | undefined>;
      create?: (input: { kind: HeaderFooterKind }) => MaybePromise<PartsReceipt>;
      delete?: (input: {
        target: { kind: 'headerFooterPart'; refId: string };
      }) => MaybePromise<DocReceipt>;
    };
  };
}

/** מה שנדרש מ-SuperDoc: רק הפאסדה של המסמך. ראו page-setup.ts. */
export interface HeaderFooterHost {
  activeEditor?: { doc?: HeaderFooterDocumentApi | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type HeaderFooterTarget = SuperDoc | HeaderFooterHost | null | undefined;

/**
 * הטיית הכשל בעברית תקנית לכל סוג. „הכותרת העליונה” נקבה, „החלק” זכר —
 * מין דקדוקי אינו נגזר ממזהה, ולכן הביטוי נשמר שלם. ראו document-api.ts.
 */
const ADD_FAILED: Record<HeaderFooterKind, string> = {
  header: 'הוספת הכותרת העליונה נכשלה',
  footer: 'הוספת הכותרת התחתונה נכשלה',
};

const REMOVE_FAILED: Record<HeaderFooterKind, string> = {
  header: 'הסרת הכותרת העליונה נכשלה',
  footer: 'הסרת הכותרת התחתונה נכשלה',
};

const TITLE_PAGE_FAILED = 'שינוי „שונה בעמוד ראשון” נכשל';
const ODD_EVEN_FAILED = 'שינוי „שונה בעמודים זוגיים ואי-זוגיים” נכשל';
const LINK_FAILED = 'שינוי הקישור לכותרת של המקטע הקודם נכשל';

function unavailable(failedAction: string, detail: string, reason: string): CommandOutcome {
  return { ok: false, message: `${failedAction}: ${detail}`, reason };
}

/** הנוסח שהתכנית קובעת ב-§12, וזהה לזה שהיכולת מחזירה. ראו footnotes.ts. */
function unsupported(failedAction: string): CommandOutcome {
  return {
    ok: false,
    message: `${failedAction}: אינו זמין בגרסה זו`,
    reason: 'command-unsupported',
  };
}

function slotOf(
  section: unknown,
  kind: HeaderFooterKind,
  variant: HeaderFooterVariant,
): SlotAddress {
  return { kind: 'headerFooterSlot', section, headerFooterKind: kind, variant };
}

/**
 * כשל הקבלה, או `null` כשהיא הצליחה. `NO_OP` נחשב הצלחה — ראו הערת הפתיחה.
 */
function failureOf(failedAction: string, receipt: DocReceipt | undefined): CommandOutcome | null {
  const code = receipt?.failure?.code;
  if (receipt?.success !== false || code === 'NO_OP') return null;
  return { ok: false, message: receiptFailureText(failedAction, receipt), reason: code };
}

/** קריאה למנוע שאינה זורקת החוצה. ראו הערת ה„לעולם לא זורקת” ב-footnotes.ts. */
async function attempt<T>(
  failedAction: string,
  call: () => MaybePromise<T>,
): Promise<{ ok: true; value: T } | { ok: false; outcome: CommandOutcome }> {
  try {
    return { ok: true, value: await call() };
  } catch (error) {
    return {
      ok: false,
      outcome: { ok: false, message: thrownText(failedAction, error), reason: 'threw' },
    };
  }
}

/** הפאסדה, או כשל מנוסח כשאין מסמך. */
function docOf(host: HeaderFooterTarget): HeaderFooterDocumentApi | null {
  return (host as HeaderFooterHost | null | undefined)?.activeEditor?.doc ?? null;
}

/** גודל העמוד בכל שאיבה, ובלם מפני מנוע שיחזיר `total` שאינו יורד. ראו footnotes.ts. */
const PAGE_SIZE = 200;
const PAGE_GUARD = 1000;

/**
 * כל הפריטים של פעולת discovery, בשאיבת עמודים עד `total`.
 *
 * `sections.list` ו-`headerFooters.parts.list` הם `DiscoveryOutput` (נמדד:
 * `sections.list` מחזיר `page:{limit:250,...}` — ראו docs/engine-gaps.md),
 * כלומר `items` הוא עמוד בלבד. קריאה יחידה הייתה מחילה פעולה על 250 המקטעים
 * או החלקים הראשונים בלבד ומדווחת „בוצע” — אותו דפוס בדיוק כמו ב-footnotes.ts.
 */
async function collectPages<T>(
  failedAction: string,
  list: (query: { limit: number; offset: number }) => MaybePromise<
    { items?: readonly T[]; total?: number } | undefined
  >,
): Promise<{ ok: true; items: T[] } | { ok: false; items: T[]; outcome: CommandOutcome }> {
  const items: T[] = [];
  let offset = 0;
  let guard = 0;

  for (;;) {
    const listed = await attempt(failedAction, () => list({ limit: PAGE_SIZE, offset }));
    if (!listed.ok) return { ok: false, items, outcome: listed.outcome };

    const page = listed.value?.items ?? [];
    items.push(...page);
    if (page.length === 0) return { ok: true, items };

    offset += page.length;

    const total = listed.value?.total;
    if (!Number.isFinite(total) || offset >= (total as number)) return { ok: true, items };
    if (++guard > PAGE_GUARD) return { ok: true, items };
  }
}

/**
 * מקטעי המסמך שיש להם כתובת. מקטע בלי כתובת אינו יעד חוקי לשום מוטציה, ולכן
 * הוא מסונן כאן ולא בכל אתר קריאה.
 *
 * שואבת עמודים עד `total` ולא נעצרת בעמוד הראשון — ראו `collectPages`.
 */
async function readSections(
  doc: HeaderFooterDocumentApi,
  failedAction: string,
): Promise<{ ok: true; value: SectionItem[] } | { ok: false; outcome: CommandOutcome }> {
  const list = doc.sections?.list;
  if (typeof list !== 'function') return { ok: false, outcome: unsupported(failedAction) };

  const collected = await collectPages<SectionItem>(failedAction, (query) => list(query));
  if (!collected.ok) return { ok: false, outcome: collected.outcome };

  const items = collected.items.filter((item) => item.address !== undefined && item.address !== null);
  if (items.length === 0) {
    return { ok: false, outcome: unavailable(failedAction, 'לא נמצא מקטע במסמך', 'target-unresolved') };
  }
  return { ok: true, value: items };
}

/* ------------------------------------------------------------------ */
/* קריאת המצב                                                          */
/* ------------------------------------------------------------------ */

/**
 * מה שהממשק צריך לדעת כדי לצייר את הקבוצה: אם יש כותרת, ואיך שלושת המתגים
 * עומדים. תצלום ולא מנוי — אותה החלטה כמו ב-doc-capabilities.ts.
 */
export interface HeaderFooterState {
  hasHeader: boolean;
  hasFooter: boolean;
  /** „שונה בעמוד ראשון” — `w:titlePg`. */
  titlePage: boolean;
  /** „שונה בעמודים זוגיים ואי-זוגיים” — `w:evenAndOddHeaders`, ברמת המסמך. */
  oddEven: boolean;
  /** „קשר לקודם”: אין הפניה מפורשת באף מקטע שאחרי הראשון. */
  linkedToPrevious: boolean;
  /** מספר המקטעים. פחות משניים = „קשר לקודם” חסר משמעות. */
  sectionCount: number;
}

/** המצב שאין בו כלום. „קשור לקודם” נכון כברירת מחדל: זו ברירת המחדל של Word. */
export function emptyHeaderFooterState(): HeaderFooterState {
  return {
    hasHeader: false,
    hasFooter: false,
    titlePage: false,
    oddEven: false,
    linkedToPrevious: true,
    sectionCount: 0,
  };
}

/**
 * קוראת את מצב הכותרות. לעולם אינה זורקת ולעולם אינה מחזירה „אולי” — כשל של
 * קריאה מחזיר את המצב הריק, כלומר הממשק יציג „אין כותרת” ולא מתג שקרי.
 */
export async function readHeaderFooterState(
  host: HeaderFooterTarget,
): Promise<HeaderFooterState> {
  const doc = docOf(host);
  if (!doc) return emptyHeaderFooterState();

  const sections = await readSections(doc, 'קריאת מצב הכותרות נכשלה');
  if (!sections.ok) return emptyHeaderFooterState();

  const items = sections.value;
  const first = items[0]!;

  const state: HeaderFooterState = {
    ...emptyHeaderFooterState(),
    titlePage: first.titlePage === true,
    oddEven: first.oddEvenHeadersFooters === true,
    sectionCount: items.length,
    // הפניה מפורשת במקטע שאינו הראשון = אותו מקטע **נותק** מהקודם לו.
    //
    // נבדקים כל שלושת הווריאנטים ולא `default` בלבד: מקטע שניתק רק את כותרת
    // העמוד הראשון מדווח „קשור” אם שואלים רק על `default`, והמתג היה מציג
    // מצב הפוך מהמסמך.
    linkedToPrevious: !items
      .slice(1)
      .some((item) =>
        HEADER_FOOTER_VARIANTS.some(
          (variant) =>
            item.headerRefs?.[variant] !== undefined || item.footerRefs?.[variant] !== undefined,
        ),
      ),
  };

  // `resolve` ולא `headerRefs` של המקטע: כותרת שהמקטע יורש קיימת מבחינת
  // המשתמש גם כשאין לו הפניה משלו, ו-`status: 'inherited'` הוא בדיוק
  // ההבחנה הזאת.
  const resolve = doc.headerFooters?.resolve;
  if (typeof resolve !== 'function') return state;

  for (const kind of ['header', 'footer'] as const) {
    const resolved = await attempt('קריאת מצב הכותרות נכשלה', () =>
      resolve({ target: slotOf(first.address, kind, 'default') }),
    );
    if (!resolved.ok) continue;
    const exists = resolved.value?.status === 'explicit' || resolved.value?.status === 'inherited';
    if (kind === 'header') state.hasHeader = exists;
    else state.hasFooter = exists;
  }

  return state;
}

/* ------------------------------------------------------------------ */
/* יצירה והסרה                                                         */
/* ------------------------------------------------------------------ */

/**
 * מוודאת שיש למסמך כותרת מהסוג המבוקש, ויוצרת אותה ריקה אם אין.
 *
 * `content` אינו נשלח: `parts.create` בלי `sourceRefId` יוצר חלק ריק, וזו
 * בדיוק ההתנהגות הנכונה — אין שום טקסט שנכון לשתול במסמך של מישהו אחר.
 * אותו היגיון כמו `content: ''` ב-footnotes.ts.
 */
export async function ensureHeaderFooter(
  host: HeaderFooterTarget,
  kind: HeaderFooterKind,
): Promise<CommandOutcome> {
  const failedAction = ADD_FAILED[kind];
  const doc = docOf(host);
  if (!doc) return unavailable(failedAction, 'המסמך עדיין נטען', 'document-api-unavailable');

  const create = doc.headerFooters?.parts?.create;
  const set = doc.headerFooters?.refs?.set;
  if (typeof create !== 'function' || typeof set !== 'function') return unsupported(failedAction);

  const sections = await readSections(doc, failedAction);
  if (!sections.ok) return sections.outcome;
  const items = sections.value;

  // `resolve` הוא **תנאי** ולא שיפור אופציונלי: בלעדיו אין דרך לדעת שכבר יש
  // כותרת, וכל לחיצה על „עריכה” הייתה יוצרת חלק ריק חדש ומפנה אליו את כל
  // המקטעים — כלומר מוחקת בשקט את הכותרת הקיימת. נכשל סגור, כמו
  // doc-capabilities.ts: „אינו זמין” עדיף על פעולה הרסנית.
  const resolve = doc.headerFooters?.resolve;
  if (typeof resolve !== 'function') return unsupported(failedAction);

  const resolved = await attempt(failedAction, () =>
    resolve({ target: slotOf(items[0]!.address, kind, 'default') }),
  );
  if (!resolved.ok) return resolved.outcome;
  // כבר יש כותרת. יצירת חלק שני הייתה משאירה חלק יתום שאיש אינו מפנה אליו.
  if (resolved.value?.status === 'explicit' || resolved.value?.status === 'inherited') {
    return { ok: true };
  }

  const created = await attempt(failedAction, () => create({ kind }));
  if (!created.ok) return created.outcome;

  const failure = failureOf(failedAction, created.value);
  if (failure) return failure;

  const refId = created.value?.refId;
  if (typeof refId !== 'string' || refId === '') {
    // הקבלה הצליחה אך לא נשא בה מזהה — בלעדיו אין למה להפנות, וההפניה
    // הייתה נכתבת ריקה.
    return unavailable(failedAction, 'המנוע לא החזיר מזהה לכותרת שנוצרה', 'missing-ref-id');
  }

  for (const section of items) {
    const assigned = await attempt(failedAction, () =>
      set({ target: slotOf(section.address, kind, 'default'), refId }),
    );
    if (!assigned.ok) return assigned.outcome;
    const assignFailure = failureOf(failedAction, assigned.value);
    if (assignFailure) return assignFailure;
  }

  return { ok: true };
}

/**
 * מסירה את הכותרת מכל המקטעים ומכל שלושת הווריאנטים, ואז מוחקת את החלקים
 * שנשארו בלי הפניה.
 *
 * הסדר אינו שרירותי: מחיקת חלק שמקטע עדיין מפנה אליו הייתה משאירה הפניה
 * שבורה ב-`sectPr`. לכן קודם `refs.clear` על הכול, ורק אז `parts.delete` על
 * מה ש-`parts.list` מדווח כבלתי-מופנה.
 */
export async function removeHeaderFooter(
  host: HeaderFooterTarget,
  kind: HeaderFooterKind,
): Promise<CommandOutcome> {
  const failedAction = REMOVE_FAILED[kind];
  const doc = docOf(host);
  if (!doc) return unavailable(failedAction, 'המסמך עדיין נטען', 'document-api-unavailable');

  const clear = doc.headerFooters?.refs?.clear;
  if (typeof clear !== 'function') return unsupported(failedAction);

  const sections = await readSections(doc, failedAction);
  if (!sections.ok) return sections.outcome;

  for (const section of sections.value) {
    for (const variant of HEADER_FOOTER_VARIANTS) {
      const cleared = await attempt(failedAction, () =>
        clear({ target: slotOf(section.address, kind, variant) }),
      );
      if (!cleared.ok) return cleared.outcome;
      // סלוט שלא היה בו כלום מחזיר NO_OP, וזו הצלחה: המשתמש ביקש שלא תהיה
      // כותרת, ואין.
      const failure = failureOf(failedAction, cleared.value);
      if (failure) return failure;
    }
  }

  return deleteOrphanParts(doc, kind, failedAction);
}

/**
 * מוחקת את חלקי הכותרת שאיש אינו מפנה אליהם.
 *
 * חלק יתום אינו שובר את המסמך, אבל הוא כן נשאר ב-ZIP ומופיע שוב בכל „הוסף
 * כותרת” הבא כאילו לא נמחק. גרסה שאין בה `parts.list`/`parts.delete` אינה
 * כשל של ההסרה עצמה — ההפניות כבר נוקו, וזה מה שהמשתמש ביקש.
 *
 * שואבת עמודים עד `total` — ראו `collectPages`: מסמך עם הרבה חלקי כותרת
 * יתומים (למשל אחרי כמה סבבי „הוסף”/„הסר”) לא היה מנוקה במלואו מקריאה יחידה.
 */
async function deleteOrphanParts(
  doc: HeaderFooterDocumentApi,
  kind: HeaderFooterKind,
  failedAction: string,
): Promise<CommandOutcome> {
  const list = doc.headerFooters?.parts?.list;
  const remove = doc.headerFooters?.parts?.delete;
  if (typeof list !== 'function' || typeof remove !== 'function') return { ok: true };

  const listed = await collectPages<PartEntry>(failedAction, (query) => list({ kind, ...query }));
  if (!listed.ok) return listed.outcome;

  for (const part of listed.items) {
    if (typeof part.refId !== 'string' || part.refId === '') continue;
    if ((part.referencedBySections?.length ?? 0) > 0) continue;

    const deleted = await attempt(failedAction, () =>
      remove({ target: { kind: 'headerFooterPart', refId: part.refId as string } }),
    );
    if (!deleted.ok) return deleted.outcome;
    const failure = failureOf(failedAction, deleted.value);
    if (failure) return failure;
  }

  return { ok: true };
}

/* ------------------------------------------------------------------ */
/* שלושת המתגים                                                        */
/* ------------------------------------------------------------------ */

/** „שונה בעמוד ראשון” — `w:titlePg`, לכל מקטע בנפרד. */
export async function setDifferentFirstPage(
  host: HeaderFooterTarget,
  enabled: boolean,
): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(TITLE_PAGE_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const setTitlePage = doc.sections?.setTitlePage;
  if (typeof setTitlePage !== 'function') return unsupported(TITLE_PAGE_FAILED);

  const sections = await readSections(doc, TITLE_PAGE_FAILED);
  if (!sections.ok) return sections.outcome;

  for (const section of sections.value) {
    const applied = await attempt(TITLE_PAGE_FAILED, () =>
      setTitlePage({ target: section.address, enabled }),
    );
    if (!applied.ok) return applied.outcome;
    const failure = failureOf(TITLE_PAGE_FAILED, applied.value);
    if (failure) return failure;
  }

  return { ok: true };
}

/**
 * „שונה בעמודים זוגיים ואי-זוגיים” — `w:evenAndOddHeaders`.
 *
 * זו הפעולה היחידה כאן שאינה מקבלת `target`: ב-OOXML הדגל יושב ב-
 * `settings.xml` ולא ב-`sectPr`, ולכן הוא ברמת המסמך כולו. החוזה משקף את זה
 * (`SectionsSetOddEvenHeadersFootersInput` נושא `enabled` בלבד), ואין כאן
 * מעבר על מקטעים.
 */
export async function setDifferentOddEvenPages(
  host: HeaderFooterTarget,
  enabled: boolean,
): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(ODD_EVEN_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const setOddEven = doc.sections?.setOddEvenHeadersFooters;
  if (typeof setOddEven !== 'function') return unsupported(ODD_EVEN_FAILED);

  const applied = await attempt(ODD_EVEN_FAILED, () => setOddEven({ enabled }));
  if (!applied.ok) return applied.outcome;

  return failureOf(ODD_EVEN_FAILED, applied.value) ?? { ok: true };
}

/**
 * „קשר לקודם” — על שני הסוגים ועל הווריאנט `default`, בכל המקטעים שאחרי
 * הראשון.
 *
 * שני הסוגים יחד ולא רק אחד: הפקד אינו יודע אם הסמן בכותרת העליונה או
 * בתחתונה — `selection.current` אינו מדווח story — ומתג שמנתק רק את אחת
 * מהשתיים היה משאיר את המשתמש עם חצי ניתוק בלי לומר לו.
 */
export async function setLinkedToPrevious(
  host: HeaderFooterTarget,
  linked: boolean,
): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(LINK_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const setLinked = doc.headerFooters?.refs?.setLinkedToPrevious;
  if (typeof setLinked !== 'function') return unsupported(LINK_FAILED);

  const sections = await readSections(doc, LINK_FAILED);
  if (!sections.ok) return sections.outcome;

  const followers = sections.value.slice(1);
  if (followers.length === 0) {
    return unavailable(LINK_FAILED, 'אין במסמך מקטע קודם לקשר אליו', 'no-previous-section');
  }

  for (const section of followers) {
    for (const kind of ['header', 'footer'] as const) {
      const applied = await attempt(LINK_FAILED, () =>
        setLinked({ target: slotOf(section.address, kind, 'default'), linked }),
      );
      if (!applied.ok) return applied.outcome;
      const failure = failureOf(LINK_FAILED, applied.value);
      if (failure) return failure;
    }
  }

  return { ok: true };
}
