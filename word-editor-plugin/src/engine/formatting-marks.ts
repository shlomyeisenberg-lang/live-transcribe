/**
 * קריאת בלוקי המסמך (`doc.blocks.list`) עבור שכבת „סימני עיצוב" (¶) —
 * ui/shell/PilcrowOverlay.vue. אותה תבנית בדיוק כמו `createLineNumberingModel`/
 * `createPageBorderModel` (engine/page-setup.ts): מונה דורות שזורק תשובה של
 * מסמך שכבר נסגר, דיווח רק על שינוי אמיתי, `setEnabled` שנמנע מקריאה כשאיש
 * אינו רואה את התוצאה (בדיוק כמו `createRulerModel` — קריאת `blocks.list()`
 * סורקת את כל המסמך, ועדיף לא לעשות זאת כשסימני העיצוב כבויים).
 *
 * הקובץ הזה נפרד מ-`engine/formatting-marks-layer.ts` באותה חלוקה בדיוק כמו
 * page-setup.ts מול line-number-layer.ts: כאן א-סינכרוני ותלוי-Document API,
 * שם טהור וסינכרוני.
 */
import type { MaybePromise } from './document-api';
import type { FormattingMarksBlock } from './formatting-marks-layer';

/** צורת כניסה אחת מ-`doc.blocks.list({includeText:true})` — רק מה שנדרש כאן. */
interface FormattingMarksBlockEntry {
  nodeId: string;
  nodeType: string;
  text?: string | null;
  isEmpty: boolean;
}

/** מה שנדרש מ-SuperDoc: רק ה-Document API של הבלוקים. */
export interface FormattingMarksHost {
  activeEditor?: {
    doc?: {
      blocks?: {
        list?: (input?: {
          includeText?: boolean;
        }) => MaybePromise<{ blocks?: readonly FormattingMarksBlockEntry[] } | undefined>;
      } | null;
    } | null;
  } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל בבדיקות — אותה תבנית כמו `PageSetupTarget`. */
export type FormattingMarksTarget = FormattingMarksHost | null | undefined;

/** קוראת את כל בלוקי המסמך עם הטקסט הקנוני המלא, או `null` כשאין Document API. */
export async function readFormattingMarksBlocks(
  host: FormattingMarksTarget,
): Promise<readonly FormattingMarksBlock[] | null> {
  const list = host?.activeEditor?.doc?.blocks?.list;
  if (typeof list !== 'function') return null;

  let result: { blocks?: readonly FormattingMarksBlockEntry[] } | undefined;
  try {
    result = await list({ includeText: true });
  } catch {
    return null;
  }

  const entries = result?.blocks;
  if (!entries) return null;

  return entries.map((entry) => ({
    nodeId: entry.nodeId,
    nodeType: entry.nodeType,
    text: entry.text ?? '',
    isEmpty: entry.isEmpty,
  }));
}

/** השקטה בין קריאה לקריאה — אותו ערך כמו „מספרי שורות", מאותה סיבה. */
export const FORMATTING_MARKS_DEBOUNCE_MS = 300;

export interface FormattingMarksModelSource {
  read: () => Promise<readonly FormattingMarksBlock[] | null>;
  onChange: (blocks: readonly FormattingMarksBlock[] | null) => void;
}

export interface FormattingMarksModel {
  getState(): readonly FormattingMarksBlock[] | null;
  /**
   * סימני עיצוב מוצגים או מוסתרים. כבוי אינו קורא כלום — `doc.blocks.list()`
   * סורק את המסמך כולו, וקריאה שלו כשאיש אינו רואה את התוצאה היא עבודה
   * מיותרת. הדלקה קוראת מיד (`refreshNow`-שקול), בלי לחכות לעריכה הראשונה.
   */
  setEnabled(enabled: boolean): void;
  /** קריאה מיידית, בלי השהיה. */
  refreshNow(): void;
  /** קריאה מושהית — אחרי שינוי כלשהו במסמך (עריכה). */
  noteDocumentChanged(): void;
  dispose(): void;
}

export function createFormattingMarksModel(source: FormattingMarksModelSource): FormattingMarksModel {
  let state: readonly FormattingMarksBlock[] | null = null;
  let enabled = false;
  let disposed = false;
  let generation = 0;
  let timer: ReturnType<typeof setTimeout> | undefined;

  async function read(): Promise<void> {
    const mine = ++generation;
    if (!enabled) return;
    let next: readonly FormattingMarksBlock[] | null;
    try {
      next = await source.read();
    } catch {
      next = null;
    }
    if (disposed || mine !== generation) return;
    state = next;
    source.onChange(next);
  }

  return {
    getState: () => state,

    setEnabled(next) {
      if (enabled === next) return;
      enabled = next;
      if (!enabled) {
        // הדור עולה כדי שקריאה שבאוויר לא תדווח אחרי הכיבוי.
        generation += 1;
        clearTimeout(timer);
        state = null;
        source.onChange(null);
        return;
      }
      void read();
    },

    refreshNow: () => void read(),

    noteDocumentChanged() {
      if (disposed || !enabled) return;
      clearTimeout(timer);
      timer = setTimeout(() => void read(), FORMATTING_MARKS_DEBOUNCE_MS);
    },

    dispose() {
      disposed = true;
      generation += 1;
      clearTimeout(timer);
    },
  };
}
