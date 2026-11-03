/**
 * מה ששורת המצב מודדת.
 *
 * הבדיקות כאן נגזרות מהתקלה שהן נועדו למנוע מלחזור: שלושת המספרים בשורת המצב
 * היו `ref(1)`, `ref(1)` ו-`ref(0)` שלא התעדכנו מעולם, ולכן כל מסמך — גם בן
 * שמונים עמודים — הוצג כ„עמוד 1 מתוך 1” ו„0 מילים”. מספר שנראה כמו מדידה
 * ואינו מדידה גרוע משורת מצב ריקה, ומכאן הכלל שכל הבדיקות כאן מודדות:
 * **מה שלא נמדד מוחזר כ-null.**
 *
 * שני מסלולי הכשל של `doc.info` נבדקים במפורש, כי שניהם קיימים בפועל: בדפדפן
 * הפאסדה א-סינכרונית ומחזירה הבטחה, והיא גם עשויה לזרוק. אף אחד מהם אינו
 * מפיל את התוסף ואינו מוחק מספר תקין שכבר על המסך.
 */
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  CURRENT_PAGE_DEBOUNCE_MS,
  WORD_COUNT_DEBOUNCE_MS,
  anchorPageIndex,
  countValue,
  createDocMetrics,
  emptyDocMetrics,
  pageCountValue,
  pageNumberFromIndex,
  readDocumentInfo,
  type DocInfoLike,
  type DocMetrics,
} from '../../src/engine/doc-metrics';

interface Harness {
  metrics: ReturnType<typeof createDocMetrics>;
  states: DocMetrics[];
  infoCalls: () => number;
  /** משנה את מה שהקריאה הבאה ל-info תעשה. */
  onInfo: (fn: () => DocInfoLike | Promise<DocInfoLike> | null) => void;
  /** משנה את אינדקס העמוד שהבחירה מדווחת. */
  setAnchor: (index: number | null) => void;
  anchorThrows: () => void;
}

function harness(): Harness {
  const states: DocMetrics[] = [];
  let infoCalls = 0;
  let infoImpl: () => DocInfoLike | Promise<DocInfoLike> | null = () => ({
    counts: { words: 12, pages: 3 },
  });
  let anchor: number | null = 0;
  let anchorThrows = false;

  const metrics = createDocMetrics({
    readInfo: () => {
      infoCalls += 1;
      return infoImpl();
    },
    readAnchorPageIndex: () => {
      if (anchorThrows) throw new Error('geometry');
      return anchor;
    },
    onChange: (next) => states.push(next),
  });

  return {
    metrics,
    states,
    infoCalls: () => infoCalls,
    onInfo: (fn) => {
      infoImpl = fn;
    },
    setAnchor: (index) => {
      anchor = index;
    },
    anchorThrows: () => {
      anchorThrows = true;
    },
  };
}

/** מריק microtasks: קריאת info עוברת דרך await גם כשהיא סינכרונית. */
async function flush(): Promise<void> {
  for (let i = 0; i < 20; i += 1) await Promise.resolve();
}

beforeEach(() => {
  vi.useFakeTimers();
});

afterEach(() => {
  vi.useRealTimers();
});

describe('פענוח ערכים', () => {
  it('ספירה: שלם אי-שלילי, ואפס הוא ערך חוקי', () => {
    expect(countValue(0)).toBe(0);
    expect(countValue(12.4)).toBe(12);
    expect(countValue(-1)).toBeNull();
    expect(countValue('12')).toBeNull();
    expect(countValue(undefined)).toBeNull();
    expect(countValue(Number.NaN)).toBeNull();
    expect(countValue(Number.POSITIVE_INFINITY)).toBeNull();
  });

  it('מספר עמודים: אפס אינו מסמך', () => {
    expect(pageCountValue(3)).toBe(3);
    expect(pageCountValue(0)).toBeNull();
    expect(pageCountValue(null)).toBeNull();
  });

  it('אינדקס מאופס הופך לעמוד מספר N', () => {
    expect(pageNumberFromIndex(0)).toBe(1);
    expect(pageNumberFromIndex(7)).toBe(8);
    expect(pageNumberFromIndex(null)).toBeNull();
    expect(pageNumberFromIndex(-1)).toBeNull();
  });

  it('מצב ריק הוא שלושה null ולא 1/1/0', () => {
    expect(emptyDocMetrics()).toEqual({ words: null, totalPages: null, currentPage: null });
  });
});

describe('קריאה מהמנוע', () => {
  it('info נקרא מהפאסדה של המסמך הפתוח', () => {
    const info = vi.fn(() => ({ counts: { words: 5 } }));
    expect(readDocumentInfo({ activeEditor: { doc: { info } } })).toEqual({ counts: { words: 5 } });
    expect(info).toHaveBeenCalledWith({});
  });

  it('מסמך שעוד נטען מוחזר כ-null ולא כחריגה', () => {
    expect(readDocumentInfo(null)).toBeNull();
    expect(readDocumentInfo({})).toBeNull();
    expect(readDocumentInfo({ activeEditor: { doc: null } })).toBeNull();
    expect(readDocumentInfo({ activeEditor: { doc: {} } })).toBeNull();
  });

  it('עמוד הסמן נקרא מ-getAnchorRect', () => {
    const getAnchorRect = vi.fn(() => ({ pageIndex: 4, left: 0, top: 0 }));
    expect(anchorPageIndex({ selection: { getAnchorRect } })).toBe(4);
    expect(getAnchorRect).toHaveBeenCalledWith({ placement: 'start' });
  });

  it('היעדר גיאומטריה הוא null — לא כשל ולא עמוד מומצא', () => {
    // המנוע מחזיר rect רק כשהבחירה נפתרת לפס פריסה מצויר.
    expect(anchorPageIndex({ selection: { getAnchorRect: () => null } })).toBeNull();
    expect(anchorPageIndex({ selection: {} })).toBeNull();
    expect(anchorPageIndex(null)).toBeNull();
    expect(
      anchorPageIndex({ selection: { getAnchorRect: () => ({ pageIndex: undefined }) } }),
    ).toBeNull();
  });
});

describe('ספירת מילים', () => {
  it('אינה רצה על כל הקשה אלא אחרי השקטה', async () => {
    const h = harness();

    for (let i = 0; i < 5; i += 1) {
      h.metrics.noteDocumentChanged();
      await vi.advanceTimersByTimeAsync(WORD_COUNT_DEBOUNCE_MS - 100);
    }
    expect(h.infoCalls()).toBe(0);

    await vi.advanceTimersByTimeAsync(WORD_COUNT_DEBOUNCE_MS);
    expect(h.infoCalls()).toBe(1);
    expect(h.metrics.getState().words).toBe(12);
  });

  it('פאסדה שמחזירה הבטחה נקראת נכון', async () => {
    const h = harness();
    h.onInfo(async () => ({ counts: { words: 480, pages: 7 } }));

    h.metrics.noteDocumentChanged();
    await vi.advanceTimersByTimeAsync(WORD_COUNT_DEBOUNCE_MS);
    await flush();

    expect(h.metrics.getState().words).toBe(480);
    expect(h.metrics.getState().totalPages).toBe(7);
  });

  it('info שזורק אינו מפיל, ואינו מוחק מספר תקין שכבר נמדד', async () => {
    const h = harness();
    h.metrics.noteDocumentChanged();
    await vi.advanceTimersByTimeAsync(WORD_COUNT_DEBOUNCE_MS);
    await flush();
    expect(h.metrics.getState().words).toBe(12);

    h.onInfo(() => {
      throw new Error('worker נפל');
    });
    h.metrics.noteDocumentChanged();
    await vi.advanceTimersByTimeAsync(WORD_COUNT_DEBOUNCE_MS);
    await flush();

    expect(h.metrics.getState().words).toBe(12);
  });

  it('הבטחה שנדחית מטופלת כמו זריקה', async () => {
    const h = harness();
    h.onInfo(() => Promise.reject(new Error('נדחה')));

    h.metrics.noteDocumentChanged();
    await vi.advanceTimersByTimeAsync(WORD_COUNT_DEBOUNCE_MS);
    await flush();

    expect(h.metrics.getState().words).toBeNull();
  });

  it('מסמך ריק מדווח 0 ולא „לא ידוע”', async () => {
    const h = harness();
    h.onInfo(() => ({ counts: { words: 0 } }));

    h.metrics.measureNow();
    await flush();

    expect(h.metrics.getState().words).toBe(0);
  });

  it('אין דיווח כשאף ערך לא השתנה', async () => {
    const h = harness();
    h.metrics.measureNow();
    await flush();
    const reports = h.states.length;

    h.metrics.noteDocumentChanged();
    await vi.advanceTimersByTimeAsync(WORD_COUNT_DEBOUNCE_MS);
    await flush();

    // אותה ספירה, אותו עמוד — שורת המצב לא צריכה להתרנדר מחדש.
    expect(h.states.length).toBe(reports);
  });
});

describe('מספר עמודים', () => {
  it('מגיע מ-callback העימוד', () => {
    const h = harness();

    h.metrics.notePaginationUpdate(24);

    expect(h.metrics.getState().totalPages).toBe(24);
  });

  it('דיווח פגום אינו מוחק מספר תקין', () => {
    const h = harness();
    h.metrics.notePaginationUpdate(24);

    h.metrics.notePaginationUpdate(0);
    h.metrics.notePaginationUpdate(Number.NaN);
    h.metrics.notePaginationUpdate('12');

    expect(h.metrics.getState().totalPages).toBe(24);
  });

  it('העימוד גובר על counts.pages של info', async () => {
    const h = harness();
    h.metrics.measureNow();
    await flush();
    expect(h.metrics.getState().totalPages).toBe(3);

    // הפריסה התארכה בזמן שה-info שנקרא לפניה עוד הראה 3.
    h.metrics.notePaginationUpdate(5);
    expect(h.metrics.getState().totalPages).toBe(5);
  });

  it('מעבר הפריסה הראשון סופר מילים אם עוד לא נספרו', async () => {
    const h = harness();
    // המצב שהמנוע מגיע אליו בפתיחה: פאסדה שעוד לא הייתה מוכנה.
    h.onInfo(() => null);
    h.metrics.measureNow();
    await flush();
    expect(h.metrics.getState().words).toBeNull();

    h.onInfo(() => ({ counts: { words: 61 } }));
    h.metrics.notePaginationUpdate(2);
    await vi.advanceTimersByTimeAsync(WORD_COUNT_DEBOUNCE_MS);
    await flush();

    expect(h.metrics.getState().words).toBe(61);
  });
});

describe('עמוד הסמן', () => {
  it('נקרא אחרי תזוזת הבחירה, בהשקטה קצרה', async () => {
    const h = harness();
    h.setAnchor(3);

    h.metrics.noteSelectionChanged();
    expect(h.metrics.getState().currentPage).toBeNull();

    await vi.advanceTimersByTimeAsync(CURRENT_PAGE_DEBOUNCE_MS);
    expect(h.metrics.getState().currentPage).toBe(4);
  });

  it('אין גיאומטריה — חוזר ל„לא ידוע” ולא נשאר על עמוד ישן', async () => {
    const h = harness();
    h.setAnchor(2);
    h.metrics.noteSelectionChanged();
    await vi.advanceTimersByTimeAsync(CURRENT_PAGE_DEBOUNCE_MS);
    expect(h.metrics.getState().currentPage).toBe(3);

    h.setAnchor(null);
    h.metrics.noteSelectionChanged();
    await vi.advanceTimersByTimeAsync(CURRENT_PAGE_DEBOUNCE_MS);

    expect(h.metrics.getState().currentPage).toBeNull();
  });

  it('קריאה שזורקת אינה מפילה', async () => {
    const h = harness();
    h.anchorThrows();

    h.metrics.noteSelectionChanged();
    await vi.advanceTimersByTimeAsync(CURRENT_PAGE_DEBOUNCE_MS);

    expect(h.metrics.getState().currentPage).toBeNull();
  });

  it('עריכה קוראת גם את העמוד — היא מזיזה את הסמן', async () => {
    const h = harness();
    h.setAnchor(1);

    h.metrics.noteDocumentChanged();
    await vi.advanceTimersByTimeAsync(CURRENT_PAGE_DEBOUNCE_MS);

    expect(h.metrics.getState().currentPage).toBe(2);
  });
});

describe('פירוק', () => {
  it('מבטל מדידות ממתינות', async () => {
    const h = harness();
    h.metrics.noteDocumentChanged();

    h.metrics.dispose();
    await vi.advanceTimersByTimeAsync(WORD_COUNT_DEBOUNCE_MS * 4);

    expect(h.infoCalls()).toBe(0);
  });

  it('ספירה שבאוויר אינה מדווחת על מסמך שנסגר', async () => {
    const h = harness();
    let release: ((info: DocInfoLike) => void) | undefined;
    h.onInfo(
      () =>
        new Promise<DocInfoLike>((resolve) => {
          release = resolve;
        }),
    );

    h.metrics.measureNow();
    h.metrics.dispose();
    release?.({ counts: { words: 999 } });
    await flush();

    // המסמך הבא הוא זה שעל המסך; 999 שייך לזה שנסגר.
    expect(h.states.some((state) => state.words === 999)).toBe(false);
  });

  it('אחרי פירוק אין מדידות חדשות', async () => {
    const h = harness();
    h.metrics.dispose();

    h.metrics.noteDocumentChanged();
    h.metrics.noteSelectionChanged();
    h.metrics.notePaginationUpdate(9);
    h.metrics.measureNow();
    await vi.advanceTimersByTimeAsync(WORD_COUNT_DEBOUNCE_MS * 4);
    await flush();

    expect(h.infoCalls()).toBe(0);
    expect(h.states).toEqual([]);
  });
});
