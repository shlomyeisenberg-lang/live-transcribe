/**
 * גשר הלחיצה על קישור `otzaria://`.
 *
 * הבדיקה מקבעת את מה שנמדד על המנוע: ריצה שהיא חלק מקישור נושאת
 * `data-link-rid`; קישור בסכימה חסומה מצויר כ-`<span data-link-blocked>`
 * ולעולם לא כ-`<a>`; ו-`hyperlinks.list()` הוא המקום היחיד שבו ה-`rId` הזה
 * מתורגם ליעד — ‏`stories[].hyperlinks[]`, לא `items[]`.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { installOtzariaLinkClicks } from '../../src/engine/otzaria-link-click';

/** התשובה של `list`, בצורה שנמדדה. */
function listReply(links: Array<{ rId: string; externalTarget?: string }>) {
  return {
    stories: [{ storyId: 'main', hyperlinks: links.map((l) => ({ ...l, targetKind: 'external' })) }],
    items: links.map((l) => ({ id: l.rId, properties: { href: l.externalTarget } })),
  };
}

function setup(
  links: Array<{ rId: string; externalTarget?: string }>,
  navigateResult: { ok: boolean; message?: string } = { ok: true },
) {
  const container = document.createElement('div');
  document.body.appendChild(container);

  const list = vi.fn(() => listReply(links));
  const navigate = vi.fn(() => navigateResult);
  const onStatus = vi.fn();
  const host = { activeEditor: { doc: { hyperlinks: { list } } } };
  const handle = installOtzariaLinkClicks(container, host, { navigate, onStatus });

  return { container, list, navigate, onStatus, handle };
}

/** ריצה חסומה, בדיוק כפי שה-DomPainter מצייר אותה. */
function blockedRun(rId: string, text = 'פסחים דף לד'): HTMLElement {
  const el = document.createElement('span');
  el.className = 'superdoc-text-run';
  el.setAttribute('data-link-rid', rId);
  el.setAttribute('data-link-blocked', 'true');
  el.setAttribute('role', 'text');
  el.setAttribute('aria-label', 'Invalid link - not clickable');
  el.textContent = text;
  return el;
}

/** ריצה שהסכימה שלה עברה — המנוע מפעיל אותה בעצמו. */
function anchorRun(rId: string, href: string): HTMLElement {
  const el = document.createElement('a');
  el.className = 'superdoc-text-run superdoc-link';
  el.setAttribute('data-link-rid', rId);
  el.setAttribute('href', href);
  el.textContent = 'אתר חיצוני';
  return el;
}

const flush = () => new Promise((resolve) => setTimeout(resolve, 0));

describe('installOtzariaLinkClicks', () => {
  beforeEach(() => {
    document.body.textContent = '';
  });

  it('לחיצה על ריצה חסומה מנווטת ליעד שמאחורי ה-rId', async () => {
    const { container, navigate } = setup([
      { rId: 'rId7', externalTarget: 'otzaria://open/book/42?index=1234&uid=id%3A42' },
    ]);
    container.appendChild(blockedRun('rId7'));

    container.querySelector('span')!.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    await flush();

    expect(navigate).toHaveBeenCalledWith({ kind: 'book', id: 42, index: 1234, uid: 'id:42' });
  });

  it('Ctrl+לחיצה עוברת באותו מסלול — זה אותו אירוע', async () => {
    const { container, navigate } = setup([{ rId: 'rId7', externalTarget: 'otzaria://open/book/9?index=0' }]);
    container.appendChild(blockedRun('rId7'));

    container
      .querySelector('span')!
      .dispatchEvent(new MouseEvent('click', { bubbles: true, ctrlKey: true }));
    await flush();

    expect(navigate).toHaveBeenCalledWith({ kind: 'book', id: 9, index: 0 });
  });

  it('לחיצה על צאצא של הריצה נתפסת גם היא', async () => {
    const { container, navigate } = setup([{ rId: 'rId7', externalTarget: 'otzaria://open/detection?q=%D7%90' }]);
    const run = blockedRun('rId7');
    const inner = document.createElement('span');
    inner.textContent = 'פנימי';
    run.appendChild(inner);
    container.appendChild(run);

    inner.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    await flush();

    expect(navigate).toHaveBeenCalledWith({ kind: 'detection', query: 'א' });
  });

  /**
   * זה מה שמונע פתיחה כפולה ביום שבו הסכימה תעבור במעלה הזרם: `<a>` אמיתי
   * מופעל בידי המנוע, ו-`onActivate` שלו מנווט בעצמו.
   */
  it('אינו נוגע ב-<a> אמיתי — המנוע מפעיל אותו', async () => {
    const { container, navigate, list } = setup([
      { rId: 'rId7', externalTarget: 'otzaria://open/book/42?index=0' },
    ]);
    container.appendChild(anchorRun('rId7', 'otzaria://open/book/42?index=0'));
    // jsdom מנסה לנווט בפועל על `<a href>`; המאזין הזה רץ **אחרי** הגשר
    // (הוא על document, והגשר על ה-container), ולכן הוא משתיק את הניווט
    // בלי להסתיר מהגשר את מצב האירוע.
    const swallow = (event: Event): void => event.preventDefault();
    document.addEventListener('click', swallow);

    container.querySelector('a')!.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
    await flush();

    document.removeEventListener('click', swallow);
    expect(navigate).not.toHaveBeenCalled();
    expect(list).not.toHaveBeenCalled();
  });

  it('קישור שאינו אוצריא נשאר באחריות המנוע', async () => {
    const { container, navigate } = setup([{ rId: 'rId7', externalTarget: 'https://example.com/x' }]);
    container.appendChild(blockedRun('rId7'));

    container.querySelector('span')!.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    await flush();

    expect(navigate).not.toHaveBeenCalled();
  });

  it('לחיצה מחוץ לקישור אינה קוראת למנוע בכלל', async () => {
    const { container, list } = setup([{ rId: 'rId7', externalTarget: 'otzaria://open/book/1' }]);
    const plain = document.createElement('span');
    plain.textContent = 'טקסט רגיל';
    container.appendChild(plain);

    plain.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    await flush();

    expect(list).not.toHaveBeenCalled();
  });

  it('rId שאין לו יעד ברשימה אינו מנווט', async () => {
    const { container, navigate } = setup([{ rId: 'rId9', externalTarget: 'otzaria://open/book/1' }]);
    container.appendChild(blockedRun('rId7'));

    container.querySelector('span')!.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    await flush();

    expect(navigate).not.toHaveBeenCalled();
  });

  /** „לחצתי ולא קרה כלום, וגם לא נאמר לי למה” — זה מה שנסגר כאן. */
  it('כשל ניווט מדווח למשתמש', async () => {
    const { container, onStatus } = setup(
      [{ rId: 'rId7', externalTarget: 'otzaria://open/book/42?index=0' }],
      { ok: false, message: 'פתיחת הקישור נכשלה: הספר אינו נמצא בספרייה של אוצריא' },
    );
    container.appendChild(blockedRun('rId7'));

    container.querySelector('span')!.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    await flush();

    expect(onStatus).toHaveBeenCalledWith(
      'פתיחת הקישור נכשלה: הספר אינו נמצא בספרייה של אוצריא',
      true,
    );
  });

  it('זריקה בניווט מדווחת ואינה מתפשטת', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onStatus = vi.fn();
    installOtzariaLinkClicks(
      container,
      { activeEditor: { doc: { hyperlinks: { list: () => listReply([{ rId: 'rId7', externalTarget: 'otzaria://open/book/42' }]) } } } },
      {
        navigate: () => {
          throw new Error('boom');
        },
        onStatus,
      },
    );
    container.appendChild(blockedRun('rId7'));

    container.querySelector('span')!.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    await flush();

    expect(onStatus).toHaveBeenCalledWith('פתיחת הקישור נכשלה', true);
  });

  /**
   * גרירה לסימון טקסט מסתיימת גם היא ב-`click`. בלי ההבחנה אי אפשר לסמן או
   * לתקן את הטקסט של הקישור בלי שהעורך יקפוץ לקורא.
   */
  it('גרירה לסימון אינה מנווטת', async () => {
    const { container, navigate } = setup([{ rId: 'rId7', externalTarget: 'otzaria://open/book/1' }]);
    const el = blockedRun('rId7');
    container.appendChild(el);

    el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, clientX: 100, clientY: 50 }));
    el.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 160, clientY: 50 }));
    await flush();

    expect(navigate).not.toHaveBeenCalled();
  });

  it('תזוזה זעירה בין הלחיצה לשחרור עדיין נחשבת לחיצה', async () => {
    const { container, navigate } = setup([{ rId: 'rId7', externalTarget: 'otzaria://open/book/1?index=0' }]);
    const el = blockedRun('rId7');
    container.appendChild(el);

    el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, clientX: 100, clientY: 50 }));
    el.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: 102, clientY: 51 }));
    await flush();

    expect(navigate).toHaveBeenCalledTimes(1);
  });

  it('לחיצה כפולה היא בחירת מילה, לא פתיחה', async () => {
    const { container, navigate } = setup([{ rId: 'rId7', externalTarget: 'otzaria://open/book/1' }]);
    container.appendChild(blockedRun('rId7'));

    container
      .querySelector('span')!
      .dispatchEvent(new MouseEvent('click', { bubbles: true, detail: 2 }));
    await flush();

    expect(navigate).not.toHaveBeenCalled();
  });

  it('לחיצה ימנית אינה מנווטת', async () => {
    const { container, list } = setup([{ rId: 'rId7', externalTarget: 'otzaria://open/book/1' }]);
    container.appendChild(blockedRun('rId7'));

    container.querySelector('span')!.dispatchEvent(new MouseEvent('click', { bubbles: true, button: 2 }));
    await flush();

    expect(list).not.toHaveBeenCalled();
  });

  it('אחרי dispose אין יותר האזנה', async () => {
    const { container, handle, list } = setup([{ rId: 'rId7', externalTarget: 'otzaria://open/book/1' }]);
    container.appendChild(blockedRun('rId7'));
    handle.dispose();

    container.querySelector('span')!.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    await flush();

    expect(list).not.toHaveBeenCalled();
  });

  /**
   * ה-handler קורא ל-`activate` כ-`void`, ולכן `list` שזורק אינו „לחיצה
   * שנכשלה” אלא **דחייה שאין לה תופס** — ובדפדפן זה `unhandledrejection` על
   * כל לחיצה. הענף הזה שרד סבב מוטציה בלי הבדיקה הזאת.
   */
  it('זריקה של list נבלעת, ואינה מנווטת לניחוש', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const navigate = vi.fn();
    const onStatus = vi.fn();
    // ‏`unhandledrejection` אינו אמין ב-jsdom, ולכן הסימן הנמדד הוא העקבה
    // שהענף משאיר: בלי ה-catch אין אזהרה והדחייה נשארת בלי תופס.
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    installOtzariaLinkClicks(
      container,
      {
        activeEditor: {
          doc: {
            hyperlinks: {
              list: () => {
                throw new Error('busy');
              },
            },
          },
        },
      },
      { navigate, onStatus },
    );
    container.appendChild(blockedRun('rId7'));

    container.querySelector('span')!.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    await flush();

    expect(navigate).not.toHaveBeenCalled();
    expect(onStatus).not.toHaveBeenCalled();
    expect(warn).toHaveBeenCalledWith(
      '[otzaria-word] קריאת הקישורים במסמך נכשלה',
      expect.any(Error),
    );
    warn.mockRestore();
  });

  it('מנוע בלי hyperlinks.list אינו מפיל את הלחיצה', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const navigate = vi.fn();
    installOtzariaLinkClicks(container, { activeEditor: { doc: {} } }, { navigate });
    container.appendChild(blockedRun('rId7'));

    container.querySelector('span')!.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    await flush();

    expect(navigate).not.toHaveBeenCalled();
  });
});
