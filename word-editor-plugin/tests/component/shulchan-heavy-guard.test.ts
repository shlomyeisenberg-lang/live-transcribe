/**
 * לשונית „שולחן העורך” היא הרצועה, והרצועה **עוקפת** את רישום הכלים על ה-kit:
 * `runTool` המקומי מייבא את הפונקציות ישירות, ויש בו שמונה-עשר אתרי קריאה.
 * לכן העטיפה של רישום הכלים אינה מכסה אותה, וצריך את שתיהן.
 *
 * מה שנמדד כאן הוא בדיוק זה: פעולה של הלשונית עוברת בשומר שהמעטפת מזריקה.
 */
import { describe, expect, it } from 'vitest';
import ShulchanTab from '../../src/ui/ribbon/tabs/ShulchanTab.vue';
import { createSuperdocDouble, mountUi, settle } from './harness';

/** מסמך עם בחירה חיה — בלעדיה חלק מהכלים אינם פעילים. */
const withSelection = () => createSuperdocDouble({ selection: { hasRange: true, text: 'טקסט נבחר' } });

describe('שולחן העורך — כל כלי עובר בשומר של הפעולה הכבדה', () => {
  it('לחיצה על כלי מריצה אותו בתוך השומר, ולא במקומו', async () => {
    const order: string[] = [];
    const harness = mountUi(ShulchanTab, {
      superdoc: withSelection(),
      heavyActionGuard: async (action) => {
        order.push('enter');
        const result = await action();
        order.push('exit');
        return result;
      },
    });
    await settle();

    // „תיקון העתקה” ולא הכפתור הראשון שנמצא: חלק מהפקדים פותחים דיאלוג
    // ואינם מריצים כלי, ובדיקה שנופלת על אחד מהם עוברת מהסיבה הלא נכונה.
    const button = harness.wrapper
      .findAll('button')
      .find((candidate) => candidate.text().includes('תיקון העתקה'));
    expect(button, 'הכפתור „תיקון העתקה” חייב להיות בלשונית').toBeDefined();

    await button!.trigger('click');
    await settle(6);

    // „enter” לפני „exit”, ושניהם קרו: פעולה שרצה **מחוץ** לשומר הייתה
    // משאירה מערך ריק, ופעולה שהשומר בלע הייתה משאירה רק „enter”.
    expect(order).toEqual(['enter', 'exit']);
  });

  it('בלי מעטפת — הלשונית עדיין עובדת', async () => {
    // ברירת המחדל של ה-inject קיימת בשביל בדיקות שמרכיבות לשונית לבדה.
    const harness = mountUi(ShulchanTab, { superdoc: withSelection() });
    await settle();

    const button = harness.wrapper
      .findAll('button')
      .find((candidate) => candidate.text().includes('תיקון העתקה'));
    await button!.trigger('click');
    await settle(6);

    expect(harness.wrapper.exists()).toBe(true);
  });

  it('מפעילה את הכלי על המסמך שנבחר לפני שהשומר ממתין', async () => {
    const source = withSelection();
    const other = withSelection();
    let entered!: () => void;
    let release!: () => void;
    const enteredGuard = new Promise<void>((resolve) => {
      entered = resolve;
    });
    const releaseGuard = new Promise<void>((resolve) => {
      release = resolve;
    });
    const harness = mountUi(ShulchanTab, {
      superdoc: source,
      heavyActionGuard: async (action) => {
        entered();
        await releaseGuard;
        return action();
      },
    });
    await settle();

    const button = harness.wrapper
      .findAll('button')
      .find((candidate) => candidate.text().includes('תיקון העתקה'));
    expect(button, 'הכפתור „תיקון העתקה” חייב להיות בלשונית').toBeDefined();

    const click = button!.trigger('click');
    await enteredGuard;
    await harness.setSuperdoc(other);
    // ה-watch של הלשונית קורא את סימני החיתוך במסמך שנעשה פעיל; מאפסים את
    // המדידות האלה כדי שהטענה למטה תמדוד רק את הכלי שהמתין בשומר.
    source.reset();
    other.reset();
    release();
    await click;
    await settle(6);

    // הצ׳קפוינט של המעטפת ממתין לפני הפעולה; בלי צילום המסמך ב-`runTool`
    // הלחיצה הייתה מפעילה את „תיקון העתקה” על `other`, הטאב שאליו עברנו.
    expect(source.ops()).toContain('blocks.list');
    expect(other.ops()).not.toContain('blocks.list');
  });
});
