/**
 * כפתור האפקט התלת-מצבי, בפני עצמו.
 *
 * למה קובץ נפרד מבדיקת „גופן מתקדם”: שם נמדד מה שיוצא ב-patch, וכאן נמדד
 * החוזה של הפקד — מה שכל צרכן שני יסתמך עליו. שלושת הדברים שאסור שיישברו
 * בשקט:
 *
 *   1. **הסדר.** „ללא שינוי” ← „דלוק” ← „כבוי” ← „ללא שינוי”. מחזור שמתחיל
 *      מ„כבוי” הוא פקד שמסיר עיצוב בלחיצה אחת מיותרת.
 *   2. **ההכרזה.** `aria-pressed="mixed"` הוא מה שאומר „אינו לחוץ ואינו
 *      משוחרר”; בלעדיו מי שאינו רואה את הכפתור מקבל „לא לחוץ”, כלומר „כבוי”.
 *   3. **אין מוטציה של ה-prop.** הפקד פולט, וההורה מחליט — אחרת „נקה הכל”
 *      היה מנקה את המצב בהורה בלי שהכפתור ידע.
 */
import { describe, expect, it } from 'vitest';
import { mount } from '@vue/test-utils';
import TriToggle from '../../src/ui/panels/common/TriToggle.vue';

describe('TriToggle', () => {
  it('הלחיצות מחזוריות: ללא שינוי ← דלוק ← כבוי ← ללא שינוי', async () => {
    const wrapper = mount(TriToggle, { props: { modelValue: '', label: 'צל' } });

    await wrapper.trigger('click');
    expect(wrapper.emitted('update:modelValue')?.[0]).toEqual(['yes']);

    await wrapper.setProps({ modelValue: 'yes' });
    await wrapper.trigger('click');
    expect(wrapper.emitted('update:modelValue')?.[1]).toEqual(['no']);

    await wrapper.setProps({ modelValue: 'no' });
    await wrapper.trigger('click');
    expect(wrapper.emitted('update:modelValue')?.[2]).toEqual(['']);
  });

  it('שלושת המצבים מוכרזים, ולא רק נצבעים', async () => {
    const wrapper = mount(TriToggle, { props: { modelValue: '', label: 'חרוט' } });
    expect(wrapper.attributes('aria-pressed')).toBe('mixed');
    expect(wrapper.attributes('aria-label')).toBe('חרוט — ללא שינוי');

    await wrapper.setProps({ modelValue: 'yes' });
    expect(wrapper.attributes('aria-pressed')).toBe('true');
    expect(wrapper.attributes('aria-label')).toBe('חרוט — דלוק');

    await wrapper.setProps({ modelValue: 'no' });
    expect(wrapper.attributes('aria-pressed')).toBe('false');
    expect(wrapper.attributes('aria-label')).toBe('חרוט — כבוי');
  });

  /**
   * הסימן מצויר בכל שלושת המצבים — גם בנייטרלי. סימן שמופיע ונעלם היה משנה
   * את רוחב הכפתור באמצע מחזור, כלומר מזיז את השורה מתחת לאצבע.
   */
  it('הסימן קיים תמיד, ונבדל בין המצבים', async () => {
    const wrapper = mount(TriToggle, { props: { modelValue: '', label: 'צל' } });
    const mark = () => wrapper.find('.tri-mark').text();

    const neutral = mark();
    expect(neutral).not.toBe('');

    await wrapper.setProps({ modelValue: 'yes' });
    const on = mark();
    await wrapper.setProps({ modelValue: 'no' });
    const off = mark();

    expect(new Set([neutral, on, off]).size).toBe(3);
  });

  it('אינו משנה את ה-prop בעצמו — ההורה מחליט', async () => {
    const wrapper = mount(TriToggle, { props: { modelValue: '', label: 'צל' } });

    await wrapper.trigger('click');
    await wrapper.trigger('click');

    // ההורה לא עדכן, ולכן הפקד עדיין בנייטרלי — ושתי הפליטות זהות.
    expect(wrapper.attributes('aria-pressed')).toBe('mixed');
    expect(wrapper.emitted('update:modelValue')).toEqual([['yes'], ['yes']]);
  });
});
