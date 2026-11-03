/**
 * דיאלוג התיקון האוטומטי: עריכה על עותק, ושמירה רק באישור.
 */
import { describe, expect, it } from 'vitest';
import { mount } from '@vue/test-utils';
import AutoCorrectDialog from '../../src/ui/panels/AutoCorrectDialog.vue';
import {
  BUILT_IN_REPLACEMENTS,
  DEFAULT_AUTOCORRECT,
  type AutoCorrectOptions,
} from '../../src/engine/autocorrect';

/**
 * הסמלים המובנים בלבד. הרשימה שנשלחת עם המוצר (`src/data/autocorrect-word.tsv`) היא
 * נתונים שמתחלפים, ובדיקה שמניחה מה יש בה — שאין בה `teh`, כמה שורות — נשברת
 * כשמחליפים אותה.
 */
const BASE: AutoCorrectOptions = { ...DEFAULT_AUTOCORRECT, replacements: BUILT_IN_REPLACEMENTS };
const DEFAULT_REPLACEMENTS = BUILT_IN_REPLACEMENTS;

function open(options: AutoCorrectOptions = BASE) {
  return mount(AutoCorrectDialog, {
    props: { isOpen: true, options },
    attachTo: document.body,
  });
}

const dialog = () => document.body.querySelector('.acorr-dialog') as HTMLElement;
const inputs = () => [...dialog().querySelectorAll<HTMLInputElement>('.acorr-input')];
const buttonByText = (text: string) =>
  [...dialog().querySelectorAll<HTMLButtonElement>('.acorr-btn')].find((b) => b.textContent?.trim() === text)!;
const rows = () => [...dialog().querySelectorAll('.acorr-row')];

async function typeInto(input: HTMLInputElement, value: string) {
  input.value = value;
  input.dispatchEvent(new Event('input', { bubbles: true }));
  await Promise.resolve();
}

describe('AutoCorrectDialog', () => {
  it('מציג את הטבלה הקיימת ומתג לכל סוג תיקון', () => {
    const wrapper = open();

    expect(rows()).toHaveLength(DEFAULT_REPLACEMENTS.length);
    expect(dialog().querySelectorAll('.acorr-check input[type="checkbox"]').length).toBeGreaterThanOrEqual(9);
    wrapper.unmount();
  });

  it('הוספת רשומה: מופיעה בטבלה, והכפתור הופך ל„עדכן” כשהדפוס קיים', async () => {
    const wrapper = open();
    const [from, to] = inputs();

    await typeInto(from, 'teh');
    await typeInto(to, 'the');
    expect(buttonByText('הוסף')).toBeTruthy();
    buttonByText('הוסף').click();
    await wrapper.vm.$nextTick();

    expect(rows()).toHaveLength(DEFAULT_REPLACEMENTS.length + 1);
    expect(buttonByText('עדכן')).toBeTruthy();
    wrapper.unmount();
  });

  it('דפוס שאינו תקין מציג הודעה ואינו מוסיף', async () => {
    const wrapper = open();
    const [from] = inputs();

    await typeInto(from, ' x');
    buttonByText('הוסף').click();
    await wrapper.vm.$nextTick();

    expect(dialog().querySelector('.acorr-message')).not.toBeNull();
    expect(rows()).toHaveLength(DEFAULT_REPLACEMENTS.length);
    wrapper.unmount();
  });

  it('שורת ההסבר: מיד, בתו שאחריו כשדפוס ארוך מתחיל בו, או בסוף המילה', async () => {
    const wrapper = open();
    const [from] = inputs();
    const hint = () => dialog().querySelector('.acorr-timing')?.textContent ?? '';

    await typeInto(from, '/');
    await wrapper.vm.$nextTick();
    expect(hint()).toContain('מיד');
    await typeInto(from, '--');
    await wrapper.vm.$nextTick();
    expect(hint()).toContain(`„${String.fromCodePoint(0x2066)}-->${String.fromCodePoint(0x2069)}” מתחיל בו`);
    await typeInto(from, 'teh');
    await wrapper.vm.$nextTick();
    expect(hint()).toContain('כשהמילה תסתיים');
    wrapper.unmount();
  });

  it('מחיקה: בחירת שורה ואז „מחק” מסירה אותה', async () => {
    const wrapper = open();

    (rows()[0] as HTMLButtonElement).click();
    await wrapper.vm.$nextTick();
    buttonByText('מחק').click();
    await wrapper.vm.$nextTick();

    expect(rows()).toHaveLength(DEFAULT_REPLACEMENTS.length - 1);
    wrapper.unmount();
  });

  it('אישור פולט עותק מעודכן, ואינו משנה את ההגדרות שהתקבלו', async () => {
    const wrapper = open();
    const [from, to] = inputs();
    await typeInto(from, 'teh');
    await typeInto(to, 'the');
    buttonByText('הוסף').click();
    await wrapper.vm.$nextTick();

    buttonByText('אישור').click();
    await wrapper.vm.$nextTick();

    const emitted = wrapper.emitted('apply')![0][0] as AutoCorrectOptions;
    expect(emitted.replacements).toContainEqual({ from: 'teh', to: 'the' });
    expect(BASE.replacements).toBe(BUILT_IN_REPLACEMENTS);
    wrapper.unmount();
  });

  it('ביטול סוגר בלי להחיל', async () => {
    const wrapper = open();

    buttonByText('ביטול').click();
    await wrapper.vm.$nextTick();

    expect(wrapper.emitted('apply')).toBeUndefined();
    expect(wrapper.emitted('close')).toBeTruthy();
    wrapper.unmount();
  });

  it('רשימת הקיצורים נטענת רק בבחירה מפורשת ונשמרת רק באישור', async () => {
    const wrapper = open({ ...BASE, replacements: [...BASE.replacements, { from: 'למד', to: 'בחירה אישית' }] });
    expect(rows().some((row) => row.textContent?.includes('טס'))).toBe(false);
    buttonByText('טען קיצורים בעברית').click();
    await wrapper.vm.$nextTick();
    expect(rows().some((row) => row.textContent?.includes('טס'))).toBe(true);
    expect(wrapper.emitted('apply')).toBeUndefined();
    buttonByText('אישור').click();
    const options = wrapper.emitted('apply')![0][0] as AutoCorrectOptions;
    expect(options.replacements).toContainEqual({ from: 'למד', to: 'בחירה אישית' });
    wrapper.unmount();
  });

  it('מתג ראשי כבוי מנטרל את שתי הקבוצות', async () => {
    const wrapper = open({ ...BASE, enabled: false });

    const groups = [...dialog().querySelectorAll<HTMLFieldSetElement>('fieldset')];
    expect(groups.length).toBe(2);
    expect(groups.every((group) => group.disabled)).toBe(true);
    wrapper.unmount();
  });

  it('פתיחה מחדש מאפסת טיוטה שלא אושרה', async () => {
    const wrapper = open();
    const [from] = inputs();
    await typeInto(from, 'abc');

    await wrapper.setProps({ isOpen: false });
    await wrapper.setProps({ isOpen: true });

    expect(inputs()[0].value).toBe('');
    wrapper.unmount();
  });
});
