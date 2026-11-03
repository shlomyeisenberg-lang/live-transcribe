import { describe, expect, it, vi } from 'vitest';
import HomeTab from '../../src/ui/ribbon/tabs/HomeTab.vue';
import { autoUnmount, mountUi, settle } from './harness';

const preference = vi.hoisted(() => ({ remember: vi.fn(async () => false) }));
vi.mock('../../src/composables/new-document-defaults', () => ({
  templateHasStyle: async () => true,
  rememberStyleDefault: preference.remember,
  UNSUPPORTED_STYLE_DEFAULTS_MESSAGE: 'עיצוב הסגנונות אינו נתמך במסמכים חדשים בלי הבדל בין העורך ל-Word; ההעדפה לא נשמרה',
}));
autoUnmount();

describe('remembering an unchanged imported style', () => {
  it('shows an unsupported-default refusal in the still-open dialog and reports no saved success', async () => {
    const harness = mountUi(HomeTab, { styleGallery: {
      items: [{ id: 'Normal', label: 'רגיל', previewText: 'AaBbCc', previewStyle: {}, custom: false }],
      activeId: 'Normal', fromDocument: true,
    } });
    await settle();
    await harness.wrapper.find('.style-card').trigger('contextmenu', { clientX: 500, clientY: 100 });
    await settle();
    document.querySelector<HTMLButtonElement>('[data-action="modify"]')!.click();
    await settle(4);
    const checkbox = document.querySelector<HTMLInputElement>('[data-testid="style-new-documents"]')!;
    expect(checkbox).not.toBeNull();
    checkbox.checked = true;
    checkbox.dispatchEvent(new Event('change', { bubbles: true }));
    await settle();
    document.querySelector<HTMLButtonElement>('[data-testid="style-dialog"] .sd-btn-primary')!.click();
    await settle(4);
    expect(preference.remember).toHaveBeenCalledOnce();
    expect(harness.styleEdits).toHaveLength(0);
    expect(document.querySelector('[data-testid="style-dialog"]')?.textContent).toContain('ההעדפה לא נשמרה');
    expect(harness.failures().some(item => item.commandId === 'style-new-documents')).toBe(true);
  });
});
