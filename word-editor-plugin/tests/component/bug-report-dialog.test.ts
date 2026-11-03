/**
 * „דיווח על באג” — מה הדיאלוג שולח למאחז, ומה קורה כשהמאחז מסרב.
 *
 * הבניה של הכתובת עצמה נבדקת ב-tests/unit/bug-report.test.ts. כאן נבדק
 * החיבור: שהלחיצה מגיעה ל-`app.openUrl` עם מה שהמשתמש כתב, שסירוב של המאחז
 * מעתיק ללוח ומציג את הכתובת, ושהתצוגה המקדימה היא מה שנשלח.
 */
import { afterEach, describe, expect, it, vi } from 'vitest';
import { DOMWrapper } from '@vue/test-utils';
import BugReportDialog from '../../src/ui/panels/BugReportDialog.vue';
import FileTab from '../../src/ui/ribbon/tabs/FileTab.vue';
import { BUG_REPORT_NEW_ISSUE } from '../../src/host/bug-report';
import { clearRecentErrors, recordError } from '../../src/host/recent-errors';
import { autoUnmount, mountUi, settle, tipSelector } from './harness';

autoUnmount();

function teleported(selector: string): DOMWrapper<Element> {
  const element = document.querySelector(selector);
  if (!element) throw new Error(`לא נמצא ${selector} בגוף הדף`);
  return new DOMWrapper(element);
}

function footerButton(label: string): DOMWrapper<Element> {
  const found = [...document.querySelectorAll('.bug-dialog .br-btn')].find(
    (button) => button.textContent?.trim() === label,
  );
  if (!found) throw new Error(`לא נמצא הכפתור „${label}”`);
  return new DOMWrapper(found);
}

type Reply = { success: boolean; data: unknown; error: { code: string; message: string } | null };

/** מאחז מדומה שעונה על `app.openUrl` בלבד, ומקליט. */
function installHost(reply: Reply): Array<{ method: string; payload: Record<string, unknown> | undefined }> {
  const calls: Array<{ method: string; payload: Record<string, unknown> | undefined }> = [];
  (window as unknown as { Otzaria: unknown }).Otzaria = {
    call: async (method: string, payload?: Record<string, unknown>) => {
      calls.push({ method, payload });
      return reply;
    },
    on: () => {},
    off: () => {},
  };
  return calls;
}

function installClipboard(): ReturnType<typeof vi.fn> {
  const writeText = vi.fn(async () => {});
  Object.defineProperty(navigator, 'clipboard', { value: { writeText }, configurable: true });
  return writeText;
}

afterEach(() => {
  delete (window as unknown as { Otzaria?: unknown }).Otzaria;
  clearRecentErrors();
});

async function openAndFill(): Promise<void> {
  const harness = mountUi(BugReportDialog, { props: { isOpen: false } });
  await harness.wrapper.setProps({ isOpen: true });
  await settle();
  await teleported('#br-title-input').setValue('הסמן קופץ');
  await teleported('#br-desc-input').setValue('הדבקתי טקסט ולחצתי Enter');
  await settle();
}

describe('BugReportDialog', () => {
  it('פתיחה ממקדת את הנושא, והתצוגה המקדימה מראה מה ייצא', async () => {
    recordError('TypeError: boom');
    await openAndFill();

    expect(document.activeElement).toBe(document.querySelector('#br-title-input'));
    const preview = teleported('.br-preview-text').text();
    expect(preview).toContain('הסמן קופץ');
    expect(preview).toContain('הדבקתי טקסט ולחצתי Enter');
    expect(preview).not.toContain('TypeError: boom');
    expect((teleported('.br-check input').element as HTMLInputElement).checked).toBe(false);
    await teleported('.br-check input').setValue(true);
    expect(teleported('.br-preview-text').text()).toContain('TypeError: boom');
  });

  it('ביטול הסימון מוציא את השגיאות מהדיווח', async () => {
    recordError('TypeError: boom');
    await openAndFill();

    await teleported('.br-check input').setValue(true);
    await teleported('.br-check input').setValue(false);
    await settle();
    expect(teleported('.br-preview-text').text()).not.toContain('TypeError: boom');
  });

  it('„פתח דיווח” קורא ל-app.openUrl עם issue ממולא', async () => {
    const calls = installHost({ success: true, data: true, error: null });
    await openAndFill();

    await footerButton('פתח דיווח').trigger('click');
    await settle();

    expect(calls).toHaveLength(1);
    expect(calls[0].method).toBe('app.openUrl');
    const url = new URL(String(calls[0].payload?.url));
    expect(`${url.origin}${url.pathname}`).toBe(BUG_REPORT_NEW_ISSUE);
    expect(url.searchParams.get('title')).toBe('הסמן קופץ');
    expect(url.searchParams.get('body')).toContain('הדבקתי טקסט ולחצתי Enter');
    expect(teleported('.br-status').text()).toContain('הטופס נפתח בדפדפן');
  });

  it('סירוב של המאחז מעתיק ללוח ומציג את הכתובת', async () => {
    installHost({
      success: false,
      data: null,
      error: { code: 'error.permission_denied', message: 'permission_denied' },
    });
    const writeText = installClipboard();
    await openAndFill();

    await footerButton('פתח דיווח').trigger('click');
    await settle();

    expect(writeText).toHaveBeenCalledTimes(1);
    expect(String(writeText.mock.calls[0][0])).toContain('הדבקתי טקסט ולחצתי Enter');
    expect(teleported('.br-status').text()).toContain('הועתק ללוח');
    expect((teleported('.br-url').element as HTMLInputElement).value).toBe(BUG_REPORT_NEW_ISSUE);
  });

  it('גם הלוח חסום: הטקסט המלא מוצג לבחירה ידנית', async () => {
    installHost({ success: false, data: null, error: { code: 'error.forbidden', message: 'x' } });
    Object.defineProperty(navigator, 'clipboard', {
      value: { writeText: vi.fn(async () => { throw new Error('denied'); }) },
      configurable: true,
    });
    await openAndFill();

    await footerButton('פתח דיווח').trigger('click');
    await settle();

    expect(teleported('.br-status').text()).toContain('ההעתקה ללוח נכשלה');
    expect((teleported('.br-manual').element as HTMLTextAreaElement).value).toContain('הדבקתי טקסט ולחצתי Enter');
  });

  it('Enter על „מה יישלח” אינו שולח את הדיווח', async () => {
    const calls = installHost({ success: true, data: true, error: null });
    await openAndFill();

    await teleported('.br-preview summary').trigger('keydown', { key: 'Enter' });
    await settle();
    expect(calls).toHaveLength(0);

    // ובשדה הנושא Enter הוא „פתח דיווח”.
    await teleported('#br-title-input').trigger('keydown', { key: 'Enter' });
    await settle();
    expect(calls.map((call) => call.method)).toEqual(['app.openUrl']);
  });

  it('דיווח שקוצר מועתק ללוח לפני פתיחת הדפדפן, כשהדף עוד ממוקד', async () => {
    const order: string[] = [];
    (window as unknown as { Otzaria: unknown }).Otzaria = {
      call: async (method: string) => {
        order.push(method);
        return { success: true, data: true, error: null };
      },
      on: () => {},
      off: () => {},
    };
    const writeText = vi.fn(async () => {
      order.push('clipboard');
    });
    Object.defineProperty(navigator, 'clipboard', { value: { writeText }, configurable: true });
    const harness = mountUi(BugReportDialog, { props: { isOpen: false } });
    await harness.wrapper.setProps({ isOpen: true });
    await settle();
    await teleported('#br-desc-input').setValue('א'.repeat(3000));
    await settle();

    await footerButton('פתח דיווח').trigger('click');
    await settle();

    expect(order).toEqual(['clipboard', 'app.openUrl']);
    expect(teleported('.br-status').text()).toContain('הנוסח המלא הועתק ללוח');
  });

  it('דיווח שקוצר, והלוח נכשל — ההודעה אינה מבטיחה לוח', async () => {
    installHost({ success: true, data: true, error: null });
    Object.defineProperty(navigator, 'clipboard', {
      value: { writeText: vi.fn(async () => { throw new Error('Document is not focused'); }) },
      configurable: true,
    });
    const harness = mountUi(BugReportDialog, { props: { isOpen: false } });
    await harness.wrapper.setProps({ isOpen: true });
    await settle();
    await teleported('#br-desc-input').setValue('א'.repeat(3000));
    await settle();

    await footerButton('פתח דיווח').trigger('click');
    await settle();

    const status = teleported('.br-status').text();
    expect(status).toContain('הטופס נפתח בדפדפן');
    expect(status).not.toContain('הנוסח המלא הועתק ללוח');
    expect(status).toContain('נכשלה');
    expect((teleported('.br-manual').element as HTMLTextAreaElement).value).toContain('א'.repeat(3000));
    const writeText = installClipboard();
    await footerButton('העתק שוב').trigger('click');
    await settle();
    expect(String(writeText.mock.calls[0][0])).toContain('א'.repeat(3000));
    expect(teleported('.br-status').text()).toContain('הטופס נפתח בדפדפן');
    expect(teleported('.br-status').text()).toContain('הנוסח המלא הועתק ללוח');
  });

  it('host false is refusal: fallback copies and never claims the browser opened', async () => {
    installHost({ success: true, data: false, error: null });
    const writeText = installClipboard();
    await openAndFill();
    await footerButton('פתח דיווח').trigger('click');
    await settle();
    expect(writeText).toHaveBeenCalledTimes(1);
    expect(teleported('.br-status').text()).not.toContain('הטופס נפתח בדפדפן');
    expect(teleported('.br-status').text()).toContain('הועתק ללוח');
  });

  it('Tab stays inside the dialog and closing restores the opener', async () => {
    const opener = document.createElement('button');
    document.body.append(opener);
    opener.focus();
    const harness = mountUi(BugReportDialog, { props: { isOpen: false } });
    await harness.wrapper.setProps({ isOpen: true });
    await settle();
    const first = teleported('.br-close-btn');
    const last = footerButton('ביטול');
    (first.element as HTMLElement).focus();
    await first.trigger('keydown', { key: 'Tab', shiftKey: true });
    expect(document.activeElement).toBe(last.element);
    await last.trigger('keydown', { key: 'Tab' });
    expect(document.activeElement).toBe(first.element);
    await harness.wrapper.setProps({ isOpen: false });
    await settle();
    expect(document.activeElement).toBe(opener);
    opener.remove();
  });

  it('a pending host response cannot replace a newly reopened report', async () => {
    let finish!: (reply: Reply) => void;
    (window as unknown as { Otzaria: unknown }).Otzaria = {
      call: () => new Promise<Reply>((resolve) => { finish = resolve; }), on() {}, off() {},
    };
    const harness = mountUi(BugReportDialog, { props: { isOpen: false } });
    await harness.wrapper.setProps({ isOpen: true });
    await settle();
    await footerButton('פתח דיווח').trigger('click');
    await harness.wrapper.setProps({ isOpen: false });
    await harness.wrapper.setProps({ isOpen: true });
    await settle();
    await teleported('#br-desc-input').setValue('new report');
    finish({ success: true, data: true, error: null });
    await settle();
    expect((teleported('#br-desc-input').element as HTMLTextAreaElement).value).toBe('new report');
    expect(document.querySelector('.br-status')).toBeNull();
  });

  it('„סגור” ו-Escape סוגרים', async () => {
    const harness = mountUi(BugReportDialog, { props: { isOpen: true } });
    await settle();
    await teleported('.bug-dialog').trigger('keydown', { key: 'Escape' });
    expect(harness.wrapper.emitted('close')).toHaveLength(1);
  });
});

describe('FileTab — דיווח על באג', () => {
  it('הכפתור פולט report-bug', async () => {
    const harness = mountUi(FileTab);
    await settle();
    await harness.wrapper.find(tipSelector('דיווח על באג')).trigger('click');
    expect(harness.wrapper.emitted('report-bug')).toHaveLength(1);
  });
});
