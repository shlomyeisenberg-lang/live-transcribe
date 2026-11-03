/**
 * „דיווח על באג” — הכתובת, הגוף, והניקוי של השגיאות.
 *
 * מה שנבדק כאן הוא מה שיוצא מהמחשב של המשתמש, ולכן כל בדיקה כאן היא בדיקת
 * פרטיות לא פחות מבדיקת פורמט: מה **לא** בגוף חשוב בדיוק כמו מה שכן.
 */
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import {
  BUG_REPORT_NEW_ISSUE,
  MAX_ISSUE_URL_LENGTH,
  buildIssueBody,
  buildIssueUrl,
  collectEnvironment,
  issueTitle,
  type BugReportEnvironment,
} from '../../src/host/bug-report';
import {
  DIAGNOSTIC_INPUT_LIMIT,
  RECENT_ERRORS_LIMIT,
  clearRecentErrors,
  recentErrors,
  recordError,
  scrubDiagnostic,
} from '../../src/host/recent-errors';
import { setHostAppVersion } from '../../src/host/host-capabilities';

const ENV: BugReportEnvironment = {
  pluginVersion: '6.3.0',
  engineVersion: '2.23.0',
  appVersion: '0.9.97',
  userAgent: 'Mozilla/5.0 (Windows NT 10.0) Chrome/130',
};

function parse(url: string): { base: string; title: string; body: string } {
  const parsed = new URL(url);
  return {
    base: `${parsed.origin}${parsed.pathname}`,
    title: parsed.searchParams.get('title') ?? '',
    body: parsed.searchParams.get('body') ?? '',
  };
}

afterEach(() => {
  clearRecentErrors();
  setHostAppVersion(null);
});

describe('collectEnvironment', () => {
  it('הגרסאות הן אלה שבקבצי החבילה, וגרסת אוצריא מה-boot', () => {
    const manifest = JSON.parse(readFileSync(join(process.cwd(), 'public', 'manifest.json'), 'utf8')) as {
      version: string;
    };
    const pkg = JSON.parse(readFileSync(join(process.cwd(), 'package.json'), 'utf8')) as {
      dependencies: Record<string, string>;
    };
    setHostAppVersion('0.9.99');

    const env = collectEnvironment({ userAgent: 'UA' });
    expect(env).toEqual({
      pluginVersion: manifest.version,
      engineVersion: pkg.dependencies.superdoc,
      appVersion: '0.9.99',
      userAgent: 'UA',
    });
  });

  it('המניפסט מצהיר על app.open_url — בלעדיה המאחז דוחה את הפתיחה', () => {
    const manifest = JSON.parse(readFileSync(join(process.cwd(), 'public', 'manifest.json'), 'utf8')) as {
      permissions: string[];
    };
    expect(manifest.permissions).toContain('app.open_url');
  });
});

describe('buildIssueUrl', () => {
  it('כתובת issues/new של מאגר התוסף, עם הכותרת והגוף כפי שנכתבו', () => {
    const { url, truncated } = buildIssueUrl({
      title: 'הסמן קופץ',
      description: 'הדבקתי & לחצתי Enter\nשורה שנייה + 100%',
      environment: ENV,
      errors: [],
    });
    const { base, title, body } = parse(url);

    expect(base).toBe(BUG_REPORT_NEW_ISSUE);
    expect(BUG_REPORT_NEW_ISSUE).toBe('https://github.com/Y-PLONI/otzaria-word-editor/issues/new');
    expect(title).toBe('הסמן קופץ');
    // תווים שמשמעותם שונה ב-query string חוזרים כמו שהם.
    expect(body).toContain('הדבקתי & לחצתי Enter\nשורה שנייה + 100%');
    expect(truncated).toBe(false);
  });

  it('הגוף מכיל את פרטי הסביבה, ולא שום דבר שלא נמסר לו', () => {
    const body = buildIssueBody({ description: 'תיאור', environment: ENV, errors: [] });
    expect(body).toContain('| גרסת התוסף | 6.3.0 |');
    expect(body).toContain('| מנוע SuperDoc | 2.23.0 |');
    expect(body).toContain('| גרסת אוצריא | 0.9.97 |');
    expect(body).toContain(ENV.userAgent);
    expect(body).not.toContain('### שגיאות אחרונות');
  });

  it('גרסת אוצריא חסרה מוצגת כ„לא ידועה” ולא כ-null', () => {
    const body = buildIssueBody({ description: '', environment: { ...ENV, appVersion: null }, errors: [] });
    expect(body).toContain('| גרסת אוצריא | לא ידועה |');
    expect(body).not.toContain('null');
  });

  it('השגיאות נכנסות רק כשנמסרו', () => {
    const body = buildIssueBody({
      description: '',
      environment: ENV,
      errors: [{ time: '10:00:00', message: 'TypeError: x is undefined', count: 3 }],
    });
    expect(body).toContain('### שגיאות אחרונות');
    expect(body).toContain('10:00:00 TypeError: x is undefined (×3)');
  });

  it('דיווח ארוך מקוצר מתחת לגבול; הטקסט המלא נשמר להעתקה', () => {
    const description = 'א'.repeat(3000);
    const errors = Array.from({ length: 8 }, (_, i) => ({ time: '10:00:00', message: `שגיאה ${i} ${'ב'.repeat(200)}`, count: 1 }));
    const link = buildIssueUrl({ title: 'ארוך', description, environment: ENV, errors });

    expect(link.truncated).toBe(true);
    expect(link.url.length).toBeLessThanOrEqual(MAX_ISSUE_URL_LENGTH);
    expect(link.fullText).toContain(description);
    const { body } = parse(link.url);
    // הפרטים הטכניים שורדים את הקיצור.
    expect(body).toContain('| גרסת התוסף | 6.3.0 |');
    expect(body).toContain('קוצר');
    // ממצה את הגבול: לא קיצור גס מדי.
    expect(link.url.length).toBeGreaterThan(MAX_ISSUE_URL_LENGTH - 200);
  });

  it('הקיצור מוותר קודם על השגיאות הישנות', () => {
    const errors = Array.from({ length: 8 }, (_, i) => ({ time: '10:00:00', message: `E${i} ${'ג'.repeat(120)}`, count: 1 }));
    const link = buildIssueUrl({ title: 'ת', description: 'קצר', environment: ENV, errors });
    const { body } = parse(link.url);
    expect(link.truncated).toBe(true);
    expect(body).toContain('E7');
    expect(body).not.toContain('E0 ');
    expect(body).toContain('קצר');
  });
});

describe('issueTitle', () => {
  it('נושא ריק → השורה הראשונה של התיאור, ואם גם היא ריקה — כללי', () => {
    expect(issueTitle('  ', 'נתקע בשמירה\nפרטים')).toBe('נתקע בשמירה');
    expect(issueTitle('', '')).toBe('דיווח על באג');
    expect(issueTitle('נושא  עם   רווחים', 'x')).toBe('נושא עם רווחים');
  });
});

describe('recent-errors', () => {
  it('מנקה שם משתמש, מייל, תיקיית התקנה ושם מסמך', () => {
    const scrubbed = scrubDiagnostic(
      'failed C:\\Users\\yossi\\Documents\\x and /home/yossi/a, mail yossi.k@gmail.com ' +
        'at file:///C:/Users/yossi/AppData/Roaming/otzaria/plugins/word/assets/app.js:12:3 ' +
        'opening „חידושים על הפרשה.docx” and שיעור.docx',
    );
    expect(scrubbed).not.toContain('yossi');
    expect(scrubbed).not.toContain('חידושים');
    expect(scrubbed).toContain('file:///…/assets/app.js:12:3');
    expect(scrubbed).toContain('<מייל>');
    expect(scrubbed).not.toContain('שיעור');
    expect(scrubbed).toContain('„<מסמך>.docx”');
  });

  it('נתיב עם רווחים, בלי מירכאות — גם התיקיות וגם השם יורדים', () => {
    const scrubbed = [
      'open D:\\הספרים שלי\\חידושים על התורה.docx',
      'at C:\\Users\\yossi cohen\\Desktop\\x.js:3:4',
      'open \\\\server\\share\\Rabbi Name\\file.docx',
    ].map(scrubDiagnostic);
    for (const leak of ['הספרים', 'חידושים', 'התורה', 'yossi', 'cohen', 'Rabbi', 'server']) {
      expect(scrubbed.join(' ')).not.toContain(leak);
    }
    expect(scrubbed[0]).toBe('open <מסמך>.docx');
    // שם הסקריפט ומיקום השורה נשארים — הם מה שמאתר את התקלה.
    expect(scrubbed[1]).toContain('x.js:3:4');
  });

  it.each([
    'open /Users/yossi/Documents/חידושים על התורה.docx',
    'open /home/yossi/private/client projects/book.docx',
    'open /Users/yossi cohen/Documents/חידושים על התורה.docx',
    'at file:///Users/yossi cohen/Library/otzaria/assets/app.js:12:3',
    'open file:///Users/yossi%20cohen/private/חידושים%20על%20התורה.docx',
    'read /srv/private/client projects/book.bin',
  ])('conservatively removes complete private paths: %s', (input) => {
    const output = scrubDiagnostic(input);
    for (const privatePart of ['yossi', 'cohen', 'חידושים', 'התורה', 'private', 'client', 'projects']) {
      expect(output).not.toContain(privatePart);
    }
  });

  it('bounds input before processing; huge messages and argument lists stay bounded', () => {
    const inspected = scrubDiagnostic('x'.repeat(DIAGNOSTIC_INPUT_LIMIT * 100));
    expect(inspected.length).toBe(DIAGNOSTIC_INPUT_LIMIT + 1);
    const unreadable = { get constructor(): never { throw new Error('must not inspect argument 33'); } };
    recordError(...Array(32).fill(''), unreadable);
    expect(recentErrors()).toEqual([]);
    recordError('x'.repeat(1_000_000));
    expect(recentErrors()[0].message.length).toBeLessThanOrEqual(301);
  });

  it('אובייקט נרשם כשם הסוג בלבד — לא כתוכנו', () => {
    recordError('[otzaria-word] שגיאה:', { text: 'תוכן המסמך' }, new TypeError('boom'));
    const [entry] = recentErrors();
    expect(entry.message).toBe('[otzaria-word] שגיאה: [Object] TypeError: boom');
    expect(entry.message).not.toContain('תוכן המסמך');
  });

  it('אותה הודעה ברצף נספרת ולא משוכפלת; הרשימה תחומה', () => {
    recordError('אותו דבר');
    recordError('אותו דבר');
    expect(recentErrors()).toHaveLength(1);
    expect(recentErrors()[0].count).toBe(2);

    for (let i = 0; i < RECENT_ERRORS_LIMIT + 5; i += 1) recordError(`שגיאה ${i}`);
    const list = recentErrors();
    expect(list).toHaveLength(RECENT_ERRORS_LIMIT);
    expect(list[list.length - 1].message).toBe(`שגיאה ${RECENT_ERRORS_LIMIT + 4}`);
  });
});
