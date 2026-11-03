/**
 * „דיווח על באג” — בניית ה-issue ופתיחתו.
 *
 * ## המסלול
 *
 * הדיווח הוא issue חדש במאגר התוסף ב-GitHub, שנפתח בדפדפן של המשתמש עם כותרת
 * וגוף ממולאים מראש (`issues/new?title=…&body=…`). התוסף אינו שולח דבר בעצמו:
 * אין לו רשת, ואין לו חשבון. המשתמש רואה את הטופס של GitHub, יכול לערוך אותו,
 * ורק הלחיצה שלו שם יוצרת את ה-issue.
 *
 * הפתיחה עוברת ב-`app.openUrl` של אוצריא (SDK 0.9.95, הרשאה `app.open_url`):
 * זו הדרך המתועדת של ה-SDK לפתוח קישור בדפדפן ברירת המחדל של המערכת, והמאחז
 * מגביל אותה ל-http/https. כשהקריאה נכשלת (אין הרשאה, מחוץ לאוצריא) הדיווח מועתק ללוח, והמשתמש מקבל
 * את הכתובת להדביק אליה.
 *
 * ## מה נשלח
 *
 * רק מה שהמשתמש כתב, ופרטי סביבה: גרסת התוסף, גרסת SuperDoc, גרסת אוצריא,
 * ה-user agent, ובאישורו — השגיאות האחרונות מהקונסול (מנוקות, ר'
 * `recent-errors.ts`). אין קריאה של המסמך; הודעות שגיאה עשויות להכיל מידע פרטי ודורשות בדיקה. אין כאן
 * מייל — אף שאוצריא מציעה `app.getUserEmail` בדיוק לדיווחי שגיאות, issue
 * ב-GitHub הוא ציבורי.
 */
import { version as manifestVersion } from '../../public/manifest.json';
import { dependencies } from '../../package.json';
import { call } from './otzaria-client';
import { hostAppVersion } from './host-capabilities';
import type { RecentError } from './recent-errors';

/** המאגר שה-issues שלו הם ערוץ הדיווח. */
export const BUG_REPORT_REPO = 'https://github.com/Y-PLONI/otzaria-word-editor';
export const BUG_REPORT_NEW_ISSUE = `${BUG_REPORT_REPO}/issues/new`;

/**
 * אורך מרבי לכתובת כולה.
 *
 * GitHub מחזיר 414 על כתובת issue ארוכה מדי (סביב 8KB), ועברית מקודדת
 * בשישה תווים לאות (`%D7%90`). 6000 משאיר מרווח, וזה כ-900 אותיות עבריות.
 * דיווח ארוך יותר מקוצר בכתובת, והמלא מועתק ללוח — ר' `buildIssueUrl`.
 */
export const MAX_ISSUE_URL_LENGTH = 6000;

export interface BugReportEnvironment {
  pluginVersion: string;
  engineVersion: string;
  /** `null` מחוץ לאוצריא, או כשה-boot לא דיווח. */
  appVersion: string | null;
  userAgent: string;
}

export interface BugReportInput {
  title: string;
  description: string;
  environment: BugReportEnvironment;
  /** ריק כשהמשתמש ביטל את צירוף השגיאות. */
  errors: readonly RecentError[];
}

/** הסביבה הנוכחית. `navigator` מוזרק לבדיקות. */
export function collectEnvironment(nav: Pick<Navigator, 'userAgent'> = navigator): BugReportEnvironment {
  return {
    pluginVersion: manifestVersion,
    engineVersion: dependencies.superdoc,
    appVersion: hostAppVersion(),
    userAgent: nav.userAgent,
  };
}

/** סימון `` ` `` בתוך תא טבלה שובר אותה; `|` מפצל אותה לעמודות. */
function cell(value: string): string {
  return value.replace(/\|/g, '\\|').replace(/`/g, "'");
}

function environmentTable(env: BugReportEnvironment): string {
  return [
    '| | |',
    '|---|---|',
    `| גרסת התוסף | ${cell(env.pluginVersion)} |`,
    `| מנוע SuperDoc | ${cell(env.engineVersion)} |`,
    `| גרסת אוצריא | ${cell(env.appVersion ?? 'לא ידועה')} |`,
    `| דפדפן | ${cell(env.userAgent)} |`,
  ].join('\n');
}

function errorsBlock(errors: readonly RecentError[]): string {
  const lines = errors.map(
    (entry) => `${entry.time} ${entry.message}${entry.count > 1 ? ` (×${entry.count})` : ''}`,
  );
  // גדר של ארבעה גרשים: הודעת שגיאה עשויה להכיל שלושה.
  return ['````', ...lines, '````'].join('\n');
}

/** הגוף המלא, ב-Markdown של GitHub. */
export function buildIssueBody(input: Omit<BugReportInput, 'title'>, options: { truncated?: boolean } = {}): string {
  const description = input.description.trim();
  const sections = [
    '### מה קרה',
    description || '_לא נכתב תיאור._',
  ];
  if (options.truncated) {
    sections.push('_(הדיווח קוצר כדי להיכנס לכתובת. הנוסח המלא זמין בחלון הדיווח בעורך להעתקה.)_');
  }
  sections.push('', '### פרטים טכניים', environmentTable(input.environment));
  if (input.errors.length > 0) {
    sections.push('', '### שגיאות אחרונות', errorsBlock(input.errors));
  }
  sections.push('', '<sub>נשלח מתוך „דיווח על באג” בוורד לאוצריא. לא בוצעה קריאה של המסמך לצורך הדיווח.</sub>');
  return sections.join('\n');
}

/** כותרת ה-issue. ריקה → השורה הראשונה של התיאור, ואם גם הוא ריק — כללית. */
export function issueTitle(title: string, description: string): string {
  const own = title.replace(/\s+/g, ' ').trim();
  if (own) return own.slice(0, 120);
  const firstLine = description.trim().split('\n')[0]?.trim() ?? '';
  if (firstLine) return firstLine.length > 80 ? `${firstLine.slice(0, 80)}…` : firstLine;
  return 'דיווח על באג';
}

function urlFor(title: string, body: string): string {
  const query = new URLSearchParams({ title, body });
  // URLSearchParams מקודד רווח כ-`+`, ו-GitHub מפענח אותו נכון; אבל `%20`
  // שורד גם העתקה ידנית של הכתובת לכל מקום אחר.
  return `${BUG_REPORT_NEW_ISSUE}?${query.toString().replace(/\+/g, '%20')}`;
}

export interface IssueLink {
  url: string;
  /** הגוף המלא, בלי קיצור — מה שמועתק ללוח. */
  fullText: string;
  /** האם הגוף שבכתובת קוצר. */
  truncated: boolean;
}

/**
 * הכתובת, בגבול האורך.
 *
 * הסדר שבו מוותרים: קודם השגיאות הישנות (החדשה ביותר היא בדרך כלל זו
 * שמעניינת), ואז זנב התיאור. הפרטים הטכניים נשארים תמיד — בלעדיהם הדיווח
 * חסר ערך, והם קצרים.
 */
export function buildIssueUrl(input: BugReportInput): IssueLink {
  const title = issueTitle(input.title, input.description);
  const fullBody = buildIssueBody(input);
  const fullText = `${title}\n\n${fullBody}`;

  const full = urlFor(title, fullBody);
  if (full.length <= MAX_ISSUE_URL_LENGTH) return { url: full, fullText, truncated: false };

  let errors = [...input.errors];
  while (errors.length > 0) {
    errors = errors.slice(1);
    const url = urlFor(title, buildIssueBody({ ...input, errors }, { truncated: true }));
    if (url.length <= MAX_ISSUE_URL_LENGTH) return { url, fullText, truncated: true };
  }

  // חיפוש בינארי על אורך התיאור: הקידוד אינו ליניארי (עברית מול לטינית).
  const description = input.description.trim();
  let low = 0;
  let high = description.length;
  let best = urlFor(title, buildIssueBody({ ...input, errors: [], description: '' }, { truncated: true }));
  while (low <= high) {
    const mid = Math.floor((low + high) / 2);
    const url = urlFor(
      title,
      buildIssueBody({ ...input, errors: [], description: `${description.slice(0, mid)}…` }, { truncated: true }),
    );
    if (url.length <= MAX_ISSUE_URL_LENGTH) {
      best = url;
      low = mid + 1;
    } else {
      high = mid - 1;
    }
  }
  return { url: best, fullText, truncated: true };
}

/** פתיחה בדפדפן דרך אוצריא. זורקת כשהמאחז סירב או אינו קיים. */
export async function openInBrowser(url: string): Promise<void> {
  if (await call<boolean>('app.openUrl', { url }) !== true) {
    throw new Error('Opening the report was declined');
  }
}

/**
 * העתקה ללוח. אין גיבוי של `execCommand('copy')`: הוא עריכת DOM, שהמאגר
 * אוסר (tests/unit/engine-boundaries.test.ts), וכשל כאן אינו מבוי סתום —
 * הדיאלוג מציג את הטקסט לבחירה ידנית.
 */
export async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    return false;
  }
}
