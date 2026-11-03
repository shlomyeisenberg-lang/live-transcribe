/**
 * בניית נכס רשימות ההשלמה הסטטיות — צד Node בלבד (vite.config.ts והבדיקות).
 * ההיגיון שצורך אותו ב-src/engine/static-completion.ts.
 *
 * האימות נעשה **בבניית האינדקס עצמו**, ולא בספירה מקבילה כאן: `buildStaticIndex`
 * סופר ריצות `WORD_INNER`, וספירה של `split(/\s+/)` אינה אותו דבר („אב-גד” הוא
 * טוקן אחד ושתי ריצות). שער שסופר אחרת מהאינדקס היה מכשיר רשומה שהאינדקס
 * זורק עליה בזמן ריצה. מעליו נבדקות כאן שתי טענות שהאינדקס אינו מתלונן
 * עליהן — רשומה ריקה וכפילות — כי נתונים שנערכים ביד הם בדיוק המקום שבו הן
 * נכנסות בלי שדבר ייכשל.
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { buildStaticIndex } from '../src/engine/static-completion';
import { STATIC_COMPLETION_GLOBAL } from '../src/engine/static-completion-constants';

/** מ-`cwd` ולא מ-`import.meta.url`: תחת jsdom הכתובת אינה file://. */
const SOURCES = {
  phrases: resolve('src/data/talmudic-phrases.json'),
  authors: resolve('src/data/authors.json'),
} as const;

function readList(name: string, path: string): string[] {
  const parsed = JSON.parse(readFileSync(path, 'utf8')) as unknown;
  if (!Array.isArray(parsed)) throw new Error(`${name}: הקובץ אינו מערך.`);

  const seen = new Set<string>();
  for (const entry of parsed) {
    if (typeof entry !== 'string' || entry.trim() === '') {
      throw new Error(`${name}: יש רשומה שאינה מחרוזת לא-ריקה (${JSON.stringify(entry)}).`);
    }
    if (seen.has(entry)) throw new Error(`${name}: הרשומה "${entry}" מופיעה יותר מפעם אחת.`);
    seen.add(entry);
  }

  const entries = parsed as string[];
  // זורק על רשומה שאינה נכנסת לאריזת המיקומים — אותו קוד שירוץ בזמן ריצה.
  const index = buildStaticIndex(name, entries);
  if (index.entries.length !== entries.length) {
    throw new Error(`${name}: ${entries.length - index.entries.length} רשומות אינן מכילות אף מילה.`);
  }

  return entries;
}

export function readStaticCompletionSource(): { phrases: string[]; authors: string[] } {
  return {
    phrases: readList('talmudic-phrases.json', SOURCES.phrases),
    authors: readList('authors.json', SOURCES.authors),
  };
}

/** תוכן `assets/static-completion.js`: השמה יחידה של ליטרל אובייקט ל-`window`. */
export function buildStaticCompletionAsset(): string {
  return `window.${STATIC_COMPLETION_GLOBAL} = ${JSON.stringify(readStaticCompletionSource())};\n`;
}
