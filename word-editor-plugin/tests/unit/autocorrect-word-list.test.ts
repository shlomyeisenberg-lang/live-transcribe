/**
 * הרשימה שיוצאה מ-Word, כפי שהיא בקובץ שנשלח עם המוצר.
 *
 * כל שורה בקובץ צריכה להיכנס לטבלה. שורה שנזרקת בשקט — בלי טאב, פגומה, או כפולה —
 * היא רשומה שהמשתמש חשב שיש לו ואין, ולכן הבדיקה נכשלת עליה ואומרת איזו. אותו
 * רעיון כמו `acronyms.test.ts`: החלפת נתונים שמזיזה מספרים צריכה להיראות, לא
 * לעבור בשקט.
 */
import { describe, expect, it } from 'vitest';
import wordListRaw from '../../src/data/autocorrect-word.tsv?raw';
import { DEFAULT_AUTOCORRECT, WORD_REPLACEMENTS, parseWordList, planAutoCorrect } from '../../src/engine/autocorrect';

describe('src/data/autocorrect-word.tsv', () => {
  const { entries, skipped } = parseWordList(wordListRaw);

  it('כל שורה בקובץ נכנסת לטבלה — אין שורות פגומות או כפולות', () => {
    const rejected = wordListRaw
      .split(/\r?\n/)
      .filter((line) => line.trim() !== '' && !line.startsWith('#'))
      .filter((line) => {
        const tab = line.indexOf('\t');
        return tab < 0 || parseWordList(line).entries.length === 0;
      });

    expect(rejected, `שורות שנדחו: ${JSON.stringify(rejected.slice(0, 5))}`).toEqual([]);
    expect(skipped).toBe(0);
  });

  it('כל רשומה בקובץ זמינה ברשימה האישית שניתן לטעון', () => {
    const byFrom = new Map(WORD_REPLACEMENTS.map((entry) => [entry.from, entry.to]));

    for (const entry of entries) expect(byFrom.get(entry.from), entry.from).toBe(entry.to);
  });
  it('כל כלל מילה באמת מתקן כשמקלידים אותו ורווח — אין כללים שלא יפעלו לעולם', () => {
    const options = {
      ...DEFAULT_AUTOCORRECT,
      enabled: true,
      replaceAsYouType: true,
      replacements: entries,
      capitalizeSentences: false,
      capitalizeDays: false,
      twoInitialCaps: false,
    };
    const dead = entries
      .filter((entry) => entry.from !== entry.to)
      .filter((entry) => !planAutoCorrect(`${entry.from} `, options, {}).some((fix) => fix.replacement === entry.to))
      .map((entry) => entry.from);

    expect(dead).toEqual([]);
  });

  it('בנתהמ משמר את ה-ב', () => {
    expect(new Map(entries.map((entry) => [entry.from, entry.to])).get('בנתהמ')).toBe('בנתה"מ');
  });

  it('מילים עבריות רגילות אינן מוחלפות', () => {
    expect(planAutoCorrect('הילד למד תורה והמטוס טס ליד סל עם עש ', DEFAULT_AUTOCORRECT)).toEqual([]);
  });
});
