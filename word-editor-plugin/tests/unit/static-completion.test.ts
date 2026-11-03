/**
 * השלמה מרשימות סטטיות: האינדקס (engine/static-completion.ts), הטעינה העצלה
 * (engine/static-completion-dictionary.ts), והנכס שהבנייה מייצרת.
 *
 * הבדיקה המרכזית כאן היא **גבול הרשומה**. המסלול המתבקש —
 * `buildSectionCache(entries.join('\n'))` — מאבד אותו: המטמון של הספר מחזיק
 * מערך שטוח של מילים, ולכן הצעה יכולה להתחיל בביטוי אחד ולהמשיך לזה שאחריו.
 * הקבוצה „גבול הרשומה על הנתונים האמיתיים” מריצה **כל** שאילתה אפשרית על שתי
 * הרשימות, ומודדת את שני המסלולים באותה לולאה: החדש 0 חציות, הקודם מאות.
 * בלי הצד השני, „אין חציות” הייתה טענה שאינה יכולה להיכשל.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { buildStaticIndex, matchStaticCompletion } from '../../src/engine/static-completion';
import { buildSectionCache, matchAtCursor } from '../../src/engine/book-completion';
import { loadStaticSources, resetStaticSources } from '../../src/engine/static-completion-dictionary';
import { STATIC_COMPLETION_GLOBAL } from '../../src/engine/static-completion-constants';
import {
  buildStaticCompletionAsset,
  readStaticCompletionSource,
} from '../../scripts/static-completion-asset';

/** שתי רשומות שחולקות את המילה הראשונה — בדיוק המקרה ששובר גבולות. */
const PHRASES = ['אבן השתיה', 'אבן שאין לה הופכין', 'אבן שואבת'];
const AUTHORS = ['אבן גבאי, מאיר בן יחזקאל', 'אבן עזרא'];

const phrases = buildStaticIndex('talmudic-phrases', PHRASES);
const authors = buildStaticIndex('authors', AUTHORS);
const sources = [phrases, authors];

const match = (precedingWords: string[], partialWord: string) =>
  matchStaticCompletion(sources, { precedingWords, partialWord });

describe('matchStaticCompletion', () => {
  it('משלימה מילה חלקית בלי הקשר, משתי אותיות', () => {
    expect(match([], 'אב')?.text).toBe('אבן השתיה');
  });

  it('אות אחת בלי הקשר אינה מספיקה', () => {
    expect(match([], 'א')).toBeNull();
  });

  it('הקשר מכוון את ההתאמה לרשומה הנכונה', () => {
    expect(match(['אבן'], 'שו')?.text).toBe('שואבת');
    expect(match(['אבן'], 'שא')?.text).toBe('שאין לה הופכין');
  });

  it('משלימה עד סוף הרשומה, לא חמש מילים', () => {
    expect(match(['אבן'], 'שאין')?.text).toBe('שאין לה הופכין');
  });

  it('פיסוק במקור נשמר', () => {
    expect(match(['אבן'], 'גב')?.text).toBe('גבאי, מאיר בן יחזקאל');
  });

  /**
   * הקצה נמשך עד סוף הטקסט ולא עד סוף המילה האחרונה. נמדד: „בן יוסף הכ”
   * החזיר „הכהן (בעל הקצות” — סוגריים לא מאוזנים אל תוך ה-DOCX.
   */
  it('סוגר שאחרי המילה האחרונה אינו נחתך', () => {
    const withParens = buildStaticIndex('authors', ['אריה ליב בן יוסף הכהן (בעל הקצות)']);
    expect(matchStaticCompletion([withParens], { precedingWords: ['יוסף'], partialWord: 'הכ' })?.text).toBe(
      'הכהן (בעל הקצות)',
    );
  });

  it('אין הצעה עם סוגריים לא מאוזנים באף מקום בנתונים', () => {
    const { phrases: p, authors: a } = readStaticCompletionSource();
    const unbalanced: string[] = [];

    for (const [name, entries] of [['talmudic-phrases', p], ['authors', a]] as [string, string[]][]) {
      const index = buildStaticIndex(name, entries);
      for (const entry of entries) {
        const words = entry.trim().split(/\s+/);
        for (let i = 0; i < words.length; i += 1) {
          const found = matchStaticCompletion([index], {
            precedingWords: words.slice(Math.max(0, i - 2), i),
            partialWord: words[i]!.slice(0, 2),
          });
          if (!found) continue;
          const opens = (found.text.match(/\(/g) ?? []).length;
          const closes = (found.text.match(/\)/g) ?? []).length;
          if (opens !== closes) unbalanced.push(found.text);
        }
      }
    }

    expect(unbalanced).toEqual([]);
  });

  it('הביטויים נבדקים לפני המחברים, באותו אורך הקשר', () => {
    // „אבן ה…” קיים רק בביטויים, „אבן ג…” רק במחברים.
    expect(match(['אבן'], 'ה')?.source).toBe('talmudic-phrases');
    expect(match(['אבן'], 'גב')?.source).toBe('authors');
  });

  /**
   * הסדר בין השניים: הקשר תואם גובר על מקור מועדף. נמדד בשער —
   * „אבן גב” החזיר את הביטוי „גבור הכובש את יצרו” (התאמה בלי הקשר, במקור
   * הראשון) במקום את המחבר „אבן גבאי” (התאמה עם הקשר, במקור השני).
   */
  it('התאמה עם הקשר במקור מאוחר גוברת על התאמה בלי הקשר במקור מוקדם', () => {
    const withDecoy = [
      buildStaticIndex('talmudic-phrases', [...PHRASES, 'גבור הכובש את יצרו']),
      authors,
    ];
    const found = matchStaticCompletion(withDecoy, { precedingWords: ['אבן'], partialWord: 'גב' });
    expect(found?.source).toBe('authors');
    expect(found?.text).toBe('גבאי, מאיר בן יחזקאל');
  });

  it('null כשאין רשומה תואמת', () => {
    expect(match(['זזזזז'], 'טט')).toBeNull();
    expect(match([], 'זזזזז')).toBeNull();
  });

  it('ניקוד וגרשיים שהוקלדו אינם מונעים התאמה', () => {
    expect(match(['אֶבֶן'], 'שו')?.text).toBe('שואבת');
  });

  /**
   * הגבול, על הצורה המדויקת ששוברת אותו: „אבן הש” אחרי חיבור הרשומות היה
   * יכול להחזיר „השתיה אבן שאין לה” — סוף הרשומה הראשונה מודבק לתחילת הבאה.
   */
  it('אינה חוצה מרשומה לרשומה', () => {
    const found = match(['אבן'], 'הש');
    expect(found?.text).toBe('השתיה');
    expect(found?.text).not.toContain('אבן');
  });
});

describe('גבול הרשומה על הנתונים האמיתיים', () => {
  const { phrases: realPhrases, authors: realAuthors } = readStaticCompletionSource();
  const LISTS: [string, string[]][] = [
    ['talmudic-phrases', realPhrases],
    ['authors', realAuthors],
  ];

  /**
   * כל שאילתה אפשרית על הנתונים: כל מילה בכל רשומה, עם עד שתי מילות ההקשר
   * שלפניה. תוצאה שחצתה גבול אינה רצף באף רשומה בודדת.
   */
  function countCrossings(entries: string[], probe: (context: string[], partial: string) => string | null): number {
    let crossings = 0;
    for (const entry of entries) {
      const words = entry.trim().split(/\s+/);
      for (let i = 0; i < words.length; i += 1) {
        const text = probe(words.slice(Math.max(0, i - 2), i), words[i]!.slice(0, 2));
        if (text !== null && !entries.some((candidate) => candidate.includes(text))) crossings += 1;
      }
    }
    return crossings;
  }

  it.each(LISTS)('%s: אף תוצאה אינה חוצה רשומה', (name, entries) => {
    const index = buildStaticIndex(name, entries);
    expect(
      countCrossings(
        entries,
        (precedingWords, partialWord) =>
          matchStaticCompletion([index], { precedingWords, partialWord })?.text ?? null,
      ),
    ).toBe(0);
  });

  /**
   * ואותה מדידה בדיוק על המסלול של PR ‏#46 — `buildSectionCache` על הרשומות
   * המחוברות ב-`\n`, ומנוע ההתאמה של הספר.
   *
   * בלי הבדיקה הזאת „אף תוצאה אינה חוצה” היא טענה שאינה יכולה להיכשל: באינדקס
   * החדש חציה אינה אפשרית מבנית. כאן נמדד שהבאג שהיא מגנה מפניו אמיתי, ומה
   * גודלו — כלומר גם המספר שמצוטט בתיעוד נשען על הרצה ולא על זיכרון.
   */
  it.each(LISTS)('%s: המסלול הקודם כן חצה, והרבה', (_name, entries) => {
    const cache = buildSectionCache(entries.join('\n'));
    const crossings = countCrossings(
      entries,
      (precedingWords, partialWord) =>
        matchAtCursor(cache, { precedingWords, partialWord }, { minStandalonePartial: 2 })?.text ?? null,
    );
    expect(crossings).toBeGreaterThan(100);
  });
});

describe('loadStaticSources', () => {
  const PACKED = { phrases: PHRASES, authors: AUTHORS };

  beforeEach(() => resetStaticSources());

  it('נטען פעם אחת בלבד, גם בקריאות מקבילות', async () => {
    let calls = 0;
    const loader = async () => {
      calls += 1;
      return PACKED;
    };

    const [first, second] = await Promise.all([loadStaticSources(loader), loadStaticSources(loader)]);
    expect(calls).toBe(1);
    expect(first).toBe(second);
  });

  it('כשל מחזיר null, וההמתנה מונעת ניסיון חוזר בכל השהיה', async () => {
    let calls = 0;
    const failing = async () => {
      calls += 1;
      return null;
    };

    expect(await loadStaticSources(failing)).toBeNull();
    expect(await loadStaticSources(failing)).toBeNull();
    expect(calls).toBe(1);
  });

  it('loader שזורק נחשב כשל ואינו מפיל את ההשלמה', async () => {
    const thrower = async () => {
      throw new Error('נכס פגום');
    };
    expect(await loadStaticSources(thrower)).toBeNull();
  });
});

describe('הנכס שהבנייה מייצרת', () => {
  const globals = globalThis as Record<string, unknown>;

  beforeEach(() => resetStaticSources());
  afterEach(() => {
    delete globals[STATIC_COMPLETION_GLOBAL];
  });

  it('נטען דרך המסלול האמיתי ומשלים ביטוי ומחבר', async () => {
    new Function('window', buildStaticCompletionAsset())(globals);

    const indexes = await loadStaticSources();
    expect(indexes).not.toBeNull();
    expect(indexes?.map((index) => index.name)).toEqual(['talmudic-phrases', 'authors']);

    const found = matchStaticCompletion(indexes!, { precedingWords: ['אבן'], partialWord: 'הש' });
    expect(found?.text).toBe('השתיה');
    expect(found?.source).toBe('talmudic-phrases');
  });

  it('הספירה מדויקת', () => {
    const { phrases: p, authors: a } = readStaticCompletionSource();
    expect(p.length).toBe(325);
    expect(a.length).toBe(651);
  });
});
