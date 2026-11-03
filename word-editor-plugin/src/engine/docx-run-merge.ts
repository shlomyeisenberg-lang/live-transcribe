/**
 * איחוד ריצות שאין ביניהן שום הבדל, לפני שהמנוע רואה את המסמך.
 *
 * ## למה
 *
 * Word מפצל טקסט לריצות גם כשהעיצוב אחיד: כל שמירה מסמנת את מה שנכתב בה
 * במזהה עריכה משלה (`w:rsidR`, `w:rsidRPr`), וכל סימון של שגיאת כתיב
 * (`<w:proofErr/>`) חותך את הריצה שהוא יושב בה. מסמך שנערך לאורך זמן מגיע כך
 * לעשרות ריצות לפסקה, שכולן נראות אותו דבר.
 *
 * המנוע משלם על כל ריצה בכל הקשה. דווח (10.2026) על מסמך של משתמש — 42
 * פסקאות, 1,217 ריצות ו-270 `proofErr` — שבו הקלדה רצופה הייתה „איטית עד כדי
 * לא שמיש”: האותיות נערמו והופיעו יחד. ניסוי מקדים שאיחד גם דרך סימוני
 * האיות צמצם אותו ל-264 ריצות; כאן הסימונים נשמרים כגבולות, ולכן המספר
 * הסופי יכול להיות גבוה יותר. הסרת `proofErr` לבדה לא שינתה את התנהגות
 * ההקלדה בניסוי. מספרי הזמנים פנימיים — רישיון המנוע אוסר לפרסם מדידות שלו.
 *
 * ## מה מאוחד, ובדיוק מה
 *
 * שתי ריצות סמוכות מתאחדות רק כשכל התנאים מתקיימים:
 *
 * - שתיהן ריצות WordprocessingML (`wordPrefix`), לא `m:r` של נוסחה ולא `a:r`.
 * - בכל אחת יש לכל היותר `rPr` אחד, ו**בדיוק** `t` אחד, ושום ילד אחר. ריצה
 *   עם טאב, שבירה, הפניה להערה, `lastRenderedPageBreak` או טקסט מחוק אינה
 *   נוגעת — היא נשארת כמות שהיא, והשכנות שלה אינן מתאחדות דרכה.
 * - ה-`rPr` של שתיהן זהה **תו-בתו**. אין כאן השוואה סמנטית של עיצוב: מה ששונה
 *   במחרוזת נשאר נפרד, גם אם היה מוצג אותו דבר.
 * - המאפיינים על התג `r` עצמו הם מזהי עריכה בלבד (`rsid*`). כל מאפיין אחר
 *   עוצר.
 * - בין הריצות יש רק רווחים. `proofErr`, סימנייה, הערת XML, תחילת הערה
 *   (`commentRangeStart`), `ins`/`del` — כל דבר אחר באמצע עוצר.
 *
 * ## מה זה עושה לקובץ שיישמר
 *
 * המנוע מייצא מהמודל שלו, ולכן מזהי העריכה של הריצות שאוחדו אינם נשמרים
 * מלבד זה של הראשונה. הם משמשים את „השוואת מסמכים” של Word לשיוך שינויים
 * לשמירה; על מסמך שנפתח ונשמר בתוסף השיוך הזה ממילא אינו מדויק.
 *
 * `proofErr` נשמרים כגבול איחוד. מחיקת רק אחד מזוג start/end משאירה סימון
 * תלוש, ואין תמיד settings.xml קריא כדי להכריח בדיקת איות מחדש. איחוד
 * שכנים עם מזהי עריכה שונים עדיין מצמצם את הריצות בלי לשנות את גבולות האיות.
 *
 * ## רווחים
 *
 * `w:t` בלי `xml:space="preserve"` מאבד את הרווחים שבקצותיו — אלא אם אלמנט
 * אב הצהיר `preserve`, שהוא תכונה עוברת בירושה ב-XML. הריצה המאוחדת נכתבת
 * תמיד עם `preserve`, ולכן ריצה בלי `preserve` שיש רווח בקצה הטקסט שלה **אינה
 * מתאחדת**: קיצוץ היה מוחק רווח שאב שמר, והשארה הייתה מוסיפה רווח ש-Word לא
 * הציג. בקורפוס של 600 מסמכי Word המקרה לא הופיע אף פעם.
 *
 * ## סופי שורה
 *
 * XML מנרמל CR ו-CRLF ל-LF בכל צומת בנפרד. מנרמלים לפני חיבור הטקסטים,
 * כדי ש-CR בקצה ריצה ו-LF בתחילת הבאה יישארו שתי ירידות שורה. הפניות לתווים
 * כמו `&#13;` נשארות כמות שהן: הן מפוענחות אחרי נרמול סופי השורה.
 *
 * ## `]]>`
 *
 * `>` גלוי מותר בטקסט XML, אבל הרצף `]]>` אסור. ריצה שנגמרת ב-`]]` וריצה
 * שמתחילה ב-`>` יוצרות אותו בהדבקה, ולכן ה-`>` שבו נכתב כ-`&gt;`.
 */
import { skippedCloser, wordPrefix } from './docx-parts';

/** ריצה שאפשר לאחד: איפה היא, ומה שצריך כדי לכתוב אותה מחדש. */
interface SimpleRun {
  start: number;
  end: number;
  /** מאפייני התג `r`, כפי שנכתבו (כולל הרווח המוביל). */
  attributes: string;
  /** ה-`rPr` כולו, מחרוזת גולמית; ריק כשאין. */
  props: string;
  /** הטקסט כפי שהוא ב-XML (ישויות כמו `&amp;` נשארות ישויות). */
  text: string;
}

/** כולל אבות בלי קידומת, כדי לזהות גם הכרזת namespace מקומית עליהם. */
const RUN_TOKEN_SOURCE =
  /<!--|<!\[CDATA\[|<\?|<(\/?)(?:([\w.-]+):)?([\w.-]+)((?:[^>"']|"[^"]*"|'[^']*')*)>/;

/** שמות וערכי מאפיינים, בלי להתבלבל ב־xmlns שכתוב בתוך ערך מצוטט. */
const ATTRIBUTE_SOURCE = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

/** `xml:space="preserve"` בדיוק, או שום מאפיין. כל מאפיין אחר על `t` עוצר. */
const PRESERVE = /^\s+xml:space\s*=\s*(?:"preserve"|'preserve')\s*$/;

/**
 * רווח של XML (לא של Unicode) בקצה הטקסט — מה שנקצץ בלי preserve. גם כשהוא
 * כתוב כהפניית תו (`&#32;`, `&#x9;`): Word מקצץ גם אותו (ממצא QA, נבדק ב-Word).
 */
const SPACE_CHAR = String.raw`(?:[ \t\r\n]|&#(?:0*(?:32|9|10|13)|[xX]0*(?:20|9|[aA]|[dD]));)`;
const XML_EDGE_SPACE = new RegExp(`^${SPACE_CHAR}|${SPACE_CHAR}$`);

/**
 * `document.xml` (או הערות שוליים, כותרות…) עם ריצות זהות מאוחדות, או `null`
 * כשאין מה לאחד.
 */
export function mergeFragmentedRuns(xml: string): string | null {
  const prefix = wordPrefix(xml);
  if (prefix === null) return null;
  // חיפוש מחרוזת לפני הסריקה: חלק בלי שתי ריצות אינו יכול להתאחד.
  const runOpen = `<${prefix}:r`;
  const first = xml.indexOf(runOpen);
  if (first < 0 || xml.indexOf(runOpen, first + runOpen.length) < 0) return null;

  const between = /^\s*$/;
  const runs = collectSimpleRuns(xml, prefix);

  let out = '';
  let copiedTo = 0;
  let merged = 0;
  for (let i = 0; i < runs.length; ) {
    let j = i;
    while (
      j + 1 < runs.length &&
      runs[j + 1].props === runs[i].props &&
      between.test(xml.slice(runs[j].end, runs[j + 1].start))
    ) {
      j += 1;
    }
    if (j > i) {
      const group = runs.slice(i, j + 1);
      out += xml.slice(copiedTo, group[0].start);
      out +=
        `<${prefix}:r${group[0].attributes}>${group[0].props}` +
        `<${prefix}:t xml:space="preserve">${group
          .map((run) => run.text.replace(/\r\n?/g, '\n'))
          .join('')
          .split(']]>')
          .join(']]&gt;')}</${prefix}:t>` +
        `</${prefix}:r>`;
      copiedTo = group[group.length - 1].end;
      merged += group.length - 1;
    }
    i = j + 1;
  }
  if (merged === 0) return null;
  return out + xml.slice(copiedTo);
}

/** מספר הריצות שנחסכו, לשורת היומן. */
export function countRuns(xml: string): number {
  const prefix = wordPrefix(xml);
  if (prefix === null) return 0;
  return xml.match(new RegExp(`<${prefix}:r[\\s>]`, 'g'))?.length ?? 0;
}

/**
 * כל הריצות במסמך שמתאימות לאיחוד, לפי הסדר. ריצה שאינה מתאימה אינה נכנסת
 * לרשימה, ולכן הבדיקה של מה שבין שתי ריצות סמוכות ברשימה (`between`) נכשלת
 * עליה — היא יושבת ביניהן כתג ולא כרווח.
 */
function collectSimpleRuns(xml: string, prefix: string): SimpleRun[] {
  const runs: SimpleRun[] = [];
  const token = new RegExp(RUN_TOKEN_SOURCE.source, 'g');
  const escapedPrefix = prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const rsidOnly = new RegExp(`^(?:\\s+${escapedPrefix}:rsid\\w*\\s*=\\s*(?:"[^"]*"|'[^']*'))*\\s*$`);
  let rootSeen = false;

  /** הריצה הפתוחה כרגע, או `null` מחוץ לריצה. */
  let run: {
    start: number;
    attributes: string;
    ok: boolean;
    props: string | null;
    text: string | null;
    /** טקסט פתוח נקרא עד הטוקן הבא, בלי חיפוש חוזר בשאר המסמך. */
    textStart: number;
    preserved: boolean;
    /** עומק אלמנטים בתוך הריצה. 0 = ילדים ישירים. */
    depth: number;
    /** איפה התחיל ה-`rPr` הפתוח, כשהוא הילד הנוכחי. */
    propsStart: number;
    /** איפה נגמר התג הקודם — כדי לוודא שבין ילדים יש רק רווחים. */
    lastEnd: number;
  } | null = null;

  for (let match = token.exec(xml); match; match = token.exec(xml)) {
    const closer = skippedCloser(match[0]);
    if (closer !== undefined) {
      const end = xml.indexOf(closer, token.lastIndex);
      // הערה שאינה נסגרת: מה שאחריה אינו אמין. מה שכבר נאסף לפניה בטוח.
      if (end < 0) break;
      // הערה או CDATA בתוך ריצה: לא מאחדים אותה.
      if (run) run.ok = false;
      token.lastIndex = end + closer.length;
      continue;
    }

    const [, closing, tagPrefix, name, rawAttributes] = match;
    const selfClosing = rawAttributes.endsWith('/');
    const attributes = selfClosing ? rawAttributes.slice(0, -1) : rawAttributes;
    const ours = tagPrefix === prefix;

    // wordPrefix קורא את השורש בלבד. הכרזה חוזרת עמוק יותר יכולה לשנות גם
    // את משמעות ה-rPr: לא מאחדים שום ריצה בחלק כזה, אפילו לפני ההכרזה.
    if (!closing) {
      let declaredWordPrefix = false;
      for (const attribute of attributes.matchAll(ATTRIBUTE_SOURCE)) {
        if (attribute[1] !== `xmlns:${prefix}`) continue;
        if (rootSeen) return [];
        declaredWordPrefix = (attribute[2] ?? attribute[3]) ===
          'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      }
      // wordPrefix אינו מפרש מאפיינים. אימות נוסף מונע הכרזה מדומה בתוך ערך.
      if (!rootSeen && !declaredWordPrefix) return [];
      rootSeen = true;
    }

    if (run && run.textStart >= 0) {
      const raw = xml.slice(run.textStart, match.index);
      if (closing && ours && name === 't' && attributes.trim() === '' && !raw.includes('<')) {
        if (!run.preserved && XML_EDGE_SPACE.test(raw)) run.ok = false;
        run.text = raw;
        run.textStart = -1;
        run.lastEnd = token.lastIndex;
        continue;
      }
      run.ok = false;
      run.textStart = -1;
      // הילד t עדיין פתוח; הסריקה הרגילה מטפלת במה שנמצא בתוכו.
      run.depth = 1;
    }

    if (!run) {
      if (ours && name === 'r' && !closing && !selfClosing) {
        run = {
          start: match.index,
          attributes,
          ok: rsidOnly.test(attributes),
          props: null,
          text: null,
          textStart: -1,
          preserved: false,
          depth: 0,
          propsStart: -1,
          lastEnd: token.lastIndex,
        };
      }
      continue;
    }

    // בין ילדים ישירים של הריצה מותרים רק רווחים.
    if (run.depth === 0 && xml.slice(run.lastEnd, match.index).trim() !== '') run.ok = false;

    if (run.depth === 0 && closing && ours && name === 'r') {
      if (run.ok && run.text !== null) {
        runs.push({
          start: run.start,
          end: token.lastIndex,
          attributes: run.attributes,
          props: run.props ?? '',
          text: run.text,
        });
      }
      run = null;
      continue;
    }

    if (run.depth === 0 && !closing) {
      if (ours && name === 'rPr' && run.props === null && run.text === null) {
        if (selfClosing) {
          run.props = match[0];
        } else {
          run.propsStart = match.index;
          run.depth = 1;
        }
      } else if (ours && name === 't' && !selfClosing && run.text === null) {
        run.preserved = PRESERVE.test(attributes);
        if (attributes.trim() !== '' && !run.preserved) run.ok = false;
        run.textStart = token.lastIndex;
      } else {
        // כל ילד אחר — טאב, שבירה, הפניה, טקסט מחוק, `rPr` שני — הריצה נשארת.
        run.ok = false;
        if (!selfClosing) run.depth = 1;
      }
      run.lastEnd = token.lastIndex;
      continue;
    }

    // בתוך ילד: רק עומק נספר. סגירה בעומק 1 מחזירה לרמת הילדים הישירים.
    if (closing) {
      run.depth -= 1;
      if (run.depth === 0) {
        if (run.propsStart >= 0) {
          run.props = xml.slice(run.propsStart, token.lastIndex);
          run.propsStart = -1;
        }
        run.lastEnd = token.lastIndex;
      }
      // סגירה בלי פתיחה (מסמך קטוע): לא יורדים מתחת לאפס, והריצה אינה מאוחדת.
      if (run.depth < 0) {
        run.depth = 0;
        run.ok = false;
      }
    } else if (!selfClosing) {
      run.depth += 1;
    }
  }
  return runs;
}
