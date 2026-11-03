import { describe, expect, it } from 'vitest';
import { markNeutralParagraphEnds } from '../../src/engine/docx-neutral-mark';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const MATH_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const DRAW_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

/** ‏U+200F. כתוב כ-escape בכוונה — בקובץ מקור הוא בלתי-נראה. */
const RLM = '‏';

function doc(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS} ${MATH_NS} ${DRAW_NS}><w:body>${body}</w:body></w:document>`;
}

/** פסקה עם `pPr` ריקה — הכיוון אינו מוצהר בה, כמו במסמך אמיתי. */
function para(...runs: string[]): string {
  return `<w:p><w:pPr/>${runs.join('')}</w:p>`;
}

function run(text: string, attributes = ''): string {
  return `<w:r><w:t${attributes}>${text}</w:t></w:r>`;
}

/** מה שבין `<w:body>` ל-`</w:body>` בפלט, או `null`. */
function bodyOf(xml: string): string | null {
  const out = markNeutralParagraphEnds(xml);
  if (out === null) return null;
  return out.slice(out.indexOf('<w:body>') + '<w:body>'.length, out.lastIndexOf('</w:body>'));
}

describe('markNeutralParagraphEnds — מה שמסומן', () => {
  it('נקודה שהיא ריצה בפני עצמה — המקרה שדווח', () => {
    expect(bodyOf(doc(para(run('שלום עולם '), run('.'))))).toBe(
      para(run('שלום עולם '), run(`.${RLM}`)),
    );
  });

  it('נקודה בתוך ריצת הטקסט — אותו באג, צורה שנייה', () => {
    expect(bodyOf(doc(para(run('שלום עולם.'))))).toBe(para(run(`שלום עולם.${RLM}`)));
  });

  it('פסיק סוגר', () => {
    expect(bodyOf(doc(para(run('שלום עולם ועוד,'))))).toBe(para(run(`שלום עולם ועוד,${RLM}`)));
  });

  it('סוגר סוגר', () => {
    expect(bodyOf(doc(para(run('מילה ואחריה סוגר (כך)'))))).toBe(
      para(run(`מילה ואחריה סוגר (כך)${RLM}`)),
    );
  });

  /**
   * ‏Word מוחק רווח סופי ב-`<w:t>` בלי `xml:space="preserve"`. הצמדת ה-RLM
   * בסוף הייתה הופכת אותו לפנימי ומחיה אותו — שינוי טקסט. נמדד: הטבלה
   * בהערת הפתיחה של המודול.
   */
  it('רווח סופי — ההכנסה לפניו, והרווח נשאר סופי', () => {
    expect(bodyOf(doc(para(run('שלום עולם. ', ' xml:space="preserve"'))))).toBe(
      para(run(`שלום עולם.${RLM} `, ' xml:space="preserve"')),
    );
  });

  it('רווחים סופיים אחדים — ההכנסה לפני כולם', () => {
    expect(bodyOf(doc(para(run('שלום.   ', ' xml:space="preserve"'))))).toBe(
      para(run(`שלום.${RLM}   `, ' xml:space="preserve"')),
    );
  });

  it('הרווח הסופי בריצה נפרדת משלו', () => {
    expect(bodyOf(doc(para(run('שלום.'), run(' ', ' xml:space="preserve"'))))).toBe(
      para(run(`שלום.${RLM}`), run(' ', ' xml:space="preserve"')),
    );
  });

  it('ישות לפני הניטרלי — ההכנסה אינה חוצה אותה', () => {
    expect(bodyOf(doc(para(run('שלום &amp; עולם.'))))).toBe(para(run(`שלום &amp; עולם.${RLM}`)));
  });

  it('הניטרלי עצמו כתוב כישות מספרית', () => {
    expect(bodyOf(doc(para(run('שלום&#46;'))))).toBe(para(run(`שלום&#46;${RLM}`)));
  });

  it('כל פסקה לעצמה', () => {
    expect(bodyOf(doc(para(run('אחת.')) + para(run('שתיים.'))))).toBe(
      para(run(`אחת.${RLM}`)) + para(run(`שתיים.${RLM}`)),
    );
  });

  it('פסקה בתיבת טקסט — פסקה בתוך ריצה', () => {
    const inner = `<w:r><w:drawing><wp:txbxContent>${para(run('בתיבה.'))}</wp:txbxContent></w:drawing></w:r>`;
    const marked = `<w:r><w:drawing><wp:txbxContent>${para(run(`בתיבה.${RLM}`))}</wp:txbxContent></w:drawing></w:r>`;
    expect(bodyOf(doc(`<w:p><w:pPr/>${inner}</w:p>`))).toBe(`<w:p><w:pPr/>${marked}</w:p>`);
  });

  it('טקסט התוצאה של שדה נספר — כותרת תחתונה „עמוד 1.”', () => {
    const field =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
    // הספרה היא התו המכריע האחרון, ולכן הפסקה יוצאת מהתחום הצר — ולא מסומנת.
    expect(bodyOf(doc(para(run('עמוד '), field, run('1'), run('.'))))).toBeNull();
    // אבל כשהעברית היא האחרונה, התוצאה של השדה אינה מפריעה.
    expect(bodyOf(doc(para(field, run('1'), run(' מתוך המסמך.'))))).toBe(
      para(field, run('1'), run(` מתוך המסמך.${RLM}`)),
    );
  });
});

describe('markNeutralParagraphEnds — מה שאינו נגע', () => {
  it('פסקה שאינה נגמרת בניטרלי', () => {
    expect(markNeutralParagraphEnds(doc(para(run('שלום עולם'))))).toBeNull();
  });

  it('פסקה לטינית', () => {
    expect(markNeutralParagraphEnds(doc(para(run('hello world.'))))).toBeNull();
  });

  it('פסקה שכולה ניטרלית — אין בה עדות לכיוון', () => {
    expect(markNeutralParagraphEnds(doc(para(run('שלום'), run('...'))))).not.toBeNull();
    expect(markNeutralParagraphEnds(doc(para(run('...'))))).toBeNull();
  });

  it('פסקה ריקה', () => {
    expect(markNeutralParagraphEnds(doc('<w:p><w:pPr/></w:p>'))).toBeNull();
  });

  /** התחום הצר, במפורש: ההכרעה שם דורשת מדידה משלה. */
  it('ספרה אחרי האות העברית האחרונה — מחוץ לתחום', () => {
    expect(markNeutralParagraphEnds(doc(para(run('שלום 12.'))))).toBeNull();
  });

  it('אות לטינית אחרי האות העברית האחרונה — מחוץ לתחום', () => {
    expect(markNeutralParagraphEnds(doc(para(run('שלום abc.'))))).toBeNull();
  });

  it('אין הצהרת מרחב שמות — אין על מה לסמוך', () => {
    expect(markNeutralParagraphEnds('<document><body><p><t>שלום.</t></p></body></document>')).toBeNull();
  });

  it('‏`w:rPr` של סימן הפסקה אינה טקסט', () => {
    const withMark = `<w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr>${run('שלום.')}</w:p>`;
    expect(bodyOf(doc(withMark))).toBe(
      `<w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr>${run(`שלום.${RLM}`)}</w:p>`,
    );
  });

  it('ריצה של נוסחה (OMML) — שם מקומי זהה, מרחב שמות אחר', () => {
    expect(
      markNeutralParagraphEnds(doc('<w:p><w:pPr/><m:oMath><m:r><m:t>שלום.</m:t></m:r></m:oMath></w:p>')),
    ).toBeNull();
  });

  it('ריצה של DrawingML — אותו דבר', () => {
    expect(
      markNeutralParagraphEnds(doc('<w:p><w:pPr/><a:p><a:r><a:t>שלום.</a:t></a:r></a:p></w:p>')),
    ).toBeNull();
  });

  it('מה שבתוך הערה אינו פסקה', () => {
    expect(markNeutralParagraphEnds(doc(`<!-- ${para(run('שלום.'))} -->`))).toBeNull();
  });

  it('מה שבתוך CDATA אינו פסקה', () => {
    expect(markNeutralParagraphEnds(doc(`<![CDATA[ ${para(run('שלום.'))} ]]>`))).toBeNull();
  });

  it('טקסט מחוק (`w:delText`) אינו מצויר, ולכן אינו נספר', () => {
    const deleted = '<w:del><w:r><w:delText>שלום.</w:delText></w:r></w:del>';
    expect(markNeutralParagraphEnds(doc(`<w:p><w:pPr/>${deleted}</w:p>`))).toBeNull();
  });
});

describe('markNeutralParagraphEnds — הערובות', () => {
  it('long single text node does not overflow the argument stack', () => {
    const input = doc(para(run('א'.repeat(300_000) + '.')));
    expect(markNeutralParagraphEnds(input)).toBe(input.replace('.</w:t>', `.${RLM}</w:t>`));
  });

  it('decodes a paragraph written entirely as numeric entities', () => {
    expect(bodyOf(doc(para(run('&#x5E9;&#1500;&#1493;&#1501;&#46;'))))).toBe(
      para(run(`&#x5E9;&#1500;&#1493;&#1501;&#46;${RLM}`)),
    );
  });

  it('does not override digits or an explicit left-to-right mark', () => {
    for (const text of ['שלום ١٢.', 'שלום ۱۲.', 'שלום\u200e.', 'שלום.\u200e']) {
      expect(markNeutralParagraphEnds(doc(para(run(text))))).toBeNull();
    }
  });

  it('finds the last directional character across neutral-only text nodes', () => {
    expect(bodyOf(doc(para(run('שלום'), run('&amp;'), run('&#46;'), run('  '))))).toBe(
      para(run('שלום'), run('&amp;'), run(`&#46;${RLM}`), run('  ')),
    );
    expect(markNeutralParagraphEnds(doc(para(run('שלום'), run('abc'), run(' .'))))).toBeNull();
  });

  it('הרצה שנייה אינה מוסיפה דבר', () => {
    const once = markNeutralParagraphEnds(doc(para(run('שלום עולם.'))));
    expect(once).not.toBeNull();
    expect(markNeutralParagraphEnds(once!)).toBeNull();
  });

  it('גם כשיש רווח סופי — ‏RLM הוא עצמו אות ימנית, וזו המלכודת', () => {
    const once = markNeutralParagraphEnds(doc(para(run('שלום. ', ' xml:space="preserve"'))));
    expect(once).not.toBeNull();
    expect(markNeutralParagraphEnds(once!)).toBeNull();
  });

  it('הפלט הוא הקלט ועוד RLM בלבד — שום בייט לא נמחק ולא זז', () => {
    const input = doc(para(run('שלום עולם.')) + para(run('עוד אחת,')) + para(run('hello.')));
    const out = markNeutralParagraphEnds(input)!;
    expect(out).not.toBeNull();
    expect(out.split(RLM).join('')).toBe(input);
    expect(out.length).toBe(input.length + 2);
  });

  it('‏`<w:t>` שלא נסגרה אינה מייצרת פלט פגום', () => {
    const broken = doc(`<w:p><w:pPr/><w:r><w:t>שלום.</w:r></w:p>`);
    const out = markNeutralParagraphEnds(broken);
    if (out !== null) expect(out.split(RLM).join('')).toBe(broken);
  });

  it('הערה שאינה נסגרת — מה שנאסף לפניה תקף, והשאר אינו', () => {
    const cut = doc(`${para(run('שלום.'))}<!-- ${para(run('אחריה.'))}`);
    const out = markNeutralParagraphEnds(cut);
    expect(out).not.toBeNull();
    expect(out!.split(RLM).join('')).toBe(cut);
    expect(out!.length).toBe(cut.length + 1);
  });

  /**
   * הפסקה הפנימית של תיבת טקסט נסגרת **ראשונה**, ולכן היא דוחפת את ההיסט
   * הגדול לפני הקטן. בלי המיון ההכנסות יוצאות בסדר יורד, והשער מחזיר `null`.
   */
  it('פסקה מקוננת דוחפת היסט גדול לפני קטן — והמיון מטפל', () => {
    const inner = `<w:r><w:drawing><wp:txbxContent>${para(run('בתיבה.'))}</wp:txbxContent></w:drawing></w:r>`;
    const input = doc(`<w:p><w:pPr/>${run('חוץ.')}${inner}</w:p>`);
    const out = markNeutralParagraphEnds(input);
    expect(out).not.toBeNull();
    expect(out!.split(RLM).join('')).toBe(input);
    expect(out!.length).toBe(input.length + 2);
  });

  it('פסקה שנגמרת בספרה אינה מסומנת', () => {
    expect(markNeutralParagraphEnds(doc(para(run('סעיף 5'))))).toBeNull();
    // ובקרה: אותה פסקה עם ניטרלי אחרי הספרה גם היא מחוץ לתחום.
    expect(markNeutralParagraphEnds(doc(para(run('סעיף 5.'))))).toBeNull();
  });

  /**
   * ההצהרה נקראת מתג השורש. חיפוש חופשי במחרוזת היה קורא הצהרה שיושבת
   * בהערה — נמדד שהוא מכניס RLM לתוך טקסט של נוסחה.
   */
  it('הצהרת מרחב שמות בתוך הערה אינה נקראת', () => {
    const trap = `<!-- xmlns:m="${'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}" -->`;
    const input = `<?xml version="1.0"?>${trap}<w:document ${NS} ${MATH_NS}>` +
      `<w:body><w:p><w:pPr/><m:oMath><m:r><m:t>שלום.</m:t></m:r></m:oMath></w:p></w:body></w:document>`;
    expect(markNeutralParagraphEnds(input)).toBeNull();
  });

  it('פיסוק עברי הוא חזק ימני, ואינו ניטרלי', () => {
    // גרש, גרשיים, מקף, סוף פסוק — מחלקת ה-bidi שלהם R. הפסקה כבר תקינה.
    for (const ending of ['ר׳', 'רש״י', 'בן־אדם', 'דבר׃']) {
      expect(markNeutralParagraphEnds(doc(para(run(`שלום ${ending}`)))), ending).toBeNull();
    }
  });

  it('תו על-בסיסי אינו נחצה', () => {
    const input = doc(para(run('שלום 𝕏.')));
    const out = markNeutralParagraphEnds(input);
    // האות הלטינית המתמטית היא התו המכריע האחרון — מחוץ לתחום.
    expect(out).toBeNull();
    const withHebrew = doc(para(run('𝕏 שלום.')));
    expect(markNeutralParagraphEnds(withHebrew)).toBe(doc(para(run(`𝕏 שלום.${RLM}`))));
  });
});
