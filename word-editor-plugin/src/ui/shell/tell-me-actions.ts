/**
 * קטלוג הפקודות והאפשרויות עבור מנגנון ה-Tell Me (חיפוש פקודות בסגנון Word).
 *
 * ## התוויות של הקיצורים
 *
 * ארבע רשומות כאן הבטיחו למשתמש צירוף שאין לו מאזין: „הגדלת גופן” ו„הקטנת
 * גופן” הראו `Ctrl+Shift+.` ו-`Ctrl+Shift+,` בעוד הרג'יסטרי קושר `Ctrl+]`
 * ו-`Ctrl+[`, „הפעלת מאקרו” הראתה `Alt+F9` שאינו קיים כלל, ו„הקלטת מאקרו”
 * הראתה `Alt+F8` — הצירוף של **ניהול** מאקרו, כלומר פעולה אחרת. ארבעתן עוברות
 * מעכשיו דרך `shortcutLabel`, ולכן אינן יכולות להיפרד מהמאזין.
 *
 * **שאר 39 השורות עדיין כתובות ביד.** הן נכונות היום (נמדד), אבל המזהים כאן
 * אינם מזהי הרג'יסטרי (`file-save` מול `save`), ולכן ההמרה שלהן דורשת טבלת
 * מיפוי ולא החלפה מכנית. זה נשאר חוב מתועד.
 *
 * המטרה: לאפשר למשתמש להקליד מונח בעברית (כגון: "מרכז", "טבלה", "גופן", "הדפסה",
 * "שמור", "סגנון", "מעקב", "אוצריא"), לקבל תוצאות מדויקות ומדורגות, ולהפעיל את
 * הפקודה ישירות מהמקלדת או העכבר.
 */

import { shortcutLabel, type ShellAction } from '../shortcuts/registry';
import { RIBBON_TAB_LABELS, type RibbonTabId } from '../ribbon/tabs';
import { alignmentPayload, lineHeightPayload, stylePayload, zoomPayload } from '../../engine/payloads';

/**
 * הפעולות הייעודיות. איחוד מוקלד ולא `string`: מזהה שאינו כאן נופל בבנייה,
 * ולא נראה כפריט תקין שלחיצה עליו אינה עושה דבר.
 *
 * `clipboard-*` — הלוח אינו משטח פקודות של המנוע (אין `copy`/`cut`/`paste`
 * ברג׳יסטרי שלו); הוא עובר ב-engine/clipboard.ts, כמו בכפתורי „בית”.
 *
 * מי שכל תפקידו להביא את המשתמש ללשונית שבה הפקד יושב אינו כאן אלא בשדה
 * `ribbonTab` — ראו ההערה עליו.
 */
export type TellMeCustomAction =
  | 'export-pdf'
  | 'export-otzaria'
  | 'about'
  | 'report-bug'
  | 'exit-app'
  | 'toggle-book-completion'
  | 'clipboard-copy'
  | 'clipboard-cut'
  | 'clipboard-paste'
  | 'clipboard-paste-merge'
  | 'clipboard-paste-text'
  | 'shortcuts-manage';

export interface TellMeAction {
  /** מזהה ייחודי לפעולה */
  id: string;
  /** כותרת ראשית בעברית שמוצגת למשתמש */
  title: string;
  /** נתיב מיקום ברצועה (Breadcrumb) או תחום פעולה */
  category: string;
  /** תיאור קצר אופציונלי */
  description?: string;
  /** מילות מפתח, נרדפות ומונחים באנגלית ובעברית לחיפוש מהיר */
  keywords: string[];
  /** צירוף מקשים מוצג (אם קיים) */
  shortcut?: string;
  /** שם האייקון בספריית ה-SvgIcon */
  icon: string;
  /** פקודת מנוע עריכה להרצה ישירה דרך commandAdapter */
  command?: { id: string; payload?: unknown };
  /** פעולת מעטפת להרצה דרך runShellAction */
  shellAction?: ShellAction;
  /** פעולה ייעודית למעטפת שאינה ב-ShellAction — מטופלת ב-`onCustomActionFromTellMe` ב-App.vue */
  customAction?: TellMeCustomAction;
  /**
   * „הפקד יושב בלשונית הזאת” — פותח את הלשונית במקום להריץ משהו.
   *
   * רוב הפקדים בלשוניות „פריסה”, „הפניות”, „סקירה” ו„שולחן העורך” פותחים
   * דיאלוג שחי בתוך קומפוננטת הלשונית ואינו נחשף החוצה: אין להם `command`
   * ואין להם `shellAction`, ולכן הם היו **בלתי ניתנים למציאה** בחיפוש. פריט
   * שמוליך אל הפקד עדיף על פריט שאינו קיים — וזו גם ההתנהגות של Word עצמו
   * כשהפריט הוא גלריה.
   */
  ribbonTab?: RibbonTabId;
}

/** התיאור האחיד לפריט שכל תפקידו להוליך אל הלשונית שבה הפקד יושב. */
function inTab(tab: RibbonTabId, control: string): string {
  return `פתיחת לשונית „${RIBBON_TAB_LABELS[tab]}” — הפקד „${control}”`;
}

/**
 * רשימת הפקודות המלאה הזמינה ב-Tell Me.
 */
export const TELL_ME_ACTIONS: readonly TellMeAction[] = [
  // --- קובץ ---
  {
    id: 'file-save',
    title: 'שמירה',
    category: 'קובץ',
    description: 'שמירת השינויים במסמך הנוכחי',
    keywords: ['שמור', 'שמירה', 'שמירת מסמך', 'save'],
    shortcut: 'Ctrl+S',
    icon: 'save',
    shellAction: 'save',
  },
  {
    id: 'file-save-as',
    title: 'שמירה בשם',
    category: 'קובץ',
    description: 'שמירת המסמך בעותק חדש או בשם אחר',
    keywords: ['שמור בשם', 'שמירה בשם', 'עותק', 'ייצוא', 'וורד', 'docx', 'word', 'save as'],
    shortcut: 'Ctrl+Shift+S',
    icon: 'saveAs',
    shellAction: 'save-as',
  },
  {
    id: 'file-new',
    title: 'מסמך חדש',
    category: 'קובץ',
    description: 'יצירת מסמך חדש ריק',
    keywords: ['חדש', 'מסמך חדש', 'דף חדש', 'קובץ חדש', 'new'],
    shortcut: 'Ctrl+N',
    icon: 'newDoc',
    shellAction: 'new-document',
  },
  {
    id: 'file-open',
    title: 'פתיחת קובץ',
    category: 'קובץ',
    description: 'פתיחת מסמך קיים מהמחשב',
    keywords: ['פתח', 'פתח קובץ', 'פתיחה', 'קובץ', 'טעינה', 'open', 'load'],
    shortcut: 'Ctrl+O',
    icon: 'folder',
    shellAction: 'open-document',
  },
  {
    id: 'file-print',
    title: 'הדפסה',
    category: 'קובץ',
    description: 'הדפסת המסמך',
    keywords: ['הדפס', 'הדפסה', 'מדפסת', 'print'],
    shortcut: 'Ctrl+P',
    icon: 'print',
    shellAction: 'print',
  },
  {
    id: 'file-export-pdf',
    title: 'ייצוא ל-PDF',
    category: 'קובץ',
    description: 'שמירת המסמך כקובץ PDF',
    keywords: ['פי די אף', 'ייצוא', 'pdf', 'export'],
    icon: 'exportPdf',
    customAction: 'export-pdf',
  },
  {
    id: 'file-exit',
    title: 'יציאה',
    category: 'קובץ',
    description: 'סגירת העורך',
    keywords: ['יציאה', 'צא', 'סגור', 'סגירה', 'exit', 'quit', 'close'],
    icon: 'exit',
    customAction: 'exit-app',
  },
  {
    id: 'help-shortcuts',
    title: 'קיצורי מקלדת',
    category: 'עזרה',
    description: 'הצגת רשימת כל קיצורי המקלדת',
    keywords: ['קיצורים', 'מקלדת', 'עזרה', 'קיצור', 'shortcuts', 'help'],
    shortcut: 'Ctrl+/',
    icon: 'info',
    shellAction: 'shortcuts-help',
  },
  {
    id: 'help-shortcuts-manage',
    title: 'ניהול קיצורים',
    category: 'עזרה',
    description: 'הצמדת צירוף מקשים לערכת עיצוב',
    keywords: [
      'ניהול',
      'קיצורים',
      'קיצור',
      'הצמדה',
      'התאמה',
      'ערכה',
      'עיצוב',
      'customize',
      'shortcuts',
      'keyboard',
    ],
    icon: 'keyboard',
    customAction: 'shortcuts-manage',
  },
  {
    id: 'help-about',
    title: 'אודות וורד לאוצריא',
    category: 'עזרה',
    description: 'מידע על התוכנה וגרסתה',
    keywords: ['אודות', 'גרסה', 'מידע', 'about'],
    icon: 'info',
    customAction: 'about',
  },
  {
    id: 'help-report-bug',
    title: 'דיווח על באג',
    category: 'עזרה',
    description: 'פתיחת דיווח על תקלה במאגר התוסף ב-GitHub',
    keywords: ['באג', 'תקלה', 'דיווח', 'שגיאה', 'bug', 'issue'],
    icon: 'bug',
    customAction: 'report-bug',
  },

  // --- לוח (Clipboard) ---
  {
    id: 'edit-undo',
    title: 'בטל פעולה',
    category: 'עריכה',
    description: 'ביטול הפעולה האחרונה',
    keywords: ['בטל', 'ביטול', 'אחורה', 'undo'],
    shortcut: 'Ctrl+Z',
    icon: 'undo',
    command: { id: 'undo' },
  },
  {
    id: 'edit-redo',
    title: 'בצע שוב',
    category: 'עריכה',
    description: 'חזרה על הפעולה שבוטלה',
    keywords: ['בצע שוב', 'חזור', 'קדימה', 'redo'],
    shortcut: 'Ctrl+Y',
    icon: 'redo',
    command: { id: 'redo' },
  },
  {
    id: 'clipboard-paste',
    title: 'הדבק',
    category: 'בית > לוח',
    description: 'הדבקת תוכן מלוח הגזירים',
    keywords: ['הדבק', 'הדבקה', 'לוח', 'paste'],
    shortcut: 'Ctrl+V',
    icon: 'paste',
    customAction: 'clipboard-paste',
  },
  {
    id: 'clipboard-paste-merge',
    title: 'הדבקה עם מיזוג עיצוב',
    category: 'בית > לוח',
    description: 'הדבקה בעיצוב של המקום שאליו מדביקים, עם המודגש, הנטוי והקו התחתון של המקור',
    keywords: ['מזג עיצוב', 'מיזוג עיצוב', 'הדבקה', 'הדבק', 'עיצוב יעד', 'merge formatting', 'paste'],
    icon: 'pasteMerge',
    customAction: 'clipboard-paste-merge',
  },
  {
    id: 'clipboard-paste-text',
    title: 'הדבקת טקסט בלבד',
    category: 'בית > לוח',
    description: 'הדבקה בלי שום עיצוב מהמקור — הטקסט מקבל את העיצוב של המקום שאליו מדביקים',
    keywords: ['טקסט בלבד', 'שמור טקסט בלבד', 'הדבקה ללא עיצוב', 'הדבקה', 'הדבק', 'plain text', 'paste'],
    icon: 'pasteText',
    customAction: 'clipboard-paste-text',
  },
  {
    id: 'clipboard-copy',
    title: 'העתק',
    category: 'בית > לוח',
    description: 'העתקת הטקסט המסומן ללוח',
    keywords: ['העתק', 'העתקה', 'לוח', 'copy'],
    shortcut: 'Ctrl+C',
    icon: 'copy',
    customAction: 'clipboard-copy',
  },
  {
    id: 'clipboard-cut',
    title: 'גזור',
    category: 'בית > לוח',
    description: 'גזירת הטקסט המסומן ללוח',
    keywords: ['גזור', 'גזירה', 'לוח', 'cut'],
    shortcut: 'Ctrl+X',
    icon: 'cut',
    customAction: 'clipboard-cut',
  },
  {
    id: 'clipboard-format-painter',
    title: 'מברשת עיצוב',
    category: 'בית > לוח',
    description: 'העתקת עיצוב ממקום אחד והחלתו במקום אחר',
    keywords: ['מברשת', 'מברשת עיצוב', 'העתק עיצוב', 'format painter'],
    shortcut: 'Ctrl+Shift+C',
    icon: 'formatPainter',
    command: { id: 'copy-format' },
  },

  // --- גופן ועיצוב תו ---
  {
    id: 'font-bold',
    title: 'מודגש (Bold)',
    category: 'בית > גופן',
    description: 'הדגשת הטקסט המסומן',
    keywords: ['מודגש', 'הדגשה', 'עבה', 'בולד', 'bold'],
    shortcut: 'Ctrl+B',
    icon: 'bold',
    command: { id: 'bold' },
  },
  {
    id: 'font-italic',
    title: 'נטוי (Italic)',
    category: 'בית > גופן',
    description: 'הטיית הטקסט המסומן',
    keywords: ['נטוי', 'הטיה', 'איטליק', 'italic'],
    shortcut: 'Ctrl+I',
    icon: 'italic',
    command: { id: 'italic' },
  },
  {
    id: 'font-underline',
    title: 'קו תחתי (Underline)',
    category: 'בית > גופן',
    description: 'הוספת קו מתחת לטקסט',
    keywords: ['קו תחתי', 'פס תחתי', 'קו למטה', 'underline'],
    shortcut: 'Ctrl+U',
    icon: 'underline',
    command: { id: 'underline' },
  },
  {
    id: 'font-strikethrough',
    title: 'קו חוצה (Strikethrough)',
    category: 'בית > גופן',
    description: 'העברת קו באמצע הטקסט',
    keywords: ['קו חוצה', 'מחיקה', 'חוצה', 'strike', 'strikethrough'],
    icon: 'strikethrough',
    command: { id: 'strikethrough' },
  },
  {
    id: 'font-grow',
    title: 'הגדלת גופן',
    category: 'בית > גופן',
    description: 'הגדלת גודל האותיות בדרגה אחת',
    keywords: ['הגדל', 'הגדלת גופן', 'גודל', 'אותיות גדולות', 'grow font', 'bigger'],
    shortcut: shortcutLabel('font-grow'),
    icon: 'growFont',
    shellAction: 'font-grow',
  },
  {
    id: 'font-shrink',
    title: 'הקטנת גופן',
    category: 'בית > גופן',
    description: 'הקטנת גודל האותיות בדרגה אחת',
    keywords: ['הקטן', 'הקטנת גופן', 'גודל', 'אותיות קטנות', 'shrink font', 'smaller'],
    shortcut: shortcutLabel('font-shrink'),
    icon: 'shrinkFont',
    shellAction: 'font-shrink',
  },
  {
    id: 'font-superscript',
    title: 'כתב עילי (Superscript)',
    category: 'בית > גופן',
    description: 'הקטנת הטקסט והצבתו מעל גובה השורה',
    keywords: ['כתב עילי', 'כתב עליון', 'חזקה', 'למעלה', 'superscript'],
    shortcut: 'Ctrl+Shift+=',
    icon: 'superscript',
    shellAction: 'superscript',
  },
  {
    id: 'font-subscript',
    title: 'כתב תחתי (Subscript)',
    category: 'בית > גופן',
    description: 'הקטנת הטקסט והצבתו מתחת לגובה השורה',
    keywords: ['כתב תחתי', 'אינדקס', 'למטה', 'subscript'],
    shortcut: 'Ctrl+=',
    icon: 'subscript',
    shellAction: 'subscript',
  },
  {
    id: 'font-clear-formatting',
    title: 'נקה עיצוב',
    category: 'בית > גופן',
    description: 'הסרת כל עיצובי התווים והחזרה לעיצוב ברירת המחדל',
    keywords: ['נקה', 'ניקוי עיצוב', 'אפס עיצוב', 'איפוס', 'clear formatting'],
    shortcut: 'Ctrl+Space',
    icon: 'clearFormatting',
    command: { id: 'clear-formatting' },
  },

  // --- פסקה ויישור ---
  {
    id: 'para-align-right',
    title: 'יישור לימין',
    category: 'בית > פסקה',
    description: 'יישור שורות הפסקה לצד ימין',
    keywords: ['ימין', 'יישור לימין', 'align right'],
    shortcut: 'Ctrl+R',
    icon: 'alignRight',
    command: { id: 'text-align', payload: alignmentPayload('right') },
  },
  {
    id: 'para-align-center',
    title: 'יישור למרכז (מרכוז)',
    category: 'בית > פסקה',
    description: 'מרכוז שורות הפסקה באמצע הדף',
    keywords: ['מרכז', 'מרכוז', 'אמצע', 'יישור למרכז', 'center', 'align center'],
    shortcut: 'Ctrl+E',
    icon: 'alignCenter',
    command: { id: 'text-align', payload: alignmentPayload('center') },
  },
  {
    id: 'para-align-left',
    title: 'יישור לשמאל',
    category: 'בית > פסקה',
    description: 'יישור שורות הפסקה לצד שמאל',
    keywords: ['שמאל', 'יישור לשמאל', 'align left'],
    shortcut: 'Ctrl+L',
    icon: 'alignLeft',
    command: { id: 'text-align', payload: alignmentPayload('left') },
  },
  {
    id: 'para-align-justify',
    title: 'יישור לשני הצדדים (בלוק)',
    category: 'בית > פסקה',
    description: 'יישור שורות הפסקה לשני הצדדים באופן שווה',
    keywords: ['יישור לשני הצדדים', 'מיושר', 'בלוק', 'חסימה', 'justify'],
    shortcut: 'Ctrl+J',
    icon: 'alignJustify',
    command: { id: 'text-align', payload: alignmentPayload('justify') },
  },
  {
    id: 'para-dir-rtl',
    title: 'כיוון פסקה מימין לשמאל (RTL)',
    category: 'בית > פסקה',
    description: 'קביעת כיוון הפסקה לשפות הנכתבות מימין לשמאל',
    keywords: ['כיוון', 'ימין לשמאל', 'עברית', 'rtl', 'right to left'],
    shortcut: 'Ctrl+RightShift',
    icon: 'dirRtl',
    command: { id: 'direction-rtl' },
  },
  {
    id: 'para-dir-ltr',
    title: 'כיוון פסקה משמאל לימין (LTR)',
    category: 'בית > פסקה',
    description: 'קביעת כיוון הפסקה לשפות הנכתבות משמאל לימין',
    keywords: ['כיוון', 'שמאל לימין', 'אנגלית', 'ltr', 'left to right'],
    shortcut: 'Ctrl+LeftShift',
    icon: 'dirLtr',
    command: { id: 'direction-ltr' },
  },
  {
    id: 'para-bullets',
    title: 'רשימת תבליטים',
    category: 'בית > פסקה',
    description: 'הוספה או הסרה של רשימת תבליטים',
    keywords: ['תבליטים', 'נקודות', 'רשימה', 'bullets', 'list'],
    icon: 'bulletList',
    command: { id: 'bullet-list' },
  },
  {
    id: 'para-numbering',
    title: 'רשימה ממוספרת',
    category: 'בית > פסקה',
    description: 'הוספה או הסרה של מספור פסקאות',
    keywords: ['מספור', 'מספרים', 'רשימה ממוספרת', 'numbers', 'numbered list'],
    icon: 'numberList',
    command: { id: 'numbered-list' },
  },
  {
    id: 'para-indent-increase',
    title: 'הגדלת כניסת פסקה',
    category: 'בית > פסקה',
    description: 'הזזת הפסקה פנימה',
    keywords: ['כניסה', 'הזחה', 'הגדל כניסה', 'טאב', 'indent', 'increase indent'],
    icon: 'indentIncrease',
    command: { id: 'indent-increase' },
  },
  {
    id: 'para-indent-decrease',
    title: 'הקטנת כניסת פסקה',
    category: 'בית > פסקה',
    description: 'הזזת הפסקה החוצה לכיוון השוליים',
    keywords: ['כניסה', 'הזחה', 'הקטן כניסה', 'הזחה אחורה', 'outdent', 'decrease indent'],
    icon: 'indentDecrease',
    command: { id: 'indent-decrease' },
  },
  {
    /*
     * הפקד עצמו, ולא ערך מסוים בו: הוא מחזיק גם את מרווח השורות וגם את
     * הרווח בין הפסקאות, ומי שמחפש „רווח בין פסקאות” מחפש אותו ולא את 1.5.
     * `ribbonTab` ולא `command` — תפריט אינו פקודה שאפשר להריץ מכאן.
     */
    id: 'para-spacing-menu',
    title: 'מרווח שורות וריווח',
    category: 'בית > פסקה',
    description: inTab('home', 'מרווח שורות וריווח'),
    keywords: [
      'מרווח שורות וריווח',
      'רווח בין פסקאות',
      'ריווח בין פסקאות',
      'רווח לפני הפסקה',
      'רווח אחרי הפסקה',
      'paragraph spacing',
    ],
    icon: 'lineSpacing',
    ribbonTab: 'home',
  },
  {
    id: 'para-spacing-1',
    title: 'מרווח שורות 1.0 (יחיד)',
    category: 'בית > פסקה',
    description: 'קביעת מרווח שורות רגיל',
    keywords: ['מרווח', 'מרווח שורות', 'רווח שורות', 'line spacing'],
    icon: 'lineSpacing',
    command: { id: 'line-height', payload: lineHeightPayload(1.0) },
  },
  {
    id: 'para-spacing-115',
    title: 'מרווח שורות 1.15',
    category: 'בית > פסקה',
    description: 'קביעת מרווח שורות 1.15',
    keywords: ['מרווח', '1.15', 'רווח בין שורות'],
    icon: 'lineSpacing',
    command: { id: 'line-height', payload: lineHeightPayload(1.15) },
  },
  {
    id: 'para-spacing-15',
    title: 'מרווח שורות 1.5',
    category: 'בית > פסקה',
    description: 'קביעת מרווח שורות שורה וחצי',
    keywords: ['מרווח', '1.5', 'שורה וחצי', 'רווח שורות'],
    icon: 'lineSpacing',
    command: { id: 'line-height', payload: lineHeightPayload(1.5) },
  },
  {
    id: 'para-spacing-2',
    title: 'מרווח שורות 2.0 (כפול)',
    category: 'בית > פסקה',
    description: 'קביעת מרווח שורות כפול',
    keywords: ['מרווח', '2.0', 'כפול', 'מרווח כפול'],
    icon: 'lineSpacing',
    command: { id: 'line-height', payload: lineHeightPayload(2.0) },
  },

  // --- סגנונות (Styles) ---
  {
    id: 'style-normal',
    title: 'סגנון: רגיל',
    category: 'בית > סגנונות',
    description: 'החלת סגנון טקסט רגיל',
    keywords: ['סגנון רגיל', 'טקסט רגיל', 'פסקה רגילה', 'normal', 'style'],
    shortcut: 'Ctrl+Shift+N',
    icon: 'pilcrow',
    command: { id: 'linked-style', payload: stylePayload('Normal') },
  },
  {
    id: 'style-h1',
    title: 'סגנון: כותרת 1',
    category: 'בית > סגנונות',
    description: 'החלת סגנון כותרת ראשית ברמה 1',
    keywords: ['כותרת 1', 'h1', 'heading 1'],
    shortcut: 'Ctrl+Alt+1',
    icon: 'pilcrow',
    command: { id: 'linked-style', payload: stylePayload('Heading1') },
  },
  {
    id: 'style-h2',
    title: 'סגנון: כותרת 2',
    category: 'בית > סגנונות',
    description: 'החלת סגנון כותרת משנית ברמה 2',
    keywords: ['כותרת 2', 'h2', 'heading 2'],
    shortcut: 'Ctrl+Alt+2',
    icon: 'pilcrow',
    command: { id: 'linked-style', payload: stylePayload('Heading2') },
  },
  {
    id: 'style-h3',
    title: 'סגנון: כותרת 3',
    category: 'בית > סגנונות',
    description: 'החלת סגנון כותרת רמה 3',
    keywords: ['כותרת 3', 'h3', 'heading 3'],
    shortcut: 'Ctrl+Alt+3',
    icon: 'pilcrow',
    command: { id: 'linked-style', payload: stylePayload('Heading3') },
  },

  // --- עריכה וחיפוש ---
  {
    id: 'edit-find',
    title: 'חיפוש במסמך',
    category: 'בית > עריכה',
    description: 'פתיחת חלונית חיפוש טקסט במסמך',
    keywords: ['חפש', 'חיפוש', 'מצא', 'איתור', 'find', 'search'],
    shortcut: 'Ctrl+F',
    icon: 'search',
    shellAction: 'find',
  },
  {
    id: 'edit-replace',
    title: 'חיפוש והחלפה',
    category: 'בית > עריכה',
    description: 'פתיחת חלונית חיפוש והחלפת טקסט במסמך',
    keywords: ['החלף', 'החלפה', 'חיפוש והחלפה', 'replace'],
    shortcut: 'Ctrl+H',
    icon: 'replace',
    shellAction: 'replace',
  },
  {
    id: 'edit-select-all',
    title: 'בחר הכל',
    category: 'בית > עריכה',
    description: 'בחירת כל הטקסט במסמך',
    keywords: ['בחר הכל', 'סימון הכל', 'סמן הכל', 'select all'],
    shortcut: 'Ctrl+A',
    icon: 'select',
    shellAction: 'select-all',
  },

  // --- הוספה ---
  {
    id: 'insert-page-break',
    title: 'מעבר עמוד',
    category: 'הוספה > עמודים',
    description: 'התחלת עמוד חדש במיקום הסמן',
    keywords: ['מעבר עמוד', 'התחל בעמוד חדש', 'עמוד חדש', 'דף חדש', 'שבירת עמוד', 'page break'],
    shortcut: 'Ctrl+Enter',
    icon: 'pageBreak',
    shellAction: 'page-break',
  },
  {
    id: 'insert-table',
    title: 'הוספת טבלה (3x3)',
    category: 'הוספה > טבלאות',
    description: 'הוספת טבלה בסיסית במסמך',
    keywords: ['טבלה', 'הוסף טבלה', 'טבלאות', 'table'],
    icon: 'table',
    command: { id: 'table-insert', payload: { rows: 3, cols: 3 } },
  },
  {
    id: 'insert-link',
    title: 'הוספת קישור',
    category: 'הוספה > קישורים',
    description: 'יצירת היפר-קישור בטקסט המסומן',
    keywords: ['קישור', 'לינק', 'היפר קישור', 'link', 'hyperlink'],
    shortcut: 'Ctrl+K',
    icon: 'link',
    shellAction: 'link',
  },
  {
    id: 'insert-footnote',
    title: 'הוספת הערת שוליים',
    category: 'הפניות > הערות',
    description: 'הוספת הערה בתחתית העמוד הנוכחי',
    keywords: ['הערת שוליים', 'הערה', 'שוליים', 'footnote'],
    shortcut: 'Ctrl+Alt+F',
    icon: 'footnote',
    shellAction: 'footnote',
  },
  {
    id: 'insert-endnote',
    title: 'הוספת הערת סיום',
    category: 'הפניות > הערות',
    description: 'הוספת הערה בסוף המסמך',
    keywords: ['הערת סיום', 'סיום', 'endnote'],
    shortcut: 'Ctrl+Alt+D',
    icon: 'footnote',
    shellAction: 'endnote',
  },
  {
    id: 'insert-toc',
    title: 'תוכן עניינים',
    category: 'הפניות > תוכן עניינים',
    description: 'יצירת תוכן עניינים אוטומטי מהכותרות',
    keywords: ['תוכן', 'תוכן עניינים', 'אינדקס', 'toc', 'table of contents'],
    icon: 'toc',
    command: { id: 'table-of-contents-insert' },
  },

  // --- הוספה: פקדים שהדיאלוג שלהם חי בתוך הלשונית ---
  {
    id: 'insert-image',
    title: 'תמונות',
    category: 'הוספה > איורים',
    description: inTab('insert', 'תמונות'),
    keywords: ['תמונה', 'תמונות', 'איור', 'צילום', 'image', 'picture'],
    icon: 'image',
    ribbonTab: 'insert',
  },
  {
    id: 'insert-bookmark',
    title: 'סימנייה',
    category: 'הוספה > קישורים',
    description: inTab('insert', 'סימנייה'),
    keywords: ['סימנייה', 'סימניה', 'סימניות', 'bookmark'],
    icon: 'bookmark',
    ribbonTab: 'insert',
  },
  {
    id: 'insert-remove-link',
    title: 'הסר קישור',
    category: 'הוספה > קישורים',
    description: inTab('insert', 'הסר קישור'),
    keywords: ['הסר קישור', 'הסרת קישור', 'ביטול קישור', 'unlink', 'remove link'],
    icon: 'link',
    ribbonTab: 'insert',
  },
  {
    id: 'insert-header',
    title: 'כותרת עליונה',
    category: 'הוספה > כותרת עליונה ותחתונה',
    description: inTab('insert', 'כותרת עליונה'),
    keywords: ['כותרת עליונה', 'כותרת', 'ראש עמוד', 'header'],
    icon: 'header',
    ribbonTab: 'insert',
  },
  {
    id: 'insert-footer',
    title: 'כותרת תחתונה',
    category: 'הוספה > כותרת עליונה ותחתונה',
    description: inTab('insert', 'כותרת תחתונה'),
    keywords: ['כותרת תחתונה', 'תחתית', 'רגל עמוד', 'footer'],
    icon: 'footer',
    ribbonTab: 'insert',
  },
  {
    id: 'insert-page-number',
    title: 'מספר עמוד',
    category: 'הוספה > כותרת עליונה ותחתונה',
    description: inTab('insert', 'מספר עמוד'),
    keywords: ['מספר עמוד', 'מספרי עמודים', 'page number'],
    icon: 'pageNumber',
    ribbonTab: 'insert',
  },
  {
    id: 'insert-hf-first-page',
    title: 'שונה בעמוד ראשון',
    category: 'הוספה > כותרת עליונה ותחתונה',
    description: inTab('insert', 'שונה בעמוד ראשון'),
    keywords: ['שונה בעמוד ראשון', 'עמוד ראשון', 'כותרת ראשונה', 'first page'],
    icon: 'firstPageHeader',
    ribbonTab: 'insert',
  },
  {
    id: 'insert-hf-odd-even',
    title: 'שונה בעמודים זוגיים ואי-זוגיים',
    category: 'הוספה > כותרת עליונה ותחתונה',
    description: inTab('insert', 'שונה בעמודים זוגיים ואי-זוגיים'),
    keywords: ['שונה בעמודים זוגיים ואי-זוגיים', 'זוגי', 'אי-זוגי', 'odd even'],
    icon: 'oddEvenPages',
    ribbonTab: 'insert',
  },
  {
    id: 'insert-hf-link-previous',
    title: 'קשר לקודם',
    category: 'הוספה > כותרת עליונה ותחתונה',
    description: inTab('insert', 'קשר לקודם'),
    keywords: ['קשר לקודם', 'קישור למקטע הקודם', 'link to previous'],
    icon: 'link',
    ribbonTab: 'insert',
  },
  {
    id: 'insert-date-time',
    title: 'תאריך ושעה',
    category: 'הוספה > טקסט',
    description: inTab('insert', 'תאריך ושעה'),
    keywords: ['תאריך ושעה', 'תאריך', 'שעה', 'date', 'time'],
    icon: 'dateTime',
    ribbonTab: 'insert',
  },
  {
    id: 'insert-update-fields',
    title: 'עדכן שדות',
    category: 'הוספה > טקסט',
    description: inTab('insert', 'עדכן שדות'),
    keywords: ['עדכן שדות', 'שדות', 'רענון שדות', 'update fields'],
    icon: 'updateFields',
    ribbonTab: 'insert',
  },

  // --- פריסה ---
  {
    id: 'layout-margins',
    title: 'שוליים',
    category: 'פריסה > הגדרת עמוד',
    description: inTab('layout', 'שוליים'),
    keywords: ['שוליים', 'שולי הדף', 'margins'],
    icon: 'margins',
    ribbonTab: 'layout',
  },
  {
    id: 'layout-orientation',
    title: 'כיוון הדף',
    category: 'פריסה > הגדרת עמוד',
    description: inTab('layout', 'כיוון'),
    keywords: ['כיוון', 'כיוון הדף', 'לאורך', 'לרוחב', 'orientation'],
    icon: 'orientation',
    ribbonTab: 'layout',
  },
  {
    id: 'layout-paper-size',
    title: 'גודל הדף',
    category: 'פריסה > הגדרת עמוד',
    description: inTab('layout', 'גודל'),
    keywords: ['גודל', 'גודל הדף', 'גודל נייר', 'a4', 'letter', 'paper size'],
    icon: 'paperSize',
    ribbonTab: 'layout',
  },
  {
    id: 'layout-columns',
    title: 'עמודות',
    category: 'פריסה > הגדרת עמוד',
    description: inTab('layout', 'עמודות'),
    keywords: ['עמודות', 'טורים', 'שתי עמודות', 'columns'],
    icon: 'columns',
    ribbonTab: 'layout',
  },
  {
    id: 'layout-line-numbers',
    title: 'מספרי שורות',
    category: 'פריסה > הגדרת עמוד',
    description: inTab('layout', 'מספרי שורות'),
    keywords: ['מספרי שורות', 'מספור שורות', 'line numbers'],
    icon: 'numberList',
    ribbonTab: 'layout',
  },
  {
    id: 'layout-page-borders',
    title: 'גבולות עמוד',
    category: 'פריסה > הגדרת עמוד',
    description: inTab('layout', 'גבולות עמוד'),
    keywords: ['גבולות עמוד', 'מסגרת', 'גבול', 'page borders'],
    icon: 'borders',
    ribbonTab: 'layout',
  },
  {
    id: 'layout-vertical-align',
    title: 'יישור אנכי',
    category: 'פריסה > מקטע',
    description: inTab('layout', 'יישור אנכי'),
    keywords: ['יישור אנכי', 'מיקום אנכי', 'גובה העמוד', 'vertical align'],
    icon: 'lineSpacing',
    ribbonTab: 'layout',
  },
  {
    id: 'layout-page-numbering',
    title: 'מספור עמודים',
    category: 'פריסה > מקטע',
    description: inTab('layout', 'מספור עמודים'),
    keywords: ['מספור עמודים', 'תבנית מספור', 'מספר התחלה', 'page numbering'],
    icon: 'pageNumber',
    ribbonTab: 'layout',
  },
  {
    id: 'layout-header-distance',
    title: 'מרחק הכותרת',
    category: 'פריסה > מקטע',
    description: inTab('layout', 'מרחק הכותרת'),
    keywords: ['מרחק הכותרת', 'מרחק כותרת', 'שולי כותרת', 'header distance'],
    icon: 'header',
    ribbonTab: 'layout',
  },
  {
    id: 'layout-doc-defaults',
    title: 'ברירות מחדל',
    category: 'פריסה > מקטע',
    description: inTab('layout', 'ברירות מחדל'),
    keywords: ['ברירות מחדל', 'ברירת מחדל', 'גופן המסמך', 'defaults'],
    icon: 'fontColor',
    ribbonTab: 'layout',
  },

  // --- הפניות ---
  {
    id: 'refs-toc-mark-entry',
    title: 'סמן ערך',
    category: 'הפניות > תוכן עניינים',
    description: inTab('references', 'סמן ערך'),
    keywords: ['סמן ערך', 'סימון ערך', 'תוכן עניינים', 'mark entry'],
    icon: 'bookmark',
    ribbonTab: 'references',
  },
  {
    id: 'refs-toc-update',
    title: 'עדכן טבלה',
    category: 'הפניות > תוכן עניינים',
    description: inTab('references', 'עדכן טבלה'),
    keywords: ['עדכן טבלה', 'רענון תוכן עניינים', 'update table'],
    icon: 'updateFields',
    ribbonTab: 'references',
  },
  {
    id: 'refs-toc-custom',
    title: 'התאמה אישית',
    category: 'הפניות > תוכן עניינים',
    description: inTab('references', 'התאמה אישית'),
    keywords: ['התאמה אישית', 'התאמה', 'תוכן עניינים מותאם', 'custom'],
    icon: 'toc',
    ribbonTab: 'references',
  },
  {
    id: 'refs-toc-remove',
    title: 'הסר תוכן עניינים',
    category: 'הפניות > תוכן עניינים',
    description: inTab('references', 'הסר'),
    keywords: ['הסר', 'הסר תוכן עניינים', 'מחיקת תוכן עניינים', 'remove toc'],
    icon: 'reject',
    ribbonTab: 'references',
  },
  {
    id: 'refs-manage-notes',
    title: 'נהל הערות',
    category: 'הפניות > הערות',
    description: inTab('references', 'נהל הערות'),
    keywords: ['נהל הערות', 'ניהול הערות', 'הערות שוליים', 'manage notes'],
    icon: 'book',
    ribbonTab: 'references',
  },
  {
    id: 'refs-index-mark',
    title: 'סמן ערך למפתח',
    category: 'הפניות > מפתח',
    description: inTab('references', 'סמן ערך למפתח'),
    keywords: ['סמן ערך למפתח', 'מפתח', 'אינדקס', 'mark index entry'],
    icon: 'bookmark',
    ribbonTab: 'references',
  },
  {
    id: 'refs-index-insert',
    title: 'הוסף מפתח',
    category: 'הפניות > מפתח',
    description: inTab('references', 'הוסף מפתח'),
    keywords: ['הוסף מפתח', 'מפתח', 'אינדקס', 'insert index'],
    icon: 'book',
    ribbonTab: 'references',
  },
  {
    id: 'refs-index-update',
    title: 'עדכן מפתח',
    category: 'הפניות > מפתח',
    description: inTab('references', 'עדכן מפתח'),
    keywords: ['עדכן מפתח', 'רענון מפתח', 'update index'],
    icon: 'updateFields',
    ribbonTab: 'references',
  },
  {
    id: 'refs-index-settings',
    title: 'הגדרות מפתח',
    category: 'הפניות > מפתח',
    description: inTab('references', 'הגדרות מפתח'),
    keywords: ['הגדרות מפתח', 'מפתח', 'index settings'],
    icon: 'toc',
    ribbonTab: 'references',
  },
  {
    id: 'refs-index-remove',
    title: 'הסר מפתח',
    category: 'הפניות > מפתח',
    description: inTab('references', 'הסר מפתח'),
    keywords: ['הסר מפתח', 'מחיקת מפתח', 'remove index'],
    icon: 'reject',
    ribbonTab: 'references',
  },
  {
    id: 'refs-citation-insert',
    title: 'הוסף ציטוט',
    category: 'הפניות > ציטוטים',
    description: inTab('references', 'הוסף ציטוט'),
    keywords: ['הוסף ציטוט', 'ציטוט', 'מקור', 'citation'],
    icon: 'comment',
    ribbonTab: 'references',
  },
  {
    id: 'refs-bibliography',
    title: 'ביבליוגרפיה',
    category: 'הפניות > ציטוטים',
    description: inTab('references', 'ביבליוגרפיה'),
    keywords: ['ביבליוגרפיה', 'רשימת מקורות', 'bibliography'],
    icon: 'toc',
    ribbonTab: 'references',
  },
  {
    id: 'refs-sources',
    title: 'נהל מקורות',
    category: 'הפניות > ציטוטים',
    description: inTab('references', 'נהל מקורות'),
    keywords: ['נהל מקורות', 'ניהול מקורות', 'מקורות', 'manage sources'],
    icon: 'book',
    ribbonTab: 'references',
  },
  {
    id: 'refs-bibliography-update',
    title: 'עדכן ביבליוגרפיה',
    category: 'הפניות > ציטוטים',
    description: inTab('references', 'עדכן ביבליוגרפיה'),
    keywords: ['עדכן ביבליוגרפיה', 'רענון ביבליוגרפיה', 'update bibliography'],
    icon: 'updateFields',
    ribbonTab: 'references',
  },
  {
    id: 'refs-bibliography-remove',
    title: 'הסר ביבליוגרפיה',
    category: 'הפניות > ציטוטים',
    description: inTab('references', 'הסר ביבליוגרפיה'),
    keywords: ['הסר ביבליוגרפיה', 'מחיקת ביבליוגרפיה', 'remove bibliography'],
    icon: 'reject',
    ribbonTab: 'references',
  },
  {
    id: 'refs-caption',
    title: 'הוסף כיתוב',
    category: 'הפניות > כיתובים',
    description: inTab('references', 'הוסף כיתוב'),
    keywords: ['הוסף כיתוב', 'כיתוב', 'תיאור תמונה', 'caption'],
    icon: 'image',
    ribbonTab: 'references',
  },
  {
    id: 'refs-cross-refs-update',
    title: 'עדכן הפניות',
    category: 'הפניות > כיתובים',
    description: inTab('references', 'עדכן הפניות'),
    keywords: ['עדכן הפניות', 'הפניה מקושרת', 'cross reference'],
    icon: 'updateFields',
    ribbonTab: 'references',
  },

  // --- תצוגה וסקירה ---
  {
    id: 'view-focus-mode',
    title: 'מצב מיקוד',
    category: 'תצוגה',
    description: 'תצוגת מסך מלא נקייה ללא הסחות דעת',
    keywords: ['מיקוד', 'מצב מיקוד', 'מסך מלא', 'קריאה', 'focus', 'focus mode'],
    shortcut: 'F11',
    icon: 'focusMode',
    shellAction: 'focus-mode',
  },
  {
    id: 'view-navigation-pane',
    title: 'חלונית ניווט',
    category: 'תצוגה',
    description: 'הצגת כותרות המסמך בצד, ומעבר מכותרת לכותרת בלחיצה',
    keywords: ['ניווט', 'חלונית ניווט', 'כותרות', 'מבנה', 'ראשי פרקים', 'navigation', 'outline'],
    icon: 'navigationPane',
    shellAction: 'navigation-pane',
  },
  {
    id: 'view-ruler',
    title: 'סרגל מידות',
    category: 'תצוגה',
    description: 'הצגה או הסתרה של סרגלי המידות (אופקי ואנכי)',
    keywords: ['סרגל', 'סרגלים', 'מידות', 'סרגל שוליים', 'ruler'],
    icon: 'ruler',
    command: { id: 'ruler' },
  },
  {
    id: 'view-formatting-marks',
    title: 'סימני עיצוב (פילקרו)',
    category: 'תצוגה',
    description: 'הצגת סימני פסקאות, רווחים ומעברי שורות מוסתרים',
    keywords: ['סימני עיצוב', 'רווחים', 'פילקרו', 'תווים נסתרים', 'pilcrow', 'formatting marks'],
    shortcut: 'Ctrl+Shift+8',
    icon: 'pilcrow',
    command: { id: 'formatting-marks' },
  },
  {
    id: 'review-track-changes',
    title: 'מעקב אחר שינויים',
    category: 'סקירה',
    description: 'הפעלה או כיבוי של מעקב אחר עריכות במסמך',
    keywords: ['מעקב', 'שינויים', 'עקוב אחר שינויים', 'track changes'],
    shortcut: 'Ctrl+Shift+E',
    icon: 'trackChanges',
    shellAction: 'track-changes',
  },
  {
    id: 'review-accept-change',
    title: 'קבל שינוי',
    category: 'סקירה > שינויים',
    description: 'קבלת השינוי הנוכחי במעקב אחר שינויים',
    keywords: ['קבל שינוי', 'קבלת שינוי', 'אשר שינוי', 'accept change'],
    icon: 'accept',
    command: { id: 'acceptChange' },
  },
  {
    id: 'review-reject-change',
    title: 'דחה שינוי',
    category: 'סקירה > שינויים',
    description: 'דחיית השינוי הנוכחי במעקב אחר שינויים',
    keywords: ['דחה שינוי', 'דחיית שינוי', 'בטל שינוי', 'reject change'],
    icon: 'reject',
    command: { id: 'rejectChange' },
  },
  {
    id: 'review-accept-all',
    title: 'קבל את כל השינויים',
    category: 'סקירה > שינויים',
    description: 'קבלת כל השינויים במסמך',
    keywords: ['קבל את כל השינויים', 'קבל הכל', 'accept all'],
    icon: 'accept',
    command: { id: 'acceptAllChanges' },
  },
  {
    id: 'review-reject-all',
    title: 'דחה את כל השינויים',
    category: 'סקירה > שינויים',
    description: 'דחיית כל השינויים במסמך',
    keywords: ['דחה את כל השינויים', 'דחה הכל', 'reject all'],
    icon: 'reject',
    command: { id: 'rejectAllChanges' },
  },
  {
    id: 'review-spellcheck',
    title: 'בדיקת איות',
    category: 'סקירה > הגהה',
    description: inTab('review', 'בדיקת איות'),
    keywords: ['בדיקת איות', 'איות', 'שגיאות כתיב', 'הגהה', 'spellcheck'],
    icon: 'proofing',
    ribbonTab: 'review',
  },
  {
    id: 'review-autocorrect',
    title: 'תיקון אוטומטי',
    category: 'סקירה > הגהה',
    description: inTab('review', 'תיקון אוטומטי'),
    keywords: ['תיקון אוטומטי', 'תיקון בהקלדה', 'החלפה אוטומטית', 'autocorrect', 'auto correct'],
    icon: 'replace',
    ribbonTab: 'review',
  },
  {
    id: 'review-comment-new',
    title: 'תגובה חדשה',
    category: 'סקירה > תגובות',
    description: inTab('review', 'תגובה חדשה'),
    keywords: ['תגובה חדשה', 'תגובה', 'הערה', 'comment'],
    icon: 'comment',
    ribbonTab: 'review',
  },
  {
    id: 'review-protect',
    title: 'הגבל עריכה',
    category: 'סקירה > הגנה',
    description: inTab('review', 'הגבל עריכה'),
    keywords: ['הגבל עריכה', 'הגנה', 'נעילה', 'קריאה בלבד', 'protect'],
    icon: 'proofing',
    ribbonTab: 'review',
  },
  {
    id: 'view-actual-size',
    title: 'גודל אמיתי',
    category: 'תצוגה',
    description: 'הצגת המסמך בגודלו האמיתי (100%)',
    keywords: ['גודל אמיתי', 'מאה אחוז', '100%', 'זום', 'zoom', 'actual size'],
    icon: 'zoom',
    command: { id: 'zoom', payload: zoomPayload(100) },
  },
  {
    id: 'view-page-width',
    title: 'רוחב עמוד',
    category: 'תצוגה',
    description: inTab('view', 'רוחב עמוד'),
    keywords: ['רוחב עמוד', 'התאם לרוחב', 'זום', 'fit width'],
    icon: 'fitWidth',
    ribbonTab: 'view',
  },
  {
    /*
     * „רקע עמוד” הוא מה שרוב המשתמשים יחפשו, וזה **לא** מה שהפקד עושה: הצבע
     * חל על הבד — המשטח שסביב הדף — והוא העדפת תצוגה שאינה נכתבת למסמך. לכן
     * הוא מילת מפתח ולא הכותרת. `ribbonTab` ולא `command`: מה שנפתח הוא פלטה.
     */
    id: 'view-canvas-color',
    title: 'צבע רקע',
    category: 'תצוגה',
    description: inTab('view', 'צבע רקע'),
    keywords: [
      'צבע רקע',
      'רקע',
      'רקע עמוד',
      'צבע רקע העורך',
      'background color',
      'page color',
    ],
    icon: 'shading',
    ribbonTab: 'view',
  },
  {
    id: 'font-advanced',
    title: 'גופן מתקדם',
    category: 'בית > גופן',
    description: inTab('home', 'מתקדם'),
    keywords: ['מתקדם', 'גופן מתקדם', 'ריווח תווים', 'דיאלוג גופן', 'advanced font'],
    icon: 'fontColor',
    ribbonTab: 'home',
  },

  // --- אוצריא ---
  {
    id: 'otzaria-citation',
    title: 'ציטוט מהקורא של אוצריא',
    category: 'אוצריא',
    description: 'הדבקת הקטע והמקור הפתוח כעת באוצריא אל תוך המסמך',
    keywords: ['ציטוט', 'ציטוט מהקורא', 'אוצריא', 'מקור', 'קורא', 'citation', 'otzaria'],
    shortcut: 'Ctrl+Shift+Q',
    icon: 'book',
    shellAction: 'insert-citation',
  },
  {
    id: 'otzaria-search',
    title: 'חיפוש הטקסט בספריית אוצריא',
    category: 'אוצריא',
    description: 'חיפוש הקטע המסומן בספרי התוכנה',
    keywords: ['חיפוש באוצריא', 'ספרייה', 'חיפוש ספרים', 'אוצריא'],
    shortcut: 'Ctrl+Shift+G',
    icon: 'search',
    shellAction: 'search-otzaria',
  },
  {
    id: 'otzaria-library',
    title: 'פתיחת ספריית אוצריא',
    category: 'אוצריא',
    description: 'פתיחת חלון ספריית הספרים',
    keywords: ['ספרייה', 'ספרים', 'פתח ספרייה', 'פתח ספר', 'אוצריא', 'library'],
    icon: 'book',
    shellAction: 'open-library',
  },
  {
    id: 'otzaria-export',
    title: 'ייצוא לאוצריא',
    category: 'אוצריא',
    description: 'ייצוא המסמך בפורמט מותאם לספריית אוצריא',
    keywords: ['ייצוא לאוצריא', 'שמירה לאוצריא', 'אוצריא'],
    icon: 'export',
    customAction: 'export-otzaria',
  },

  // --- מאקרו ---
  {
    id: 'macro-record',
    title: 'הקלטת מאקרו',
    category: 'מאקרו',
    description: 'התחלה או עצירה של הקלטת רצף פעולות',
    keywords: ['מאקרו', 'הקלטה', 'הקלטת מאקרו', 'הקלט מאקרו', 'עצור הקלטה', 'record', 'macro'],
    shortcut: shortcutLabel('macro-record'),
    icon: 'macro',
    shellAction: 'macro-record',
  },
  {
    id: 'macro-play',
    title: 'הפעלת מאקרו אחרון',
    category: 'מאקרו',
    description: 'ביצוע חוזר של המאקרו שהוקלט לאחרונה',
    keywords: ['הפעל מאקרו', 'נגן מאקרו', 'נגן אחרון', 'מאקרו', 'play macro'],
    shortcut: shortcutLabel('macro-play'),
    icon: 'macro',
    shellAction: 'macro-play',
  },
  {
    id: 'macro-manage',
    title: 'ניהול מאקרו',
    category: 'מאקרו',
    description: 'פתיחת חלון ניהול ועריכת המאקרו',
    keywords: ['ניהול מאקרו', 'רשימת מאקרו', 'מאקרו'],
    icon: 'macro',
    shellAction: 'macro-manage',
  },

  {
    id: 'otzaria-book-completion',
    title: 'השלמה מהספר',
    category: 'אוצריא',
    description: 'בזמן הקלדה, אם הטקסט תואם את הספר הפתוח בקורא — Tab משלים מהמקור',
    keywords: ['השלמה', 'השלמה מהספר', 'השלמה אוטומטית', 'ghost', 'אוצריא', 'ספר', 'קורא'],
    icon: 'highlight',
    customAction: 'toggle-book-completion',
  },

  // --- שולחן עורך ---
  {
    id: 'shulchan-first-word',
    title: 'מילה ראשונה מוגדלת ומודגשת',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'מילה ראשונה'),
    keywords: ['מילה ראשונה', 'שולחן עורך', 'פתיח', 'מודגשת', 'ראשונה'],
    icon: 'growFont',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-unclosed',
    title: 'חיפוש סוגריים לא סגורים',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'סוגריים לא סגורים'),
    keywords: ['סוגריים', 'סוגריים לא סגורים', 'שולחן עורך', 'הגהה'],
    icon: 'search',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-uniform',
    title: 'אחידות מסמך (שולחן עורך)',
    category: 'שולחן עורך',
    description: 'פתיחת לשונית שולחן העורך — גודל עמוד, שוליים ורוחב טורים אחידים',
    keywords: ['אחידות', 'אחיד', 'שוליים', 'טורים', 'שולחן עורך'],
    icon: 'alignJustify',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-common-errors',
    title: 'שגיאות מצויות',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'שגיאות מצויות'),
    keywords: ['שגיאות מצויות', 'שגיאות', 'הגהה', 'תיקון', 'שולחן עורך'],
    icon: 'proofing',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-alt-text',
    title: 'טקסט מתחלף',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'טקסט מתחלף'),
    keywords: ['טקסט מתחלף', 'מתחלף', 'החלפת טקסט', 'שולחן עורך'],
    icon: 'bold',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-paste-fix',
    title: 'תיקון העתקה',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'תיקון העתקה'),
    keywords: ['תיקון העתקה', 'ניקוי הדבקה', 'הדבקה', 'שולחן עורך'],
    icon: 'paste',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-parens-to-notes',
    title: 'סוגריים ⟵ הערות',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'סוגריים ⟵ הערות'),
    keywords: ['סוגריים להערות', 'סוגריים ⟵ הערות', 'הערות שוליים', 'שולחן עורך'],
    icon: 'footnote',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-notes-to-parens',
    title: 'הערות ⟵ סוגריים',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'הערות ⟵ סוגריים'),
    keywords: ['הערות לסוגריים', 'הערות ⟵ סוגריים', 'הערות שוליים', 'שולחן עורך'],
    icon: 'footnote',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-uniform-line-spacing',
    title: 'מרווח שורות אחיד',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'מרווח שורות אחיד'),
    keywords: ['מרווח שורות אחיד', 'מרווח אחיד', 'ריווח אחיד', 'שולחן עורך'],
    icon: 'lineSpacing',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-clear-line-spacing',
    title: 'בטל מרווח אחיד',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'בטל מרווח אחיד'),
    keywords: ['בטל מרווח אחיד', 'ביטול מרווח', 'שולחן עורך'],
    icon: 'lineSpacing',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-page-size',
    title: 'גודל עמוד ושוליים',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'גודל עמוד ושוליים'),
    keywords: ['גודל עמוד ושוליים', 'גודל עמוד', 'שוליים', 'אחידות', 'שולחן עורך'],
    icon: 'paperSize',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-column-width',
    title: 'רוחב טורים',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'רוחב טורים'),
    keywords: ['רוחב טורים', 'טורים', 'עמודות', 'אחידות', 'שולחן עורך'],
    icon: 'columns',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-shrink',
    title: 'צמצום מסמך',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'צמצום מסמך'),
    keywords: ['צמצום מסמך', 'צמצום', 'כיווץ', 'עמודים', 'שולחן עורך'],
    icon: 'shrinkFont',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-mark-pages',
    title: 'סמן עמודים',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'סמן עמודים'),
    keywords: ['סמן עמודים', 'סימון עמודים', 'דפוס', 'שולחן עורך'],
    icon: 'pageNumber',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-check-pages',
    title: 'בדוק עמודים',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'בדוק עמודים'),
    keywords: ['בדוק עמודים', 'בדיקת עמודים', 'דפוס', 'שולחן עורך'],
    icon: 'proofing',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-clear-marks',
    title: 'הסר סימון',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'הסר סימון'),
    keywords: ['הסר סימון', 'ניקוי סימון', 'דפוס', 'שולחן עורך'],
    icon: 'clearFormatting',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-crop-marks',
    title: 'סימני חיתוך',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'סימני חיתוך'),
    keywords: ['סימני חיתוך', 'חיתוך', 'דפוס', 'שולחן עורך'],
    icon: 'borders',
    ribbonTab: 'shulchan',
  },
  {
    id: 'shulchan-split-doc',
    title: 'פירוק מסמך',
    category: 'שולחן עורך',
    description: inTab('shulchan', 'פירוק מסמך'),
    keywords: ['פירוק מסמך', 'פיצול מסמך', 'חלוקה', 'דפוס', 'שולחן עורך'],
    icon: 'export',
    ribbonTab: 'shulchan',
  },
];

/**
 * מזהי הפעולות המוצעות כברירת מחדל כאשר תיבת החיפוש ריקה.
 */
export const DEFAULT_SUGGESTED_IDS: readonly string[] = [
  'file-save',
  'edit-find',
  'edit-replace',
  'insert-table',
  'font-bold',
  'para-align-center',
  'insert-link',
  'view-focus-mode',
  'help-shortcuts',
];

/**
 * אותיות סופיות → הצורה הרגילה שלהן.
 *
 * בלי זה התאמת התחילית נשברת בדיוק במקום שבו המשתמש מקליד: „מסמך” אינה
 * תחילית של „מסמכים”, „עמוד” אינה תחילית של „עמודים במסמך”, ו„טור” אינה
 * תחילית של „טורים” — אף שלמשתמש אלה אותן מילים. הנרמול חל על השאילתה
 * ועל האינדקס כאחד, ולכן שני הצדדים נפגשים באותה צורה.
 */
const FINAL_LETTERS: Record<string, string> = {
  'ך': 'כ',
  'ם': 'מ',
  'ן': 'נ',
  'ף': 'פ',
  'ץ': 'צ',
};

/**
 * מנרמל מחרוזת עברית לחיפוש: מסיר ניקוד, משווה אותיות סופיות לרגילות,
 * מוריד אותיות לועזיות לאותיות קטנות, ומסיר רווחים מיותרים.
 */
export function normalizeSearchTerm(text: string): string {
  return text
    .replace(/[\u0591-\u05C7]/g, '') // הסרת טעמים וניקוד
    .replace(/[ךםןףץ]/g, (letter) => FINAL_LETTERS[letter])
    .trim()
    .toLowerCase();
}

/** פעולה עם השדות שלה מנורמלים — הצורה שהחיפוש באמת קורא. */
interface IndexedAction {
  readonly action: TellMeAction;
  readonly title: string;
  readonly titleWords: readonly string[];
  readonly category: string;
  readonly description: string;
  readonly keywords: readonly string[];
  /** כל המילים שבמילות המפתח, שטוחות — לחיפוש מילה-מילה. */
  readonly keywordWords: readonly string[];
}

interface CatalogIndex {
  readonly entries: readonly IndexedAction[];
  readonly suggested: readonly TellMeAction[];
}

/**
 * האינדקס המנורמל של קטלוג. הנרמול קבוע — ולכן הוא נעשה פעם אחת לקטלוג ולא
 * מחדש בכל הקשה: בלעדיו כל תו שהמשתמש מקליד היה מריץ מאות `replace` של regex
 * על כותרות, תיאורים ומילות מפתח שאינם משתנים.
 */
const CATALOG_INDEX = new WeakMap<readonly TellMeAction[], CatalogIndex>();

function catalogIndex(actions: readonly TellMeAction[]): CatalogIndex {
  const cached = CATALOG_INDEX.get(actions);
  if (cached) return cached;

  const entries = actions.map((action): IndexedAction => {
    const title = normalizeSearchTerm(action.title);
    const keywords = action.keywords.map(normalizeSearchTerm);
    return {
      action,
      title,
      titleWords: title.split(/\s+/).filter(Boolean),
      category: normalizeSearchTerm(action.category),
      description: action.description ? normalizeSearchTerm(action.description) : '',
      keywords,
      keywordWords: [...new Set(keywords.flatMap((kw) => kw.split(/\s+/).filter(Boolean)))],
    };
  });
  const byId = new Map(actions.map((action) => [action.id, action]));
  const suggested = DEFAULT_SUGGESTED_IDS.map((id) => byId.get(id)).filter(
    (action): action is TellMeAction => action !== undefined,
  );

  const index: CatalogIndex = { entries, suggested };
  CATALOG_INDEX.set(actions, index);
  return index;
}

/**
 * מחפש ומדרג פעולות מתוך קטלוג ה-Tell Me.
 *
 * @param query שאילתת החיפוש של המשתמש
 * @param actions רשימת הפעולות לחיפוש (ברירת מחדל: TELL_ME_ACTIONS)
 * @returns רשימת פעולות ממוינת לפי ציון רלוונטיות
 */
export function searchTellMeActions(
  query: string,
  actions: readonly TellMeAction[] = TELL_ME_ACTIONS,
): TellMeAction[] {
  const index = catalogIndex(actions);
  const normalized = normalizeSearchTerm(query);
  // כשאין שאילתה, מחזירים את הפעולות המוצעות לפי סדר ההגדרה שלהן
  if (!normalized) return [...index.suggested];

  const queryWords = normalized.split(/\s+/).filter(Boolean);

  const scored: { action: TellMeAction; score: number }[] = [];

  for (const entry of index.entries) {
    let score = 0;

    // התאמה מדויקת לכותרת
    if (entry.title === normalized) {
      score += 120;
    } else if (entry.title.startsWith(normalized)) {
      // תחילית של הכותרת
      score += 90;
    } else if (entry.title.includes(normalized)) {
      // מוכלת בכותרת
      score += 60;
    }

    // בדיקת מילים בתוך הכותרת ובתוך מילות המפתח.
    //
    // מילות המפתח נבדקו עד עכשיו מול השאילתה **השלמה** בלבד, ולכן „השלמה ספר”
    // לא מצא את „השלמה מהספר”: אף מילת מפתח אינה מכילה את שתי המילים ברצף.
    // המשתמש מקליד שתיים-שלוש מילים מתוך שם הפקד, לא את שמו המדויק.
    //
    // **וכולן, לא אחת מהן.** מילה אחת קצרה ונפוצה שמתאימה הספיקה כדי לתת ניקוד:
    // „זזזזז לא קיים” החזיר תוצאות, כי „לא” הוא תחילית של מילה במילות המפתח של
    // „סוגריים לא סגורים”. מי שמקליד שלוש מילים מתכוון לשלושתן.
    let matchedWords = 0;
    let wordScore = 0;
    for (const qWord of queryWords) {
      if (entry.titleWords.some((w) => w.startsWith(qWord))) {
        matchedWords += 1;
        wordScore += 35;
      } else if (entry.keywordWords.some((w) => w.startsWith(qWord))) {
        matchedWords += 1;
        wordScore += 25;
      }
    }
    if (matchedWords === queryWords.length) {
      score += wordScore;
    }

    // מילות מפתח
    for (const kw of entry.keywords) {
      if (kw === normalized) {
        score += 80;
      } else if (kw.startsWith(normalized)) {
        score += 50;
      } else if (kw.includes(normalized)) {
        score += 30;
      }
    }

    // קטגוריה
    if (entry.category.includes(normalized)) {
      score += 20;
    }

    // תיאור
    if (entry.description.includes(normalized)) {
      score += 15;
    }

    if (score > 0) {
      scored.push({ action: entry.action, score });
    }
  }

  // מיון לפי ציון יורד
  scored.sort((a, b) => b.score - a.score);

  return scored.slice(0, 15).map((item) => item.action);
}
