/**
 * לשוניות הרצועה — מזהים ותוויות, במקום אחד.
 *
 * הרשימה ישבה בתוך `Ribbon.vue` ולכן לא הייתה ניתנת לייבוא: `Tell Me` היה
 * צריך להפנות ללשונית („הפקד יושב ב„פריסה””), ובלי מודול היה עושה זאת
 * במחרוזת חופשית — בדיוק סוג הקישור שנקרע בשקט כששם לשונית משתנה.
 */

/** `as const` הוא מה שהופך `ribbonTab` שגוי לשגיאת בנייה, ולא לפריט מת בתפריט. */
const TABS = [
  { id: 'file', label: 'קובץ' },
  { id: 'home', label: 'בית' },
  { id: 'insert', label: 'הוספה' },
  { id: 'layout', label: 'פריסה' },
  { id: 'references', label: 'הפניות' },
  { id: 'review', label: 'סקירה' },
  { id: 'view', label: 'תצוגה' },
  // „מפתחים” יושבת אחרי „תצוגה”, במקום שבו Word מציב אותה, ומחזיקה את המאקרו
  // שישבו עד עכשיו ב„אוצריא” — ראו DeveloperTab.vue.
  { id: 'developer', label: 'מפתחים' },
  { id: 'shulchan', label: 'שולחן העורך' },
  { id: 'otzaria', label: '✦ אוצריא', className: 'otzaria-tab' },
] as const;

export type RibbonTabId = (typeof TABS)[number]['id'];

export interface RibbonTabDefinition {
  id: RibbonTabId;
  label: string;
  className?: string;
}

/** הסדר כאן הוא הסדר שבו הלשוניות מוצגות. */
export const RIBBON_TABS: readonly RibbonTabDefinition[] = TABS;

/** התווית בעברית של לשונית, לשימוש בטקסט שמוצג למשתמש. */
export const RIBBON_TAB_LABELS: Record<RibbonTabId, string> = Object.fromEntries(
  TABS.map((tab) => [tab.id, tab.label]),
) as Record<RibbonTabId, string>;
