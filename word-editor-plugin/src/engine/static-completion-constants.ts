/**
 * שם הנכס והמפתח הגלובלי של רשימות ההשלמה הסטטיות (ביטויים תלמודיים
 * ומחברים) — מודול בלי import, כמו `acronyms-constants.ts`, כדי
 * ש-`vite.config.ts` יוכל לייבא אותו.
 *
 * שני המקורות בנכס אחד ולא שניים: הם נטענים תמיד יחד (מקור שני הוא ה-fallback
 * של הראשון לאותה הקשה), ו-32KB יחד אינם מצדיקים שתי הזרקות.
 */
export const STATIC_COMPLETION_FILE = 'assets/static-completion.js';
export const STATIC_COMPLETION_GLOBAL = '__OTZARIA_STATIC_COMPLETION__';
