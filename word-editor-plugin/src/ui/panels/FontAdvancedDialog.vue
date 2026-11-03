<template>
  <Teleport to="body">
    <div
      v-if="isOpen"
      ref="rootRef"
      class="fontadv-dialog"
      :style="dragStyle"
      role="dialog"
      aria-modal="true"
      :aria-label="DIALOG_TITLE"
      :aria-describedby="blockingText === '' ? undefined : 'fa-blocking-note'"
      tabindex="-1"
      @keydown.esc.stop="$emit('close')"
      @keydown.enter="onDialogEnter"
    >
      <div
        class="fa-header dialog-drag-handle"
        @pointerdown="startDialogDrag"
      >
        <span class="fa-title">{{ DIALOG_TITLE }}</span>
        <button
          type="button"
          class="fa-close-btn"
          data-tip-title="סגור"
          data-tip-shortcut="Esc"
          aria-label="סגור את תפריט הגופן המתקדם"
          @click="$emit('close')"
        >
          ✕
        </button>
      </div>

      <div class="fa-body">
        <!--
          מה שחוסם נאמר **בפתיחה**, ולא אחרי שהדיאלוג נסגר ובלע את מה שהוקלד.
          שתי סיבות שונות ושתי הודעות שונות — ראו `hasRangeSelection`
          ב-engine/font-advanced.ts. `aria-describedby` מהשורש: „אישור” נעול
          אינו ממוקד, ולכן הנימוק לא היה מגיע למי שקורא מסך דרכו.
        -->
        <p
          v-if="blockingText !== ''"
          id="fa-blocking-note"
          class="fa-notice fa-notice-blocking"
          role="alert"
        >
          {{ blockingText }}
        </p>

        <div class="fa-columns">
          <!-- עמודה א׳: מספרים ואפקטים -->
          <div class="fa-column">
            <fieldset class="fa-group">
              <legend>מרווחים ומיקום</legend>
              <div class="fa-fields">
                <label
                  for="fa-scale"
                  class="fa-label"
                >מתיחה אופקית</label>
                <span class="fa-input">
                  <input
                    id="fa-scale"
                    v-model="charScale"
                    class="fa-number"
                    type="number"
                    min="1"
                    max="600"
                    step="1"
                    :placeholder="UNCHANGED"
                    aria-label="מתיחה אופקית של התווים, באחוזים"
                  >
                  <span class="fa-unit">%</span>
                </span>

                <label
                  for="fa-spacing"
                  class="fa-label"
                >ריווח תווים</label>
                <span class="fa-input">
                  <input
                    id="fa-spacing"
                    v-model="letterSpacing"
                    class="fa-number"
                    type="number"
                    step="1"
                    :placeholder="UNCHANGED"
                    aria-label="ריווח בין תווים, בנקודות. שלילי = מכווץ"
                  >
                  <span class="fa-unit">נק'</span>
                </span>

                <label
                  for="fa-kerning"
                  class="fa-label"
                >קרנינג מעל</label>
                <span class="fa-input">
                  <input
                    id="fa-kerning"
                    v-model="kerning"
                    class="fa-number"
                    type="number"
                    min="0"
                    step="1"
                    :placeholder="UNCHANGED"
                    aria-label="גודל מינימלי לקרנינג, בנקודות"
                  >
                  <span class="fa-unit">נק'</span>
                </span>

                <label
                  for="fa-position"
                  class="fa-label"
                >הרמה/הנמכה</label>
                <span class="fa-input">
                  <input
                    id="fa-position"
                    v-model="position"
                    class="fa-number"
                    type="number"
                    step="1"
                    :placeholder="UNCHANGED"
                    aria-label="מיקום התו בנקודות. חיובי = מוגבה, שלילי = מונמך"
                  >
                  <span class="fa-unit">נק'</span>
                </span>
              </div>
            </fieldset>

            <fieldset class="fa-group">
              <legend>אפקטים</legend>
              <div class="fa-toggles">
                <TriToggle
                  v-for="effect in EFFECTS"
                  :key="effect.key"
                  v-model="effects[effect.key]"
                  :label="effect.label"
                  :description="effect.description"
                />
              </div>
              <!-- vanish מסתיר תוכן: האזהרה כאן ולא ב-tooltip בלבד, כי זו הפעולה
                   היחידה בדיאלוג שהמשתמש עלול לחשוב ש„לא עבדה". -->
              <p
                v-if="vanish === 'yes'"
                class="fa-warning"
                role="note"
              >
                הטקסט המסומן יוסתר מעיני הקורא. „✕” מחזיר אותו.
              </p>
            </fieldset>
          </div>

          <!-- עמודה ב׳: הליבה העברית -->
          <div class="fa-column">
            <fieldset class="fa-group">
              <legend>גופן מורכב (עברית)</legend>
              <div class="fa-fields">
                <span class="fa-label">גופן</span>
                <span class="fa-input">
                  <!--
                    אותו בורר של הרצועה, ומאותה רשימה — ראו
                    composables/font-family-options.ts. `focus-return="stay"`
                    מפני שהדיאלוג עדיין פתוח: מיקוד שחוזר למסמך היה מוציא את
                    ה-Escape שלו מכלל פעולה.
                  -->
                  <RibbonCombo
                    v-model="complexFontName"
                    class="fa-combo"
                    :options="familyOptions"
                    width="150px"
                    title="גופן מורכב"
                    :placeholder="UNCHANGED"
                    focus-return="stay"
                  />
                </span>

                <label
                  for="fa-sizecs"
                  class="fa-label"
                >גודל</label>
                <span class="fa-input">
                  <input
                    id="fa-sizecs"
                    v-model="fontSizeCs"
                    class="fa-number"
                    type="number"
                    min="0.5"
                    step="0.5"
                    :placeholder="UNCHANGED"
                    aria-label="גודל הגופן המורכב, בנקודות"
                  >
                  <span class="fa-unit">נק'</span>
                </span>

                <label
                  for="fa-lang"
                  class="fa-label"
                >שפת הגהה</label>
                <span class="fa-input">
                  <!--
                    בורר ולא כפתור, וזו אינה חוסר-עקביות: „עברית” ו„אנגלית”
                    אינן הדלקה וכיבוי של אותו דבר אלא שתי בחירות, ומיתוג
                    ביניהן היה מסתיר את השלישית — „אל תיגע”.
                  -->
                  <select
                    id="fa-lang"
                    v-model="proofingLang"
                    class="fa-select"
                  >
                    <option value="">{{ UNCHANGED }}</option>
                    <option value="he-IL">עברית (he-IL)</option>
                    <option value="en-US">אנגלית (en-US)</option>
                  </select>
                </span>
              </div>

              <div class="fa-toggles">
                <TriToggle
                  v-for="effect in COMPLEX"
                  :key="effect.key"
                  v-model="effects[effect.key]"
                  :label="effect.label"
                  :description="effect.description"
                />
              </div>
            </fieldset>
          </div>
        </div>

        <!--
          פס התצוגה המקדימה. `aria-hidden` כמו פס הדגימה של בורר הגופן: הוא
          חוזר על טקסט שהמשתמש עצמו סימן, ואין בו מה להכריז.
        -->
        <div class="fa-preview">
          <span class="fa-preview-caption">{{ PREVIEW_CAPTION }}</span>
          <div
            class="fa-preview-strip"
            aria-hidden="true"
            dir="auto"
          >
            <span
              class="fa-preview-text"
              :style="previewStyle"
            >{{ sampleText }}</span>
          </div>

          <!--
            מתחת לפס, ולא בתוך אחד המקטעים: הרשימה חוצה את שתי העמודות
            (אפקטים, קרנינג, ומחסנית הכתב המורכב), והיא ההמשך הישיר של מה
            שהכיתוב מעליה מתחייב עליו — „כך ייראה ב-Word”, ומה שהמסך לא
            יראה. מופיעה רק כשנבחר משהו מהרשימה: הערת קבע היא הערה שלא נקראת.
          -->
          <p
            v-if="undrawnText !== ''"
            class="fa-notice"
            role="note"
          >
            {{ undrawnText }}
          </p>
        </div>

        <p
          v-if="showError"
          class="fa-error"
          role="alert"
        >
          {{ INVALID_HINT }}
        </p>
      </div>

      <div class="fa-footer">
        <span
          class="fa-count"
          role="status"
        >{{ countText }}</span>
        <button
          type="button"
          class="fa-btn fa-btn-quiet"
          :disabled="changeCount === 0"
          data-tip-title="נקה הכל"
          data-tip-desc="מחזיר כל שדה ל„ללא שינוי”"
          @pointerdown.prevent
          @click="resetFields"
        >
          נקה הכל
        </button>
        <span class="fa-spacer" />
        <button
          type="button"
          class="fa-btn fa-btn-primary"
          data-default-action
          :disabled="busy || !canSubmit"
          @pointerdown.prevent
          @click="onSubmit"
        >
          אישור
        </button>
        <button
          type="button"
          class="fa-btn"
          @pointerdown.prevent
          @click="$emit('close')"
        >
          ביטול
        </button>
      </div>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
/**
 * „גופן מתקדם" — ריווח תווים, מיקום, אפקטים, טקסט מוסתר והליבה העברית (CS).
 * ההנמקות של מה שנשלח למנוע ב-engine/font-advanced.ts.
 *
 * ## שלוש ההכרעות שמחזיקות את הדיאלוג
 *
 * 1. **אין מילוי מוקדם** — ואין סיכון הרסני כמו ב„פסקה". `format.apply`
 *    הוא patch לפי מפתח: מפתח שלא נשלח אינו נוגע בעיצוב קיים. לכן כל שדה
 *    פותח ריק/„ללא שינוי", ורק מה שהמשתמש מילא יוצא למנוע. זה שונה לגמרי
 *    מ-setIndentation, שמחליף אלמנט שלם.
 * 2. **הבוליאנים תלת-מצביים.** ההנמקה המלאה — ומה שיסיר את המצב השלישי —
 *    ב-`common/TriToggle.vue`. בקצרה: המנוע אינו מדווח את מצב האפקטים על
 *    הבחירה, ולכן כפתור דו-מצבי היה מסיר צל וחריטה מטקסט שאיש לא ביקש לגעת בו.
 * 3. **בורר גופן ולא תיבת טקסט.** „גופן מורכב" ביקש עד כה להקליד שם מהזיכרון,
 *    בעורך שהרשימה שלו יודעת מה מותקן במכונה ומה מכסה עברית. הוא מקבל עכשיו
 *    את **אותה** רשימה של בורר הרצועה — composables/font-family-options.ts.
 *
 * ## הפריסה: שתי עמודות, ולמה
 *
 * שבעה-עשר הפקדים היו עמודה אחת גבוהה מהמסך, כלומר גוף שגולל — והפוטר, שהוא
 * „אישור"/„ביטול", הגיע רק אחרי גלילה (scripts/qa/dialog-drag-qa.mjs מדד את
 * זה). עשרת הבוררים שהפכו לכפתורים הם מה שקיצר את הרשימה מספיק כדי ששתי
 * עמודות יכניסו את הכול בלי גלילה בחלון סביר. הגלילה עצמה לא ירדה: היא
 * נשארת ב-`.fa-body` בלבד, בשביל חלון נמוך במיוחד.
 *
 * ## פס התצוגה המקדימה
 *
 * מצייר את **הטקסט שהמשתמש סימן** (`readSelectionText` — קריאה בלבד, בלי מגע
 * במסמך; ההנמקה ב-composables/font-sample.ts) עם מה שנבחר בדיאלוג. הוא
 * `CSS` ולא המנוע, ולכן הוא **קירוב**: המסגרת, החריטה והשקיעה של Word אינן
 * `text-shadow`, והקרנינג אינו ניתן לציור כלל. הכיתוב אומר את זה במפורש —
 * פס שמתיימר להיות Word הוא פס שמשקר, ופס שאומר „כך ייראה” על אפקט שלא
 * צויר גרוע מאין פס.
 */
import { computed, inject, nextTick, onBeforeUnmount, reactive, ref, shallowRef, watch, type CSSProperties } from 'vue';
import type { SuperDoc } from 'superdoc';
import { hasRangeSelection, type FontAdvancedPatch, type SelectionReadiness } from '../../engine/font-advanced';
import { ACTIVE_SUPERDOC } from '../../engine/document-api';
import { readSelectionText } from '../../engine/font-preview';
import { createFontSample } from '../../composables/font-sample';
import { useDialogDrag } from '../../composables/dialog-drag';
import { useDialogDefaultAction } from '../../composables/dialog-default-action';
import { useFamilyPicker } from '../../composables/font-family-options';
import RibbonCombo from '../ribbon/common/RibbonCombo.vue';
import TriToggle, { type TriState } from './common/TriToggle.vue';

/* הדיאלוג נגרר בכותרת שלו — composables/dialog-drag.ts. */
const { dragStyle, startDialogDrag } = useDialogDrag();
/* Enter = הכפתור הראשי — composables/dialog-default-action.ts. */
const { onDialogEnter } = useDialogDefaultAction();

const DIALOG_TITLE = 'גופן מתקדם';
const INVALID_HINT = 'הערכים שהוקלדו אינם בטווח המותר — עיין בשדות המסומנים.';
const UNCHANGED = 'ללא שינוי';
/*
 * „ב-Word” ולא „תצוגה מקדימה” סתם — וזו הכרעה שנגזרת מהמדידה.
 *
 * הפס מצייר קירוב CSS גם לארבעת האפקטים שהעורך **אינו** מצייר (ראו `UNDRAWN`).
 * כיתוב שאומר „תצוגה מקדימה” לבדו הבטיח מראה שהמסמך שמתחתיו לעולם לא יקבל,
 * וההודעה שנוספה מתחת לרשת הכפתורים סותרת אותו במפורש.
 *
 * ההכרעה היא **לא** להוריד מהפס את האפקטים האלה: הוא המקום היחיד בעורך שבו
 * אפשר לראות מה נבחר, וזה שווה יותר דווקא כשהעורך בולע אותם. מה שמשתנה הוא
 * מה שהכיתוב מתחייב עליו — הקובץ, ולא המסך.
 */
const PREVIEW_CAPTION = 'כך ייראה ב-Word (קירוב)';
/*
 * „סמן טקסט במסמך” ולא „סגור, סמן, ופתח שוב”: הדיאלוג אינו חוסם את המסמך,
 * והוא שואל שוב בכל שינוי בחירה (ראו `watchSelection`) — כלומר הסימון פותח את
 * „אישור” בלי לסגור דבר, והשדות שכבר מולאו נשארים. עצה שמורה לסגור הייתה
 * אומרת למשתמש לזרוק את עבודתו בדיוק כשאין בכך צורך.
 */
const NO_SELECTION_HINT = 'העיצוב חל על הטקסט המסומן — סמן טקסט במסמך, והכפתור ייפתח.';
const NO_SELECTION_COUNT = 'אין טקסט מסומן';
const UNAVAILABLE_HINT = 'עיצוב גופן מתקדם אינו זמין בגרסה זו של המנוע.';
const UNAVAILABLE_COUNT = 'אינו זמין';

const props = defineProps<{
  isOpen: boolean;
  busy: boolean;
}>();

const emit = defineEmits<{
  close: [];
  submit: [patch: FontAdvancedPatch];
}>();

const rootRef = ref<HTMLElement | null>(null);

/** מחרוזות ריקות = „ללא שינוי". */
const charScale = ref('');
const letterSpacing = ref('');
const kerning = ref('');
const position = ref('');
const fontSizeCs = ref('');
const complexFontName = ref('');
const proofingLang = ref('');

/**
 * עשרת הבוליאנים כמפה, ולא עשרה `ref` נפרדים.
 *
 * המפה היא מה שמאפשר לתבנית לצייר אותם ב-`v-for` מתוך `EFFECTS`/`COMPLEX`,
 * ול„נקה הכל" לאפס את כולם בלולאה — עשר שורות שחוזרות על עצמן הן עשר
 * הזדמנויות לשכוח אחת בדיוק במקום שבו זה לא נראה (איפוס בפתיחה).
 */
type EffectKey =
  | 'dstrike'
  | 'outline'
  | 'shadow'
  | 'emboss'
  | 'imprint'
  | 'vanish'
  | 'boldCs'
  | 'italicCs'
  | 'complexScript'
  | 'rtl';

/** אפקט אחד: המפתח ב-patch, מה שכתוב על הכפתור, ומה שנאמר בטולטיפ. */
interface EffectControl {
  key: EffectKey;
  label: string;
  description: string;
}

/**
 * התוויות והסדר — עמודת האפקטים.
 *
 * התוויות קצרות מפני שהן ברשת כפתורים ולא בשורת תווית; ההסבר עובר לטולטיפ
 * (`description`), ושם גם מה שהתווית אינה יכולה לומר.
 */
const EFFECTS: readonly EffectControl[] = [
  { key: 'dstrike', label: 'קו חוצה כפול', description: 'שני קווים חוצים במקום אחד' },
  { key: 'outline', label: 'מסגרת לתו', description: 'קו מתאר במקום מילוי — אותיות חלולות' },
  { key: 'shadow', label: 'צל', description: 'צל נופל מאחורי האותיות' },
  { key: 'emboss', label: 'חרוט', description: 'האותיות נראות מורמות מהדף' },
  { key: 'imprint', label: 'שקוע', description: 'האותיות נראות שקועות בדף' },
  { key: 'vanish', label: 'טקסט מוסתר', description: 'הטקסט לא יוצג ולא יודפס' },
];

/**
 * התוויות והסדר — הליבה העברית.
 *
 * „מודגש”/„נטוי” כאן הם `w:bCs`/`w:iCs` ולא `w:b`/`w:i`, וזה ההבדל שקובע:
 * Word קורא הדגשה של כתב מורכב מ-`bCs`, וריצה עברית שנושאת `b` בלבד אינה
 * מוצגת מודגשת. הפער עצמו מתועד ב-docs/engine-gaps.md.
 */
const COMPLEX: readonly EffectControl[] = [
  { key: 'boldCs', label: 'מודגש', description: 'הדגשה של כתב מורכב (bCs) — מה ש-Word קורא ממנו בעברית' },
  { key: 'italicCs', label: 'נטוי', description: 'נטייה של כתב מורכב (iCs)' },
  { key: 'complexScript', label: 'כתב מורכב', description: 'מסמן את הריצה ככתב מורכב (cs)' },
  { key: 'rtl', label: 'מימין לשמאל', description: 'כיוון הריצה (rtl)' },
];

const effects = reactive<Record<EffectKey, TriState>>({
  dstrike: '',
  outline: '',
  shadow: '',
  emboss: '',
  imprint: '',
  vanish: '',
  boldCs: '',
  italicCs: '',
  complexScript: '',
  rtl: '',
});

/** האזהרה בתבנית קוראת אותו ישירות — זה המצב היחיד שמסתיר תוכן. */
const vanish = computed(() => effects.vanish);

/**
 * רשימת הגופנים, עם „ללא שינוי" בראשה. `complexFontName` הוא גם המצב הנוכחי
 * של הבורר וגם מה שנשלח, ולכן `''` הוא בדיוק „לא ייגע".
 */
const familyOptions = useFamilyPicker(() => complexFontName.value, UNCHANGED);

/**
 * הטקסט שבפס — מה שהמשתמש סימן, או פסוק כשאין בחירה.
 *
 * אותו מנגנון בדיוק של פס הדגימה בבורר הגופן, ומאותו טעם: `selection.current`
 * היא קריאה בלבד (1ms, `history` זהה לפניה ואחריה — נמדד), ולכן הפס אינו
 * נוגע במסמך ואינו קונה דרגת undo. ההנמקה המלאה ב-composables/font-sample.ts.
 */
const superdoc = inject(ACTIVE_SUPERDOC, shallowRef<SuperDoc | null>(null));
const sample = createFontSample({ read: () => readSelectionText(superdoc.value) });
/** ref ברמה העליונה — כך התבנית כותבת `sampleText` ולא `sample.text.value`. */
const sampleText = sample.text;

/**
 * מצב הבחירה — ההנמקה של ארבעת הערכים ב-`hasRangeSelection`.
 *
 * `'unknown'` **אינו** נועל דבר: דיאלוג שננעל על תשובה שטרם הגיעה גרוע
 * מדיאלוג שמסתמך על השער שב-`applyFontAdvanced` ממילא.
 */
const selectionState = shallowRef<SelectionReadiness>('unknown');

/** מונה סבבים — תשובה שאיחרה שייכת לסבב שלה, לא לזה שעל המסך. */
let selectionRound = 0;

async function probeSelection(refreshSample = false): Promise<void> {
  selectionRound += 1;
  const mine = selectionRound;
  const answer = await hasRangeSelection(superdoc.value);
  if (mine !== selectionRound) return;
  selectionState.value = answer;

  /*
   * ופס התצוגה המקדימה נקרא מחדש יחד איתו.
   *
   * `createFontSample` הוא תפס לסבב אחד — `begin()` שני באותו סבב אינו עושה
   * דבר — ולכן בלי השחרור כאן הפס היה נשאר על מה שנקרא ברגע הפתיחה. בדיוק
   * בזרימה שההודעה החוסמת מבקשת („סמן טקסט במסמך, והכפתור ייפתח”) הכפתור היה
   * נפתח והפס היה ממשיך להראות את פסוק ברירת המחדל, מתחת לכיתוב שמתחייב „כך
   * ייראה ב-Word”. וגם בכיוון השני: נפתח על בחירה א׳, המשתמש סימן ב׳, והפס
   * מראה את א׳ בזמן שההחלה תיפול על ב׳.
   */
  if (refreshSample) {
    sample.end();
    sample.begin();
  }
}

/**
 * ונשאל **שוב** בכל שינוי בחירה, לא רק בפתיחה.
 *
 * זה אינו ליטוש: הדיאלוג הזה אינו חוסם את המסמך — אין מאחוריו רקע, הוא נגרר
 * בכוונה, והמשתמש יכול ללחוץ בטקסט בזמן שהוא פתוח. תשובה שנקראה פעם אחת
 * בפתיחה הייתה מתיישנת בשני הכיוונים, ושניהם מחזירים בדיוק את הבאג שדווח:
 *
 * 1. נפתח **עם** בחירה, המשתמש הזיז את הדיאלוג ולחץ בטקסט כדי לקרוא אותו —
 *    הבחירה התכווצה, „אישור” עדיין פתוח, וההחלה נכשלת אחרי שהדיאלוג נסגר.
 * 2. נפתח **בלי** בחירה, המשתמש סימן טקסט כמו שההודעה ביקשה — והכפתור
 *    היה נשאר נעול לנצח על הודעה שהוא רואה שאינה נכונה עוד.
 *
 * `selectionchange` ולא מאזין על העורך: זה האירוע שהדפדפן מפעיל על כל שינוי
 * בבחירת המסמך, בעכבר ובמקלדת כאחד, ואינו מחייב להכיר את פנים המנוע.
 * ההשהיה מקבצת את הרצף שגרירה מייצרת — עשרות אירועים לגרירה אחת — לקריאה
 * אחת בסופה.
 */
const SELECTION_SETTLE_MS = 120;
let selectionTimer: ReturnType<typeof setTimeout> | null = null;

function onSelectionChanged(): void {
  if (selectionTimer !== null) clearTimeout(selectionTimer);
  selectionTimer = setTimeout(() => {
    selectionTimer = null;
    void probeSelection(true);
  }, SELECTION_SETTLE_MS);
}

function watchSelection(on: boolean): void {
  if (selectionTimer !== null) {
    clearTimeout(selectionTimer);
    selectionTimer = null;
  }
  if (on) document.addEventListener('selectionchange', onSelectionChanged);
  else document.removeEventListener('selectionchange', onSelectionChanged);
}

// מאזין שנשאר אחרי שהרכיב פורק הוא דליפה, וגם קריאה למנוע שכבר אינו שלנו.
onBeforeUnmount(() => watchSelection(false));

/*
 * והמסמך עצמו יכול להתחלף מתחת לדיאלוג הפתוח.
 *
 * `selectionchange` אינו נורה על מעבר בין לשוניות מסמך ולא על סגירת מסמך —
 * `App.vue` מאפס את המסמך הפעיל ופותח אחר, ואין בזה שינוי בבחירת ה-DOM. בלי
 * המעקב הזה „אישור” היה נשאר פתוח על תשובה ששייכת למסמך שכבר אינו על המסך,
 * וכל לחיצה הייתה נכשלת. זו אותה התיישנות שהמאזין נועד לסגור, בכניסה אחרת.
 */
watch(superdoc, () => {
  if (props.isOpen) void probeSelection(true);
});

/**
 * מה שהמנוע **כותב ואינו מצייר** — נמדד ב-Chrome אמיתי על ה-dist הארוז:
 * ה-docx יוצא קנוני, ו-`.superdoc-text-run` על המסך אינו משתנה. המדידה
 * הושוותה גם בצילום מסך בייט-בבייט, פקד אחד לכל ריצה — ארבעת האפקטים,
 * הקרנינג, וכל מחסנית הכתב המורכב. הרישום המלא ב-docs/engine-gaps.md.
 *
 * **הרשימה אינה „האפקטים”, אלא כל מה שנמדד.** גרסה קודמת מנתה את ארבעת
 * האפקטים בלבד, ובדיוק לידם ישבו „קרנינג” ו„גודל” של הגופן המורכב שגם הם
 * אינם מצוירים — כלומר הודעה שאמרה חצי אמת על אותה רשת פקדים.
 *
 * מה **שאינו** כאן, ולמה: `dstrike` מצויר (רק כקו בודד — נאמר בנפרד),
 * `rtl` מגיע ל-DOM כ-`dir="rtl"`, ומתיחה/ריווח/הרמה/טקסט מוסתר מצוירים
 * במלואם.
 */
const UNDRAWN_EFFECTS: readonly EffectKey[] = [
  'outline',
  'shadow',
  'emboss',
  'imprint',
  // מחסנית הכתב המורכב: המנוע מרנדר מ-`w:b`/`w:i` בלבד. הפער מתועד למעלה
  // בקובץ הפערים, ותיקון שלו נשלח כ-superdoc/docx-editor#3958.
  'boldCs',
  'italicCs',
  'complexScript',
];

/** פקדים שאינם כפתורי מיתוג, ולכן אינם ב-`effects` — אותה מדידה בדיוק. */
const UNDRAWN_FIELDS: readonly { chosen: () => boolean; label: string }[] = [
  { chosen: () => patch.value.kerningPt !== undefined, label: 'קרנינג' },
  { chosen: () => patch.value.fontSizeCsPt !== undefined, label: 'גודל הגופן המורכב' },
  { chosen: () => patch.value.complexFontName !== undefined, label: 'גופן מורכב' },
];

const ALL_EFFECT_LABELS = new Map<EffectKey, string>(
  [...EFFECTS, ...COMPLEX].map((effect) => [effect.key, effect.label]),
);

const undrawnChosen = computed(() => [
  ...UNDRAWN_EFFECTS.filter((key) => effects[key] === 'yes').map(
    (key) => ALL_EFFECT_LABELS.get(key) ?? key,
  ),
  ...UNDRAWN_FIELDS.filter((field) => field.chosen()).map((field) => field.label),
]);

/**
 * ההודעה מופיעה רק על מה שנבחר **עכשיו**, ובלשון „ייכתב ולא יצויר”.
 *
 * הערת קבע שיושבת מתחת לרשת הכפתורים היא הערה שנקראת פעם אחת ואז נעלמת
 * מהעין; מה שמופיע ברגע שנבחר האפקט הוא מה שנקרא. וזו אינה אזהרה על תקלה —
 * הקובץ ייצא נכון, ו-Word יראה בדיוק את מה שנבחר.
 */
const undrawnText = computed(() => {
  const parts: string[] = [];
  const chosen = undrawnChosen.value;
  if (chosen.length > 0) {
    const names = chosen.join(', ');
    /*
     * ניסוח שאינו נזקק להתאמת מין ומספר, ובכוונה.
     *
     * „ייכתב … אותו” נכון ל„צל” ושגוי ל„מסגרת לתו”, שהיא נקבה; „ייכתבו …
     * אותם” שגוי לפקד יחיד. שם הפקד הוא נתון — הוא נקרא מרשימת התוויות —
     * ולכן כל נוסח שמטה פועל אחריו נשען על ידיעה שאין כאן. הכותרת „מה שנבחר
     * …:” מוציאה את השמות מהמשפט, ומשאירה משפט אחד שנכון לכל הצירופים.
     *
     * ‏„נשמר בקובץ” ולא „יוצג ב-Word”: זה נכון גם ל„כתב מורכב” (`w:cs`), שאין
     * לו מראה משלו כלל — הוא מסמן לקורא איך לקרוא את שאר המאפיינים.
     */
    parts.push(`מה שנבחר נשמר בקובץ, והעורך אינו מצייר אותו: ${names}.`);
  }
  if (effects.dstrike === 'yes') parts.push('הקו החוצה הכפול מצויר בעורך כקו בודד.');
  return parts.join(' ');
});

function resetFields(): void {
  charScale.value = '';
  letterSpacing.value = '';
  kerning.value = '';
  position.value = '';
  fontSizeCs.value = '';
  complexFontName.value = '';
  proofingLang.value = '';
  for (const key of Object.keys(effects) as EffectKey[]) effects[key] = '';
}

watch(
  () => props.isOpen,
  async (open) => {
    if (!open) {
      // הפס משחרר את מה שקרא: פתיחה הבאה תקרא את הבחירה **שלה**, ולא תציג
      // לרגע את הטקסט של הפעם הקודמת.
      sample.end();
      watchSelection(false);
      return;
    }
    // איפוס בכל פתיחה: הדיאלוג אינו זוכר ערכים בין פעמים, כדי שאישור
    // לא-מכוון לא יחזור על עיצוב של פעם קודמת על בחירה חדשה.
    resetFields();
    sample.begin();
    // „לא ידוע” עד שהתשובה מגיעה: פתיחה אינה יורשת את מצב הפתיחה הקודמת.
    selectionState.value = 'unknown';
    void probeSelection();
    watchSelection(true);

    await nextTick();
    rootRef.value?.focus();
  },
  /*
   * `immediate`: הרכבה שנולדת פתוחה (בדיקת רכיב, ומצב שאין לו מניעה בקוד)
   * הייתה מדלגת על האיפוס ועל קריאת הטקסט לפס — כלומר דיאלוג בלי תצוגה
   * מקדימה, ובלי שום סימן לכך.
   */
  { immediate: true },
);

/**
 * v-model על `type="number"` ממיר אוטומטית למספר — הקלט עשוי להיות שני הסוגים.
 */
function asText(value: string | number): string {
  return typeof value === 'string' ? value : String(value);
}

function parseOptionalInt(value: string | number): number | undefined | null {
  const text = asText(value);
  if (text.trim() === '') return undefined;
  const parsed = Number(text);
  return Number.isFinite(parsed) ? Math.trunc(parsed) : null;
}

function parseOptionalNumber(value: string | number): number | undefined | null {
  const text = asText(value);
  if (text.trim() === '') return undefined;
  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : null;
}

/** '' → לא נשלח; 'yes' → true; 'no' → false. */
function tri(value: TriState): boolean | undefined {
  if (value === 'yes') return true;
  if (value === 'no') return false;
  return undefined;
}

const showError = computed(() => {
  // ארבעת השדות הראשונים מקבלים שלמים בלבד (המנוע כותב twips/חצאי-נקודות
  // מהכפלה); fontSizeCs מקבל חצאי נקודות (נמדד: 12.5 → szCs 25).
  for (const raw of [charScale.value, letterSpacing.value, kerning.value, position.value]) {
    if (asText(raw).trim() !== '' && parseOptionalInt(raw) === null) return true;
  }
  if (asText(fontSizeCs.value).trim() !== '' && parseOptionalNumber(fontSizeCs.value) === null) return true;
  return false;
});

/** „לא לשלוח / לשלוח"; הערך עבר כבר את שער ה-`canSubmit`. */
function isPresent(value: number | undefined | null): value is number {
  return value !== undefined && value !== null;
}

/**
 * ה-patch — מקור אחד גם ל„אישור" וגם למונה שבפוטר.
 *
 * זה מה שהופך את „N שינויים יוחלו" למדידה ולא להערכה: המספר הוא בדיוק מספר
 * המפתחות שיישלחו. שני חישובים נפרדים היו נפרדים גם כשמישהו יוסיף פקד.
 */
const patch = computed<FontAdvancedPatch>(() => {
  const built: FontAdvancedPatch = {};

  const charScaleValue = parseOptionalInt(charScale.value);
  if (isPresent(charScaleValue)) built.charScale = charScaleValue;
  const spacingValue = parseOptionalInt(letterSpacing.value);
  if (isPresent(spacingValue)) built.letterSpacingPt = spacingValue;
  const kerningValue = parseOptionalInt(kerning.value);
  if (isPresent(kerningValue)) built.kerningPt = kerningValue;
  const positionValue = parseOptionalInt(position.value);
  if (isPresent(positionValue)) built.positionPt = positionValue;
  const sizeCsValue = parseOptionalNumber(fontSizeCs.value);
  if (isPresent(sizeCsValue)) built.fontSizeCsPt = sizeCsValue;

  const dstrikeValue = tri(effects.dstrike);
  if (dstrikeValue !== undefined) built.dstrike = dstrikeValue;
  const outlineValue = tri(effects.outline);
  if (outlineValue !== undefined) built.outline = outlineValue;
  const shadowValue = tri(effects.shadow);
  if (shadowValue !== undefined) built.shadow = shadowValue;
  const embossValue = tri(effects.emboss);
  if (embossValue !== undefined) built.emboss = embossValue;
  const imprintValue = tri(effects.imprint);
  if (imprintValue !== undefined) built.imprint = imprintValue;
  const vanishValue = tri(effects.vanish);
  if (vanishValue !== undefined) built.vanish = vanishValue;
  const boldCsValue = tri(effects.boldCs);
  if (boldCsValue !== undefined) built.boldCs = boldCsValue;
  const italicCsValue = tri(effects.italicCs);
  if (italicCsValue !== undefined) built.italicCs = italicCsValue;
  const csValue = tri(effects.complexScript);
  if (csValue !== undefined) built.complexScript = csValue;
  const rtlValue = tri(effects.rtl);
  if (rtlValue !== undefined) built.rtl = rtlValue;

  if (complexFontName.value.trim() !== '') built.complexFontName = complexFontName.value.trim();
  if (proofingLang.value !== '') built.proofingLangBidi = proofingLang.value;

  return built;
});

const changeCount = computed(() => Object.keys(patch.value).length);

/**
 * „אישור" נעול על דיאלוג שאין בו מה להחיל — אותה הכרעה שכבר תוקנה
 * ב„ברירות מחדל למסמך": כפתור שנלחץ וסוגר בלי לעשות דבר הוא „ביטול" בתחפושת.
 */
/** מה שחוסם את הדיאלוג כולו, אם יש כזה. `''` = אין. */
const blockingText = computed(() => {
  if (selectionState.value === 'none') return NO_SELECTION_HINT;
  if (selectionState.value === 'unavailable') return UNAVAILABLE_HINT;
  return '';
});

const canSubmit = computed(
  () => !showError.value && changeCount.value > 0 && blockingText.value === '',
);

const countText = computed(() => {
  // הסיבה החוסמת גוברת על מניין השינויים: כשאין על מה להחיל, מספר השדות
  // שמולאו אינו המידע שחסר.
  if (selectionState.value === 'none') return NO_SELECTION_COUNT;
  if (selectionState.value === 'unavailable') return UNAVAILABLE_COUNT;
  if (changeCount.value === 0) return 'אין מה להחיל';
  return changeCount.value === 1 ? 'שינוי אחד יוחל' : `${changeCount.value} שינויים יוחלו`;
});

/**
 * ה-CSS של הפס — הקירוב, וגבולותיו.
 *
 * מה שנאמן: הגופן, הגודל, הריווח, המתיחה, ההרמה, ההדגשה, הנטייה, הכיוון והקו
 * החוצה — לכולם יש מקבילה ישירה ב-CSS. מה ש**מקורב**: מסגרת, צל, חריטה
 * ושקיעה, שב-Word הם רינדור של מנוע הטיפוגרפיה ולא הצללה. ומה שאינו מצויר
 * כלל: קרנינג — „מגודל X ומעלה" אינו מצב שאפשר להראות על מילה אחת.
 *
 * „כבוי" אינו מצייר דבר, וזה נכון: הפס מראה כיצד ייראה טקסט **רגיל** אחרי
 * החלת מה שנבחר, וטקסט רגיל ממילא אינו נושא צל.
 */
const previewStyle = computed<CSSProperties>(() => {
  const style: CSSProperties = {};

  if (complexFontName.value.trim() !== '') style.fontFamily = complexFontName.value.trim();

  const sizeCs = parseOptionalNumber(fontSizeCs.value);
  if (isPresent(sizeCs) && sizeCs > 0) style.fontSize = `${sizeCs}pt`;

  const spacing = parseOptionalInt(letterSpacing.value);
  if (isPresent(spacing)) style.letterSpacing = `${spacing}pt`;

  const scale = parseOptionalInt(charScale.value);
  if (isPresent(scale) && scale > 0) {
    style.display = 'inline-block';
    style.transform = `scaleX(${scale / 100})`;
    /*
     * ו-`max-width` הפוך לקנה המידה, אחרת הפס נחתך **משני** הצדדים.
     *
     * `transform` פועל אחרי הפריסה: התיבה נמדדת ברוחב הרגיל, שלוש הנקודות
     * נקבעות לפיו, ואז הציור נמתח מעבר לה — ומכיוון שהוא ממורכז הוא חורג
     * לשני הכיוונים, וה-`overflow: hidden` של המסגרת חותך את שניהם. נמדד
     * בכרום ב-140%: מצויר 686..1234 מול מסגרת 693..1227, כלומר גם ההתחלה
     * וגם הסוף נעלמו בלי שום סימן. עם התקרה ההפוכה (10000/140 = 71.43%)
     * התיבה **אחרי** המתיחה שווה בדיוק למסגרת: 702..1218.
     */
    style.maxWidth = `${(10000 / scale).toFixed(2)}%`;
  }

  const raise = parseOptionalInt(position.value);
  if (isPresent(raise) && raise !== 0) {
    style.position = 'relative';
    // חיובי = מוגבה, כמו במנוע — ולכן `top` שלילי.
    style.top = `${-raise}pt`;
  }

  if (effects.boldCs !== '') style.fontWeight = effects.boldCs === 'yes' ? 700 : 400;
  if (effects.italicCs !== '') style.fontStyle = effects.italicCs === 'yes' ? 'italic' : 'normal';
  if (effects.rtl !== '') style.direction = effects.rtl === 'yes' ? 'rtl' : 'ltr';

  if (effects.dstrike === 'yes') {
    style.textDecorationLine = 'line-through';
    style.textDecorationStyle = 'double';
  }

  if (effects.outline === 'yes') {
    style.WebkitTextStroke = '0.6px currentColor';
    style.color = 'transparent';
  }

  /*
   * שלושת ההצללות: צל נופל, חרוט מורם ושקוע שקוע. הכיוון הוא ההבדל, וזה גם
   * ההבדל שיש ל-Word עצמו — הוא פשוט מצייר אותו במנוע הטיפוגרפיה ולא בהצללה.
   */
  const shadows: string[] = [];
  if (effects.shadow === 'yes') shadows.push('1px 1px 1px var(--color-on-surface-variant)');
  if (effects.emboss === 'yes') {
    shadows.push('-1px -1px 0 var(--color-surface)', '1px 1px 0 var(--color-on-surface-variant)');
  }
  if (effects.imprint === 'yes') {
    shadows.push('1px 1px 0 var(--color-surface)', '-1px -1px 0 var(--color-on-surface-variant)');
  }
  if (shadows.length > 0) style.textShadow = shadows.join(', ');

  if (effects.vanish === 'yes') {
    // מוסתר אינו „בלתי נראה" בפס: פס ריק היה נראה כמו תקלה. זה מה ש-Word
    // מראה על טקסט מוסתר כשמסמנים „הצג הכול" — מעומעם, בקו מקווקו.
    style.opacity = 0.45;
    style.textDecorationLine = effects.dstrike === 'yes' ? 'line-through underline' : 'underline';
    style.textDecorationStyle = 'dotted';
  }

  return style;
});

function onSubmit(): void {
  if (props.busy || !canSubmit.value) return;
  emit('submit', patch.value);
}
</script>

<style scoped>
.fontadv-dialog {
  position: fixed;
  top: 140px;
  inset-inline-start: 40px;
  z-index: 2000;
  background: var(--color-surface);
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-sm);
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.16);
  /*
   * שתי עמודות במקום אחת: הרוחב גדל, והגובה — שהיה מעבר לתקרה — יורד מתחתיה.
   * ראו „הפריסה” בהערת הפתיחה.
   */
  width: 560px;
  max-width: calc(100vw - 32px);
  max-height: calc(100vh - 200px);
  /*
   * הגוף גולל, לא הדיאלוג.
   *
   * כאן היה `overflow-block: auto` על השורש. התוכן היה גבוה מהתקרה
   * `calc(100vh - 200px)`, והפוטר („אישור” / „ביטול”) הוא האחרון אחריו —
   * ולכן הוא נדחק מתחת לקצה המסך. נמדד בחלון של 600px: תחתית שורת הכפתורים
   * ב-896, כמעט 300px מתחת לתחתית המסך, והשורש לא נעשה אזור גלילה בפועל —
   * כלומר לא הייתה שום דרך להגיע אליה (scripts/qa/dialog-drag-qa.mjs מודד
   * בדיוק את זה).
   *
   * העמודה מקבעת את הכותרת (הידית לגרירה) ואת הפוטר בקצוות, ומשאירה את
   * הגלילה ל-`.fa-body` בלבד — שגם היא נדרשת עכשיו רק בחלון נמוך במיוחד.
   */
  display: flex;
  flex-direction: column;
  overflow: hidden;
  font-family: var(--font-main);
  user-select: none;
}

.fa-header {
  flex: 0 0 auto;
  display: flex;
  align-items: center;
  justify-content: space-between;
  border-block-end: 1px solid var(--color-outline-variant);
  padding: 6px 8px 6px 12px;
  background: var(--color-surface-container-high);
  border-radius: var(--radius-sm) var(--radius-sm) 0 0;
}

.fa-title {
  font-size: 12px;
  font-weight: 600;
  color: var(--color-on-surface);
}

.fa-close-btn {
  background: none;
  border: none;
  font-size: 14px;
  color: var(--color-on-surface-variant);
  cursor: pointer;
  padding: 4px 6px;
  border-radius: var(--radius-xs);
}

.fa-close-btn:hover {
  background: var(--word-btn-hover);
  color: var(--color-error);
}

.fa-body {
  /* `min-height: 0` הוא מה שמתיר לפריט flex להתכווץ מתחת לגובה תוכנו — בלעדיו
     העמודה הייתה נמתחת והפוטר היה יוצא מהמסגרת שוב. */
  flex: 1 1 auto;
  min-height: 0;
  overflow-y: auto;
  padding: 10px 12px;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.fa-columns {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 8px;
  align-items: start;
}

.fa-column {
  display: flex;
  flex-direction: column;
  gap: 8px;
  min-width: 0;
}

.fa-warning {
  margin: 0;
  font-size: 10.5px;
  line-height: 1.4;
  color: var(--color-error);
}

/*
  „הערה” ולא „אזהרה”, ולכן לא `--color-error`: שום דבר לא נכשל — הקובץ ייצא
  נכון, וזה מה שהעורך אינו מצייר. הצבע המשני הוא מה שמפריד בין השתיים,
  ו-`.fa-warning` שמעליה נשארת אדומה מפני שהיא כן על תוכן שנעלם.
*/
.fa-notice {
  margin: 0;
  font-size: 10.5px;
  line-height: 1.4;
  color: var(--color-on-surface-variant);
}

/*
  זו כן חוסמת — „אישור” נעול מאחוריה — ולכן היא נושאת את משקל האזהרה, ויושבת
  בראש הגוף ולא בתוך אחד המקטעים: היא נכונה לדיאלוג כולו.
*/
.fa-notice-blocking {
  padding: 6px 8px;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface-container-high);
  color: var(--color-on-surface);
  font-size: 11px;
}

.fa-group {
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  padding: 4px 8px 8px;
  margin: 0;
  min-width: 0;
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.fa-group legend {
  font-size: 11px;
  font-weight: 600;
  color: var(--color-on-surface-variant);
  padding-inline: 4px;
}

/* תווית ופקד, בשתי עמודות שנשארות מיושרות בין הקבוצות. */
.fa-fields {
  display: grid;
  grid-template-columns: auto 1fr;
  gap: 5px 8px;
  align-items: center;
}

.fa-input {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  min-width: 0;
}

/* הכפתורים זורמים וממלאים את הרוחב — שתיים או שלוש בשורה, לפי מה שנכנס. */
.fa-toggles {
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
}

.fa-label {
  font-size: 11px;
  color: var(--color-on-surface);
  flex-shrink: 0;
}

.fa-unit {
  font-size: 11px;
  color: var(--color-on-surface-variant);
}

.fa-number,
.fa-select {
  padding: 3px 6px;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  color: var(--color-on-surface);
  font-family: var(--font-main);
  font-size: 12px;
  outline: none;
  width: 72px;
}

.fa-select {
  width: auto;
  min-width: 110px;
}

.fa-number:focus,
.fa-select:focus {
  border-color: var(--word-blue);
  box-shadow: 0 0 0 1px var(--word-blue);
}

/*
  הבורר מגיע מהרצועה עם הגובה שלה (22px). כאן הוא יושב בשורת שדות של דיאלוג,
  וההשוואה היא מול `.fa-number` שלידו.
*/
.fa-combo :deep(.ribbon-combo-input) {
  height: 24px;
  font-size: 12px;
}

.fa-preview {
  display: flex;
  flex-direction: column;
  gap: 3px;
}

.fa-preview-caption {
  font-size: 10px;
  color: var(--color-on-surface-variant);
}

.fa-preview-strip {
  /*
   * גובה קבוע: הפס משנה גופן וגודל תוך כדי בחירה, וגובה שנגזר מהתוכן היה
   * מזיז את הפוטר בכל לחיצה — כלומר „אישור” שבורח מתחת לעכבר.
   */
  height: 46px;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 0 8px;
  overflow: hidden;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface-container-high);
  color: var(--color-on-surface);
}

.fa-preview-text {
  font-size: 15px;
  line-height: 1.3;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  /*
   * `min-width: 0` הוא מה שמפעיל את שלוש הנקודות: פריט flex אינו מתכווץ
   * מתחת לרוחב תוכנו (`min-width: auto`), ולכן הפס — שממורכז — גלש לשני
   * הצדדים ונחתך **בשניהם**. נמדד בכרום: הפסוק הופיע כ„סף וזהב … יהוצ”,
   * בלי התחלה ובלי סוף ובלי שום סימן לכך שהוא נחתך.
   */
  min-width: 0;
}

.fa-error {
  margin: 0;
  font-size: 11px;
  color: var(--color-error);
}

.fa-footer {
  flex: 0 0 auto;
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 8px 12px;
  border-block-start: 1px solid var(--color-outline-variant);
  background: var(--color-surface-container-high);
  border-radius: 0 0 var(--radius-sm) var(--radius-sm);
}

.fa-count {
  font-size: 11px;
  color: var(--color-on-surface-variant);
}

.fa-spacer {
  flex: 1 1 auto;
}

.fa-btn {
  padding: 4px 10px;
  font-size: 11px;
  font-family: var(--font-main);
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  color: var(--color-on-surface);
  cursor: pointer;
  transition: all 0.08s;
}

.fa-btn:hover:not(:disabled) {
  background: var(--word-btn-hover);
  border-color: var(--word-blue);
}

/* „נקה הכל” אינו פעולה על המסמך — הוא מנקה את הטופס, ולכן שקט משניהם. */
.fa-btn-quiet {
  border-color: transparent;
  color: var(--color-on-surface-variant);
}

/* הטקסט על הכפתור הממולא הוא `--color-on-primary` ולא לבן קבוע — ראו
   LinkDialog.vue. */
.fa-btn-primary {
  background: var(--word-blue);
  color: var(--color-on-primary);
  border-color: var(--word-blue);
  font-weight: 600;
}

.fa-btn-primary:hover:not(:disabled) {
  background: var(--word-blue-dark);
}

.fa-btn:disabled {
  opacity: 0.4;
  cursor: default;
}
</style>
