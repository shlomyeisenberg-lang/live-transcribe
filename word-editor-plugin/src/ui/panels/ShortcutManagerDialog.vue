<template>
  <Teleport to="body">
    <div
      v-if="isOpen"
      ref="rootRef"
      class="shortmgr-dialog"
      :style="dragStyle"
      role="dialog"
      aria-modal="true"
      :aria-label="DIALOG_TITLE"
      tabindex="-1"
      @keydown.esc.stop="onEsc"
      @keydown.enter="onDialogEnter"
    >
      <div
        class="sm-header dialog-drag-handle"
        @pointerdown="startDialogDrag"
      >
        <span class="sm-title">{{ DIALOG_TITLE }}</span>
        <button
          type="button"
          class="sm-close-btn"
          data-tip-title="סגור"
          data-tip-shortcut="Esc"
          aria-label="סגור את ניהול הקיצורים"
          @click="$emit('close')"
        >
          ✕
        </button>
      </div>

      <div class="sm-body">
        <!-- הרשימה -->
        <div class="sm-list-pane">
          <div
            class="sm-list"
            role="listbox"
            aria-label="הקיצורים שהגדרתם"
          >
            <button
              v-for="entry in list"
              :key="entry.id"
              type="button"
              class="sm-item"
              :class="{ 'sm-item--active': entry.id === editingId }"
              role="option"
              :aria-selected="entry.id === editingId"
              @pointerdown.prevent
              @click="edit(entry)"
            >
              <span class="sm-item-name">{{ entry.name }}</span>
              <span
                class="sm-item-combo"
                dir="ltr"
              >{{ comboLabel(entry.combo) }}</span>
              <span class="sm-item-preset">{{ shortcutSummary(entry) }}</span>
            </button>
            <p
              v-if="list.length === 0"
              class="sm-empty"
              role="note"
            >
              עדיין לא הגדרתם קיצורים. „קיצור חדש” מוסיף אחד.
            </p>
          </div>
          <button
            type="button"
            class="sm-btn sm-btn-wide"
            @pointerdown.prevent
            @click="startNew"
          >
            קיצור חדש
          </button>
        </div>

        <!-- הטופס -->
        <div class="sm-form">
          <div class="sm-row">
            <label
              for="sm-name"
              class="sm-label"
            >שם:</label>
            <input
              id="sm-name"
              v-model="name"
              class="sm-text"
              type="text"
              :maxlength="MAX_NAME_LENGTH"
              placeholder="למשל: כותרת קטע"
            >
          </div>

          <div class="sm-row">
            <span class="sm-label">צירוף:</span>
            <!--
              לוכד ולא שדה טקסט: מה שנשמר הוא המקש הפיזי (`event.code`), ואין
              דרך להקליד אותו — ראו ui/shortcuts/combo.ts.
            -->
            <button
              ref="recorderRef"
              type="button"
              class="sm-recorder"
              :class="{ 'sm-recorder--armed': recording }"
              :aria-label="recording ? 'ממתין לצירוף מקשים' : 'שינוי צירוף המקשים'"
              @pointerdown.prevent
              @click="arm"
              @keydown="onRecorderKey"
              @blur="recording = false"
            >
              <span
                v-if="recording"
                class="sm-recorder-hint"
              >לחצו על הצירוף…</span>
              <span
                v-else-if="combo"
                dir="ltr"
              >{{ comboLabel(combo) }}</span>
              <span
                v-else
                class="sm-recorder-hint"
              >לחצו כאן, ואז על הצירוף</span>
            </button>
          </div>

          <!-- קיצור „החל סגנון” אין לו ערכת עיצוב: הסגנון עצמו נערך בגלריה. -->
          <p
            v-if="editingStyle"
            class="sm-style-note"
            role="note"
          >
            הקיצור מחיל את הסגנון „{{ editingStyle.name }}”. את הסגנון עצמו משנים בגלריית הסגנונות.
          </p>

          <template v-else>
          <div class="sm-section">עיצוב</div>

          <div class="sm-row">
            <span class="sm-label">גופן:</span>
            <RibbonCombo
              v-model="fontFamily"
              class="sm-combo"
              :options="familyOptions"
              width="160px"
              title="הגופן שהקיצור יחיל"
              :placeholder="UNCHANGED"
              focus-return="stay"
            />
          </div>

          <div class="sm-row">
            <label
              for="sm-size"
              class="sm-label"
            >גודל:</label>
            <input
              id="sm-size"
              v-model="fontSize"
              class="sm-number"
              type="number"
              min="1"
              max="1638"
              step="0.5"
              :placeholder="UNCHANGED"
              aria-label="הגודל שהקיצור יחיל, בנקודות"
            >
            <span class="sm-unit">נק'</span>
          </div>

          <div
            v-for="slot in COLOR_SLOTS"
            :key="slot.field"
            class="sm-row"
          >
            <span class="sm-label">{{ slot.title }}:</span>
            <select
              v-model="colorMode[slot.field]"
              class="sm-select"
              :aria-label="`${slot.title} שהקיצור יחיל`"
            >
              <option value="keep">{{ UNCHANGED }}</option>
              <option value="none">{{ slot.clearText }}</option>
              <option value="pick">צבע:</option>
            </select>
            <input
              v-if="colorMode[slot.field] === 'pick'"
              v-model="colorValue[slot.field]"
              class="sm-color"
              type="color"
              :aria-label="`בחירת ${slot.title}`"
            >
          </div>

          <div class="sm-toggles">
            <div
              v-for="field in TOGGLE_FIELDS"
              :key="field"
              class="sm-row sm-row--toggle"
            >
              <span class="sm-label sm-label--short">{{ FIELD_TITLES[field] }}:</span>
              <select
                v-model="toggles[field]"
                class="sm-select sm-select--short"
                :aria-label="`${FIELD_TITLES[field]} — מה שהקיצור יקבע`"
              >
                <option value="keep">{{ UNCHANGED }}</option>
                <option value="on">כן</option>
                <option value="off">לא</option>
              </select>
            </div>
          </div>
          </template>

          <!--
            שלושה מצבים ולא שניים — ראו „שלושת המצבים של השורה” ב-<script>.
            אדום **רק** על תשובה שלילית למשהו שהמשתמש עשה.
          -->
          <p
            v-if="issue && issue.kind === 'conflict'"
            class="sm-error"
            role="alert"
          >
            {{ issue.message }}
          </p>
          <p
            v-else-if="issue && started"
            class="sm-note"
            role="note"
          >
            {{ issue.message }}
          </p>
          <p
            v-else
            class="sm-note"
            role="note"
          >
            לחיצה על הצירוף מחילה את העיצוב; לחיצה נוספת מחזירה את מה שהיה.
          </p>
        </div>
      </div>

      <div class="sm-footer">
        <button
          type="button"
          class="sm-btn sm-btn-primary"
          data-default-action
          :disabled="problem !== null"
          @pointerdown.prevent
          @click="onSave"
        >
          {{ editingId === null ? 'הוספה' : 'שמירה' }}
        </button>
        <button
          type="button"
          class="sm-btn"
          :disabled="editingId === null"
          @pointerdown.prevent
          @click="onDelete"
        >
          מחיקה
        </button>
        <button
          type="button"
          class="sm-btn"
          @pointerdown.prevent
          @click="$emit('close')"
        >
          סגירה
        </button>
      </div>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
/**
 * „ניהול קיצורים” — הדיאלוג שמגדיר קיצור אישי לערכת עיצוב.
 *
 * ## מה כאן ומה לא
 *
 * הקומפוננטה מחזיקה **טופס**, ולא את הרשימה: היא מקבלת את הקיצורים כ-prop
 * ופולטת כוונה (`save` / `remove`). האוסף, האימות והכתיבה לאחסון יושבים
 * ב-ui/shortcuts/custom-shortcuts.ts וב-App.vue — אותה הפרדה שכל דיאלוג אחר
 * במאגר עושה, ומאותו טעם: ההכרעות נבדקות כקלט-פלט ולא דרך הרכבת DOM.
 *
 * האימות אינו משוכפל כאן. `draftProblem` היא **אותה** פונקציה שהשמירה אוכפת,
 * ולכן אין מצב שבו הכפתור פעיל והשמירה מסרבת — או להפך.
 *
 * ## שלושת המצבים של כל שדה
 *
 * לכל תכונת עיצוב יש „ללא שינוי”, ולא רק ערך: ערכה שמגדילה גופן בלבד אינה
 * אמורה לגעת בצבע. לצבע יש מצב שלישי — „ללא/אוטומטי” — ולמתגים יש „כן” ו„לא”
 * מפורשים, כי ערכה שאינה מכבה מודגש נראית שבורה כשמחילים אותה בתוך טקסט
 * מודגש. שלושת המצבים הם הסיבה ש-`<select>` ולא תיבת סימון: תיבה שלישית-מצב
 * אינה קיימת בטופס נגיש.
 */
import { computed, nextTick, reactive, ref, watch } from 'vue';
import { useDialogDrag } from '../../composables/dialog-drag';
import { useDialogDefaultAction } from '../../composables/dialog-default-action';
import { useFamilyPicker } from '../../composables/font-family-options';
import RibbonCombo from '../ribbon/common/RibbonCombo.vue';
import { comboFromEvent, comboLabel, type KeyCombo } from '../shortcuts/combo';
import {
  FIELD_TITLES,
  TOGGLE_FIELDS,
  isEmptyPreset,
  type FormatPreset,
  type ToggleField,
} from '../shortcuts/format-preset';
import {
  MAX_NAME_LENGTH,
  draftIssue,
  type CustomShortcut,
  type CustomShortcutDraft,
  type TakenCombos,
  shortcutSummary,
  type StyleRef,
} from '../shortcuts/custom-shortcuts';

const { dragStyle, startDialogDrag } = useDialogDrag();
const { onDialogEnter } = useDialogDefaultAction();

const DIALOG_TITLE = 'ניהול קיצורים';
const UNCHANGED = 'ללא שינוי';

/** שני שדות הצבע, כדי שהשורה תיכתב פעם אחת. */
const COLOR_SLOTS = [
  { field: 'color' as const, title: 'צבע גופן', clearText: 'אוטומטי' },
  { field: 'highlight' as const, title: 'הדגשה', clearText: 'ללא הדגשה' },
];

type ColorField = (typeof COLOR_SLOTS)[number]['field'];
type ColorMode = 'keep' | 'none' | 'pick';
type ToggleMode = 'keep' | 'on' | 'off';

const props = defineProps<{
  isOpen: boolean;
  list: readonly CustomShortcut[];
  /**
   * צירופים שתפוסים בידי מערכת המאקרו — חתימה → שם הפריט.
   *
   * נמסר מבחוץ מפני שהוא תלוי במסמך הפתוח (`MacroKit` הוא של ה-session).
   * בלעדיו הדיאלוג היה מאשר צירוף שכבר תפוס, והתוצאה אינה שגיאה אלא שאחד
   * מהשניים שותק — ראו `signaturesOfShortcutText`.
   */
  taken?: TakenCombos;
}>();

const emit = defineEmits<{
  close: [];
  save: [draft: CustomShortcutDraft];
  remove: [id: string];
}>();

const rootRef = ref<HTMLElement | null>(null);
const recorderRef = ref<HTMLButtonElement | null>(null);

/** `null` = הוספה; מזהה = עריכה של רשומה קיימת. */
const editingId = ref<string | null>(null);
/** הסגנון של הרשומה שבעריכה, כשהיא קיצור „החל סגנון”. */
const editingStyle = ref<StyleRef | null>(null);
const name = ref('');
const combo = ref<KeyCombo | null>(null);
const fontFamily = ref('');
const fontSize = ref<string | number>('');
const colorMode = reactive<Record<ColorField, ColorMode>>({ color: 'keep', highlight: 'keep' });
/**
 * הצבע שנבחר בבורר. ברירת המחדל שחור/צהוב — הצבעים שלחיצה על „צבע גופן”
 * ו„הדגשה” ברצועה מחילה, כלומר מה שהמשתמש מצפה לו כשהוא עובר למצב „צבע”.
 * אינם CSS ולכן אינם נסרקים בשער הצבע הקשיח; הם נתון, בדיוק כמו פלטת הבורר.
 */
const colorValue = reactive<Record<ColorField, string>>({
  color: '#000000',
  highlight: '#FFFF00',
});
const toggles = reactive<Record<ToggleField, ToggleMode>>({
  bold: 'keep',
  italic: 'keep',
  underline: 'keep',
  strikethrough: 'keep',
});

const recording = ref(false);

const familyOptions = useFamilyPicker(() => fontFamily.value, UNCHANGED);

/** `type="number"` ב-v-model מחזיר מספר — ראו DocDefaultsDialog.vue. */
function asText(value: string | number): string {
  return typeof value === 'string' ? value : String(value);
}

/** הערכה שהטופס מתאר כרגע. */
const preset = computed<FormatPreset>(() => {
  const next: FormatPreset = {};

  if (fontFamily.value.trim() !== '') next.fontFamily = fontFamily.value.trim();

  const sizeText = asText(fontSize.value).trim();
  if (sizeText !== '') {
    const parsed = Number(sizeText);
    // ערך לא תקין אינו נכנס לערכה, ולכן `draftProblem` יאמר „לא נבחרה תכונה”
    // כשזה כל מה שהוקלד. עדיף על הודעה שנייה על אותו שדה.
    if (Number.isFinite(parsed) && parsed > 0) next.fontSizePt = parsed;
  }

  for (const slot of COLOR_SLOTS) {
    const mode = colorMode[slot.field];
    if (mode === 'none') next[slot.field] = null;
    else if (mode === 'pick') next[slot.field] = colorValue[slot.field].toUpperCase();
  }

  for (const field of TOGGLE_FIELDS) {
    const mode = toggles[field];
    if (mode !== 'keep') next[field] = mode === 'on';
  }

  return next;
});

const draft = computed<CustomShortcutDraft>(() => ({
  ...(editingId.value === null ? {} : { id: editingId.value }),
  name: name.value,
  combo: combo.value,
  preset: preset.value,
  ...(editingStyle.value ? { style: editingStyle.value } : {}),
}));

/**
 * אותו אימות שהשמירה אוכפת — ראו את ראש הקומפוננטה.
 *
 * ## שלושת המצבים של שורת ההודעה
 *
 * `draftIssue` מסווג בין „הטופס לא הושלם” ל„הבקשה נדחתה”, וההפרדה הזאת היא
 * מה שמונע את התקלה שנמדדה: הדיאלוג נפתח עם „יש לתת שם לקיצור” **באדום**
 * ובתוך `role="alert"`, לפני שאיש עשה כלום. זו הפרה של כלל מתועד — ראו
 * LinkDialog.vue: „השגיאה מוצגת רק אחרי שהמשתמש הקליד משהו”.
 *
 * לכן: התנגשות היא תמיד אדומה (היא תשובה למשהו שנעשה), חוסר מוצג כהנחיה
 * שקטה ורק אחרי שהמשתמש התחיל, ועל טופס נקי מוצג ההסבר מה הקיצור עושה.
 */
const issue = computed(() => draftIssue(draft.value, props.list, props.taken));
const problem = computed(() => issue.value?.message ?? null);

/**
 * האם המשתמש התחיל למלא. נגזר מהטופס עצמו ואינו דגל נפרד: דגל היה מצב שני
 * לאותה שאלה, ומצב שני צריך איפוס — עוד מקום לשכוח בו.
 */
const started = computed(
  () => name.value.trim() !== '' || combo.value !== null || !isEmptyPreset(preset.value),
);

/* ------------------------------------------------------------------ */
/* לכידת הצירוף                                                        */
/* ------------------------------------------------------------------ */

function arm(): void {
  recording.value = true;
  recorderRef.value?.focus();
}

/**
 * ההקשה בזמן לכידה.
 *
 * `Escape` מבטל את הלכידה **ואינו** סוגר את הדיאלוג — מי שנכנס ללכידה בטעות
 * צריך דרך לצאת ממנה, וסגירה של כל הדיאלוג הייתה מוחקת גם את הטופס. `Tab`
 * נשאר מקש ניווט: בלעדיו אין דרך לצאת מהלוכד במקלדת.
 *
 * `stopPropagation` **ו**-`preventDefault`: ההקשה לא אמורה להגיע לא למנתב
 * הקיצורים שעל `window` ולא לדפדפן. `Ctrl+O` בזמן לכידה הוא צירוף שנלכד, לא
 * פתיחת קובץ.
 */
function onRecorderKey(event: KeyboardEvent): void {
  if (!recording.value) return;
  if (event.key === 'Tab') return;

  event.stopPropagation();
  event.preventDefault();

  if (event.key === 'Escape') {
    recording.value = false;
    return;
  }

  const captured = comboFromEvent(event);
  // מודיפייר לבדו, או אות בלי מודיפייר: אין מה לשמור, והלכידה נשארת פתוחה
  // כדי שהמשתמש יוסיף Ctrl ולא ייצא בלי להבין למה לא נקלט דבר.
  if (!captured) return;

  combo.value = captured;
  recording.value = false;
}

function onEsc(): void {
  if (recording.value) {
    recording.value = false;
    return;
  }
  emit('close');
}

/* ------------------------------------------------------------------ */
/* מצב הטופס                                                          */
/* ------------------------------------------------------------------ */

function reset(): void {
  editingId.value = null;
  editingStyle.value = null;
  name.value = '';
  combo.value = null;
  fontFamily.value = '';
  fontSize.value = '';
  colorMode.color = 'keep';
  colorMode.highlight = 'keep';
  for (const field of TOGGLE_FIELDS) toggles[field] = 'keep';
  recording.value = false;
}

function startNew(): void {
  reset();
}

/** טוען רשומה קיימת לטופס. */
function edit(entry: CustomShortcut): void {
  reset();
  editingId.value = entry.id;
  name.value = entry.name;
  combo.value = entry.combo;
  if (entry.kind === 'apply-style') {
    editingStyle.value = entry.style;
    return;
  }
  fontFamily.value = entry.preset.fontFamily ?? '';
  fontSize.value = entry.preset.fontSizePt === undefined ? '' : String(entry.preset.fontSizePt);

  for (const slot of COLOR_SLOTS) {
    const value = entry.preset[slot.field];
    if (value === undefined) colorMode[slot.field] = 'keep';
    else if (value === null) colorMode[slot.field] = 'none';
    else {
      colorMode[slot.field] = 'pick';
      // בורר הצבע של הדפדפן דורש אותיות קטנות; ערכה נשמרת ברישיות גדולות
      // (`parseColor`), ובלי ההמרה הבורר נפתח על שחור במקום על הצבע שנשמר.
      colorValue[slot.field] = value.toLowerCase();
    }
  }

  for (const field of TOGGLE_FIELDS) {
    const value = entry.preset[field];
    toggles[field] = value === undefined ? 'keep' : value ? 'on' : 'off';
  }
}

function onSave(): void {
  if (problem.value !== null) return;
  emit('save', draft.value);
  reset();
}

function onDelete(): void {
  const id = editingId.value;
  if (id === null) return;
  emit('remove', id);
  reset();
}

watch(
  () => props.isOpen,
  async (open) => {
    if (!open) {
      recording.value = false;
      return;
    }
    reset();
    await nextTick();
    rootRef.value?.focus();
  },
);
</script>

<style scoped>
.shortmgr-dialog {
  position: fixed;
  top: 120px;
  inset-inline-start: 40px;
  z-index: 2000;
  background: var(--color-surface);
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-sm);
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.16);
  width: 580px;
  font-family: var(--font-main);
  user-select: none;
}

.sm-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  border-block-end: 1px solid var(--color-outline-variant);
  padding: 6px 8px 6px 12px;
  background: var(--color-surface-container-high);
  border-radius: var(--radius-sm) var(--radius-sm) 0 0;
}

.sm-title {
  font-size: 12px;
  font-weight: 600;
  color: var(--color-on-surface);
}

.sm-close-btn {
  background: none;
  border: none;
  font-size: 14px;
  color: var(--color-on-surface-variant);
  cursor: pointer;
  padding: 4px 6px;
  border-radius: var(--radius-xs);
}

.sm-close-btn:hover {
  background: var(--word-btn-hover);
  color: var(--color-error);
}

.sm-body {
  display: flex;
  gap: 12px;
  padding: 12px;
}

.sm-list-pane {
  display: flex;
  flex-direction: column;
  gap: 6px;
  width: 210px;
  flex-shrink: 0;
}

.sm-list {
  display: flex;
  flex-direction: column;
  gap: 2px;
  height: 236px;
  overflow-y: auto;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  padding: 4px;
  background: var(--color-surface-container-low);
}

.sm-item {
  display: grid;
  grid-template-columns: 1fr max-content;
  gap: 0 6px;
  text-align: start;
  background: none;
  border: 1px solid transparent;
  border-radius: var(--radius-xs);
  padding: 4px 6px;
  cursor: pointer;
  color: var(--color-on-surface);
  font-family: var(--font-main);
}

.sm-item:hover {
  background: var(--word-btn-hover);
}

.sm-item--active {
  background: var(--word-btn-hover);
  border-color: var(--word-blue);
}

.sm-item-name {
  font-size: 11.5px;
  font-weight: 600;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.sm-item-combo {
  font-size: 10.5px;
  color: var(--color-on-surface-variant);
  white-space: nowrap;
}

.sm-item-preset {
  grid-column: 1 / -1;
  font-size: 10px;
  color: var(--color-on-surface-variant);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.sm-empty {
  margin: 4px 6px;
  font-size: 10.5px;
  line-height: 1.4;
  color: var(--color-on-surface-variant);
}

.sm-form {
  display: flex;
  flex-direction: column;
  gap: 6px;
  flex: 1;
  min-width: 0;
}

.sm-section {
  font-size: 11px;
  font-weight: 600;
  color: var(--color-primary);
  margin-block-start: 2px;
}

.sm-row {
  display: flex;
  align-items: center;
  gap: 8px;
}

.sm-row--toggle {
  gap: 4px;
}

.sm-label {
  font-size: 11px;
  color: var(--color-on-surface);
  flex-shrink: 0;
  width: 62px;
}

.sm-label--short {
  width: 58px;
}

.sm-unit {
  font-size: 11px;
  color: var(--color-on-surface-variant);
}

.sm-text,
.sm-number {
  padding: 4px 8px;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  color: var(--color-on-surface);
  font-family: var(--font-main);
  font-size: 12px;
  outline: none;
}

.sm-text {
  flex: 1;
  min-width: 0;
}

.sm-number {
  width: 74px;
}

.sm-text:focus,
.sm-number:focus,
.sm-select:focus {
  border-color: var(--word-blue);
  box-shadow: 0 0 0 1px var(--word-blue);
}

.sm-select {
  padding: 3px 6px;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  color: var(--color-on-surface);
  font-family: var(--font-main);
  font-size: 11.5px;
  outline: none;
}

.sm-select--short {
  width: 92px;
}

.sm-color {
  width: 34px;
  height: 24px;
  padding: 0;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  cursor: pointer;
}

.sm-toggles {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 4px 10px;
}

/* הלוכד: תיבה בגובה שדה, שמשתנה גלויות כשהיא ממתינה להקשה. */
.sm-recorder {
  flex: 1;
  min-width: 0;
  text-align: start;
  padding: 4px 8px;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  color: var(--color-on-surface);
  font-family: var(--font-main);
  font-size: 12px;
  cursor: pointer;
}

.sm-recorder--armed {
  border-color: var(--word-blue);
  box-shadow: 0 0 0 1px var(--word-blue);
  background: var(--color-surface-container-high);
}

.sm-recorder-hint {
  color: var(--color-on-surface-variant);
}

.sm-combo :deep(.ribbon-combo-input) {
  height: 26px;
  font-size: 12px;
}

.sm-note,
.sm-error {
  margin: 2px 0 0;
  font-size: 10.5px;
  line-height: 1.4;
}

.sm-note {
  color: var(--color-on-surface-variant);
}

.sm-error {
  color: var(--color-error);
}

.sm-footer {
  display: flex;
  align-items: center;
  justify-content: flex-end;
  gap: 6px;
  padding: 8px 12px;
  border-block-start: 1px solid var(--color-outline-variant);
  background: var(--color-surface-container-high);
  border-radius: 0 0 var(--radius-sm) var(--radius-sm);
}

.sm-btn {
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

.sm-btn:hover:not(:disabled) {
  background: var(--word-btn-hover);
  border-color: var(--word-blue);
}

.sm-btn-wide {
  width: 100%;
}

.sm-btn-primary {
  background: var(--word-blue);
  color: var(--color-on-primary);
  border-color: var(--word-blue);
}

.sm-btn-primary:hover:not(:disabled) {
  background: var(--word-blue-dark);
}

.sm-btn:disabled {
  opacity: 0.4;
  cursor: default;
}

.sm-style-note {
  margin: 4px 0;
  font-size: 11px;
  line-height: 1.5;
  color: var(--color-on-surface-variant);
}
</style>
