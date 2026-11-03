<template>
  <Teleport to="body">
    <div
      v-if="isOpen"
      class="acorr-dialog"
      :style="dragStyle"
      role="dialog"
      aria-modal="true"
      :aria-label="TITLE"
      tabindex="-1"
      @keydown.esc.stop="$emit('close')"
      @keydown.enter="onDialogEnter"
    >
      <div
        class="acorr-header dialog-drag-handle"
        @pointerdown="startDialogDrag"
      >
        <span class="acorr-title">{{ TITLE }}</span>
        <button
          type="button"
          class="acorr-close-btn"
          data-tip-title="סגור"
          data-tip-shortcut="Esc"
          aria-label="סגור את אפשרויות התיקון האוטומטי"
          @click="$emit('close')"
        >
          ✕
        </button>
      </div>

      <div class="acorr-body">
        <label class="acorr-check acorr-master">
          <input
            v-model="draft.enabled"
            type="checkbox"
          >
          <span>הפעל תיקון אוטומטי בהקלדה</span>
        </label>

        <fieldset
          class="acorr-group"
          :disabled="!draft.enabled"
        >
          <legend>תיקון תוך כדי הקלדה</legend>
          <label
            v-for="option in TOGGLES"
            :key="option.key"
            class="acorr-check"
          >
            <input
              v-model="draft[option.key]"
              type="checkbox"
            >
            <span>{{ option.label }}</span>
            <small
              v-if="option.hint"
              class="acorr-hint"
              dir="auto"
            >{{ option.hint }}</small>
          </label>
        </fieldset>

        <fieldset
          class="acorr-group"
          :disabled="!draft.enabled"
        >
          <legend>
            <label class="acorr-check">
              <input
                v-model="draft.replaceAsYouType"
                type="checkbox"
                :disabled="!draft.enabled"
              >
              <span>החלף טקסט תוך כדי הקלדה</span>
            </label>
          </legend>

          <div class="acorr-edit">
            <label class="acorr-field">
              <span>החלף:</span>
              <input
                v-model="draftFrom"
                class="acorr-input"
                type="text"
                dir="auto"
                maxlength="64"
                aria-label="הטקסט להחלפה"
                :disabled="!draft.replaceAsYouType"
                @keydown.enter.stop.prevent="onUpsert"
              >
            </label>
            <label class="acorr-field">
              <span>ב:</span>
              <input
                v-model="draftTo"
                class="acorr-input"
                type="text"
                dir="auto"
                maxlength="255"
                aria-label="הטקסט שיבוא במקומו"
                :disabled="!draft.replaceAsYouType"
                @keydown.enter.stop.prevent="onUpsert"
              >
            </label>
          </div>

          <p
            v-if="message"
            class="acorr-message"
            role="alert"
          >
            {{ message }}
          </p>
          <small
            v-else-if="draftFrom !== ''"
            class="acorr-timing"
          >{{ timingHint }}</small>

          <div
            class="acorr-list"
            role="listbox"
            aria-label="טבלת ההחלפות"
            :aria-disabled="!draft.replaceAsYouType"
          >
            <button
              v-for="entry in draft.replacements"
              :key="entry.from"
              type="button"
              role="option"
              class="acorr-row"
              :class="{ 'is-selected': entry.from === selectedFrom }"
              :aria-selected="entry.from === selectedFrom"
              :disabled="!draft.replaceAsYouType"
              @click="onSelect(entry)"
            >
              <span
                class="acorr-cell"
                dir="auto"
              >{{ entry.from }}</span>
              <span
                class="acorr-arrow"
                aria-hidden="true"
              >←</span>
              <span
                class="acorr-cell"
                dir="auto"
              >{{ entry.to === '' ? '(מחיקה)' : entry.to }}</span>
              <span
                v-if="isImmediatePattern(entry.from)"
                class="acorr-mode"
                data-tip-title="מוחלף מיד כשהוקלד, גם בתוך מילה"
              >מיד</span>
            </button>
            <p
              v-if="draft.replacements.length === 0"
              class="acorr-empty"
            >
              הטבלה ריקה.
            </p>
          </div>

          <div class="acorr-table-actions">
            <button
              type="button"
              class="acorr-btn"
              :disabled="!draft.replaceAsYouType || draftFrom === ''"
              @pointerdown.prevent
              @click="onUpsert"
            >
              {{ exists ? 'עדכן' : 'הוסף' }}
            </button>
            <button
              type="button"
              class="acorr-btn"
              :disabled="!draft.replaceAsYouType || !exists"
              @pointerdown.prevent
              @click="onRemove"
            >
              מחק
            </button>
            <button
              type="button"
              class="acorr-btn"
              :disabled="!draft.replaceAsYouType"
              @pointerdown.prevent
              @click="onResetTable"
            >
              שחזר ברירת מחדל
            </button>
            <button
              type="button"
              class="acorr-btn acorr-load-hebrew"
              :disabled="!draft.replaceAsYouType"
              @pointerdown.prevent
              @click="onLoadHebrew"
            >
              טען קיצורים בעברית
            </button>
          </div>
          <small class="acorr-hint">רשימת הקיצורים אישית וכוללת גם למד ← למ״ד וטס ← ט״ס. הטעינה תחול רק לאחר אישור.</small>
        </fieldset>
      </div>

      <div class="acorr-footer">
        <button
          type="button"
          class="acorr-btn acorr-btn-primary"
          data-default-action
          @pointerdown.prevent
          @click="onOk"
        >
          אישור
        </button>
        <button
          type="button"
          class="acorr-btn"
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
 * אפשרויות התיקון האוטומטי, כמו „אפשרויות תיקון אוטומטי” של Word: מתג ראשי,
 * סוגי תיקון, וטבלת „החלף טקסט תוך כדי הקלדה” שאפשר לערוך.
 *
 * ## עריכה על עותק, ושמירה רק באישור
 *
 * הדיאלוג עובד על `draft` — עותק של ההגדרות שנלקח בפתיחה — ופולט `apply` רק
 * ב„אישור”. „ביטול” וסגירה זורקים אותו. הסיבה שונה מ-`ShulchanTyposDialog`
 * (שם גם ביטול שומר, כי זה מה שהמקור עושה): כאן ההגדרות **משפיעות מיד על כל
 * ההקלדה הבאה**, ושינוי חצי-גמור בטבלה היה מתקן טקסט לפי דפוס שהמשתמש עוד לא
 * סיים להקליד.
 *
 * ## מה הדיאלוג אינו מחליט
 *
 * מה תקין ומה לא — `normalizeEntries` ב-`engine/autocorrect.ts`, והוא גם מה
 * ששומר את הטבלה בעת הטעינה. כאן הוא נקרא פעם נוספת על הרשומה שהמשתמש מקליד,
 * כדי שהודעת השגיאה תגיע לפני האישור ולא כרשומה שנעלמה אחריו.
 */
import { computed, reactive, ref, watch } from 'vue';
import {
  DEFAULT_REPLACEMENTS,
  WORD_REPLACEMENTS,
  isImmediatePattern,
  isSelfFeedingEntry,
  mergeEntries,
  normalizeEntries,
  type AutoCorrectEntry,
  type AutoCorrectOptions,
} from '../../engine/autocorrect';
import { useDialogDrag } from '../../composables/dialog-drag';
import { useDialogDefaultAction } from '../../composables/dialog-default-action';

const { dragStyle, startDialogDrag } = useDialogDrag();
const { onDialogEnter } = useDialogDefaultAction();

const props = withDefaults(
  defineProps<{
    isOpen?: boolean;
    options: AutoCorrectOptions;
  }>(),
  { isOpen: false },
);

const emit = defineEmits<{
  (e: 'close'): void;
  (e: 'apply', options: AutoCorrectOptions): void;
}>();

const TITLE = 'תיקון אוטומטי';

type ToggleKey =
  | 'dashes'
  | 'spacedHyphen'
  | 'gershayim'
  | 'finalLetters'
  | 'twoInitialCaps'
  | 'capitalizeSentences'
  | 'capitalizeDays';

const TOGGLES: ReadonlyArray<{ key: ToggleKey; label: string; hint?: string }> = [
  { key: 'dashes', label: 'מקף כפול ← קו מפריד', hint: 'א--ב ← א—ב' },
  { key: 'spacedHyphen', label: 'מקף עם רווחים ← קו מפריד בינוני', hint: 'א - ב ← א – ב' },
  { key: 'gershayim', label: 'גרש וגרשיים בעברית', hint: 'רש"י ← רש״י' },
  { key: 'finalLetters', label: 'אותיות סופיות', hint: 'שלומ ← שלום' },
  { key: 'twoInitialCaps', label: 'תיקון שתי אותיות גדולות ראשונות', hint: 'TWo ← Two' },
  { key: 'capitalizeSentences', label: 'אות גדולה בתחילת משפט (באנגלית)' },
  { key: 'capitalizeDays', label: 'אות גדולה בשמות ימי השבוע (באנגלית)' },
];

/** עותק שאפשר לערוך: האפשרויות ב-`AutoCorrectOptions` קריאות בלבד, והטבלה בו משתנה. */
type Draft = {
  -readonly [K in Exclude<keyof AutoCorrectOptions, 'replacements'>]: AutoCorrectOptions[K];
} & { replacements: AutoCorrectEntry[] };

function copyOf(options: AutoCorrectOptions): Draft {
  return { ...options, replacements: options.replacements.map((entry) => ({ ...entry })) };
}

const draft = reactive(copyOf(props.options));
const draftFrom = ref('');
const draftTo = ref('');
const selectedFrom = ref<string | null>(null);
const message = ref('');

watch(
  () => props.isOpen,
  (open) => {
    if (!open) return;
    Object.assign(draft, copyOf(props.options));
    draftFrom.value = '';
    draftTo.value = '';
    selectedFrom.value = null;
    message.value = '';
  },
  { immediate: true },
);

/** האם הדפוס שבשדה כבר בטבלה — קובע אם הכפתור „יוסיף” או „יעדכן”. */
const exists = computed(() => draft.replacements.some((entry) => entry.from === draftFrom.value));

/**
 * מתי הדפוס שבשדה יוחלף. הכלל נגזר מהדפוס (`isImmediatePattern`) ולא נבחר, ולכן
 * הדיאלוג אומר אותו לפני האישור: מי שמגדיר `/` ← `"` צריך לדעת שזה יפעל גם
 * בתוך מילה, ומי שמגדיר `תשו` — שזה יחכה לרווח.
 */
/** בידוד כיוון (LRI…PDI): רצף סימנים בתוך משפט עברי מתהפך בלי בידוד — `-->` נראה `<--`. */
const LRI = String.fromCodePoint(0x2066);
const PDI = String.fromCodePoint(0x2069);

const timingHint = computed(() => {
  const from = draftFrom.value;
  if (!isImmediatePattern(from)) return 'יוחלף כשהמילה תסתיים — ברווח או בסימן פיסוק.';
  // דפוס שהוא תחילתו של דפוס סמל ארוך ממנו ממתין לתו הבא (`mayStillGrow`).
  const longer = draft.replacements.find(
    (entry) => entry.from.length > from.length && entry.from.startsWith(from) && isImmediatePattern(entry.from),
  );
  return longer
    ? `יוחלף בתו שאחריו, כי „${LRI}${longer.from}${PDI}” מתחיל בו. Ctrl+Z מיד אחרי ההחלפה מחזיר את מה שהוקלד.`
    : 'יוחלף מיד כשיוקלד, גם בתוך מילה. Ctrl+Z מיד אחרי ההחלפה מחזיר את מה שהוקלד.';
});

function onSelect(entry: AutoCorrectEntry): void {
  selectedFrom.value = entry.from;
  draftFrom.value = entry.from;
  draftTo.value = entry.to;
  message.value = '';
}

function onUpsert(): void {
  const candidate = { from: draftFrom.value, to: draftTo.value };
  const [valid] = normalizeEntries([candidate]);
  if (!valid) {
    message.value =
      draftFrom.value.trim() !== draftFrom.value
        ? 'הדפוס אינו יכול להתחיל או להסתיים ברווח.'
        : 'הדפוס אינו תקין: חובה לכתוב טקסט להחלפה, בשורה אחת.';
    return;
  }
  if (isSelfFeedingEntry(valid)) {
    message.value = 'ההחלפה מכילה את הטקסט שמוחלף, והייתה חוזרת בכל הקשה.';
    return;
  }

  message.value = '';
  const index = draft.replacements.findIndex((entry) => entry.from === valid.from);
  if (index >= 0) draft.replacements.splice(index, 1, valid);
  else draft.replacements.push(valid);
  selectedFrom.value = valid.from;
}

function onRemove(): void {
  const index = draft.replacements.findIndex((entry) => entry.from === draftFrom.value);
  if (index < 0) return;
  draft.replacements.splice(index, 1);
  draftFrom.value = '';
  draftTo.value = '';
  selectedFrom.value = null;
  message.value = '';
}

function onResetTable(): void {
  draft.replacements.splice(0, draft.replacements.length, ...DEFAULT_REPLACEMENTS.map((entry) => ({ ...entry })));
  draftFrom.value = '';
  draftTo.value = '';
  selectedFrom.value = null;
  message.value = '';
}

function onLoadHebrew(): void {
  const existing = new Set(draft.replacements.map((entry) => entry.from));
  draft.replacements = mergeEntries(draft.replacements, WORD_REPLACEMENTS.filter((entry) => !existing.has(entry.from)));
  message.value = '';
}

function onOk(): void {
  emit('apply', copyOf(draft));
}
</script>

<style scoped>
.acorr-dialog {
  position: fixed;
  top: 100px;
  inset-inline-start: 40px;
  z-index: 2000;
  background: var(--color-surface);
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-sm);
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.16);
  width: 440px;
  max-height: calc(100vh - 140px);
  display: flex;
  flex-direction: column;
  font-family: var(--font-main);
  user-select: none;
}

.acorr-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  border-block-end: 1px solid var(--color-outline-variant);
  padding: 6px 8px 6px 12px;
  background: var(--color-surface-container-high);
  border-radius: var(--radius-sm) var(--radius-sm) 0 0;
}

.acorr-title {
  font-size: 12px;
  font-weight: 600;
  color: var(--color-on-surface);
}

.acorr-close-btn {
  background: none;
  border: none;
  font-size: 14px;
  color: var(--color-on-surface-variant);
  cursor: pointer;
  padding: 4px 6px;
  border-radius: var(--radius-xs);
}

.acorr-close-btn:hover {
  background: var(--word-btn-hover);
  color: var(--color-error);
}

.acorr-body {
  padding: 12px;
  display: flex;
  flex-direction: column;
  gap: 10px;
  overflow-y: auto;
}

.acorr-group {
  margin: 0;
  padding: 8px 10px 10px;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  display: flex;
  flex-direction: column;
  gap: 5px;
  min-width: 0;
}

.acorr-group legend {
  padding: 0 4px;
  font-size: 11px;
  color: var(--color-on-surface-variant);
}

.acorr-group:disabled {
  opacity: 0.5;
}

.acorr-check {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 12px;
  color: var(--color-on-surface);
  cursor: pointer;
}

.acorr-master {
  font-weight: 600;
}

.acorr-hint {
  margin-inline-start: auto;
  font-size: 11px;
  color: var(--color-on-surface-variant);
}

.acorr-edit {
  display: flex;
  gap: 8px;
}

.acorr-field {
  display: flex;
  align-items: center;
  gap: 6px;
  flex: 1 1 0;
  min-width: 0;
  font-size: 12px;
  color: var(--color-on-surface);
}

.acorr-input {
  flex: 1 1 auto;
  min-width: 0;
  padding: 3px 6px;
  font-size: 12px;
  font-family: var(--font-main);
  color: var(--color-on-surface);
  background: var(--color-surface);
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  user-select: text;
}

.acorr-input:focus-visible {
  outline: 2px solid var(--color-primary);
  outline-offset: -1px;
}

.acorr-message {
  margin: 0;
  font-size: 11px;
  color: var(--color-error);
}

.acorr-list {
  height: 150px;
  overflow-y: auto;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
}

.acorr-row {
  display: flex;
  align-items: center;
  gap: 8px;
  width: 100%;
  padding: 3px 8px;
  border: none;
  background: transparent;
  font-family: var(--font-main);
  font-size: 12px;
  color: var(--color-on-surface);
  text-align: start;
  cursor: pointer;
}

.acorr-row:hover:not(:disabled) {
  background: var(--word-btn-hover);
}

.acorr-row.is-selected {
  background: var(--word-btn-active);
}

.acorr-row:focus-visible {
  outline: 2px solid var(--color-primary);
  outline-offset: -2px;
}

.acorr-cell {
  flex: 1 1 0;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.acorr-mode {
  flex: 0 0 auto;
  padding: 0 4px;
  font-size: 10px;
  color: var(--color-on-surface-variant);
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
}

.acorr-timing {
  font-size: 11px;
  color: var(--color-on-surface-variant);
}

.acorr-arrow {
  flex: 0 0 auto;
  color: var(--color-on-surface-variant);
}

.acorr-empty {
  margin: 0;
  padding: 10px;
  font-size: 11px;
  color: var(--color-on-surface-variant);
}

.acorr-table-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
}

.acorr-footer {
  display: flex;
  align-items: center;
  justify-content: flex-end;
  gap: 6px;
  padding: 8px 12px;
  border-block-start: 1px solid var(--color-outline-variant);
  background: var(--color-surface-container-high);
  border-radius: 0 0 var(--radius-sm) var(--radius-sm);
}

.acorr-btn {
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

.acorr-btn:hover:not(:disabled) {
  background: var(--word-btn-hover);
  border-color: var(--word-blue);
}

.acorr-btn-primary {
  background: var(--word-blue);
  color: var(--color-on-primary);
  border-color: var(--word-blue);
}

.acorr-btn-primary:hover:not(:disabled) {
  background: var(--word-blue-dark);
}

.acorr-btn:disabled {
  opacity: 0.4;
  cursor: default;
}
</style>
