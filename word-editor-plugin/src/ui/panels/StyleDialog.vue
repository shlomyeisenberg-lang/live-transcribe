<template>
  <Teleport to="body">
    <div
      v-if="isOpen"
      ref="rootRef"
      class="style-dialog"
      :class="{ 'style-dialog--short': short }"
      :style="dragStyle"
      role="dialog"
      aria-modal="true"
      :aria-label="title"
      tabindex="-1"
      data-testid="style-dialog"
      @keydown.esc.stop="onEscape"
      @keydown.enter="onDialogEnter"
    >
      <div
        class="sd-header dialog-drag-handle"
        @pointerdown="startDialogDrag"
      >
        <span class="sd-title">{{ title }}</span>
        <button
          type="button"
          class="sd-close-btn"
          data-tip-title="סגור"
          data-tip-shortcut="Esc"
          aria-label="סגור"
          @click="$emit('close')"
        >
          ✕
        </button>
      </div>

      <div class="sd-body">
          <div class="sd-row">
            <label
              for="sd-name"
              class="sd-label"
            >שם:</label>
            <input
              id="sd-name"
              ref="nameRef"
              v-model="name"
              class="sd-input sd-input--wide"
              type="text"
              :readonly="mode === 'modify'"
              :data-tip-title="mode === 'modify' ? 'לשינוי השם: לחצן ימני על הסגנון בגלריה' : undefined"
            >
          </div>

          <template v-if="!short">
            <div class="sd-row">
              <label
                for="sd-based"
                class="sd-label"
              >מבוסס על:</label>
              <select
                id="sd-based"
                v-model="basedOn"
                class="sd-input sd-input--wide"
              >
                <option value="">(ללא סגנון)</option>
                <option
                  v-for="style in basedOnChoices"
                  :key="style.id"
                  :value="style.id"
                >
                  {{ style.label }}
                </option>
              </select>
            </div>
            <div
              v-if="styleType === 'paragraph'"
              class="sd-row"
            >
              <label
                for="sd-next"
                class="sd-label"
              >הפסקה הבאה:</label>
              <select
                id="sd-next"
                v-model="next"
                class="sd-input sd-input--wide"
              >
                <option
                  v-if="mode === 'create'"
                  value=""
                >
                  (הסגנון הזה)
                </option>
                <option
                  v-for="style in paragraphStyles"
                  :key="style.id"
                  :value="style.id"
                >
                  {{ style.label }}
                </option>
              </select>
            </div>

            <fieldset class="sd-group">
              <legend>גופן</legend>
              <div class="sd-row">
                <RibbonCombo
                  v-model="fontFamily"
                  class="sd-combo"
                  :options="familyOptions"
                  width="160px"
                  title="גופן"
                  placeholder="גופן ערכת הנושא"
                  focus-return="stay"
                />
                <input
                  v-model="fontSize"
                  class="sd-input sd-number"
                  type="number"
                  min="1"
                  max="400"
                  step="0.5"
                  aria-label="גודל, בנקודות"
                >
                <button
                  v-for="toggle in TOGGLES"
                  :key="toggle.key"
                  type="button"
                  class="sd-toggle"
                  :class="{ active: toggles[toggle.key] }"
                  :aria-pressed="toggles[toggle.key]"
                  :data-tip-title="toggle.label"
                  :aria-label="toggle.label"
                  @pointerdown.prevent
                  @click="toggles[toggle.key] = !toggles[toggle.key]"
                >
                  <SvgIcon
                    :name="toggle.icon"
                    :size="14"
                  />
                </button>
                <input
                  v-model="color"
                  class="sd-color"
                  type="color"
                  aria-label="צבע גופן"
                  data-tip-title="צבע גופן"
                >
              </div>
            </fieldset>

            <fieldset
              v-if="styleType === 'paragraph'"
              class="sd-group"
            >
              <legend>פסקה</legend>
              <div class="sd-row">
                <span class="sd-label">יישור:</span>
                <button
                  v-for="align in ALIGNMENTS"
                  :key="align.value"
                  type="button"
                  class="sd-toggle"
                  :class="{ active: alignment === align.value }"
                  :aria-pressed="alignment === align.value"
                  :data-tip-title="align.label"
                  :aria-label="align.label"
                  @pointerdown.prevent
                  @click="alignment = align.value"
                >
                  <SvgIcon
                    :name="align.icon"
                    :size="14"
                  />
                </button>
              </div>
              <div class="sd-row">
                <label
                  for="sd-before"
                  class="sd-label"
                >ריווח לפני:</label>
                <input
                  id="sd-before"
                  v-model="spaceBefore"
                  class="sd-input sd-number"
                  type="number"
                  min="0"
                  step="1"
                >
                <span class="sd-unit">נק'</span>
                <label
                  for="sd-after"
                  class="sd-label sd-label--inline"
                >אחרי:</label>
                <input
                  id="sd-after"
                  v-model="spaceAfter"
                  class="sd-input sd-number"
                  type="number"
                  min="0"
                  step="1"
                >
                <span class="sd-unit">נק'</span>
              </div>
              <div class="sd-row">
                <label
                  for="sd-line"
                  class="sd-label"
                >מרווח שורות:</label>
                <select
                  id="sd-line"
                  v-model="lineSpacing"
                  class="sd-input"
                >
                  <option
                    v-for="option in lineOptions"
                    :key="option"
                    :value="option"
                  >
                    {{ option }}
                  </option>
                </select>
              </div>
              <div class="sd-row">
                <label
                  for="sd-start"
                  class="sd-label"
                >הזחה לפני טקסט:</label>
                <input
                  id="sd-start"
                  v-model="indentStart"
                  class="sd-input sd-number"
                  type="number"
                  min="0"
                  step="0.25"
                >
                <span class="sd-unit">ס"מ</span>
                <label
                  for="sd-end"
                  class="sd-label sd-label--inline"
                >אחרי:</label>
                <input
                  id="sd-end"
                  v-model="indentEnd"
                  class="sd-input sd-number"
                  type="number"
                  min="0"
                  step="0.25"
                >
                <span class="sd-unit">ס"מ</span>
              </div>
              <div class="sd-row">
                <label
                  for="sd-first"
                  class="sd-label"
                >שורה ראשונה:</label>
                <input
                  id="sd-first"
                  v-model="firstLine"
                  class="sd-input sd-number"
                  type="number"
                  step="0.25"
                >
                <span class="sd-unit">ס"מ (שלילי = תלויה)</span>
              </div>
              <div class="sd-row">
                <label
                  for="sd-outline"
                  class="sd-label"
                >רמת חלוקה:</label>
                <select
                  id="sd-outline"
                  v-model.number="outlineLevel"
                  class="sd-input"
                  data-tip-title="רמת חלוקה"
                  data-tip-desc="מה ש-Word קורא כדי לזהות כותרת — בחלונית הניווט ובתוכן העניינים"
                >
                  <option :value="BODY_TEXT_LEVEL">
                    גוף טקסט
                  </option>
                  <option
                    v-for="level in 9"
                    :key="level"
                    :value="level - 1"
                  >
                    רמה {{ level }}
                  </option>
                </select>
              </div>
            </fieldset>

            <div class="sd-row">
              <span class="sd-label">קיצור מקשים:</span>
              <!-- לוכד ולא שדה טקסט, כמו ב„ניהול קיצורים”: נשמר המקש הפיזי. -->
              <button
                ref="recorderRef"
                type="button"
                class="sd-input sd-recorder"
                :class="{ 'sd-recorder--armed': recording }"
                data-testid="style-shortcut-recorder"
                :aria-label="recording ? 'ממתין לצירוף מקשים' : 'קיצור מקשים לסגנון'"
                @pointerdown.prevent
                @click="arm"
                @keydown="onRecorderKey"
                @blur="recording = false"
              >
                <span
                  v-if="recording"
                  class="sd-recorder-hint"
                >לחצו על הצירוף…</span>
                <span
                  v-else-if="combo"
                  dir="ltr"
                >{{ comboLabel(combo) }}</span>
                <span
                  v-else
                  class="sd-recorder-hint"
                >ללא — לחצו כאן כדי להגדיר</span>
              </button>
              <button
                v-if="combo"
                type="button"
                class="sd-toggle"
                data-tip-title="הסר את הקיצור"
                aria-label="הסר את הקיצור"
                @pointerdown.prevent
                @click="combo = null"
              >
                ✕
              </button>
            </div>

            <div
              class="sd-preview"
              aria-hidden="true"
            >
              <div
                class="sd-preview-text"
                :style="previewStyle"
              >
                {{ PREVIEW_TEXT }}
              </div>
            </div>
            <p class="sd-note">
              השינוי חל על כל הפסקאות בסגנון ונשמר במסמך. אפשר לבטל אותו באמצעות „בטל”.
            </p>
            <!--
              „מסמכים חדשים המבוססים על תבנית זו” של Word: העיצוב נשמר גם לכל
              מסמך חדש — composables/new-document-defaults.ts. רק לסגנון שקיים
              בתבנית; לסגנון שנוצר במסמך הזה אין מה לשנות שם.
            -->
            <label
              v-if="newDocumentsOption"
              class="sd-check"
            >
              <input
                v-model="forNewDocuments"
                type="checkbox"
                data-testid="style-new-documents"
              >
              גם לכל מסמך חדש
            </label>
          </template>

        <p
          v-if="error || invalid"
          class="sd-error"
          role="alert"
        >
          {{ error || invalid }}
        </p>
      </div>

      <div class="sd-footer">
        <button
          type="button"
          class="sd-btn sd-btn-primary"
          data-default-action
          :disabled="busy || invalid !== null"
          @pointerdown.prevent
          @click="onSubmit"
        >
          אישור
        </button>
        <button
          type="button"
          class="sd-btn"
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
 * „שנה סגנון” / „צור סגנון” / שינוי שם — הדיאלוג בלבד.
 *
 * הדיאלוג מציג ערכים שלמים (העיצוב האפקטיבי, כולל ירושה) ומחזיר ערכים
 * שלמים. מה מתוכם נכתב לסגנון — רק מה ששונה מהירושה — מחושב אצל מי שפתח
 * אותו (engine/style-from-selection.ts, `formatPatch`), כי רק הוא מחזיק את
 * גיליון הסגנונות ויודע מה הבסיס שנבחר.
 */
import { computed, nextTick, reactive, ref, watch, type CSSProperties } from 'vue';
import SvgIcon from '../icons/SvgIcon.vue';
import RibbonCombo from '../ribbon/common/RibbonCombo.vue';
import { useDialogDrag } from '../../composables/dialog-drag';
import { useDialogDefaultAction } from '../../composables/dialog-default-action';
import { useFamilyPicker } from '../../composables/font-family-options';
import { BODY_TEXT_LEVEL, type StyleAlignment, type StyleFormat } from '../../engine/docx-style-edit';
import { comboFromEvent, comboLabel, type KeyCombo } from '../shortcuts/combo';

export type StyleDialogMode = 'modify' | 'create' | 'rename';

export interface StyleChoice {
  id: string;
  label: string;
  type: string;
}

export interface StyleDialogResult {
  name: string;
  basedOn: string | null;
  next: string | null;
  format: StyleFormat;
  /** קיצור המקשים שנבחר (`null` = אין). נבדק מול `shortcut` בפתיחה כדי לדעת אם שונה. */
  combo: KeyCombo | null;
  /** לשמור את העיצוב גם לכל מסמך חדש. */
  forNewDocuments: boolean;
}

const { dragStyle, startDialogDrag } = useDialogDrag();
const { onDialogEnter } = useDialogDefaultAction();

const props = withDefaults(
  defineProps<{
    isOpen: boolean;
    busy?: boolean;
    mode: StyleDialogMode;
    styleId?: string | null;
    styleName?: string;
    styleType?: string;
    styles?: readonly StyleChoice[];
    basedOn?: string | null;
    next?: string | null;
    format?: StyleFormat;
    error?: string | null;
    /** קיצור המקשים של הסגנון כרגע. */
    shortcut?: KeyCombo | null;
    /** למה צירוף אינו זמין (תפוס, או בלי Ctrl/Alt), או `null`. */
    comboProblem?: (combo: KeyCombo) => string | null;
    /** להציע „גם לכל מסמך חדש” — רק בשינוי של סגנון שקיים בתבנית. */
    newDocumentsOption?: boolean;
  }>(),
  {
    busy: false,
    styleId: null,
    styleName: '',
    styleType: 'paragraph',
    styles: () => [],
    basedOn: null,
    next: null,
    format: () => ({}),
    error: null,
    shortcut: null,
    comboProblem: () => null,
    newDocumentsOption: false,
  },
);

const emit = defineEmits<{
  close: [];
  submit: [result: StyleDialogResult];
}>();

const PREVIEW_TEXT = 'בראשית ברא אלהים את השמים ואת הארץ';

const TOGGLES = [
  { key: 'bold', icon: 'bold', label: 'מודגש' },
  { key: 'italic', icon: 'italic', label: 'נטוי' },
  { key: 'underline', icon: 'underline', label: 'קו תחתון' },
] as const;

const ALIGNMENTS: readonly { value: StyleAlignment; icon: string; label: string }[] = [
  { value: 'right', icon: 'alignRight', label: 'יישור לימין' },
  { value: 'center', icon: 'alignCenter', label: 'מרכז' },
  { value: 'left', icon: 'alignLeft', label: 'יישור לשמאל' },
  { value: 'justify', icon: 'alignJustify', label: 'יישור לשני הצדדים' },
];

const BASE_LINE_OPTIONS = ['1', '1.15', '1.5', '2', '2.5', '3'];

const rootRef = ref<HTMLElement | null>(null);
const nameRef = ref<HTMLInputElement | null>(null);

const name = ref('');
const basedOn = ref('');
const next = ref('');
const fontFamily = ref('');
const fontSize = ref<string | number>('');
const toggles = reactive({ bold: false, italic: false, underline: false });
const color = ref('#000000');
const alignment = ref<StyleAlignment>('right');
const spaceBefore = ref<string | number>('0');
const spaceAfter = ref<string | number>('0');
const lineSpacing = ref('1');
const indentStart = ref<string | number>('0');
const indentEnd = ref<string | number>('0');
const firstLine = ref<string | number>('0');
const outlineLevel = ref<number>(BODY_TEXT_LEVEL);
const combo = ref<KeyCombo | null>(null);
const forNewDocuments = ref(false);
const recording = ref(false);
const recorderRef = ref<HTMLButtonElement | null>(null);

/**
 * מדליקה את הלכידה **ומעבירה מיקוד ללוכד**. `@pointerdown.prevent` (שלא
 * ייגנב המיקוד מהמסמך) מונע גם מהכפתור עצמו לקבל אותו, וההקשה הלכה לשורש
 * הדיאלוג — כלומר לחיצה בעכבר לא הקליטה דבר (נמצא ב-QA). כמו `arm` ב„ניהול
 * קיצורים”.
 */
function arm(): void {
  recording.value = true;
  recorderRef.value?.focus();
}

function onRecorderKey(event: KeyboardEvent): void {
  if (!recording.value || event.key === 'Tab') return;
  event.stopPropagation();
  event.preventDefault();
  if (event.key === 'Escape') {
    recording.value = false;
    return;
  }
  const captured = comboFromEvent(event);
  // מודיפייר לבדו: הלכידה נשארת פתוחה עד שיגיע מקש.
  if (!captured) return;
  combo.value = captured;
  recording.value = false;
}

/**
 * מה שהוצג בפתיחה לשדות שאין להם ערך „לא מוצהר” בפקד: צבע ויישור. סגנון
 * שאינו מצהיר צבע מוצג שחור, ובלי ההשוואה הזאת „אישור” היה כותב לו שחור
 * מפורש — כלומר משנה את הסגנון בלי שהמשתמש נגע בדבר.
 */
let shownColor = '';
let shownAlignment: StyleAlignment = 'right';

const short = computed(() => props.mode === 'rename');

const title = computed(() => {
  switch (props.mode) {
    case 'create':
      return 'סגנון חדש';
    case 'rename':
      return 'שינוי שם סגנון';
    default:
      return `שינוי הסגנון „${props.styleName}”`;
  }
});

const familyOptions = useFamilyPicker(() => fontFamily.value);

const paragraphStyles = computed(() => props.styles.filter((style) => style.type === 'paragraph'));

/** סגנון אינו מבוסס על עצמו, ומבוסס רק על סגנון מאותו סוג. */
const basedOnChoices = computed(() =>
  props.styles.filter((style) => style.type === props.styleType && style.id !== props.styleId),
);

/** ערך שאינו בסולם (1.08 של Word) נשאר בחירה — לא נעלם בפתיחה. */
const lineOptions = computed(() =>
  BASE_LINE_OPTIONS.includes(lineSpacing.value) ? BASE_LINE_OPTIONS : [lineSpacing.value, ...BASE_LINE_OPTIONS],
);

const text = (value: number | null | undefined, fallback: string): string =>
  value === null || value === undefined ? fallback : String(value);

watch(
  () => props.isOpen,
  async (open) => {
    if (!open) return;
    const format = props.format;
    name.value = props.styleName;
    basedOn.value = props.basedOn ?? '';
    next.value = props.next ?? '';
    fontFamily.value = format.fontFamily ?? '';
    fontSize.value = text(format.fontSizePt, '');
    toggles.bold = format.bold === true;
    toggles.italic = format.italic === true;
    toggles.underline = format.underline === true;
    color.value = `#${format.color ?? '000000'}`;
    alignment.value = format.alignment ?? 'right';
    shownColor = color.value;
    shownAlignment = alignment.value;
    spaceBefore.value = text(format.spaceBeforePt, '0');
    spaceAfter.value = text(format.spaceAfterPt, '0');
    lineSpacing.value = text(format.lineSpacing, '1');
    indentStart.value = text(format.indentStartCm, '0');
    indentEnd.value = text(format.indentEndCm, '0');
    firstLine.value = text(format.firstLineCm, '0');
    outlineLevel.value = format.outlineLevel ?? BODY_TEXT_LEVEL;
    combo.value = props.shortcut;
    forNewDocuments.value = false;
    recording.value = false;
    await nextTick();
    if (props.mode === 'create' || props.mode === 'rename') {
      nameRef.value?.focus();
      nameRef.value?.select();
    } else {
      rootRef.value?.focus();
    }
  },
  { immediate: true },
);

/** ‏v-model על `type="number"` מחזיר מספר — ראו DocDefaultsDialog.vue. */
function parse(value: string | number): number | null {
  const raw = typeof value === 'number' ? value : value.trim() === '' ? NaN : Number(value);
  return Number.isFinite(raw) ? raw : null;
}

const invalid = computed<string | null>(() => {
  if (name.value.trim() === '') return 'לסגנון חייב להיות שם.';
  if (short.value) return null;
  const size = parse(fontSize.value);
  if (fontSize.value !== '' && (size === null || size <= 0)) return 'הגודל אינו תקין.';
  for (const value of [spaceBefore.value, spaceAfter.value, indentStart.value, indentEnd.value]) {
    const parsed = parse(value);
    if (parsed === null || parsed < 0) return 'הערכים שהוקלדו אינם תקינים.';
  }
  if (parse(firstLine.value) === null) return 'הערכים שהוקלדו אינם תקינים.';
  if (combo.value && combo.value !== props.shortcut) return props.comboProblem(combo.value);
  return null;
});

const previewStyle = computed<CSSProperties>(() => {
  const size = parse(fontSize.value);
  return {
    fontFamily: fontFamily.value ? `"${fontFamily.value}"` : undefined,
    // התצוגה המקדימה מוגבלת לגובה התיבה; גודל אמיתי של 36 נקודות היה בולע אותה.
    fontSize: size ? `${Math.min(Math.max(size, 8), 24)}pt` : undefined,
    fontWeight: toggles.bold ? '700' : '400',
    fontStyle: toggles.italic ? 'italic' : 'normal',
    textDecoration: toggles.underline ? 'underline' : 'none',
    color: color.value,
    textAlign: alignment.value === 'justify' ? 'justify' : alignment.value,
  };
});

function onEscape(): void {
  if (recording.value) {
    recording.value = false;
    return;
  }
  emit('close');
}

function onSubmit(): void {
  if (props.busy || invalid.value !== null) return;
  const format: StyleFormat = short.value
    ? {}
    : {
        fontFamily: fontFamily.value.trim() || undefined,
        fontSizePt: parse(fontSize.value) ?? undefined,
        bold: toggles.bold,
        italic: toggles.italic,
        underline: toggles.underline,
        color:
          props.format.color !== undefined || color.value !== shownColor
            ? color.value.replace(/^#/, '').toUpperCase()
            : undefined,
        ...(props.styleType === 'paragraph'
          ? {
              alignment:
                props.format.alignment !== undefined || alignment.value !== shownAlignment
                  ? alignment.value
                  : undefined,
              spaceBeforePt: parse(spaceBefore.value) ?? 0,
              spaceAfterPt: parse(spaceAfter.value) ?? 0,
              lineSpacing: Number(lineSpacing.value),
              indentStartCm: parse(indentStart.value) ?? 0,
              indentEndCm: parse(indentEnd.value) ?? 0,
              firstLineCm: parse(firstLine.value) ?? 0,
              outlineLevel: outlineLevel.value,
            }
          : {}),
      };
  emit('submit', {
    name: name.value.trim(),
    basedOn: basedOn.value || null,
    next: next.value || null,
    format,
    combo: short.value ? props.shortcut : combo.value,
    forNewDocuments: props.newDocumentsOption && !short.value && forNewDocuments.value,
  });
}
</script>

<style scoped>
.style-dialog {
  position: fixed;
  --dialog-top: min(120px, 15vh);
  top: var(--dialog-top);
  display: flex;
  flex-direction: column;
  max-height: calc(100vh - var(--dialog-top) - 16px);
  overflow: hidden;
  inset-inline-start: 40px;
  z-index: 2000;
  background: var(--color-surface);
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-sm);
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.16);
  width: 420px;
  max-width: calc(100vw - 32px);
  font-family: var(--font-main);
  user-select: none;
}

.style-dialog--short {
  width: 340px;
}

.sd-header {
  flex-shrink: 0;
  display: flex;
  align-items: center;
  justify-content: space-between;
  border-block-end: 1px solid var(--color-outline-variant);
  padding: 6px 8px 6px 12px;
  background: var(--color-surface-container-high);
  border-radius: var(--radius-sm) var(--radius-sm) 0 0;
}

.sd-title {
  font-size: 12px;
  font-weight: 600;
  color: var(--color-on-surface);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.sd-close-btn {
  background: none;
  border: none;
  font-size: 14px;
  color: var(--color-on-surface-variant);
  cursor: pointer;
  padding: 4px 6px;
  border-radius: var(--radius-xs);
}

.sd-close-btn:hover {
  background: var(--word-btn-hover);
  color: var(--color-error);
}

.sd-body {
  min-height: 0;
  overflow-y: auto;
  padding: 12px;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.sd-row {
  display: flex;
  align-items: center;
  gap: 6px;
  flex-wrap: wrap;
}

.sd-label {
  font-size: 11px;
  color: var(--color-on-surface);
  flex-shrink: 0;
  width: 92px;
}

.sd-label--inline {
  width: auto;
  margin-inline-start: 8px;
}

.sd-line {
  margin: 0;
  font-size: 12px;
  color: var(--color-on-surface);
}

.sd-unit {
  font-size: 11px;
  color: var(--color-on-surface-variant);
}

.sd-input {
  padding: 3px 6px;
  height: 26px;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  color: var(--color-on-surface);
  font-family: var(--font-main);
  font-size: 12px;
  outline: none;
}

.sd-input:focus {
  border-color: var(--word-blue);
  box-shadow: 0 0 0 1px var(--word-blue);
}

.sd-input--wide {
  flex: 1 1 auto;
  min-width: 0;
}

.sd-input[readonly] {
  background: var(--color-surface-container-high);
}

.sd-number {
  width: 64px;
}

.sd-combo :deep(.ribbon-combo-input) {
  height: 26px;
  font-size: 12px;
}

.sd-group {
  margin: 0;
  padding: 6px 8px 8px;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.sd-group legend {
  font-size: 11px;
  color: var(--color-on-surface-variant);
  padding: 0 4px;
}

.sd-toggle {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 26px;
  height: 26px;
  padding: 0;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  color: var(--color-on-surface);
  cursor: pointer;
}

.sd-toggle:hover {
  background: var(--word-btn-hover);
  border-color: var(--word-blue);
}

.sd-toggle.active {
  background: var(--word-btn-active);
  border-color: var(--word-blue);
}

.sd-color {
  width: 30px;
  height: 26px;
  padding: 1px;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  cursor: pointer;
}

.sd-preview {
  height: 56px;
  overflow: hidden;
  padding: 6px 10px;
  display: flex;
  flex-direction: column;
  justify-content: center;
  border: 1px solid var(--color-outline-variant);
  border-radius: var(--radius-xs);
  /* הדף עצמו: התצוגה המקדימה היא טקסט על נייר, כמו ב-Word. */
  background: var(--color-surface);
  direction: rtl;
  white-space: nowrap;
}

.sd-preview-text {
  overflow: hidden;
  text-overflow: ellipsis;
}

.sd-note {
  margin: 0;
  font-size: 10.5px;
  line-height: 1.4;
  color: var(--color-on-surface-variant);
}

.sd-check {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 12px;
  color: var(--color-on-surface);
}

.sd-error {
  margin: 0;
  font-size: 11px;
  color: var(--color-error);
}

.sd-footer {
  flex-shrink: 0;
  display: flex;
  align-items: center;
  justify-content: flex-end;
  gap: 6px;
  padding: 8px 12px;
  border-block-start: 1px solid var(--color-outline-variant);
  background: var(--color-surface-container-high);
  border-radius: 0 0 var(--radius-sm) var(--radius-sm);
}

.sd-btn {
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

.sd-btn:hover:not(:disabled) {
  background: var(--word-btn-hover);
  border-color: var(--word-blue);
}

.sd-btn-primary {
  background: var(--word-blue);
  color: var(--color-on-primary);
  border-color: var(--word-blue);
}

.sd-btn-primary:hover:not(:disabled) {
  background: var(--word-blue-dark);
}

.sd-btn:disabled {
  opacity: 0.4;
  cursor: default;
}

.sd-recorder {
  flex: 1 1 auto;
  min-width: 0;
  text-align: start;
  cursor: pointer;
}

.sd-recorder--armed {
  border-color: var(--word-blue);
  box-shadow: 0 0 0 1px var(--word-blue);
}

.sd-recorder-hint {
  color: var(--color-on-surface-variant);
}
</style>
