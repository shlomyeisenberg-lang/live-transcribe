<template>
  <div
    ref="containerRef"
    class="table-picker-container"
  >
    <RibbonButton
      icon="table"
      label="טבלה"
      variant="large"
      :active="isOpen"
      :disabled="disabled"
      tooltip="הוספת טבלה"
      aria-haspopup="grid"
      :aria-expanded="isOpen ? 'true' : 'false'"
      @click="toggleDropdown"
    />

    <!--
      `.prevent` ולא רק `.stop`, וזה נדרש מרגע שיש בפופאובר משטח שמקבל מיקוד:
      ברירת המחדל של pointerdown הייתה מעבירה את המיקוד מהעורך אל הגריד, הבחירה
      במסמך הייתה אובדת, והטבלה נכנסת למקום הלא נכון. אותו דפוס כמו
      RibbonMenuButton. הלחיצה עצמה אינה מבוטלת — `click` נורה גם אחרי
      preventDefault על pointerdown.
    -->
    <div
      v-if="isOpen"
      ref="popoverRef"
      class="table-picker-popover"
      :style="popoverStyle"
      @pointerdown.prevent.stop
    >
      <!-- role="status" כדי שהמידות ייקראו גם למי שאינו רואה את ההדגשה -->
      <div
        class="table-picker-header"
        role="status"
        aria-live="polite"
      >
        {{ headerText }}
      </div>
      <!--
        נקודת Tab אחת לכל הגריד, והתא הנוכחי מוכרז ב-aria-activedescendant.
        עשרה תאים בשורה × עשר שורות = 100 נקודות Tab, ולכן הדפוס הזה ולא
        tabindex לכל תא.
      -->
      <div
        ref="gridRef"
        class="table-grid"
        role="grid"
        tabindex="0"
        :aria-label="menuString('בחירת מידות הטבלה')"
        :aria-activedescendant="activeCellId"
        @mouseleave="onMouseLeave"
        @focus="onGridFocus"
        @keydown="onKeydown"
      >
        <div
          v-for="r in MAX_ROWS"
          :key="`row-${r}`"
          class="grid-row"
          role="row"
        >
          <div
            v-for="c in MAX_COLS"
            :id="cellId(r, c)"
            :key="`cell-${r}-${c}`"
            class="grid-cell"
            role="gridcell"
            :class="{ highlighted: isWithin(r, c) }"
            :aria-selected="isWithin(r, c)"
            :aria-label="cellLabel(r, c)"
            @mouseenter="setSize(r, c)"
            @click="onCellClick(r, c)"
          />
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
/**
 * בורר מידות הטבלה.
 *
 * **מה שהיה שבור:** מאה תאי `<div>` עם `@click`, בלי `role`, בלי `tabindex`
 * ובלי מקלדת — כלומר אי אפשר היה להוסיף טבלה בלי עכבר. הגריד היה גם חסר שם
 * וחסר סמנטיקה, ולכן קורא מסך לא ראה כאן שום פקד: הוא ראה מאה `div` ריקים.
 *
 * **הפתרון:** `role="grid"` עם שורות ותאים, נקודת Tab אחת (הגריד עצמו) והתא
 * הנוכחי מוכרז ב-`aria-activedescendant` — הדפוס שמונע מאה נקודות Tab. החצים
 * משנים מידות, Enter/רווח מאשרים, ו-Escape סוגר ומחזיר את המיקוד לכפתור, כמו
 * ב-RibbonMenuButton.
 *
 * **המיקוד והבחירה במסמך:** הפופאובר מסמן `@pointerdown.prevent.stop`. בלי
 * `.prevent` הלחיצה על הגריד — שהוא מעכשיו משטח שמקבל מיקוד — הייתה גוזלת את
 * המיקוד מהעורך, והבחירה במסמך שהטבלה אמורה להיכנס אליה הייתה אובדת. לכן גם
 * `onGridFocus` נוגע רק במשתמשי מקלדת: העכבר אינו יכול למקד את הגריד.
 */
import { computed, ref, onMounted, onUnmounted } from 'vue';
import RibbonButton from './RibbonButton.vue';
import { menuString } from '../i18n';
import { usePopoverPosition } from '../../../composables/popover-position';

const MAX_ROWS = 10;
const MAX_COLS = 10;

defineProps<{
  disabled?: boolean;
}>();

const emit = defineEmits<{
  (e: 'select', dimensions: { rows: number; cols: number }): void;
}>();

const containerRef = ref<HTMLElement | null>(null);
const popoverRef = ref<HTMLElement | null>(null);
const gridRef = ref<HTMLElement | null>(null);
const isOpen = ref(false);

// `.word-ribbon-body` חותך אנכית, ולכן הפופאובר `position: fixed` בקואורדינטות
// שנמדדות ולא `top: 100%`. ראו composables/popover-position.ts.
const { popoverStyle } = usePopoverPosition(containerRef, popoverRef, isOpen);

/** המידות שמעל הסמן או שהמקלדת בחרה. 0 = טרם נבחר דבר. */
const hoveredRows = ref(0);
const hoveredCols = ref(0);

const hasSize = computed(() => hoveredRows.value > 0 && hoveredCols.value > 0);

const headerText = computed(() =>
  hasSize.value
    ? `${menuString('טבלה')} ${hoveredCols.value} × ${hoveredRows.value}`
    : menuString('הוסף טבלה'),
);

function cellId(row: number, col: number): string {
  return `table-picker-cell-${row}-${col}`;
}

/** התא שהמקלדת עומדת עליו, או `undefined` כשטרם נבחר דבר. */
const activeCellId = computed(() =>
  hasSize.value ? cellId(hoveredRows.value, hoveredCols.value) : undefined,
);

function isWithin(row: number, col: number): boolean {
  return row <= hoveredRows.value && col <= hoveredCols.value;
}

/** „3 עמודות על 2 שורות” — הניסוח המלא, כי `aria-activedescendant` מכריז אותו. */
function cellLabel(row: number, col: number): string {
  const cols = col === 1 ? menuString('עמודה אחת') : `${col} ${menuString('עמודות')}`;
  const rows = row === 1 ? menuString('שורה אחת') : `${row} ${menuString('שורות')}`;
  return `${cols} ${menuString('על')} ${rows}`;
}

function setSize(row: number, col: number): void {
  hoveredRows.value = row;
  hoveredCols.value = col;
}

function toggleDropdown(): void {
  isOpen.value = !isOpen.value;
  if (isOpen.value) setSize(0, 0);
}

function close(): void {
  isOpen.value = false;
}

function onMouseLeave(): void {
  // מי שניווט במקלדת אינו מאבד את בחירתו מפני שעכבר חלף על הגריד.
  if (gridRef.value === document.activeElement) return;
  setSize(0, 0);
}

/**
 * כניסה לגריד ב-Tab מעמידה את הסמן על 1×1, כדי שיהיה תא נוכחי להכריז ומשהו
 * לאשר. העכבר אינו מגיע לכאן — `@pointerdown.prevent` על הפופאובר מונע ממנו
 * למקד את הגריד.
 */
function onGridFocus(): void {
  if (!hasSize.value) setSize(1, 1);
}

function onCellClick(row: number, col: number): void {
  setSize(row, col);
  confirm();
}

function confirm(): void {
  if (!hasSize.value) return;
  emit('select', { rows: hoveredRows.value, cols: hoveredCols.value });
  close();
}

function clamp(value: number, max: number): number {
  return Math.min(max, Math.max(1, value));
}

/**
 * הצעד שהמקש מבקש, או `null` אם אינו מקש ניווט.
 *
 * 'rtl' קבוע, כמו בסרגל הלשוניות (ui/ribbon/Ribbon.vue): המעטפת כולה
 * `dir="rtl"`, השורה נפרסת מימין לשמאל, ולכן ArrowLeft הוא **תוספת** עמודה —
 * הכיוון החזותי, כפי ש-WAI-ARIA דורש.
 */
function arrowStep(key: string): { rows: number; cols: number } | null {
  switch (key) {
    case 'ArrowLeft':
      return { rows: 0, cols: 1 };
    case 'ArrowRight':
      return { rows: 0, cols: -1 };
    case 'ArrowDown':
      return { rows: 1, cols: 0 };
    case 'ArrowUp':
      return { rows: -1, cols: 0 };
    default:
      return null;
  }
}

function onKeydown(event: KeyboardEvent): void {
  if (event.key === 'Escape') {
    // עצירת ההתפשטות רק כשהתפריט פתוח (וכאן הוא תמיד פתוח — הגריד אינו קיים
    // אחרת): Escape סוגר גם דיאלוגים ומצב מיקוד ברמת המעטפת.
    event.stopPropagation();
    close();
    containerRef.value?.querySelector('button')?.focus();
    return;
  }

  if (event.key === 'Enter' || event.key === ' ') {
    // רווח גם גולל את הפופאובר, ו-Enter מפעיל טופס עוטף.
    event.preventDefault();
    confirm();
    return;
  }

  const step = arrowStep(event.key);
  if (!step) return;

  // בלי זה החצים גם גוללים את הרצועה.
  event.preventDefault();
  // מ-0 כל חץ מגיע ל-1, ולכן הצעד הראשון תמיד נוחת על 1×1.
  setSize(clamp(hoveredRows.value + step.rows, MAX_ROWS), clamp(hoveredCols.value + step.cols, MAX_COLS));
}

function handleClickOutside(event: MouseEvent): void {
  if (containerRef.value && !containerRef.value.contains(event.target as Node)) {
    close();
  }
}

onMounted(() => {
  document.addEventListener('pointerdown', handleClickOutside);
});

onUnmounted(() => {
  document.removeEventListener('pointerdown', handleClickOutside);
});
</script>

<style scoped>
.table-picker-container {
  position: relative;
  display: inline-flex;
  flex-shrink: 0;
}

/* `top` / `left` / `max-height` מגיעים מ-`:style` — ראו popover-position.ts.
   `overflow-y: auto` הוא הצד השני של אותה החלטה: כשאין מקום לגובה המלא הבורר
   נגלל בתוך עצמו, ולא נחתך. */
.table-picker-popover {
  position: fixed;
  z-index: 1000;
  background: var(--color-surface);
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-sm);
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.16);
  padding: 8px;
  overflow-y: auto;
}

.table-picker-header {
  font-size: 11px;
  font-weight: 600;
  color: var(--color-on-surface);
  text-align: center;
  margin-bottom: 6px;
  min-height: 16px;
}

.table-grid {
  display: flex;
  flex-direction: column;
  gap: 2px;
  outline: none;
}

/* הגריד הוא פקד שמקבל מיקוד, ולכן חייב להראות זאת. */
.table-grid:focus-visible {
  outline: 2px solid var(--word-blue);
  outline-offset: 2px;
}

.grid-row {
  display: flex;
  gap: 2px;
}

.grid-cell {
  width: 16px;
  height: 16px;
  border: 1px solid var(--color-outline-variant);
  border-radius: 1px;
  background: var(--color-surface);
  cursor: pointer;
  transition: background 0.05s, border-color 0.05s;
}

.grid-cell.highlighted {
  background: var(--color-primary-subtle);
  border-color: var(--color-primary);
}
</style>
