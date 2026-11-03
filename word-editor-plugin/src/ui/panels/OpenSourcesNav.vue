<template>
  <nav
    class="src-nav"
    aria-label="מקומות"
  >
    <ul
      class="src-tree"
      role="tree"
      aria-label="מקומות"
      @keydown="onKeydown"
    >
      <li
        v-for="(row, index) in rows"
        :key="row.key"
        :ref="(el) => setRowRef(el, index)"
        class="src-item"
        :class="{
          'src-item--selected': row.selected,
          'src-item--group': row.group,
          'src-item--muted': row.muted,
        }"
        role="treeitem"
        :aria-level="row.depth"
        :aria-expanded="row.expandable ? row.expanded : undefined"
        :aria-selected="row.selectable ? row.selected : undefined"
        :aria-disabled="busy || undefined"
        :tabindex="index === activeIndex ? 0 : -1"
        :data-tip-title="row.tip || undefined"
        :style="{ '--src-depth': row.depth - 1 }"
        @click="onRowClick(row, index)"
        @focus="activeIndex = index"
      >
        <span
          class="src-twisty"
          :class="{ 'src-twisty--open': row.expanded }"
          aria-hidden="true"
          @click.stop="row.expandable && toggle(row)"
        >
          <SvgIcon
            v-if="row.expandable"
            name="chevronLeft"
            :size="12"
          />
        </span>
        <SvgIcon
          :name="row.icon"
          :size="16"
          class="src-icon"
        />
        <span
          class="src-label"
          dir="auto"
        >{{ row.label }}</span>
        <span
          v-if="row.badge"
          class="src-badge"
          aria-hidden="true"
        >{{ row.badge }}</span>
      </li>
    </ul>

    <!--
      מחוץ לעץ ובכוונה: פריט של `role="tree"` חייב להיות `treeitem`, ו„הוסף
      תיקייה” הוא פעולה ולא מקום. כפתור רגיל, בעצירת Tab משלו.
    -->
    <button
      v-if="foldersSupported !== false"
      type="button"
      class="src-add"
      :disabled="busy"
      data-tip-title="הוספת תיקייה"
      data-tip-desc="קובצי ה-Word שבה יופיעו כאן, וייפתחו בלחיצה אחת"
      @click="$emit('add-folder')"
    >
      <SvgIcon
        name="folderAdd"
        :size="16"
      />
      <span>הוסף תיקייה…</span>
    </button>
    <p
      v-else-if="folders.length === 0"
      class="src-note"
    >
      תיקיות אישיות דורשות גרסה חדשה יותר של אוצריא
    </p>
  </nav>
</template>

<script setup lang="ts">
/**
 * העמודה הימנית של „פתח מסמך”: אחרונים, ספריית אוצריא, ספרים אישיים,
 * והתיקיות שהמשתמש הוסיף.
 *
 * ## למה רשימה שטוחה של שורות
 *
 * העץ מרונדר מרשימה אחת (`rows`) של הצמתים **הנראים**, עם עומק לכל שורה, ולא
 * מקומפוננטה רקורסיבית. זה מה שהופך את הניווט במקלדת לפשוט ונכון: „למטה” הוא
 * השורה הבאה ברשימה, בלי לחפש את האח הבא של האב של הצומת. זה גם מה ש-WAI-ARIA
 * מתאר ל-tree עם `aria-level` — הקינון הוא נתון, לא מבנה DOM.
 *
 * ## הבחירה עוקבת אחרי המיקוד
 *
 * חץ למטה **בוחר** את הצומת הבא, והרשימה משמאל מתחלפת מיד — כמו חלונית הניווט
 * של סייר הקבצים. זה מה שהופך מעבר בין מקומות למהיר: אין Enter אחרי כל חץ.
 * המחיר (טעינת תיקייה על כל חץ) זול, כי פירוטים נשמרים בזיכרון.
 *
 * ## החצים ב-RTL
 *
 * ← פותח ונכנס, → סוגר ויוצא. זה השיקוף ש-WAI-ARIA קובעת לעץ בממשק מימין
 * לשמאל, ואותו כיוון ש-`nextTabIndex` נותן לרצועה: קדימה הוא שמאלה.
 */
import { computed, nextTick, ref, watch } from 'vue';
import SvgIcon from '../icons/SvgIcon.vue';
import type { LibraryState, ListingState } from '../../composables/use-open-sources';
import {
  listingKey,
  placeKey,
  type DocFolder,
  type LibraryShelf,
  type OpenPlace,
} from '../../sessions/open-sources';

const props = withDefaults(
  defineProps<{
    place: OpenPlace;
    recentCount?: number;
    library?: LibraryShelf | null;
    libraryState?: LibraryState;
    personalCount?: number;
    folders?: readonly DocFolder[];
    foldersSupported?: boolean | null;
    listings?: ReadonlyMap<string, ListingState>;
    busy?: boolean;
  }>(),
  {
    recentCount: 0,
    library: null,
    libraryState: 'idle',
    personalCount: 0,
    folders: () => [],
    foldersSupported: null,
    listings: () => new Map(),
    busy: false,
  },
);

const emit = defineEmits<{
  (e: 'select', place: OpenPlace): void;
  (e: 'expand-folder', token: string, path: string): void;
  (e: 'add-folder'): void;
}>();

interface Row {
  key: string;
  depth: number;
  label: string;
  icon: string;
  badge: string;
  tip: string;
  expandable: boolean;
  expanded: boolean;
  selectable: boolean;
  selected: boolean;
  /** צומת קיבוץ („התיקיות שלי”) — לחיצה פותחת/סוגרת, אינה בוחרת. */
  group: boolean;
  muted: boolean;
  place: OpenPlace | null;
  /** לצמתי תיקייה: מה לטעון כשנפתחים. */
  folder?: { token: string; path: string };
  parentIndex: number;
}

/**
 * מה פתוח. ספריית אוצריא והתיקיות פתוחות כברירת מחדל — הרמה הראשונה שלהן
 * היא מה שהמשתמש בא לראות, ולחיצה נוספת רק כדי לגלות אותה היא לחיצה מיותרת.
 */
const expanded = ref(new Set<string>(['library\u0000/', 'folders']));

function isOpen(key: string): boolean {
  return expanded.value.has(key);
}

const selectedKey = computed(() => placeKey(props.place));

function formatCount(count: number): string {
  return count > 0 ? String(count) : '';
}

const rows = computed<Row[]>(() => {
  const out: Row[] = [];
  const push = (row: Omit<Row, 'selected' | 'parentIndex'>, parentIndex: number): number => {
    const selected = row.selectable && row.place !== null && placeKey(row.place) === selectedKey.value;
    out.push({ ...row, selected, parentIndex });
    return out.length - 1;
  };

  push(
    {
      key: 'recent',
      depth: 1,
      label: 'אחרונים',
      icon: 'history',
      badge: formatCount(props.recentCount),
      tip: '',
      expandable: false,
      expanded: false,
      selectable: true,
      group: false,
      muted: false,
      place: { kind: 'recent' },
    },
    -1,
  );

  // הספרייה: מוסתרת רק כשאוצריא אמרה במפורש שאינה תומכת. „טוען” ו„אין הרשאה”
  // מוצגים — הם אומרים למשתמש משהו שהוא יכול לעשות איתו.
  const state = props.libraryState;
  if (state !== 'unsupported' && state !== 'idle' && !(state === 'ready' && !props.library)) {
    const root = props.library;
    const rootKey = 'library\u0000/';
    const hint =
      state === 'loading' && !root
        ? 'טוען…'
        : state === 'denied'
          ? 'לתוסף חסרה הרשאה לספרייה'
          : state === 'error' && !root
            ? 'הספרייה אינה זמינה כרגע'
            : '';
    const rootIndex = push(
      {
        key: rootKey,
        depth: 1,
        label: 'ספריית אוצריא',
        icon: 'library',
        badge: root ? formatCount(root.total) : '',
        tip: hint,
        expandable: Boolean(root && root.shelves.length > 0),
        expanded: Boolean(root && root.shelves.length > 0 && isOpen(rootKey)),
        selectable: true,
        group: false,
        muted: !root,
        place: { kind: 'library', path: '/' },
      },
      -1,
    );
    if (root && isOpen(rootKey)) {
      const walk = (shelves: readonly LibraryShelf[], depth: number, parent: number): void => {
        for (const shelf of shelves) {
          const key = `library\u0000${shelf.path}`;
          const open = shelf.shelves.length > 0 && isOpen(key);
          const index = push(
            {
              key,
              depth,
              label: shelf.title,
              icon: 'folderClosed',
              badge: formatCount(shelf.total),
              tip: '',
              expandable: shelf.shelves.length > 0,
              expanded: open,
              selectable: true,
              group: false,
              muted: false,
              place: { kind: 'library', path: shelf.path },
            },
            parent,
          );
          if (open) walk(shelf.shelves, depth + 1, index);
        }
      };
      walk(root.shelves, 2, rootIndex);
    }
  }

  if (props.personalCount > 0) {
    push(
      {
        key: 'personal',
        depth: 1,
        label: 'ספרים אישיים',
        icon: 'book',
        badge: formatCount(props.personalCount),
        tip: '',
        expandable: false,
        expanded: false,
        selectable: true,
        group: false,
        muted: false,
        place: { kind: 'personal' },
      },
      -1,
    );
  }

  if (props.folders.length > 0) {
    const groupIndex = push(
      {
        key: 'folders',
        depth: 1,
        label: 'התיקיות שלי',
        icon: 'folder',
        badge: '',
        tip: '',
        expandable: true,
        expanded: isOpen('folders'),
        selectable: false,
        group: true,
        muted: false,
        place: null,
      },
      -1,
    );
    if (isOpen('folders')) {
      const walkDir = (token: string, path: string, depth: number, parent: number): void => {
        const listing = props.listings.get(listingKey(token, path));
        if (listing?.state !== 'ready') return;
        for (const entry of listing.listing.entries) {
          if (entry.type !== 'dir') continue;
          const key = `folder\u0000${token}\u0000${entry.path}`;
          const open = isOpen(key);
          const index = push(
            {
              key,
              depth,
              label: entry.name,
              icon: 'folderClosed',
              badge: '',
              tip: '',
              expandable: true,
              expanded: open,
              selectable: true,
              group: false,
              muted: false,
              place: { kind: 'folder', token, path: entry.path },
              folder: { token, path: entry.path },
            },
            parent,
          );
          if (open) walkDir(token, entry.path, depth + 1, index);
        }
      };
      for (const folder of props.folders) {
        const key = `folder\u0000${folder.token}\u0000`;
        const open = isOpen(key);
        const listing = props.listings.get(listingKey(folder.token, ''));
        const index = push(
          {
            key,
            depth: 2,
            label: folder.name,
            icon: 'folderClosed',
            badge: '',
            tip: folder.path,
            // תיקייה שטרם נפתחה עשויה להכיל תת-תיקיות; מה שכבר ידוע שאין בו
            // — אין לו חץ, כדי שלא יזמין לחיצה שלא תעשה כלום.
            expandable: listing?.state !== 'ready' || listing.listing.entries.some((entry) => entry.type === 'dir'),
            expanded: open,
            selectable: true,
            group: false,
            muted: listing?.state === 'error',
            place: { kind: 'folder', token: folder.token, path: '' },
            folder: { token: folder.token, path: '' },
          },
          groupIndex,
        );
        if (open) walkDir(folder.token, '', 3, index);
      }
    }
  }

  return out;
});

/* ------------------------------------------------------------------ */
/* פתיחה, סגירה ובחירה                                                 */
/* ------------------------------------------------------------------ */

function setOpen(row: Row, open: boolean): void {
  if (!row.expandable || row.expanded === open) return;
  const next = new Set(expanded.value);
  if (open) next.add(row.key);
  else next.delete(row.key);
  expanded.value = next;
  if (open && row.folder) emit('expand-folder', row.folder.token, row.folder.path);
}

function toggle(row: Row): void {
  setOpen(row, !row.expanded);
}

function select(row: Row): void {
  if (props.busy) return;
  if (row.group) {
    toggle(row);
    return;
  }
  if (row.place) emit('select', row.place);
}

/**
 * לחיצה על צומת שכבר נבחר פותחת/סוגרת אותו — זו הלחיצה השנייה הטבעית על
 * מדף, ובלעדיה היה צריך לכוון בדיוק אל החץ הקטן.
 */
function onRowClick(row: Row, index: number): void {
  activeIndex.value = index;
  if (row.selected && row.expandable) toggle(row);
  else select(row);
}

/* ------------------------------------------------------------------ */
/* מקלדת                                                                */
/* ------------------------------------------------------------------ */

const activeIndex = ref(0);
const rowRefs = ref<(HTMLElement | null)[]>([]);

function setRowRef(el: unknown, index: number): void {
  rowRefs.value[index] = el instanceof HTMLElement ? el : null;
}

function focusRow(index: number, andSelect: boolean): void {
  const row = rows.value[index];
  if (!row) return;
  activeIndex.value = index;
  if (andSelect && !row.group) select(row);
  void nextTick(() => {
    const el = rowRefs.value[index];
    el?.focus();
    el?.scrollIntoView?.({ block: 'nearest' });
  });
}

function onKeydown(event: KeyboardEvent): void {
  const list = rows.value;
  const current = list[activeIndex.value];
  if (!current) return;
  let handled = true;
  switch (event.key) {
    case 'ArrowDown':
      focusRow(Math.min(activeIndex.value + 1, list.length - 1), true);
      break;
    case 'ArrowUp':
      focusRow(Math.max(activeIndex.value - 1, 0), true);
      break;
    case 'Home':
      focusRow(0, true);
      break;
    case 'End':
      focusRow(list.length - 1, true);
      break;
    case 'ArrowLeft':
      // קדימה ב-RTL: פתיחה, ואם כבר פתוח — כניסה לילד הראשון.
      if (current.expandable && !current.expanded) setOpen(current, true);
      else if (current.expanded && list[activeIndex.value + 1]?.parentIndex === activeIndex.value) {
        focusRow(activeIndex.value + 1, true);
      }
      break;
    case 'ArrowRight':
      if (current.expanded) setOpen(current, false);
      else if (current.parentIndex >= 0) focusRow(current.parentIndex, true);
      break;
    case 'Enter':
    case ' ':
      if (current.group || current.expandable) toggle(current);
      else select(current);
      break;
    default:
      handled = false;
  }
  if (handled) event.preventDefault();
}

/**
 * המקום שנבחר מבחוץ (מקום שנשמר, תיקייה שנוספה, ניווט מהרשימה) חייב להיות
 * נראה בעץ: האבות שלו נפתחים, והמיקוד הנודד עובר אליו.
 */
watch(
  () => [selectedKey.value, rows.value.length] as const,
  () => {
    const place = props.place;
    const next = new Set(expanded.value);
    let changed = false;
    const openKey = (key: string): void => {
      if (!next.has(key)) {
        next.add(key);
        changed = true;
      }
    };
    if (place.kind === 'library' && place.path !== '/') {
      openKey('library\u0000/');
      const parts = place.path.split('/').filter(Boolean);
      for (let i = 1; i < parts.length; i += 1) openKey(`library\u0000/${parts.slice(0, i).join('/')}`);
    }
    if (place.kind === 'folder') {
      openKey('folders');
      if (place.path !== '') {
        openKey(`folder\u0000${place.token}\u0000`);
        const parts = place.path.split('/').filter(Boolean);
        for (let i = 1; i < parts.length; i += 1) {
          const sub = parts.slice(0, i).join('/');
          openKey(`folder\u0000${place.token}\u0000${sub}`);
          emit('expand-folder', place.token, sub);
        }
        emit('expand-folder', place.token, '');
      }
    }
    if (changed) expanded.value = next;
    const index = rows.value.findIndex((row) => row.selected);
    if (index >= 0) activeIndex.value = index;
  },
  { immediate: true },
);

defineExpose({
  /** ממקד את הצומת הנבחר — מה שהדיאלוג קורא לו כשהוא נוחת על העץ. */
  focusSelected(): void {
    const index = rows.value.findIndex((row) => row.selected);
    focusRow(index >= 0 ? index : 0, false);
  },
});
</script>

<style scoped>
.src-nav {
  display: flex;
  flex-direction: column;
  gap: 6px;
  min-height: 0;
  padding: 6px;
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-sm);
  background: var(--color-surface-container-high);
}

/* פס הגלילה מימין, כמו ב-`.rec-list` — ראו שם את ההסבר ל-`direction: ltr`. */
.src-tree {
  list-style: none;
  margin: 0;
  padding: 0;
  flex: 1 1 auto;
  min-height: 0;
  overflow-y: auto;
  direction: ltr;
}

/*
 * ההזחה היא `padding-inline-start` לפי העומק, ולא קינון של רשימות: הצמתים
 * שטוחים ב-DOM (ראו ראש הסקריפט). 14 פיקסלים לרמה — מספיק כדי שהעין תקרא
 * היררכיה, ומעט מספיק כדי שעומק 5 לא יאכל את העמודה.
 */
.src-item {
  direction: rtl;
  display: flex;
  align-items: center;
  gap: 6px;
  min-height: 28px;
  padding-block: 2px;
  padding-inline: calc(4px + var(--src-depth, 0) * 14px) 8px;
  border-radius: var(--radius-sm);
  color: var(--color-on-surface);
  cursor: pointer;
  user-select: none;
  outline: none;
  transition: background-color 0.08s ease;
}

.src-item:hover {
  background: var(--word-btn-hover);
}

.src-item:focus-visible {
  box-shadow: inset 0 0 0 2px var(--word-blue);
}

/* הנבחר: מילוי עדין ופס בקצה ההתחלה — אותה שפה כמו הלשונית הפעילה ברצועה. */
.src-item--selected {
  background: var(--color-primary-subtle);
  box-shadow: inset -3px 0 0 var(--word-blue);
  font-weight: 600;
}

.src-item--selected:hover {
  background: var(--color-primary-selected-hover);
}

.src-item--selected:focus-visible {
  box-shadow: inset -3px 0 0 var(--word-blue), inset 0 0 0 2px var(--word-blue);
}

.src-item--group {
  margin-block-start: 6px;
  font-size: 0.86em;
  font-weight: 700;
  color: var(--color-primary);
}

.src-item--muted .src-label {
  color: var(--color-on-surface-variant);
}

.src-twisty {
  flex: 0 0 auto;
  width: 14px;
  height: 14px;
  display: flex;
  align-items: center;
  justify-content: center;
  color: var(--color-on-surface-variant);
  transition: transform 0.12s ease;
}

/* ‹ פונה שמאלה ב-RTL (קדימה); פתוח הוא ‹ מסובב כלפי מטה. */
.src-twisty--open {
  transform: rotate(-90deg);
}

.src-icon {
  flex: 0 0 auto;
  color: var(--word-blue);
}

.src-item--group .src-icon {
  color: var(--color-primary);
}

.src-label {
  flex: 1 1 auto;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.src-badge {
  flex: 0 0 auto;
  min-width: 18px;
  padding: 0 6px;
  border-radius: var(--radius-pill);
  background: var(--color-surface);
  font-size: 0.78em;
  font-weight: 600;
  line-height: 1.6;
  text-align: center;
  color: var(--color-on-surface-variant);
}

.src-add {
  flex: 0 0 auto;
  display: flex;
  align-items: center;
  gap: 6px;
  width: 100%;
  padding: 6px 8px;
  background: none;
  border: 1px dashed var(--color-outline);
  border-radius: var(--radius-sm);
  color: var(--color-primary);
  font-family: var(--font-main);
  font-size: 0.92em;
  text-align: start;
  cursor: pointer;
  transition: background-color 0.08s ease, border-color 0.08s ease;
}

.src-add:hover:not(:disabled) {
  background: var(--word-btn-hover);
  border-color: var(--color-primary);
}

.src-note {
  margin: 0;
  padding: 4px 6px;
  font-size: 0.82em;
  color: var(--color-on-surface-variant);
}

@media (prefers-reduced-motion: reduce) {
  .src-item,
  .src-twisty,
  .src-add {
    transition: none;
  }
}
</style>
