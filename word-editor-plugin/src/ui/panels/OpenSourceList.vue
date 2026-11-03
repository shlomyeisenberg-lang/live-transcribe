<template>
  <section
    class="sl"
    :aria-labelledby="HEADING_ID"
  >
    <div class="sl-head">
      <!-- שביל הלחם: כל מקטע חוץ מהאחרון הוא כפתור שחוזר אליו. -->
      <h3
        :id="HEADING_ID"
        class="sl-title"
      >
        <template
          v-for="(crumb, i) in crumbs"
          :key="i"
        >
          <SvgIcon
            v-if="i > 0"
            name="chevronLeft"
            :size="12"
            class="sl-crumb-sep"
          />
          <button
            v-if="i < crumbs.length - 1 && crumb.place"
            type="button"
            class="sl-crumb"
            :disabled="busy"
            @click="$emit('navigate', crumb.place)"
          >
            {{ crumb.label }}
          </button>
          <span
            v-else
            class="sl-crumb sl-crumb--current"
            dir="auto"
          >{{ crumb.label }}</span>
        </template>
      </h3>
      <div class="sl-search">
        <SvgIcon
          name="search"
          :size="14"
          class="sl-search__icon"
        />
        <input
          ref="searchRef"
          class="sl-search__input"
          type="text"
          :value="query"
          :placeholder="searchPlaceholder"
          :aria-label="searchPlaceholder"
          :disabled="busy"
          @input="$emit('update:query', ($event.target as HTMLInputElement).value)"
          @keydown.down.prevent="focusRow(0)"
        >
        <button
          v-if="query !== ''"
          type="button"
          class="sl-search__clear"
          aria-label="נקה את החיפוש"
          @click="$emit('update:query', '')"
        >
          ✕
        </button>
      </div>
      <button
        v-if="canRefresh"
        type="button"
        class="sl-iconbtn"
        :class="{ 'sl-iconbtn--spin': refreshing }"
        :disabled="busy || refreshing"
        aria-label="רענן"
        data-tip-title="רענן"
        @click="$emit('refresh')"
      >
        <SvgIcon
          name="updateFields"
          :size="15"
        />
      </button>
    </div>

    <!-- טעינה ראשונה בלבד. רענון של מה שכבר מוצג אינו מסתיר אותו. -->
    <div
      v-if="status === 'loading'"
      class="sl-state"
      role="status"
      aria-live="polite"
    >
      <span
        class="sl-spinner"
        aria-hidden="true"
      />
      <p class="sl-state__title">
        טוען…
      </p>
    </div>

    <div
      v-else-if="status === 'error'"
      class="sl-state"
      role="status"
    >
      <p class="sl-state__title">
        {{ errorTitle }}
      </p>
      <p class="sl-state__hint">
        {{ errorHint }}
      </p>
      <div class="sl-state__actions">
        <button
          type="button"
          class="sl-link"
          :disabled="busy"
          @click="$emit('refresh')"
        >
          נסה שוב
        </button>
        <button
          v-if="place.kind === 'folder' && place.path === ''"
          type="button"
          class="sl-link sl-link--danger"
          :disabled="busy"
          @click="$emit('remove-folder', place.token)"
        >
          הסר את התיקייה מהרשימה
        </button>
      </div>
    </div>

    <div
      v-else-if="items.length === 0"
      class="sl-state"
      role="status"
      aria-live="polite"
    >
      <SvgIcon
        name="folderClosed"
        :size="28"
        class="sl-state__icon"
      />
      <p class="sl-state__title">
        {{ emptyTitle }}
      </p>
      <p
        v-if="emptyHint"
        class="sl-state__hint"
      >
        {{ emptyHint }}
      </p>
    </div>

    <ul
      v-else
      class="sl-list"
      :aria-labelledby="HEADING_ID"
      @keydown="onKeydown"
    >
      <li
        v-for="(item, index) in items"
        :key="item.key"
        class="sl-row"
        :class="{ 'sl-row--dir': item.kind === 'dir' }"
      >
        <button
          :ref="(el) => setRowRef(el, index)"
          type="button"
          class="sl-open"
          :tabindex="index === activeIndex ? 0 : -1"
          :disabled="busy"
          :aria-label="ariaLabel(item)"
          :data-tip-title="item.tip || undefined"
          @focus="activeIndex = index"
          @click="activate(item)"
        >
          <SvgIcon
            :name="item.kind === 'dir' ? 'folderClosed' : 'docFile'"
            :size="18"
            class="sl-icon"
            :class="item.kind === 'dir' ? 'sl-icon--dir' : 'sl-icon--doc'"
          />
          <span class="sl-text">
            <span
              class="sl-name"
              dir="auto"
            >{{ item.name }}</span>
            <span
              v-if="item.sub"
              class="sl-sub"
              dir="auto"
            >{{ item.sub }}</span>
          </span>
          <span
            v-if="item.meta"
            class="sl-meta"
          >{{ item.meta }}</span>
          <SvgIcon
            v-if="item.kind === 'dir'"
            name="chevronLeft"
            :size="14"
            class="sl-chevron"
          />
        </button>
      </li>
      <li
        v-if="footnote"
        class="sl-footnote"
        aria-live="polite"
      >
        {{ footnote }}
      </li>
    </ul>
  </section>
</template>

<script setup lang="ts">
/**
 * הרשימה משמאל כשנבחר בעץ מקום שאינו „אחרונים”: מדף בספרייה, הספרים
 * האישיים, או תיקייה.
 *
 * כמו הדיאלוג שמכיל אותה — תצוגה בלבד. כל המצב נכנס ב-props, וכל פעולה יוצאת
 * כאירוע. הרשימה המוצגת נגזרת כאן, מאותן פונקציות טהורות שנבדקות ב-
 * sessions/open-sources.ts.
 *
 * ## לחיצה אחת
 *
 * לחיצה על ספר או על קובץ **פותחת** אותו, ולחיצה על מדף או על תת-תיקייה
 * נכנסת אליו. אין „בחר ואז פתח”: זו רשימה שכל מטרתה לפתוח, ושני שלבים היו
 * מכפילים את מספר הלחיצות בלי להוסיף ביטחון — הפעולה אינה הרסנית, והמסמך
 * הפתוח אינו נדרס (`ensureOpenTargetTab` פותח טאב חדש).
 *
 * ## החיפוש
 *
 * בספרייה החיפוש הוא **בכל הספרייה**, לא רק במדף הנוכחי: מי שמקליד שם ספר
 * אינו יודע ואינו צריך לדעת באיזה מדף הוא יושב. בתיקייה הוא מסנן את מה שמוצג.
 */
import { computed, nextTick, ref, watch } from 'vue';
import SvgIcon from '../icons/SvgIcon.vue';
import { draftAgeLabel } from '../../sessions/session-state';
import { foldForSearch } from '../../sessions/recent-documents';
import type { LibraryState, ListingState } from '../../composables/use-open-sources';
import {
  findShelf,
  searchLibrary,
  shelfTrail,
  type DocFolder,
  type LibraryBook,
  type LibraryShelf,
  type OpenPlace,
} from '../../sessions/open-sources';

const HEADING_ID = 'open-dialog-source-title';
/** כמה תוצאות חיפוש לצייר. מעבר לזה — שורת „צמצם את החיפוש”. */
const SEARCH_LIMIT = 60;

const props = withDefaults(
  defineProps<{
    place: OpenPlace;
    query?: string;
    library?: LibraryShelf | null;
    libraryState?: LibraryState;
    libraryMessage?: string;
    personal?: readonly LibraryBook[];
    folder?: DocFolder | null;
    listing?: ListingState | null;
    busy?: boolean;
  }>(),
  {
    query: '',
    library: null,
    libraryState: 'idle',
    libraryMessage: '',
    personal: () => [],
    folder: null,
    listing: null,
    busy: false,
  },
);

const emit = defineEmits<{
  (e: 'navigate', place: OpenPlace): void;
  (e: 'open-book', book: LibraryBook): void;
  (e: 'open-file', token: string, path: string, name: string): void;
  (e: 'remove-folder', token: string): void;
  (e: 'refresh'): void;
  (e: 'update:query', value: string): void;
}>();

interface Item {
  key: string;
  kind: 'dir' | 'doc';
  name: string;
  sub: string;
  meta: string;
  tip: string;
  run: () => void;
}

/* ------------------------------------------------------------------ */
/* שביל הלחם                                                            */
/* ------------------------------------------------------------------ */

const crumbs = computed<{ label: string; place: OpenPlace | null }[]>(() => {
  const place = props.place;
  if (place.kind === 'library') {
    const trail = shelfTrail(props.library, place.path);
    if (trail.length === 0) return [{ label: 'ספריית אוצריא', place: null }];
    return trail.map((shelf, i) => ({
      label: i === 0 ? 'ספריית אוצריא' : shelf.title,
      place: { kind: 'library', path: shelf.path },
    }));
  }
  if (place.kind === 'personal') return [{ label: 'ספרים אישיים', place: null }];
  if (place.kind === 'folder') {
    const root = props.folder?.name ?? 'תיקייה';
    const parts = place.path.split('/').filter(Boolean);
    return [
      { label: root, place: { kind: 'folder', token: place.token, path: '' } },
      ...parts.map((part, i) => ({
        label: part,
        place: { kind: 'folder', token: place.token, path: parts.slice(0, i + 1).join('/') } as OpenPlace,
      })),
    ];
  }
  return [];
});

const searchPlaceholder = computed(() => {
  if (props.place.kind === 'library') return 'חיפוש בכל ספרי ה-Word שבספרייה';
  if (props.place.kind === 'personal') return 'חיפוש בספרים האישיים';
  return 'סינון בתיקייה הזאת';
});

/* ------------------------------------------------------------------ */
/* המצב                                                                 */
/* ------------------------------------------------------------------ */

const status = computed<'loading' | 'error' | 'ready'>(() => {
  const place = props.place;
  if (place.kind === 'folder') {
    if (!props.listing || props.listing.state === 'loading') return 'loading';
    return props.listing.state === 'error' ? 'error' : 'ready';
  }
  if (props.library) return 'ready';
  if (props.libraryState === 'loading' || props.libraryState === 'idle') return 'loading';
  return props.libraryState === 'ready' ? 'ready' : 'error';
});

const refreshing = computed(() =>
  props.place.kind === 'folder'
    ? props.listing?.state === 'ready' && props.listing.refreshing
    : props.libraryState === 'loading' && props.library !== null,
);

const canRefresh = computed(() => status.value === 'ready');

const errorTitle = computed(() => {
  if (props.place.kind === 'folder') {
    const reason = props.listing?.state === 'error' ? props.listing.reason : '';
    if (reason === 'not-found') return 'התיקייה לא נמצאה';
    if (reason === 'unsupported') return 'גרסת אוצריא הזאת אינה תומכת בתיקיות';
    return 'קריאת התיקייה נכשלה';
  }
  if (props.libraryState === 'denied') return 'לתוסף חסרה הרשאה לספרייה';
  if (props.libraryState === 'unsupported') return 'גרסת אוצריא הזאת אינה מציגה את הספרייה כאן';
  return 'הספרייה אינה זמינה כרגע';
});

const errorHint = computed(() => {
  if (props.place.kind === 'folder') {
    const reason = props.listing?.state === 'error' ? props.listing.reason : '';
    if (reason === 'not-found') return 'ייתכן שהיא הוזזה, נמחקה, או שהכונן שלה אינו מחובר.';
    return props.listing?.state === 'error' ? props.listing.message : '';
  }
  if (props.libraryState === 'denied') return 'יש לאשר את ההרשאה לתוסף בהגדרות אוצריא.';
  return props.libraryMessage;
});

/* ------------------------------------------------------------------ */
/* השורות                                                               */
/* ------------------------------------------------------------------ */

function sizeLabel(size: number): string {
  if (!Number.isFinite(size) || size <= 0) return '';
  if (size < 1024) return `${size} בייט`;
  if (size < 1048576) return `${(size / 1024).toFixed(0)} ק״ב`;
  return `${(size / 1048576).toFixed(1)} מ״ב`;
}

/**
 * שם הקובץ בלי `.docx` — כמו ברשימת „פתח” של Word: כל מה שברשימה הזאת הוא
 * מסמך Word, והסיומת היא רעש שחוזר בכל שורה. `.docm` **נשאר**, כי הוא אומר
 * משהו: יש במסמך מאקרו. השם המלא נשאר בטולטיפ.
 */
function displayName(name: string): string {
  return /\.docx$/i.test(name) ? name.slice(0, -5) : name;
}

function bookItem(book: LibraryBook, where: string): Item {
  const sub = [book.author, where].filter(Boolean).join(' · ');
  return {
    key: `b:${book.key}`,
    kind: 'doc',
    name: book.title,
    sub,
    meta: book.source === 'user' ? 'אישי' : '',
    tip: book.title,
    run: () => emit('open-book', book),
  };
}

function shelfItem(shelf: LibraryShelf): Item {
  return {
    key: `s:${shelf.path}`,
    kind: 'dir',
    name: shelf.title,
    sub: '',
    meta: String(shelf.total),
    tip: '',
    run: () => emit('navigate', { kind: 'library', path: shelf.path }),
  };
}

const matches = (name: string): boolean => {
  const needle = foldForSearch(props.query);
  return needle === '' || foldForSearch(name).includes(needle);
};

/**
 * השורות, ויחד איתן — האם תוצאות החיפוש נחתכו (שורת ההערה בתחתית). computed
 * אחד ולא שניים, כדי שהחיפוש בעץ ירוץ פעם אחת לכל הקשה.
 */
const view = computed<{ items: Item[]; truncated: boolean }>(() => ({ truncated: false, ...buildItems() }));
const items = computed(() => view.value.items);

function buildItems(): { items: Item[]; truncated?: boolean } {
  const place = props.place;
  const searching = props.query.trim() !== '';

  if (place.kind === 'library') {
    if (searching) {
      const hits = searchLibrary(props.library, props.query, SEARCH_LIMIT + 1);
      return {
        items: hits.slice(0, SEARCH_LIMIT).map((hit) => bookItem(hit.book, hit.where)),
        truncated: hits.length > SEARCH_LIMIT,
      };
    }
    const shelf = findShelf(props.library, place.path);
    if (!shelf) return { items: [] };
    return { items: [...shelf.shelves.map(shelfItem), ...shelf.books.map((book) => bookItem(book, ''))] };
  }

  if (place.kind === 'personal') {
    return { items: props.personal.filter((book) => matches(book.title)).map((book) => bookItem(book, '')) };
  }

  if (place.kind === 'folder' && props.listing?.state === 'ready') {
    const token = place.token;
    const entries = props.listing.listing.entries
      .filter((entry) => matches(entry.name))
      .map<Item>((entry) =>
        entry.type === 'dir'
          ? {
              key: `d:${entry.path}`,
              kind: 'dir',
              name: entry.name,
              sub: '',
              meta: '',
              tip: '',
              run: () => emit('navigate', { kind: 'folder', token, path: entry.path }),
            }
          : {
              key: `f:${entry.path}`,
              kind: 'doc',
              name: displayName(entry.name),
              sub: '',
              meta: [draftAgeLabel(entry.modified, Date.now()), sizeLabel(entry.size)].filter(Boolean).join(' · '),
              tip: entry.name,
              run: () => emit('open-file', token, entry.path, entry.name),
            },
      );
    return { items: entries };
  }
  return { items: [] };
}

const footnote = computed(() => {
  if (view.value.truncated) return `מוצגות ${SEARCH_LIMIT} התוצאות הראשונות — אפשר לצמצם את החיפוש`;
  if (props.place.kind === 'folder' && props.listing?.state === 'ready') {
    const notes: string[] = [];
    if (props.listing.listing.truncated) notes.push('התיקייה גדולה מאוד, ומוצג רק חלק ממנה');
    if (props.listing.refreshError) notes.push('הרענון נכשל; עדיין מוצגים הנתונים הקודמים');
    return notes.join(' · ');
  }
  return '';
});

const emptyTitle = computed(() => {
  if (props.query.trim() !== '') return `לא נמצא דבר שתואם ל„${props.query.trim()}”`;
  if (props.place.kind === 'folder') return 'אין בתיקייה הזאת קובצי Word';
  if (props.place.kind === 'personal') return 'אין ספרים אישיים בפורמט Word';
  return 'אין כאן ספרי Word';
});

const emptyHint = computed(() => {
  if (props.query.trim() !== '') return '';
  if (props.place.kind === 'folder') return 'מוצגים קבצים מסוג docx ו-docm, ותת-התיקיות שבה.';
  return '';
});

function ariaLabel(item: Item): string {
  const kind = item.kind === 'dir' ? 'תיקייה' : 'מסמך';
  return [item.name, kind, item.sub, item.meta].filter(Boolean).join(', ');
}

function activate(item: Item): void {
  if (props.busy) return;
  item.run();
}

/* ------------------------------------------------------------------ */
/* מקלדת                                                                */
/* ------------------------------------------------------------------ */

const activeIndex = ref(0);
const rowRefs = ref<(HTMLElement | null)[]>([]);
const searchRef = ref<HTMLInputElement | null>(null);

function setRowRef(el: unknown, index: number): void {
  rowRefs.value[index] = el instanceof HTMLElement ? el : null;
}

function focusRow(index: number): void {
  if (items.value.length === 0) return;
  const next = Math.max(0, Math.min(index, items.value.length - 1));
  activeIndex.value = next;
  void nextTick(() => {
    const el = rowRefs.value[next];
    el?.focus();
    el?.scrollIntoView?.({ block: 'nearest' });
  });
}

/**
 * ↑↓ בין שורות, Home/End, ו-→ (אחורה ב-RTL) עולה רמה — כמו Backspace
 * בסייר. ↑ מהשורה הראשונה חוזר לחיפוש, כדי שהקלדה-חץ-Enter תהיה תנועה אחת.
 */
function onKeydown(event: KeyboardEvent): void {
  let handled = true;
  if (event.key === 'ArrowDown') focusRow(activeIndex.value + 1);
  else if (event.key === 'ArrowUp') {
    if (activeIndex.value === 0) searchRef.value?.focus();
    else focusRow(activeIndex.value - 1);
  } else if (event.key === 'Home') focusRow(0);
  else if (event.key === 'End') focusRow(items.value.length - 1);
  else if (event.key === 'ArrowRight' || event.key === 'Backspace') {
    const up = crumbs.value[crumbs.value.length - 2]?.place;
    if (up) emit('navigate', up);
    else handled = false;
  } else handled = false;
  if (handled) event.preventDefault();
}

// מקום חדש או חיפוש חדש — המצביע חוזר לראש.
watch(
  () => [props.place, props.query] as const,
  () => {
    activeIndex.value = 0;
  },
);

defineExpose({
  focusSearch(): void {
    searchRef.value?.focus();
  },
  focusFirst(): void {
    focusRow(0);
  },
});
</script>

<style scoped>
.sl {
  display: flex;
  flex-direction: column;
  gap: 8px;
  min-width: 0;
  min-height: 0;
}

.sl-head {
  display: flex;
  align-items: center;
  gap: 8px;
  min-height: 28px;
}

.sl-title {
  flex: 1 1 auto;
  min-width: 0;
  margin: 0;
  display: flex;
  align-items: center;
  gap: 2px;
  overflow: hidden;
  font-size: 0.92em;
  font-weight: 700;
  color: var(--color-primary);
  white-space: nowrap;
}

.sl-crumb {
  flex: 0 1 auto;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  padding: 1px 4px;
  background: none;
  border: none;
  border-radius: var(--radius-xs);
  font: inherit;
  font-family: var(--font-main);
  color: var(--color-on-surface-variant);
  cursor: pointer;
}

button.sl-crumb:hover:not(:disabled) {
  background: var(--word-btn-hover);
  color: var(--color-primary);
}

.sl-crumb--current {
  color: var(--color-primary);
  cursor: default;
}

.sl-crumb-sep {
  flex: 0 0 auto;
  color: var(--color-on-surface-variant);
}

.sl-search {
  position: relative;
  display: flex;
  align-items: center;
  flex: 0 1 260px;
  min-width: 140px;
}

.sl-search__input {
  width: 100%;
  padding-block: 4px;
  padding-inline: 26px 24px;
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-xs);
  background: var(--color-surface);
  color: var(--color-on-surface);
  font-family: var(--font-main);
  font-size: inherit;
  outline: none;
  transition: border-color 0.1s, box-shadow 0.1s;
}

.sl-search__input::placeholder {
  color: var(--color-on-surface-variant);
}

.sl-search__input:focus {
  border-color: var(--word-blue);
  box-shadow: 0 0 0 1px var(--word-blue);
}

.sl-search__icon {
  position: absolute;
  inset-inline-start: 6px;
  pointer-events: none;
  color: var(--color-on-surface-variant);
}

.sl-search__clear {
  position: absolute;
  inset-inline-end: 2px;
  width: 20px;
  height: 20px;
  padding: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  background: none;
  border: none;
  border-radius: var(--radius-xs);
  color: var(--color-on-surface-variant);
  cursor: pointer;
}

.sl-search__clear:hover {
  background: var(--word-btn-hover);
  color: var(--color-on-surface);
}

.sl-iconbtn {
  flex: 0 0 auto;
  width: 26px;
  height: 26px;
  padding: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  background: none;
  border: 1px solid transparent;
  border-radius: var(--radius-sm);
  color: var(--color-on-surface-variant);
  cursor: pointer;
}

.sl-iconbtn:hover:not(:disabled) {
  background: var(--word-btn-hover);
  color: var(--color-on-surface);
}

.sl-iconbtn--spin :deep(svg) {
  animation: sl-spin 0.9s linear infinite;
}

/* פס הגלילה מימין — ראו `.rec-list` בדיאלוג. */
.sl-list {
  list-style: none;
  margin: 0;
  padding: 2px 0;
  flex: 1 1 auto;
  min-height: 96px;
  overflow-y: auto;
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-sm);
  background: var(--color-surface);
  direction: ltr;
}

.sl-row {
  direction: rtl;
}

.sl-open {
  display: flex;
  align-items: center;
  gap: 10px;
  width: 100%;
  min-height: 34px;
  padding-block: 4px;
  padding-inline: 10px 8px;
  background: none;
  border: 0;
  border-radius: 0;
  font: inherit;
  font-family: var(--font-main);
  text-align: start;
  color: var(--color-on-surface);
  cursor: pointer;
  outline: none;
  transition: background-color 0.08s ease;
}

.sl-open:hover:not(:disabled) {
  background: var(--word-btn-hover);
}

.sl-open:focus-visible {
  background: var(--word-btn-hover);
  box-shadow: inset 0 0 0 2px var(--word-blue);
}

.sl-open:active:not(:disabled) {
  background: var(--word-btn-active);
}

.sl-icon {
  flex: 0 0 auto;
}

.sl-icon--doc {
  color: var(--word-blue);
}

.sl-icon--dir {
  color: var(--color-on-surface-variant);
}

.sl-text {
  flex: 1 1 auto;
  min-width: 0;
  display: flex;
  flex-direction: column;
  line-height: 1.3;
}

.sl-name {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-weight: 500;
}

.sl-row--dir .sl-name {
  font-weight: 600;
}

.sl-sub {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-size: 0.84em;
  color: var(--color-on-surface-variant);
}

.sl-meta {
  flex: 0 0 auto;
  font-size: 0.86em;
  white-space: nowrap;
  color: var(--color-on-surface-variant);
}

.sl-chevron {
  flex: 0 0 auto;
  color: var(--color-on-surface-variant);
}

.sl-footnote {
  direction: rtl;
  padding: 6px 12px;
  font-size: 0.84em;
  color: var(--color-on-surface-variant);
}

.sl-state {
  flex: 1 1 auto;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 4px;
  padding: 16px 12px;
  text-align: center;
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-sm);
  background: var(--color-surface);
}

.sl-state__icon {
  color: var(--color-on-surface-variant);
}

.sl-state__title {
  margin: 0;
  font-weight: 600;
  color: var(--color-on-surface);
}

.sl-state__hint {
  margin: 0;
  max-width: 44ch;
  font-size: 0.88em;
  color: var(--color-on-surface-variant);
}

.sl-state__actions {
  display: flex;
  gap: 8px;
  margin-block-start: 6px;
}

.sl-link {
  padding: 2px 10px;
  background: none;
  border: 1px solid var(--color-outline);
  border-radius: var(--radius-xs);
  font-family: var(--font-main);
  font-size: 0.88em;
  color: var(--color-primary);
  cursor: pointer;
}

.sl-link:hover:not(:disabled) {
  background: var(--word-btn-hover);
}

.sl-link--danger {
  color: var(--color-error);
}

.sl-spinner {
  width: 22px;
  height: 22px;
  border: 2px solid var(--color-outline-variant);
  border-block-start-color: var(--word-blue);
  border-radius: 50%;
  animation: sl-spin 0.8s linear infinite;
}

@keyframes sl-spin {
  to {
    transform: rotate(360deg);
  }
}

@media (prefers-reduced-motion: reduce) {
  .sl-open,
  .sl-search__input {
    transition: none;
  }

  .sl-spinner,
  .sl-iconbtn--spin :deep(svg) {
    animation: none;
  }
}
</style>
