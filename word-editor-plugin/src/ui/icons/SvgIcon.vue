<template>
  <!--
    ה-span מוסתר מקוראי מסך: האייקון דקורטיבי בלבד, והשם הנגיש של הפקד מגיע
    מה-`title`/`aria-label` של הכפתור שעוטף אותו. בלי זה הכפתור נקרא פעמיים.
  -->
  <span
    class="svg-icon"
    aria-hidden="true"
    :style="{ width: `${size}px`, height: `${size}px` }"
    v-html="svgContent"
  />
</template>

<script setup lang="ts">
import { computed } from 'vue';
import { ICONS } from './icons';

const props = withDefaults(
  defineProps<{
    name: string;
    size?: number;
  }>(),
  {
    size: 16,
  }
);

/**
 * `v-html` כאן אינו נתיב הזרקה: הערכים הם מחרוזות SVG קבועות מ-icons.ts,
 * שנקבעות בזמן קומפילציה ואינן נגזרות מקלט משתמש או מתוכן המסמך. שם שאינו
 * מוכר נותן מחרוזת ריקה ולא נזרק — tests/unit/icons.test.ts סורק את קובצי
 * ה-Vue ומאמת שכל שם שבשימוש קיים, כדי שכפתור לא יישאר בלי אייקון בשקט.
 */
const svgContent = computed(() => ICONS[props.name] || '');
</script>

<style scoped>
.svg-icon {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  flex-shrink: 0;
  line-height: 1;
}

.svg-icon :deep(svg) {
  width: 100%;
  height: 100%;
  fill: currentColor;
  shape-rendering: geometricPrecision;
}
</style>
