<script setup>
defineProps({ modelValue: String, disabled: Boolean })
defineEmits(['update:modelValue'])

const modes = [
  {
    id: 'receipt',
    title: '收款组',
    sub: '发票识别与信息转 Excel',
    desc: '按人员类型识别中文发票并生成对应表格',
    colorTheme: 'green'
  },
  {
    id: '3',
    title: '付款组',
    sub: '发票明细识别转 Excel',
    desc: '识别英文发票明细，支持 13 种发票版式',
    colorTheme: 'blue'
  },
  {
    id: '4',
    title: '报表组',
    sub: '表格处理',
    desc: 'Excel 报表数据清洗、整理与生成结果',
    colorTheme: 'emerald'
  },
]
</script>

<template>
  <div class="sidebar-mode-list" role="radiogroup" aria-label="工作组别">
    <div
      v-for="item in modes"
      :key="item.id"
      class="sidebar-mode-item"
      :class="{
        'is-active': modelValue === item.id,
        'is-disabled': disabled
      }"
      role="radio"
      :aria-checked="modelValue === item.id"
      tabindex="0"
      @click="!disabled && $emit('update:modelValue', item.id)"
      @keydown.enter="!disabled && $emit('update:modelValue', item.id)"
      @keydown.space.prevent="!disabled && $emit('update:modelValue', item.id)"
    >
      <!-- 图标 -->
      <div class="item-icon" :class="`theme-${item.colorTheme}`">
        <!-- 收款组图标 -->
        <svg v-if="item.id === 'receipt'" viewBox="0 0 24 24" width="20" height="20" fill="none">
          <rect x="4" y="3" width="16" height="18" rx="2.5" fill="currentColor" fill-opacity="0.16"/>
          <path d="M4 7h16" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>
          <path d="M8 12h8M8 15h5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>
          <path d="M14.5 13.5l1.5 2.5m-3 0l1.5-2.5v4" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>

        <!-- 付款组图标 -->
        <svg v-else-if="item.id === '3'" viewBox="0 0 24 24" width="20" height="20" fill="none">
          <path d="M5 8V5h3M16 5h3v3M5 16v3h3M19 16v3h-3" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
          <rect x="8" y="9" width="8" height="6" rx="1.2" fill="currentColor" fill-opacity="0.25"/>
          <path d="M9 12h6" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>
        </svg>

        <!-- 报表组图标 -->
        <svg v-else-if="item.id === '4'" viewBox="0 0 24 24" width="20" height="20" fill="none">
          <rect x="4" y="13" width="3.5" height="8" rx="1" fill="currentColor"/>
          <rect x="10.25" y="8" width="3.5" height="13" rx="1" fill="currentColor"/>
          <rect x="16.5" y="3.5" width="3.5" height="17.5" rx="1" fill="currentColor"/>
        </svg>

        <!-- 资金组图标 -->
        <svg v-else viewBox="0 0 24 24" width="20" height="20" fill="none">
          <ellipse cx="12" cy="6" rx="8" ry="3" fill="currentColor" fill-opacity="0.3" stroke="currentColor" stroke-width="1.6"/>
          <path d="M4 6v5c0 1.66 3.58 3 8 3s8-1.34 8-3V6" stroke="currentColor" stroke-width="1.6"/>
          <path d="M4 11v5c0 1.66 3.58 3 8 3s8-1.34 8-3v-5" stroke="currentColor" stroke-width="1.6"/>
        </svg>
      </div>

      <!-- 文本内容 -->
      <div class="item-content">
        <div class="item-title-row">
          <span class="item-title">{{ item.title }}</span>
        </div>
        <span class="item-sub">{{ item.sub }}</span>
      </div>

      <!-- 右侧勾选圆圈 -->
      <div class="item-check" :class="{ 'checked': modelValue === item.id }">
        <svg v-if="modelValue === item.id" viewBox="0 0 16 16" width="10" height="10" fill="none">
          <path d="M3.5 8.5L6.5 11.5L12.5 4.5" stroke="#fff" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
      </div>
    </div>
  </div>
</template>

<style scoped>
.sidebar-mode-list {
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.sidebar-mode-item {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 12px 14px;
  background: var(--surface);
  border: 1.5px solid var(--border);
  border-radius: var(--radius-m);
  cursor: pointer;
  transition: all 0.2s var(--ease-smooth);
  user-select: none;
}

.sidebar-mode-item:hover:not(.is-disabled) {
  border-color: var(--border-strong);
  background: var(--surface-hover);
  transform: translateY(-1px);
}

.sidebar-mode-item.is-active {
  border-color: var(--primary);
  background: #f4fbf8;
  box-shadow: 0 2px 10px rgba(0, 135, 101, 0.12);
}

.sidebar-mode-item.is-disabled {
  opacity: 0.55;
  cursor: not-allowed;
}

.sidebar-mode-item:focus-visible {
  outline: 2px solid var(--primary);
  outline-offset: 1px;
}

/* 图标容器 */
.item-icon {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 36px;
  height: 36px;
  border-radius: 9px;
  flex-shrink: 0;
}

.item-icon.theme-green {
  background: #e6f6f1;
  color: #008765;
}

.item-icon.theme-blue {
  background: #e8f0fe;
  color: #3b82f6;
}

.item-icon.theme-emerald {
  background: #e6f8f0;
  color: #10b981;
}

.item-icon.theme-indigo {
  background: #edf1fe;
  color: #4f46e5;
}

/* 内容文字 */
.item-content {
  flex: 1;
  min-width: 0;
  display: flex;
  flex-direction: column;
  gap: 2px;
}

.item-title-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.item-title {
  font-size: 14.5px;
  font-weight: 700;
  color: var(--text);
  line-height: 1.2;
}

.sidebar-mode-item.is-active .item-title {
  color: var(--primary-ink);
}

.item-sub {
  font-size: 11.5px;
  color: var(--text-soft);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

/* 勾选圆圈 */
.item-check {
  width: 18px;
  height: 18px;
  border-radius: 50%;
  border: 1.5px solid var(--border-strong);
  background: #fff;
  display: flex;
  align-items: center;
  justify-content: center;
  flex-shrink: 0;
  transition: all 0.18s ease;
}

.item-check.checked {
  border-color: var(--primary);
  background: var(--primary);
}
</style>
