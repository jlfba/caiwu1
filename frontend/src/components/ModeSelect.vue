<script setup>
defineProps({ modelValue: String })
defineEmits(['update:modelValue'])

const modes = [
  {
    id: '1',
    title: '收款组',
    sub: 'PDF 转图片 + 发票识别',
    desc: '逐页渲染为图片，识别发票号码、购买方、销售方、金额并重命名，输出含图 Excel。'
  },
  {
    id: '2',
    title: '付款组',
    sub: '发票明细识别转 Excel',
    desc: '识别英文发票的 INVOICE / 明细行，支持 canexs、精准、创时亚马逊卡派、创时卡派、创时清关费、创时附加费六种版式。'
  }
]
</script>

<template>
  <div class="mode-select" role="radiogroup" aria-label="选择功能">
    <button
      v-for="m in modes"
      :key="m.id"
      class="mode-tile"
      :class="{ active: modelValue === m.id }"
      role="radio"
      :aria-checked="modelValue === m.id"
      type="button"
      :disabled="disabled"
      @click="$emit('update:modelValue', m.id)"
    >
      <span class="tile-head">
        <span class="tile-title">{{ m.title }}</span>
        <span class="tile-check" :class="{ on: modelValue === m.id }">
          <svg viewBox="0 0 16 16" width="13" height="13" fill="none">
            <path d="M3 8.5l3.2 3L13 4.5" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" />
          </svg>
        </span>
      </span>
      <span class="tile-sub">{{ m.sub }}</span>
      <span class="tile-desc">{{ m.desc }}</span>
    </button>
  </div>
</template>

<style scoped>
.mode-select {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 14px;
}

.mode-tile {
  display: flex;
  flex-direction: column;
  gap: 6px;
  align-items: stretch;
  text-align: left;
  padding: 20px 20px 18px;
  background: var(--surface);
  border: 1.5px solid var(--border);
  border-radius: var(--radius);
  cursor: pointer;
  transition: transform 0.22s var(--ease-out), border-color 0.22s var(--ease-out),
    box-shadow 0.22s var(--ease-out), background 0.22s var(--ease-out);
}

.mode-tile:hover:not(:disabled) {
  transform: translateY(-2px);
  border-color: var(--border-strong);
  box-shadow: var(--shadow-md);
}

.mode-tile:focus-visible {
  outline: 2px solid var(--primary);
  outline-offset: 2px;
}

.mode-tile.active {
  border-color: var(--primary);
  background: var(--primary-soft);
  box-shadow: var(--shadow-sm);
}

.mode-tile.active .tile-title {
  color: var(--primary-ink);
}

.mode-tile:disabled {
  opacity: 0.6;
  cursor: not-allowed;
}

.tile-head {
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.tile-title {
  font-size: 17px;
  font-weight: 800;
  letter-spacing: 0.3px;
}

.tile-check {
  width: 22px;
  height: 22px;
  border-radius: 50%;
  border: 1.5px solid var(--border-strong);
  display: grid;
  place-items: center;
  color: transparent;
  transition: all 0.2s var(--ease-out);
}

.tile-check.on {
  background: var(--primary);
  border-color: var(--primary);
  color: #fff;
}

.tile-sub {
  font-size: 13px;
  font-weight: 600;
  color: var(--text);
}

.tile-desc {
  font-size: 12.5px;
  color: var(--text-soft);
  line-height: 1.6;
}

@media (max-width: 640px) {
  .mode-select {
    grid-template-columns: 1fr;
  }
}
</style>
