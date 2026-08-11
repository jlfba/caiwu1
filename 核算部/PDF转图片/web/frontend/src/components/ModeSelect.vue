<script setup>
defineProps({ modelValue: String })
defineEmits(['update:modelValue'])

const modes = [
  {
    id: '1',
    tag: '收款组',
    title: 'PDF 转图片 + 发票识别',
    desc: '逐页渲染 PDF 为图片，识别发票号码、购买方、销售方、金额并重命名，生成含图片的 Excel 表格。',
    icon: '🖼️'
  },
  {
    id: '2',
    tag: '付款组',
    title: '发票明细识别并转 Excel',
    desc: '识别英文发票的 INVOICE / TRACKING NO. / 明细行（支持 canexs、精准两种版式），生成发票明细 Excel。',
    icon: '📊'
  }
]
</script>

<template>
  <section class="mode-select">
    <button
      v-for="m in modes"
      :key="m.id"
      class="mode-card"
      :class="{ active: modelValue === m.id }"
      type="button"
      @click="$emit('update:modelValue', m.id)"
    >
      <span class="mode-icon" :class="{ active: modelValue === m.id }">{{ m.icon }}</span>
      <span class="mode-body">
        <span class="mode-head">
          <span class="mode-tag">{{ m.tag }}</span>
          <span class="mode-title">{{ m.title }}</span>
        </span>
        <span class="mode-desc">{{ m.desc }}</span>
      </span>
      <span class="mode-check" :class="{ on: modelValue === m.id }">✓</span>
    </button>
  </section>
</template>

<style scoped>
.mode-select {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 14px;
}

.mode-card {
  display: flex;
  gap: 14px;
  align-items: flex-start;
  padding: 20px;
  background: var(--card);
  border: 1.5px solid var(--border);
  border-radius: var(--radius);
  box-shadow: var(--shadow-sm);
  text-align: left;
  transition: all 0.18s ease;
}

.mode-card:hover {
  border-color: #99f6e4;
  transform: translateY(-2px);
  box-shadow: var(--shadow-md);
}

.mode-card.active {
  border-color: var(--primary);
  background: var(--primary-softer);
  box-shadow: var(--shadow-md);
}

.mode-icon {
  font-size: 26px;
  line-height: 1;
  padding: 10px;
  background: #f8fafc;
  border-radius: 12px;
}

.mode-icon.active {
  background: var(--primary-soft);
}

.mode-body {
  display: flex;
  flex-direction: column;
  gap: 6px;
  flex: 1;
  min-width: 0;
}

.mode-head {
  display: flex;
  flex-direction: column;
  gap: 2px;
}

.mode-tag {
  font-size: 11px;
  font-weight: 600;
  color: var(--primary);
  background: var(--primary-soft);
  padding: 2px 8px;
  border-radius: 999px;
  width: fit-content;
}

.mode-title {
  font-size: 15px;
  font-weight: 700;
}

.mode-desc {
  font-size: 12.5px;
  color: var(--text-soft);
  line-height: 1.55;
}

.mode-check {
  width: 20px;
  height: 20px;
  border-radius: 50%;
  border: 2px solid var(--border);
  display: grid;
  place-items: center;
  font-size: 12px;
  color: transparent;
  flex-shrink: 0;
  margin-top: 2px;
  transition: all 0.18s ease;
}

.mode-check.on {
  background: var(--primary);
  border-color: var(--primary);
  color: #fff;
}

@media (max-width: 640px) {
  .mode-select {
    grid-template-columns: 1fr;
  }
}
</style>
