<script setup>
defineProps({ modelValue: String, disabled: Boolean })
defineEmits(['update:modelValue'])

const types = [
  { id: '1', label: 'canexs', hint: '美国货运发票 · INVOICE / TRACKING NO. / 明细行' },
  { id: '2', label: '精准', hint: 'Accuracy Customs Brokers 清关发票' },
  { id: '3', label: '创时亚马逊卡派', hint: 'Invoice Number / Reference / Description 明细' },
  { id: '4', label: '创时卡派', hint: '// 链 Description · 邮编换行拼接' },
  { id: '5', label: '创时清关费', hint: '单行明细 · 海运清关 / DPD / DUTY & VAT' },
  { id: '6', label: '创时附加费', hint: '多行 34 开头单号 · 支持跨页续行' },
  { id: '7', label: 'MAX萨凡纳', hint: 'MAXPORTLINK · Ship to / Invoice details / 柜号 / 邮编' },
  { id: '8', label: 'MAX纽约', hint: 'MAX GLOBAL LOGISTICS · 同 MAX萨凡纳 布局' },
  { id: '9', label: 'AA', hint: 'TX-AA LOGISTICS · 同 MAX萨凡纳 布局' },
  { id: '10', label: 'JCK', hint: 'JCK LOGISTICS · Container No. / 十列明细' },
  { id: '11', label: 'MKK', hint: '编号 / 柜号 / 主单号 / Bill To · 四列费用明细' },
  { id: '12', label: 'DINO', hint: '发票号 / Product or service / Description 多行合并' },
  { id: '12', label: 'DINO', hint: '发票号 / Product or service / Description 多行合并' }
]
</script>

<template>
  <div class="inv-type" role="radiogroup" aria-label="选择发票类型">
    <button
      v-for="t in types"
      :key="t.id"
      class="type-option"
      :class="{ active: modelValue === t.id }"
      role="radio"
      :aria-checked="modelValue === t.id"
      type="button"
      :disabled="disabled"
      @click="$emit('update:modelValue', t.id)"
    >
      <span class="radio" :class="{ on: modelValue === t.id }">
        <span class="radio-dot"></span>
      </span>
      <span class="type-text">
        <span class="type-label">{{ t.label }}</span>
        <span class="type-hint">{{ t.hint }}</span>
      </span>
    </button>
  </div>
</template>

<style scoped>
.inv-type {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 12px;
}

.type-option {
  display: flex;
  gap: 13px;
  align-items: flex-start;
  text-align: left;
  padding: 15px 17px;
  background: var(--surface);
  border: 1.5px solid var(--border);
  border-radius: var(--radius-s);
  cursor: pointer;
  transition: border-color 0.2s var(--ease-out), background 0.2s var(--ease-out),
    transform 0.2s var(--ease-out);
}

.type-option:hover:not(:disabled) {
  border-color: var(--border-strong);
  transform: translateY(-1px);
}

.type-option:focus-visible {
  outline: 2px solid var(--primary);
  outline-offset: 2px;
}

.type-option.active {
  border-color: var(--primary);
  background: var(--primary-soft);
}

.type-option:disabled {
  opacity: 0.6;
  cursor: not-allowed;
}

.radio {
  width: 18px;
  height: 18px;
  border-radius: 50%;
  border: 1.5px solid var(--border-strong);
  display: grid;
  place-items: center;
  margin-top: 2px;
  flex-shrink: 0;
  transition: all 0.2s var(--ease-out);
}

.radio.on {
  border-color: var(--primary);
}

.radio-dot {
  width: 0;
  height: 0;
  border-radius: 50%;
  background: var(--primary);
  transition: all 0.2s var(--ease-out);
}

.radio.on .radio-dot {
  width: 8px;
  height: 8px;
}

.type-text {
  display: flex;
  flex-direction: column;
  gap: 3px;
}

.type-label {
  font-size: 14.5px;
  font-weight: 700;
  color: var(--text);
}

.type-option.active .type-label {
  color: var(--primary-ink);
}

.type-hint {
  font-size: 11.5px;
  color: var(--text-soft);
  line-height: 1.5;
}

@media (max-width: 640px) {
  .inv-type {
    grid-template-columns: 1fr;
  }
}
</style>
