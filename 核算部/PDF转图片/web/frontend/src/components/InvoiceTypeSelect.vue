<script setup>
defineProps({ modelValue: String })
defineEmits(['update:modelValue'])

const types = [
  { id: '1', label: 'canexs', hint: '美国货运发票（INVOICE / TRACKING NO. / 明细行）' },
  { id: '2', label: '精准', hint: 'Accuracy Customs Brokers 清关发票' }
]
</script>

<template>
  <section class="inv-type">
    <div
      v-for="t in types"
      :key="t.id"
      class="type-option"
      :class="{ active: modelValue === t.id }"
      role="radio"
      :aria-checked="modelValue === t.id"
      tabindex="0"
      @click="$emit('update:modelValue', t.id)"
      @keydown.enter="$emit('update:modelValue', t.id)"
    >
      <span class="radio" :class="{ on: modelValue === t.id }">
        <span class="radio-dot" v-if="modelValue === t.id"></span>
      </span>
      <span class="type-body">
        <span class="type-label">{{ t.label }}</span>
        <span class="type-hint">{{ t.hint }}</span>
      </span>
    </div>
  </section>
</template>

<style scoped>
.inv-type {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 12px;
}

.type-option {
  display: flex;
  gap: 12px;
  align-items: flex-start;
  padding: 16px 18px;
  background: var(--card);
  border: 1.5px solid var(--border);
  border-radius: 12px;
  cursor: pointer;
  transition: all 0.15s ease;
}

.type-option:hover {
  border-color: #99f6e4;
}

.type-option.active {
  border-color: var(--primary);
  background: var(--primary-softer);
}

.radio {
  width: 18px;
  height: 18px;
  border-radius: 50%;
  border: 2px solid var(--border);
  display: grid;
  place-items: center;
  margin-top: 2px;
  flex-shrink: 0;
  transition: all 0.15s ease;
}

.radio.on {
  border-color: var(--primary);
}

.radio-dot {
  width: 9px;
  height: 9px;
  border-radius: 50%;
  background: var(--primary);
}

.type-body {
  display: flex;
  flex-direction: column;
  gap: 3px;
}

.type-label {
  font-weight: 700;
  font-size: 15px;
}

.type-hint {
  font-size: 12px;
  color: var(--text-soft);
  line-height: 1.5;
}

@media (max-width: 640px) {
  .inv-type {
    grid-template-columns: 1fr;
  }
}
</style>
