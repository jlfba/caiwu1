<script setup>
defineProps({ modelValue: String, disabled: Boolean })
defineEmits(['update:modelValue'])
const types = [
  { id: 'invoice', label: '发票识别', ready: true },
  { id: 'wechat', label: '微信', ready: false },
  { id: 'alipay', label: '支付宝', ready: false },
  { id: 'huolala', label: '货拉拉', ready: false }
]
</script>

<template>
  <div class="receipt-subtype" role="radiogroup" aria-label="选择发票类型">
    <button
      v-for="type in types"
      :key="type.id"
      class="type-option"
      :class="{ active: modelValue === type.id, unavailable: !type.ready }"
      role="radio"
      :aria-checked="modelValue === type.id"
      type="button"
      :disabled="disabled"
      @click="$emit('update:modelValue', type.id)"
    >
      <span class="radio" :class="{ on: modelValue === type.id }"><span class="radio-dot"></span></span>
      <span class="type-copy"><span class="type-label">{{ type.label }}</span><span v-if="!type.ready" class="type-hint">暂未接入</span></span>
    </button>
  </div>
</template>

<style scoped>
.receipt-subtype{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}.type-option{display:flex;gap:13px;align-items:center;text-align:left;padding:15px 17px;background:var(--surface);border:1.5px solid var(--border);border-radius:var(--radius-s);cursor:pointer;transition:border-color .2s var(--ease-out),background .2s var(--ease-out),transform .2s var(--ease-out)}.type-option:hover:not(:disabled){border-color:var(--border-strong);transform:translateY(-1px)}.type-option:focus-visible{outline:2px solid var(--primary);outline-offset:2px}.type-option.active{border-color:var(--primary);background:var(--primary-soft)}.type-option.unavailable{background:var(--surface-2)}.type-option:disabled{opacity:.6;cursor:not-allowed}.radio{width:18px;height:18px;border-radius:50%;border:1.5px solid var(--border-strong);display:grid;place-items:center;flex-shrink:0}.radio.on{border-color:var(--primary)}.radio-dot{width:0;height:0;border-radius:50%;background:var(--primary);transition:all .2s var(--ease-out)}.radio.on .radio-dot{width:8px;height:8px}.type-copy{display:flex;flex-direction:column;gap:3px}.type-label{font-size:14.5px;font-weight:700;color:var(--text)}.type-option.active .type-label{color:var(--primary-ink)}.type-hint{font-size:11px;color:var(--text-faint)}@media(max-width:760px){.receipt-subtype{grid-template-columns:repeat(2,minmax(0,1fr))}@media(max-width:640px){.receipt-subtype{grid-template-columns:1fr}}}
</style>
