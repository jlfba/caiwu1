<script setup>
defineProps({ modelValue: String, disabled: Boolean })
defineEmits(['update:modelValue'])
</script>

<template>
  <div class="receipt-subtype" role="radiogroup" aria-label="选择发票类型">
    <button
      class="type-option"
      :class="{ active: modelValue === 'invoice' }"
      role="radio"
      :aria-checked="modelValue === 'invoice'"
      type="button"
      :disabled="disabled"
      @click="$emit('update:modelValue', 'invoice')"
    >
      <span class="radio" :class="{ on: modelValue === 'invoice' }"><span class="radio-dot"></span></span>
      <span class="type-label">发票识别</span>
    </button>
  </div>
</template>

<style scoped>
.receipt-subtype{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}.type-option{display:flex;gap:13px;align-items:center;text-align:left;padding:15px 17px;background:var(--surface);border:1.5px solid var(--border);border-radius:var(--radius-s);cursor:pointer;transition:border-color .2s var(--ease-out),background .2s var(--ease-out),transform .2s var(--ease-out)}.type-option:hover:not(:disabled){border-color:var(--border-strong);transform:translateY(-1px)}.type-option:focus-visible{outline:2px solid var(--primary);outline-offset:2px}.type-option.active{border-color:var(--primary);background:var(--primary-soft)}.type-option:disabled{opacity:.6;cursor:not-allowed}.radio{width:18px;height:18px;border-radius:50%;border:1.5px solid var(--border-strong);display:grid;place-items:center;flex-shrink:0}.radio.on{border-color:var(--primary)}.radio-dot{width:0;height:0;border-radius:50%;background:var(--primary);transition:all .2s var(--ease-out)}.radio.on .radio-dot{width:8px;height:8px}.type-label{font-size:14.5px;font-weight:700;color:var(--text)}.type-option.active .type-label{color:var(--primary-ink)}@media(max-width:640px){.receipt-subtype{grid-template-columns:1fr}}
</style>
