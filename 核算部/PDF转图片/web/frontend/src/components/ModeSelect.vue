<script setup>
defineProps({ modelValue: String, disabled: Boolean })
defineEmits(['update:modelValue'])
</script>

<template>
  <div class="mode-select" role="radiogroup" aria-label="选择功能">
    <div class="mode-card" :class="{ active: modelValue === '1' || modelValue === '2' }">
      <button class="mode-tile" type="button" role="radio" :aria-checked="modelValue === '1' || modelValue === '2'" :disabled="disabled" @click="$emit('update:modelValue', modelValue === '2' ? '2' : '1')">
        <span class="tile-head"><span class="tile-title">收款组</span><span class="tile-check" :class="{ on: modelValue === '1' || modelValue === '2' }"><svg viewBox="0 0 16 16" width="13" height="13" fill="none" aria-hidden="true"><path d="M3 8.5l3.2 3L13 4.5" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" /></svg></span></span>
        <span class="tile-sub">发票识别与信息转 Excel</span><span class="tile-desc">选择收款组处理模式，支持发票转图片或提取发票信息生成表格。</span>
      </button>
      <div v-if="modelValue === '1' || modelValue === '2'" class="submode-picker"><span class="submode-label">选择收款组模式</span><div class="submode-options" role="radiogroup" aria-label="选择收款组模式"><button type="button" class="submode-btn" :class="{ on: modelValue === '1' }" :disabled="disabled" role="radio" :aria-checked="modelValue === '1'" @click="$emit('update:modelValue', '1')">模式 1：含图片表格</button><button type="button" class="submode-btn" :class="{ on: modelValue === '2' }" :disabled="disabled" role="radio" :aria-checked="modelValue === '2'" @click="$emit('update:modelValue', '2')">模式 2：信息表格</button></div></div>
    </div>
    <button class="mode-tile" :class="{ active: modelValue === '3' }" role="radio" :aria-checked="modelValue === '3'" type="button" :disabled="disabled" @click="$emit('update:modelValue', '3')"><span class="tile-head"><span class="tile-title">付款组</span><span class="tile-check" :class="{ on: modelValue === '3' }"><svg viewBox="0 0 16 16" width="13" height="13" fill="none" aria-hidden="true"><path d="M3 8.5l3.2 3L13 4.5" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" /></svg></span></span><span class="tile-sub">发票明细识别转 Excel</span><span class="tile-desc">识别英文发票的 INVOICE / 明细行，支持现有十三种发票版式。</span></button>
  </div>
</template>

<style scoped>
.mode-select{display:grid;grid-template-columns:1fr 1fr;gap:14px}.mode-card{display:flex;flex-direction:column;gap:0;background:var(--surface);border:1.5px solid var(--border);border-radius:var(--radius);overflow:hidden}.mode-card.active{border-color:var(--primary);background:var(--primary-soft);box-shadow:var(--shadow-sm)}.mode-tile{display:flex;flex-direction:column;gap:6px;align-items:stretch;text-align:left;padding:20px;background:transparent;border:0;cursor:pointer}.mode-tile:hover:not(:disabled){background:color-mix(in srgb,var(--primary-soft) 45%,transparent)}.mode-tile:focus-visible,.submode-btn:focus-visible{outline:2px solid var(--primary);outline-offset:-2px}.mode-tile:disabled,.submode-btn:disabled{opacity:.6;cursor:not-allowed}.tile-head{display:flex;justify-content:space-between;align-items:center}.tile-title{font-size:17px;font-weight:800;letter-spacing:.3px}.tile-check{width:22px;height:22px;border-radius:50%;border:1.5px solid var(--border-strong);display:grid;place-items:center;color:transparent}.tile-check.on{background:var(--primary);border-color:var(--primary);color:#fff}.tile-sub{font-size:13px;font-weight:600;color:var(--text)}.tile-desc{font-size:12.5px;color:var(--text-soft);line-height:1.6}.submode-picker{display:flex;flex-direction:column;gap:8px;padding:12px 20px 16px;border-top:1px dashed var(--border-strong)}.submode-label{font-size:12px;font-weight:700;color:var(--text-soft)}.submode-options{display:grid;grid-template-columns:1fr;gap:6px}.submode-btn{min-height:36px;padding:6px 10px;border:1px solid var(--border);border-radius:8px;background:var(--surface);color:var(--text-soft);text-align:left;font-size:12px;cursor:pointer}.submode-btn.on{border-color:var(--primary);color:var(--primary-ink);font-weight:700;box-shadow:inset 3px 0 0 var(--primary)}
@media (max-width:640px){.mode-select{grid-template-columns:1fr}}
</style>
