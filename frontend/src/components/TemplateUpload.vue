<script setup>
import { ref } from 'vue'

defineProps({ disabled: Boolean })
const emit = defineEmits(['selected', 'cleared'])

const fileInput = ref(null)
const dragging = ref(false)
const file = ref(null)

function openPicker() {
  if (fileInput.value) fileInput.value.click()
}

function onPick(event) {
  handleFiles(event.target.files)
  event.target.value = ''
}

function handleFiles(list) {
  if (!list || !list.length) return
  const f = list[0]
  const name = (f.name || '').toLowerCase()
  if (!name.endsWith('.xlsx') && !name.endsWith('.xlsm')) return
  file.value = f
  emit('selected', f)
}

function onDrop(event) {
  dragging.value = false
  handleFiles(event.dataTransfer.files)
}

function onDragOver(event) {
  event.preventDefault()
  dragging.value = true
}

function onDragLeave() {
  dragging.value = false
}

function remove() {
  file.value = null
  emit('cleared')
}
</script>

<template>
  <div class="tpl">
    <div
      v-if="!file"
      class="tpl-drop"
      :class="{ dragging, disabled }"
      role="button"
      tabindex="0"
      @click="openPicker"
      @keydown.enter="openPicker"
      @keydown.space.prevent="openPicker"
      @dragover.prevent="onDragOver"
      @dragleave="onDragLeave"
      @drop.prevent="onDrop"
    >
      <input
        ref="fileInput"
        type="file"
        accept=".xlsx,.xlsm"
        hidden
        @change="onPick"
      />
      <svg viewBox="0 0 22 22" width="18" height="18" fill="none" aria-hidden="true">
        <rect x="3" y="2" width="16" height="18" rx="2.5" stroke="var(--primary)" stroke-width="1.7" />
        <path d="M8 8.5h6M8 11.5h6M8 14.5h6" stroke="var(--primary)" stroke-width="1.7" stroke-linecap="round" />
      </svg>
      <span class="tpl-title">上传表格模板（.xlsx）</span>
      <span class="tpl-hint">图片将插入到这个表格里</span>
    </div>

    <div v-else class="tpl-chip" :class="{ disabled }">
      <svg viewBox="0 0 22 22" width="17" height="17" fill="none" aria-hidden="true">
        <path d="M6 2h7l4 4v14H6V2z" stroke="var(--primary)" stroke-width="1.7" stroke-linejoin="round" />
        <path d="M13 2v4h4" stroke="var(--primary)" stroke-width="1.7" stroke-linejoin="round" />
      </svg>
      <span class="tpl-name" :title="file.name">{{ file.name }}</span>
      <button
        class="tpl-remove"
        type="button"
        :disabled="disabled"
        :aria-label="'移除模板'"
        @click="remove"
      >✕</button>
    </div>
  </div>
</template>

<style scoped>
.tpl {
  width: 100%;
}

.tpl-drop {
  display: inline-flex;
  align-items: center;
  gap: 10px;
  padding: 12px 18px;
  border: 1.5px dashed var(--border-strong);
  border-radius: var(--radius-s);
  background: var(--surface);
  cursor: pointer;
  transition: border-color 0.2s var(--ease-out), background 0.2s var(--ease-out);
}

.tpl-drop:hover,
.tpl-drop.dragging {
  border-color: var(--primary);
  background: var(--primary-soft);
}

.tpl-drop:focus-visible {
  outline: 2px solid var(--primary);
  outline-offset: 2px;
}

.tpl-drop.disabled {
  opacity: 0.55;
  pointer-events: none;
}

.tpl-title {
  font-size: 13px;
  font-weight: 700;
}

.tpl-hint {
  font-size: 12px;
  color: var(--text-faint);
}

.tpl-chip {
  display: inline-flex;
  align-items: center;
  gap: 9px;
  padding: 8px 12px;
  border: 1.5px solid var(--primary);
  border-radius: var(--radius-s);
  background: var(--primary-soft);
  max-width: 100%;
}

.tpl-chip.disabled {
  opacity: 0.6;
}

.tpl-name {
  font-size: 13px;
  font-weight: 700;
  color: var(--primary-ink);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  max-width: 320px;
}

.tpl-remove {
  border: none;
  background: transparent;
  color: var(--primary-ink);
  font-size: 12px;
  padding: 2px 5px;
  border-radius: 5px;
  cursor: pointer;
}

.tpl-remove:hover:not(:disabled) {
  background: var(--danger-soft);
  color: var(--danger);
}

.tpl-remove:disabled {
  opacity: 0.4;
  cursor: not-allowed;
}
</style>
