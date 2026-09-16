<script setup>
import { ref } from 'vue'

defineProps({ disabled: Boolean })
const emit = defineEmits(['selected', 'cleared'])

const fileInput = ref(null)
const dragging = ref(false)
const file = ref(null)
const localError = ref('')

function openPicker() {
  if (fileInput.value) fileInput.value.click()
}

function handleFiles(list) {
  if (!list || !list.length) return
  const selected = list[0]
  const name = (selected.name || '').toLowerCase()
  if (!name.endsWith('.xlsx') && !name.endsWith('.xlsm')) {
    localError.value = '请选择 .xlsx 或 .xlsm 格式的表格'
    return
  }
  localError.value = ''
  file.value = selected
  emit('selected', selected)
}

function onPick(event) {
  handleFiles(event.target.files)
  event.target.value = ''
}

function onDrop(event) {
  dragging.value = false
  handleFiles(event.dataTransfer.files)
}

function remove() {
  file.value = null
  localError.value = ''
  emit('cleared')
}
</script>

<template>
  <div class="report-uploader">
    <div
      v-if="!file"
      class="report-drop"
      :class="{ dragging, disabled }"
      role="button"
      :tabindex="disabled ? -1 : 0"
      aria-label="上传需要处理的 Excel 表格"
      @click="openPicker"
      @keydown.enter="openPicker"
      @keydown.space.prevent="openPicker"
      @dragover.prevent="dragging = true"
      @dragleave="dragging = false"
      @drop.prevent="onDrop"
    >
      <input ref="fileInput" type="file" accept=".xlsx,.xlsm" hidden :disabled="disabled" @change="onPick" />
      <svg viewBox="0 0 28 28" width="28" height="28" fill="none" aria-hidden="true">
        <path d="M7 3h10l5 5v17H7V3Z" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round" />
        <path d="M17 3v5h5M10.5 13h8M10.5 17h8M10.5 21h5" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" />
      </svg>
      <span class="report-drop-title">拖入需要处理的 Excel 表格</span>
      <span class="report-drop-hint">或点击选择文件，支持 .xlsx / .xlsm</span>
    </div>

    <div v-else class="report-file" :class="{ disabled }">
      <svg viewBox="0 0 22 22" width="20" height="20" fill="none" aria-hidden="true">
        <path d="M6 2h7l4 4v14H6V2Z" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round" />
        <path d="M13 2v4h4" stroke="currentColor" stroke-width="1.7" stroke-linejoin="round" />
      </svg>
      <span class="report-file-name" :title="file.name">{{ file.name }}</span>
      <button class="report-remove" type="button" :disabled="disabled" aria-label="移除报表文件" @click="remove">
        <svg viewBox="0 0 20 20" width="18" height="18" fill="none" aria-hidden="true">
          <path d="m6 6 8 8M14 6l-8 8" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" />
        </svg>
      </button>
    </div>
    <p v-if="localError" class="report-upload-error" role="alert">{{ localError }}</p>
  </div>
</template>

<style scoped>
.report-uploader { width: 100%; }
.report-drop {
  display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 10px;
  width: 100%; min-height: 190px; padding: 28px 24px; border: 1.5px dashed var(--border-strong);
  border-radius: var(--radius-s); background: var(--surface); color: var(--primary); cursor: pointer;
  text-align: center; transition: border-color .2s var(--ease-out), background .2s var(--ease-out);
}
.report-drop:hover, .report-drop.dragging { border-color: var(--primary); background: var(--primary-soft); }
.report-drop:focus-visible { outline: 2px solid var(--primary); outline-offset: 3px; }
.report-drop.disabled { opacity: .55; pointer-events: none; cursor: not-allowed; }
.report-drop-title { color: var(--text); font-size: 15px; font-weight: 700; }
.report-drop-hint { color: var(--text-faint); font-size: 13px; }
.report-file {
  display: flex; align-items: center; gap: 10px; min-height: 52px; padding: 10px 12px;
  border: 1.5px solid var(--primary); border-radius: var(--radius-s); background: var(--primary-soft); color: var(--primary-ink);
}
.report-file.disabled { opacity: .6; }
.report-file-name { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 13px; font-weight: 700; }
.report-remove {
  display: grid; place-items: center; width: 44px; height: 44px; border: 0; border-radius: 8px;
  background: transparent; color: currentColor; cursor: pointer;
}
.report-remove:hover:not(:disabled) { color: var(--danger); background: var(--danger-soft); }
.report-remove:focus-visible { outline: 2px solid var(--primary); outline-offset: 2px; }
.report-remove:disabled { opacity: .4; cursor: not-allowed; }
.report-upload-error { margin: 10px 0 0; color: var(--danger); font-size: 13px; }
@media (max-width: 520px) { .report-drop { min-height: 170px; padding: 24px 16px; } }
</style>
