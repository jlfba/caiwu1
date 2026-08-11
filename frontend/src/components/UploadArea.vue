<script setup>
import { ref } from 'vue'

defineProps({
  disabled: Boolean,
  count: { type: Number, default: 0 }
})
const emit = defineEmits(['add', 'remove', 'clear'])

const fileInput = ref(null)
const dragging = ref(false)

function openPicker() {
  if (!fileInput.value) return
  fileInput.value.click()
}

function onPick(event) {
  handleFiles(event.target.files)
  event.target.value = ''
}

function handleFiles(list) {
  if (!list || !list.length) return
  emit('add', Array.from(list))
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
</script>

<template>
  <div class="dropzone-wrap">
    <div
      class="dropzone"
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
        accept=".pdf"
        multiple
        hidden
        @change="onPick"
      />
      <span class="dz-icon">
        <svg viewBox="0 0 40 40" width="40" height="40" fill="none" aria-hidden="true">
          <circle cx="20" cy="20" r="19" stroke="var(--primary-soft)" stroke-width="2" />
          <path d="M20 27V13M13.5 19.5L20 13l6.5 6.5" stroke="var(--primary)" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" />
          <path d="M12 29h16" stroke="var(--border-strong)" stroke-width="2.4" stroke-linecap="round" />
        </svg>
      </span>
      <p class="dz-title">拖入 PDF，或点击选择文件</p>
      <p class="dz-hint">支持多选 · 仅接受 .pdf</p>
    </div>

    <div v-if="$slots.default" class="file-list">
      <slot />
      <button v-if="count" class="clear-all" type="button" @click="emit('clear')">
        清空全部
      </button>
    </div>
  </div>
</template>

<style scoped>
.dropzone-wrap {
  display: flex;
  flex-direction: column;
  gap: 14px;
}

.dropzone {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 4px;
  padding: 42px 24px 38px;
  border: 1.5px dashed var(--border-strong);
  border-radius: var(--radius);
  background: var(--surface);
  cursor: pointer;
  transition: border-color 0.2s var(--ease-out), background 0.2s var(--ease-out),
    transform 0.2s var(--ease-out);
}

.dropzone:hover {
  border-color: var(--primary);
  background: var(--surface-2);
}

.dropzone.dragging {
  border-color: var(--primary);
  border-style: solid;
  background: var(--primary-soft);
  transform: scale(1.006);
}

.dropzone:focus-visible {
  outline: 2px solid var(--primary);
  outline-offset: 2px;
}

.dropzone.disabled {
  opacity: 0.55;
  pointer-events: none;
}

.dz-icon {
  margin-bottom: 10px;
  display: grid;
  place-items: center;
  transition: transform 0.25s var(--ease-out);
}

.dropzone:hover .dz-icon,
.dropzone.dragging .dz-icon {
  transform: translateY(-2px);
}

.dz-title {
  margin: 0;
  font-size: 15px;
  font-weight: 700;
}

.dz-hint {
  margin: 2px 0 0;
  font-size: 12.5px;
  color: var(--text-faint);
}

.file-list {
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.clear-all {
  align-self: flex-end;
  border: none;
  background: transparent;
  color: var(--text-faint);
  font-size: 12.5px;
  padding: 4px 8px;
  border-radius: 6px;
  cursor: pointer;
  transition: color 0.15s var(--ease-out), background 0.15s var(--ease-out);
}

.clear-all:hover {
  color: var(--danger);
  background: var(--danger-soft);
}
</style>
