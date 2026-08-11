<script setup>
import { ref, computed } from 'vue'

defineProps({ disabled: Boolean })
const emit = defineEmits(['add', 'remove', 'clear'])

const fileInput = ref(null)
const dragging = ref(false)

const accept = '.pdf'

function openPicker() {
  fileInput.value?.click()
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

function formatSize(bytes) {
  if (bytes < 1024) return bytes + ' B'
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB'
  return (bytes / (1024 * 1024)).toFixed(2) + ' MB'
}

const hint = computed(() => {
  return '支持一次拖入或选择多个 PDF 文件'
})
</script>

<template>
  <div>
    <div
      class="dropzone"
      :class="{ dragging, disabled }"
      @click="openPicker"
      @dragover.prevent="onDragOver"
      @dragleave="onDragLeave"
      @drop.prevent="onDrop"
    >
      <input
        ref="fileInput"
        type="file"
        :accept="accept"
        multiple
        hidden
        @change="onPick"
      />
      <div class="dz-icon">📎</div>
      <div class="dz-title">点击选择，或将 PDF 文件拖到此处</div>
      <div class="dz-hint">{{ hint }}</div>
    </div>

    <div v-if="$slots.default" class="file-list">
      <slot />
      <button class="clear-btn" type="button" @click.stop="emit('clear')">清空全部</button>
    </div>
  </div>
</template>

<style scoped>
.dropzone {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 6px;
  padding: 40px 24px;
  border: 2px dashed var(--border);
  border-radius: var(--radius);
  background: var(--card);
  cursor: pointer;
  transition: all 0.18s ease;
}

.dropzone:hover,
.dropzone.dragging {
  border-color: var(--primary);
  background: var(--primary-softer);
}

.dropzone.dragging {
  transform: scale(1.01);
}

.dropzone.disabled {
  opacity: 0.55;
  pointer-events: none;
}

.dz-icon {
  font-size: 30px;
  line-height: 1;
  margin-bottom: 4px;
}

.dz-title {
  font-weight: 700;
  font-size: 15px;
}

.dz-hint {
  font-size: 12.5px;
  color: var(--text-soft);
}

.file-list {
  margin-top: 12px;
  display: flex;
  flex-direction: column;
  gap: 8px;
}
</style>
