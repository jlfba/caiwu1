<script setup>
import { ref, watch } from 'vue'

const props = defineProps({
  disabled: Boolean,
  count: { type: Number, default: 0 },
  allowDirectories: Boolean,
  accept: { type: String, default: '.pdf' },
  fileLabel: { type: String, default: 'PDF' }
})
const emit = defineEmits(['add', 'remove', 'clear'])

const fileInput = ref(null)
const folderInput = ref(null)
const dragging = ref(false)
const expanded = ref(false)

watch(
  () => props.count,
  (count, previousCount) => {
    if (!count || count > previousCount) expanded.value = false
  }
)

function openPicker() {
  if (!fileInput.value) return
  fileInput.value.click()
}

function openFolderPicker() {
  if (folderInput.value) folderInput.value.click()
}

function onPick(event) {
  handleFiles(event.target.files)
  event.target.value = ''
}

async function handleFiles(list) {
  if (!list || !list.length) return
  emit('add', Array.from(list))
}

async function readEntry(entry) {
  if (entry.isFile) return await new Promise((resolve) => entry.file(resolve, () => resolve([])))
  if (!entry.isDirectory) return []
  const reader = entry.createReader()
  const entries = []
  let batch
  do {
    batch = await new Promise((resolve) => reader.readEntries(resolve, () => resolve([])))
    entries.push(...batch)
  } while (batch.length)
  return (await Promise.all(entries.map(readEntry))).flat()
}

async function onDrop(event) {
  dragging.value = false
  const items = Array.from(event.dataTransfer.items || [])
  if (props.allowDirectories && items.some((item) => item.webkitGetAsEntry?.()?.isDirectory)) {
    const files = (await Promise.all(items.map((item) => {
      const entry = item.webkitGetAsEntry?.()
      return entry ? readEntry(entry) : []
    }))).flat()
    await handleFiles(files)
    return
  }
  await handleFiles(event.dataTransfer.files)
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
        :accept="accept"
        multiple
        hidden
        @change="onPick"
      />
      <input
        v-if="allowDirectories"
        ref="folderInput"
        type="file"
        webkitdirectory
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
      <p class="dz-title">{{ allowDirectories ? '拖入 ' + fileLabel + ' 或文件夹' : '拖入 ' + fileLabel + '，或点击选择文件' }}</p>
      <p class="dz-hint">{{ allowDirectories ? '支持文件夹递归识别 · 仅接受 ' + accept : '支持多选 · 仅接受 ' + accept }}</p>
      <button v-if="allowDirectories" class="folder-picker" type="button" :disabled="disabled" @click.stop="openFolderPicker">选择文件夹</button>
    </div>

    <div v-if="$slots.default && count" class="file-list-shell">
      <div class="file-list-head">
        <span class="file-count">已选 {{ count }} 个{{ fileLabel }}</span>
        <button
          class="file-list-toggle"
          type="button"
          :aria-expanded="expanded"
          @click="expanded = !expanded"
        >
          {{ expanded ? '收起列表' : '展开列表' }}
          <svg class="toggle-icon" :class="{ expanded }" viewBox="0 0 16 16" width="16" height="16" fill="none" aria-hidden="true">
            <path d="m4 6 4 4 4-4" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" />
          </svg>
        </button>
      </div>
      <div v-show="expanded" class="file-list">
        <slot />
      </div>
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

.folder-picker {
  margin-top: 8px;
  padding: 6px 10px;
  border: 1px solid var(--border-strong);
  border-radius: 7px;
  background: var(--surface);
  color: var(--primary);
  font-size: 12px;
  font-weight: 700;
  cursor: pointer;
}

.file-list-shell {
  position: relative;
  overflow: hidden;
  border: 1px solid var(--border);
  border-radius: var(--radius);
  background: var(--surface);
}

.file-list-head {
  min-height: 52px;
  padding: 4px 104px 4px 14px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
}

.file-count {
  color: var(--text-muted);
  font-size: 13px;
  font-weight: 700;
}

.file-list {
  display: flex;
  flex-direction: column;
  gap: 8px;
  max-height: clamp(180px, 30vh, 300px);
  padding: 10px;
  overflow-y: auto;
  border-top: 1px solid var(--border);
}

.file-list-toggle,
.clear-all {
  min-height: 44px;
  border: none;
  background: transparent;
  border-radius: 8px;
  cursor: pointer;
}

.file-list-toggle {
  padding: 0 10px;
  display: inline-flex;
  align-items: center;
  gap: 4px;
  color: var(--primary);
  font-size: 12.5px;
  font-weight: 700;
  transition: background 0.15s var(--ease-out);
}

.file-list-toggle:hover {
  background: var(--primary-soft);
}

.toggle-icon {
  transition: transform 0.2s var(--ease-out);
}

.toggle-icon.expanded {
  transform: rotate(180deg);
}

.clear-all {
  position: absolute;
  top: 4px;
  right: 8px;
  color: var(--text-faint);
  font-size: 12.5px;
  padding: 0 10px;
  transition: color 0.15s var(--ease-out), background 0.15s var(--ease-out);
}

.clear-all:hover {
  color: var(--danger);
  background: var(--danger-soft);
}

.file-list-toggle:focus-visible,
.clear-all:focus-visible {
  outline: 2px solid var(--primary);
  outline-offset: -2px;
}

@media (max-width: 560px) {
  .file-list-head {
    min-height: 96px;
    padding: 8px 10px 52px;
  }

  .file-list-toggle {
    position: absolute;
    bottom: 4px;
    left: 8px;
  }

  .clear-all {
    top: auto;
    bottom: 4px;
  }
}
</style>
