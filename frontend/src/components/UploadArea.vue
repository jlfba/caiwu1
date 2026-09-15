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
  if (props.allowDirectories) {
    const entries = items.map((item) => item.webkitGetAsEntry?.()).filter(Boolean)
    if (entries.some((entry) => entry.isDirectory)) {
      const files = (await Promise.all(entries.map(readEntry))).flat()
      await handleFiles(files)
      return
    }
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
    <!-- 设计图拖拽区 -->
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

      <!-- 绿色圆形上传云朵/向上箭头图标 -->
      <div class="upload-icon-circle">
        <svg viewBox="0 0 24 24" width="26" height="26" fill="none" aria-hidden="true">
          <path d="M12 15V7M8.5 10.5L12 7l3.5 3.5" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" />
          <path d="M7 17h10" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-opacity="0.8" />
        </svg>
      </div>

      <p class="dz-title">拖入 {{ fileLabel }} 文件，或点击选择文件</p>
      <p class="dz-subtitle">支持多选，单个文件不超过 20MB</p>

      <button class="dz-select-btn" type="button" :disabled="disabled" @click.stop="openPicker">
        <svg viewBox="0 0 20 20" width="16" height="16" fill="none" aria-hidden="true">
          <path d="M10 4v8m-3-3l3-3 3 3" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
          <path d="M4 14h12" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
        </svg>
        <span>选择文件</span>
      </button>

      <button v-if="allowDirectories" class="folder-picker" type="button" :disabled="disabled" @click.stop="openFolderPicker">
        <svg viewBox="0 0 16 16" width="14" height="14" fill="none">
          <path d="M2 4.5A1.5 1.5 0 013.5 3h2.88a1.5 1.5 0 011.06.44L8.5 4.5H13A1.5 1.5 0 0114.5 6v6.5a1.5 1.5 0 01-1.5 1.5h-9.5A1.5 1.5 0 012 12.5v-8z" stroke="currentColor" stroke-width="1.3"/>
        </svg>
        选择文件夹
      </button>

      <!-- 或 — 格式支持标签 -->
      <div class="divider-line">
        <span>— 或 —</span>
      </div>

      <div class="file-format-badge">
        <div class="badge-icon">
          <span class="badge-ext">PDF</span>
          <span class="badge-dot">.pdf</span>
        </div>
        <div class="badge-info">
          <span class="badge-type">仅支持 {{ fileLabel }} 格式</span>
          <span class="badge-sub">可同时上传多个文件</span>
        </div>
      </div>
    </div>

    <!-- 已选文件列表 -->
    <div v-if="$slots.default && count" class="file-list-shell">
      <div class="file-list-head">
        <span class="file-count">已选 <b>{{ count }}</b> 个文件</span>
        <div class="head-actions">
          <button
            class="file-list-toggle"
            type="button"
            :aria-expanded="expanded"
            @click="expanded = !expanded"
          >
            {{ expanded ? '收起列表' : '展开查看' }}
            <svg class="toggle-icon" :class="{ expanded }" viewBox="0 0 16 16" width="14" height="14" fill="none">
              <path d="m4 6 4 4 4-4" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" />
            </svg>
          </button>
          <button class="clear-all" type="button" @click="emit('clear')">清空全部</button>
        </div>
      </div>
      <div v-show="expanded" class="file-list">
        <slot />
      </div>
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
  position: relative;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  padding: 32px 20px 24px;
  border: 1.5px dashed var(--border-strong);
  border-radius: var(--radius-m);
  background: var(--surface);
  cursor: pointer;
  transition: all 0.2s var(--ease-smooth);
}

.dropzone:hover:not(.disabled) {
  border-color: var(--primary);
  background: var(--surface-hover);
}

.dropzone.dragging {
  border-color: var(--primary);
  border-style: solid;
  background: var(--primary-soft);
  transform: scale(1.004);
}

.dropzone.disabled {
  opacity: 0.55;
  pointer-events: none;
}

/* 绿色上传圆形图标 */
.upload-icon-circle {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 48px;
  height: 48px;
  border-radius: 50%;
  background: var(--primary);
  box-shadow: 0 4px 14px var(--primary-shadow);
  margin-bottom: 12px;
  transition: transform 0.2s ease;
}

.dropzone:hover .upload-icon-circle {
  transform: translateY(-2px);
}

.dz-title {
  margin: 0;
  font-size: 15px;
  font-weight: 700;
  color: var(--text);
  text-align: center;
}

.dz-subtitle {
  margin: 4px 0 16px;
  font-size: 12px;
  color: var(--text-faint);
  text-align: center;
}

/* 选择文件主按钮 */
.dz-select-btn {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 8px 24px;
  border: 1.5px solid var(--primary);
  background: #fff;
  color: var(--primary);
  border-radius: var(--radius-pill);
  font-size: 13.5px;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.2s ease;
}

.dz-select-btn:hover:not(:disabled) {
  background: var(--primary);
  color: #fff;
  box-shadow: 0 4px 12px var(--primary-shadow);
}

.folder-picker {
  margin-top: 8px;
  display: inline-flex;
  align-items: center;
  gap: 5px;
  padding: 5px 12px;
  border: 1px dashed var(--border-strong);
  border-radius: var(--radius-xs);
  background: var(--surface-2);
  color: var(--text-soft);
  font-size: 11.5px;
  cursor: pointer;
  transition: all 0.15s ease;
}

.folder-picker:hover {
  color: var(--primary);
  border-color: var(--primary);
  background: var(--primary-soft);
}

/* 分割线 */
.divider-line {
  display: flex;
  align-items: center;
  justify-content: center;
  margin: 18px 0 14px;
  color: var(--text-faint);
  font-size: 11px;
}

/* 格式支持卡片 */
.file-format-badge {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 8px 16px;
  background: var(--surface-2);
  border: 1px solid var(--border);
  border-radius: var(--radius-s);
}

.badge-icon {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  width: 32px;
  height: 32px;
  background: #fdf2f2;
  border: 1px solid #fed7d7;
  border-radius: 6px;
  line-height: 1;
}

.badge-ext {
  font-size: 9px;
  font-weight: 800;
  color: #e53e3e;
}

.badge-dot {
  font-size: 8px;
  color: #e53e3e;
  opacity: 0.8;
}

.badge-info {
  display: flex;
  flex-direction: column;
  gap: 1px;
  text-align: left;
}

.badge-type {
  font-size: 12px;
  font-weight: 600;
  color: var(--text-secondary);
}

.badge-sub {
  font-size: 11px;
  color: var(--text-faint);
}

/* 已选文件列表卡片 */
.file-list-shell {
  border: 1px solid var(--border);
  border-radius: var(--radius-m);
  background: var(--surface);
  overflow: hidden;
}

.file-list-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 10px 14px;
  background: var(--surface-2);
  border-bottom: 1px solid var(--border);
}

.file-count {
  font-size: 12.5px;
  color: var(--text-secondary);
}

.file-count b {
  color: var(--primary);
}

.head-actions {
  display: flex;
  align-items: center;
  gap: 12px;
}

.file-list-toggle,
.clear-all {
  border: none;
  background: transparent;
  padding: 4px 8px;
  font-size: 12px;
  font-weight: 600;
  cursor: pointer;
  border-radius: 4px;
}

.file-list-toggle {
  display: inline-flex;
  align-items: center;
  gap: 3px;
  color: var(--primary);
}

.file-list-toggle:hover {
  background: var(--primary-soft);
}

.toggle-icon {
  transition: transform 0.2s ease;
}

.toggle-icon.expanded {
  transform: rotate(180deg);
}

.clear-all {
  color: var(--text-faint);
}

.clear-all:hover {
  color: var(--danger);
  background: var(--danger-soft);
}

.file-list {
  display: flex;
  flex-direction: column;
  gap: 6px;
  max-height: 240px;
  padding: 10px;
  overflow-y: auto;
}
</style>
