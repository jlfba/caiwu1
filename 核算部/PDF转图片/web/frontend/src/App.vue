<script setup>
import { ref, computed, onUnmounted } from 'vue'
import ModeSelect from './components/ModeSelect.vue'
import InvoiceTypeSelect from './components/InvoiceTypeSelect.vue'
import UploadArea from './components/UploadArea.vue'
import ProgressPanel from './components/ProgressPanel.vue'
import ResultPanel from './components/ResultPanel.vue'
import { createTask, getTask } from './api'

const mode = ref('')
const invType = ref('1')
const files = ref([])

const status = ref('idle') // idle | processing | done | error
const taskId = ref('')
const current = ref(0)
const total = ref(0)
const message = ref('')
const filename = ref('')
const error = ref('')

let pollTimer = null

const step = computed(() => {
  if (!mode.value) return 1
  if (mode.value === '2') return 2
  return 3
})

const submitting = computed(() => status.value === 'processing')
const canSubmit = computed(() => files.value.length > 0 && !submitting.value)

function formatSize(bytes) {
  if (bytes < 1024) return bytes + ' B'
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB'
  return (bytes / (1024 * 1024)).toFixed(2) + ' MB'
}

function addFiles(list) {
  const seen = new Set(files.value.map((f) => `${f.name}:${f.size}`))
  for (const f of list) {
    if (!f.name.toLowerCase().endsWith('.pdf')) continue
    const key = `${f.name}:${f.size}`
    if (!seen.has(key)) {
      seen.add(key)
      files.value.push(f)
    }
  }
}

function removeFile(index) {
  files.value.splice(index, 1)
}

function clearFiles() {
  files.value = []
}

function stopPolling() {
  if (pollTimer) {
    clearInterval(pollTimer)
    pollTimer = null
  }
}

async function submit() {
  if (!canSubmit.value) return
  stopPolling()
  status.value = 'processing'
  current.value = 0
  total.value = 0
  message.value = '正在上传文件…'
  error.value = ''
  filename.value = ''
  try {
    const data = await createTask({
      files: files.value,
      mode: mode.value,
      invType: invType.value
    })
    taskId.value = data.task_id
    pollTimer = setInterval(poll, 1200)
    poll()
  } catch (e) {
    status.value = 'error'
    error.value = e.message
  }
}

async function poll() {
  if (!taskId.value) return
  let data
  try {
    data = await getTask(taskId.value)
  } catch (e) {
    status.value = 'error'
    error.value = e.message
    stopPolling()
    return
  }
  current.value = data.current || 0
  total.value = data.total || 0
  message.value = data.message || ''

  if (data.status === 'done') {
    filename.value = data.filename || ''
    status.value = 'done'
    stopPolling()
  } else if (data.status === 'error') {
    error.value = data.error || data.message || '处理失败'
    status.value = 'error'
    stopPolling()
  }
}

function reset() {
  stopPolling()
  status.value = 'idle'
  taskId.value = ''
  current.value = 0
  total.value = 0
  message.value = ''
  filename.value = ''
  error.value = ''
  files.value = []
}

onUnmounted(stopPolling)
</script>

<template>
  <header class="page-head">
    <div class="brand">
      <span class="brand-logo">🧾</span>
      <div>
        <h1>PDF 工具网页版</h1>
        <p>发票图片识别 · 发票明细转表格 · 在线处理</p>
      </div>
    </div>
  </header>

  <main>
    <!-- 步骤 1：选择功能 -->
    <section class="card">
      <div class="step-head">
        <span class="step-no" :class="{ on: step >= 1 }">1</span>
        <span class="step-label">选择功能</span>
      </div>
      <ModeSelect v-model="mode" :disabled="submitting" />
    </section>

    <!-- 步骤 2：选择发票类型（付款组） -->
    <section v-if="mode === '2'" class="card">
      <div class="step-head">
        <span class="step-no" :class="{ on: step >= 2 }">2</span>
        <span class="step-label">选择发票类型</span>
      </div>
      <InvoiceTypeSelect v-model="invType" :disabled="submitting" />
    </section>

    <!-- 步骤 3：上传 PDF -->
    <section class="card">
      <div class="step-head">
        <span class="step-no" :class="{ on: step >= 3 }">
          {{ mode === '2' ? '3' : '2' }}
        </span>
        <span class="step-label">上传 PDF 文件</span>
      </div>

      <UploadArea
        :disabled="submitting"
        @add="addFiles"
        @remove="removeFile"
        @clear="clearFiles"
      >
        <div v-for="(f, i) in files" :key="f.name + i" class="file-item">
          <span class="file-icon">📄</span>
          <span class="file-name" :title="f.name">{{ f.name }}</span>
          <span class="file-size">{{ formatSize(f.size) }}</span>
          <button
            class="file-remove"
            type="button"
            :disabled="submitting"
            @click="removeFile(i)"
          >✕</button>
        </div>
        <div v-if="!files.length" class="file-empty">尚未添加文件</div>
      </UploadArea>
    </section>

    <!-- 步骤 4：制作 -->
    <section class="card">
      <div class="submit-row">
        <button
          class="btn make"
          type="button"
          :disabled="!canSubmit"
          @click="submit"
        >
          <span v-if="submitting" class="spinner"></span>
          {{ submitting ? '正在制作…' : '开始制作' }}
        </button>
        <p class="submit-hint">
          已选择 {{ files.length }} 个 PDF，生成结果将直接从浏览器下载。
        </p>
      </div>

      <ProgressPanel
        v-if="submitting"
        :status="'processing'"
        :current="current"
        :total="total"
        :message="message"
      />

      <ResultPanel
        v-if="status === 'done' || status === 'error'"
        :status="status"
        :task-id="taskId"
        :filename="filename"
        :error="error"
        @reset="reset"
      />
    </section>
  </main>

  <footer class="page-foot">内部工具 · 处理数据仅保留在本次任务内，不落盘保存</footer>
</template>

<style scoped>
.page-head {
  margin-bottom: 26px;
}

.brand {
  display: flex;
  align-items: center;
  gap: 16px;
}

.brand-logo {
  font-size: 38px;
  line-height: 1;
}

.brand h1 {
  margin: 0;
  font-size: 26px;
  font-weight: 800;
  letter-spacing: 0.5px;
}

.brand p {
  margin: 4px 0 0;
  font-size: 13px;
  color: var(--text-soft);
}

main {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.card {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: var(--radius);
  padding: 22px 24px;
  box-shadow: var(--shadow-sm);
}

.step-head {
  display: flex;
  align-items: center;
  gap: 10px;
  margin-bottom: 16px;
}

.step-no {
  width: 26px;
  height: 26px;
  border-radius: 8px;
  background: #f1f5f9;
  color: var(--text-faint);
  display: grid;
  place-items: center;
  font-size: 13px;
  font-weight: 800;
  transition: all 0.2s ease;
}

.step-no.on {
  background: var(--primary);
  color: #fff;
}

.step-label {
  font-size: 16px;
  font-weight: 800;
}

.file-item {
  display: flex;
  align-items: center;
  gap: 10px;
  background: #f8fafc;
  border: 1px solid var(--border);
  border-radius: 10px;
  padding: 9px 12px;
  font-size: 13px;
}

.file-icon {
  font-size: 16px;
}

.file-name {
  flex: 1;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-weight: 600;
}

.file-size {
  color: var(--text-faint);
  font-variant-numeric: tabular-nums;
}

.file-remove {
  border: none;
  background: transparent;
  color: var(--text-faint);
  font-size: 14px;
  padding: 2px 6px;
  border-radius: 6px;
  transition: all 0.15s ease;
}

.file-remove:hover:not(:disabled) {
  background: #fee2e2;
  color: var(--danger);
}

.file-remove:disabled {
  opacity: 0.4;
  cursor: not-allowed;
}

.file-empty {
  text-align: center;
  color: var(--text-faint);
  font-size: 12.5px;
  padding: 4px 0;
}

.submit-row {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 12px;
}

.btn.make {
  border: none;
  border-radius: 12px;
  padding: 14px 56px;
  font-size: 16px;
  font-weight: 800;
  color: #fff;
  background: linear-gradient(120deg, var(--primary), var(--primary-strong));
  box-shadow: 0 10px 24px -10px rgba(15, 118, 110, 0.6);
  transition: all 0.18s ease;
  display: inline-flex;
  align-items: center;
  gap: 10px;
}

.btn.make:hover:not(:disabled) {
  transform: translateY(-1px);
  box-shadow: 0 14px 28px -10px rgba(15, 118, 110, 0.7);
}

.btn.make:disabled {
  opacity: 0.45;
  cursor: not-allowed;
  box-shadow: none;
}

.spinner {
  width: 16px;
  height: 16px;
  border: 2.5px solid rgba(255, 255, 255, 0.4);
  border-top-color: #fff;
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
}

@keyframes spin {
  to {
    transform: rotate(360deg);
  }
}

.submit-hint {
  margin: 0;
  font-size: 12.5px;
  color: var(--text-soft);
}

.page-foot {
  margin-top: 26px;
  text-align: center;
  font-size: 12px;
  color: var(--text-faint);
}
</style>
