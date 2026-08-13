<script setup>
import { ref, computed, watch, onUnmounted } from 'vue'
import ModeSelect from './components/ModeSelect.vue'
import InvoiceTypeSelect from './components/InvoiceTypeSelect.vue'
import UploadArea from './components/UploadArea.vue'
import TemplateUpload from './components/TemplateUpload.vue'
import ProgressPanel from './components/ProgressPanel.vue'
import ResultPanel from './components/ResultPanel.vue'
import { createTask, getTask, getWorksheets } from './api'

const mode = ref('')
const invType = ref('1')
const files = ref([])
const layoutDir = ref('v') // 收款组排版方向：v 纵向 | h 横向
const startCell = ref('A1') // 收款组起始格
const templateFile = ref(null) // 收款组可选表格模板
const sheets = ref([]) // 模板的工作表列表
const selectedSheet = ref('') // 选中的工作表
const sheetError = ref('')

const status = ref('idle') // idle | processing | done | error
const taskId = ref('')
const current = ref(0)
const total = ref(0)
const message = ref('')
const filename = ref('')
const error = ref('')

let pollTimer = null

const steps = computed(() => {
  const list = [
    { key: 'mode', no: 1, label: '选择功能' },
    { key: 'type', no: 2, label: '发票类型', visible: mode.value === '2' },
    { key: 'upload', no: 3, label: '上传 PDF' },
    { key: 'run', no: 4, label: '制作' }
  ]
  // 收款组少一步发票类型
  if (mode.value !== '2') list[2].no = 2
  if (mode.value !== '2') list[3].no = 3
  let stepNo = 0
  for (const s of list) {
    if (s.visible) s.cur = ++stepNo
  }
  return list
})

const currentStep = computed(() => {
  if (!mode.value) return 1
  if (mode.value === '2') return 2
  return 3
})

const submitting = computed(() => status.value === 'processing')
const startCellValid = computed(() => /^[A-Za-z]{1,3}\d{1,7}$/.test(startCell.value))
const canSubmit = computed(
  () =>
    files.value.length > 0 &&
    !submitting.value &&
    (mode.value !== '1' || startCellValid.value) &&
    (mode.value !== '1' || !templateFile.value || selectedSheet.value)
)

// 切换功能模式或发票类型时清空已上传文件，避免旧文件混入生成导致识别不到
watch(mode, (val, old) => {
  if (val !== old && !submitting.value) {
    files.value = []
    clearTemplate()
  }
})
watch(invType, (val, old) => {
  if (val !== old && !submitting.value) files.value = []
})

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

function clearTemplate() {
  templateFile.value = null
  sheets.value = []
  selectedSheet.value = ''
  sheetError.value = ''
}

async function onTemplateSelected(file) {
  templateFile.value = file
  sheets.value = []
  selectedSheet.value = ''
  sheetError.value = ''
  try {
    const list = await getWorksheets(file)
    sheets.value = list
    selectedSheet.value = list[0] || ''
  } catch (e) {
    sheetError.value = e.message
  }
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
      invType: invType.value,
      layout: mode.value === '1' ? layoutDir.value : 'v',
      startCell: mode.value === '1' ? startCell.value : 'A1',
      template: mode.value === '1' ? templateFile.value : null,
      sheetName: mode.value === '1' ? selectedSheet.value : ''
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
  clearTemplate()
}

onUnmounted(stopPolling)
</script>

<template>
  <header class="masthead">
    <div class="brand">
      <span class="brand-mark">
        <svg viewBox="0 0 32 32" width="30" height="30" fill="none" aria-hidden="true">
          <rect x="4" y="2.5" width="24" height="27" rx="5" fill="var(--primary-soft)" />
          <path d="M9 9h14M9 14h14M9 19h9" stroke="var(--primary-strong)" stroke-width="2.4" stroke-linecap="round" />
          <path d="M20 22.5l4 4 6-6" stroke="var(--primary)" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round" />
        </svg>
      </span>
      <span class="brand-text">
        <span class="brand-name">财务内部在线工具</span>
        <span class="brand-sub">发票识别 · 明细转表 · 在线导出</span>
      </span>
    </div>
    <span class="masthead-tag" v-if="status === 'idle'">网页版</span>
  </header>

  <main class="workspace">
    <div class="rail" aria-hidden="true"></div>

    <!-- 步骤 1：选择功能 -->
    <section class="step">
      <span class="step-dot" :class="{ done: mode, cur: currentStep === 1 }">
        <svg v-if="mode" viewBox="0 0 16 16" width="14" height="14" fill="none">
          <path d="M3 8.5l3.2 3L13 4.5" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" />
        </svg>
        <template v-else>1</template>
      </span>
      <div class="step-body">
        <h2 class="step-title">选择功能</h2>
        <p class="step-sub">根据要处理的内容选择模式</p>
        <ModeSelect v-model="mode" :disabled="submitting" />
      </div>
    </section>

    <!-- 步骤 2：选择发票类型（付款组） -->
    <section v-if="mode === '2'" class="step">
      <span class="step-dot" :class="{ done: false, cur: currentStep === 2 }">
        <svg v-if="currentStep > 2" viewBox="0 0 16 16" width="14" height="14" fill="none">
          <path d="M3 8.5l3.2 3L13 4.5" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" />
        </svg>
        <template v-else>2</template>
      </span>
      <div class="step-body">
        <h2 class="step-title">选择发票类型</h2>
        <p class="step-sub">四种版式，选错会识别不到明细</p>
        <InvoiceTypeSelect v-model="invType" :disabled="submitting" />
      </div>
    </section>

    <!-- 步骤 3/2：上传 PDF -->
    <section class="step">
      <span class="step-dot" :class="{ done: files.length > 0, cur: currentStep === (mode === '2' ? 3 : 2) }">
        <svg v-if="files.length > 0" viewBox="0 0 16 16" width="14" height="14" fill="none">
          <path d="M3 8.5l3.2 3L13 4.5" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" />
        </svg>
        <template v-else>{{ mode === '2' ? 3 : 2 }}</template>
      </span>
      <div class="step-body">
        <h2 class="step-title">上传 PDF 文件</h2>
        <p class="step-sub">支持多选，一次拖入全部发票</p>

        <UploadArea :disabled="submitting" :count="files.length" @add="addFiles" @remove="removeFile" @clear="clearFiles">
          <div v-for="(f, i) in files" :key="f.name + i" class="file-row">
            <svg viewBox="0 0 20 20" width="17" height="17" fill="none" class="file-glyph" aria-hidden="true">
              <path d="M6 2h5l4 4v12H6V2z" stroke="var(--primary)" stroke-width="1.6" stroke-linejoin="round" />
              <path d="M11 2v4h4" stroke="var(--primary)" stroke-width="1.6" stroke-linejoin="round" />
              <path d="M8.5 11h4M8.5 14h4" stroke="var(--primary)" stroke-width="1.6" stroke-linecap="round" />
            </svg>
            <span class="file-name" :title="f.name">{{ f.name }}</span>
            <span class="file-size">{{ formatSize(f.size) }}</span>
            <button
              class="file-remove"
              type="button"
              :disabled="submitting"
              :aria-label="'移除 ' + f.name"
              @click="removeFile(i)"
            >✕</button>
          </div>
        </UploadArea>

        <div v-if="mode === '1'" class="template-section">
          <div class="ts-head">
            <span class="ts-label">插入已有表格（可选）</span>
            <span class="ts-tip">不传模板则自动生成新表格</span>
          </div>
          <TemplateUpload
            :disabled="submitting"
            @selected="onTemplateSelected"
            @cleared="clearTemplate"
          />
          <div v-if="templateFile && !sheetError" class="sheet-pick">
            <span class="lo-label">选择工作表</span>
            <div class="sheet-btns" role="radiogroup" aria-label="选择工作表">
              <button
                v-for="s in sheets"
                :key="s"
                type="button"
                class="sheet-btn"
                :class="{ on: selectedSheet === s }"
                :disabled="submitting"
                @click="selectedSheet = s"
              >{{ s }}</button>
            </div>
          </div>
          <p v-if="sheetError" class="lo-error">{{ sheetError }}</p>
        </div>

        <div v-if="mode === '1'" class="layout-options">
          <div class="ts-head">
            <span class="ts-label">排版位置</span>
            <span class="ts-tip">图片在表格里的排布方式</span>
          </div>
          <div class="lo-row">
            <div class="lo-field">
              <span class="lo-label">排版方向</span>
              <div class="seg" role="radiogroup" aria-label="排版方向">
                <button
                  type="button"
                  class="seg-btn"
                  :class="{ on: layoutDir === 'v' }"
                  @click="layoutDir = 'v'"
                >纵向</button>
                <button
                  type="button"
                  class="seg-btn"
                  :class="{ on: layoutDir === 'h' }"
                  @click="layoutDir = 'h'"
                >横向</button>
              </div>
            </div>
            <div class="lo-field">
              <label class="lo-label" for="start-cell">起始位置</label>
              <input
                id="start-cell"
                v-model="startCell"
                class="cell-input"
                :class="{ invalid: !startCellValid }"
                spellcheck="false"
              />
              <span v-if="!startCellValid" class="lo-error">请输入如 A1、C5 的格位置</span>
            </div>
          </div>
          <p class="lo-hint">
            <template v-if="layoutDir === 'v'">纵向：图片沿列向下排，字段标在图片右侧</template>
            <template v-else>横向：图片沿行向右排，字段标在图片下方</template>
          </p>
        </div>
      </div>
    </section>

    <!-- 步骤 4/3：制作 -->
    <section class="step">
      <span class="step-dot" :class="{ cur: currentStep >= 3 }">
        <svg v-if="status === 'done'" viewBox="0 0 16 16" width="14" height="14" fill="none">
          <path d="M3 8.5l3.2 3L13 4.5" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" />
        </svg>
        <template v-else>{{ mode === '2' ? 4 : 3 }}</template>
      </span>
      <div class="step-body">
        <h2 class="step-title">制作</h2>
        <p class="step-sub">后端处理完成后，表格会直接从浏览器下载</p>

        <div class="run-area">
          <button
            class="btn-make"
            type="button"
            :disabled="!canSubmit"
            @click="submit"
          >
            <span v-if="submitting" class="spinner" aria-hidden="true"></span>
            <span>{{ submitting ? '正在制作…' : '开始制作' }}</span>
          </button>
          <p class="run-hint">
            <template v-if="files.length">
              已选 <b>{{ files.length }}</b> 个文件
            </template>
            <template v-else>请先上传 PDF</template>
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
      </div>
    </section>
  </main>

  <footer class="colophon">
    内部工具 · 文件仅在本次任务内处理，完成后清理
  </footer>
</template>

<style scoped>
.masthead {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 44px;
}

.brand {
  display: flex;
  align-items: center;
  gap: 14px;
}

.brand-mark {
  display: grid;
  place-items: center;
}

.brand-text {
  display: flex;
  flex-direction: column;
  gap: 2px;
}

.brand-name {
  font-family: var(--font-serif);
  font-size: 26px;
  font-weight: 700;
  letter-spacing: 0.5px;
  color: var(--text);
}

.brand-sub {
  font-size: 12.5px;
  color: var(--text-soft);
  letter-spacing: 0.3px;
}

.masthead-tag {
  font-size: 12px;
  color: var(--primary-ink);
  background: var(--primary-soft);
  border: 1px solid rgba(13, 138, 122, 0.18);
  padding: 4px 12px;
  border-radius: 999px;
  font-weight: 600;
  letter-spacing: 0.5px;
}

.workspace {
  position: relative;
  padding-left: 52px;
}

.rail {
  position: absolute;
  left: 15px;
  top: 10px;
  bottom: 14px;
  width: 2px;
  background: var(--border);
  border-radius: 2px;
}

.step {
  position: relative;
  padding-bottom: 46px;
}

.step:last-child {
  padding-bottom: 0;
}

.step-dot {
  position: absolute;
  left: -52px;
  top: 2px;
  width: 32px;
  height: 32px;
  border-radius: 50%;
  background: var(--bg);
  border: 2px solid var(--border-strong);
  color: var(--text-faint);
  display: grid;
  place-items: center;
  font-size: 13px;
  font-weight: 700;
  font-family: var(--font-num);
  transition: all 0.3s var(--ease-out);
}

.step-dot.cur {
  border-color: var(--primary);
  color: var(--primary);
  background: var(--surface);
  box-shadow: 0 0 0 4px var(--primary-soft);
}

.step-dot.done {
  border-color: var(--primary);
  background: var(--primary);
  color: #fff;
}

.step-body {
  animation: rise 0.5s var(--ease-out) both;
}

.step:nth-child(1) .step-body {
  animation-delay: 0.04s;
}
.step:nth-child(2) .step-body {
  animation-delay: 0.1s;
}
.step:nth-child(3) .step-body {
  animation-delay: 0.16s;
}
.step:nth-child(4) .step-body {
  animation-delay: 0.22s;
}

@keyframes rise {
  from {
    opacity: 0;
    transform: translateY(10px);
  }
  to {
    opacity: 1;
    transform: none;
  }
}

.step-title {
  margin: 0;
  font-size: 19px;
  font-weight: 700;
  letter-spacing: 0.2px;
}

.step-sub {
  margin: 4px 0 0;
  font-size: 12.5px;
  color: var(--text-soft);
}

.step-body > :deep(.mode-select),
.step-body > :deep(.inv-type),
.step-body > :deep(.dropzone-wrap),
.step-body > :deep(.progress-panel),
.step-body > :deep(.result-panel) {
  margin-top: 20px;
}

.run-area {
  margin-top: 24px;
  display: flex;
  align-items: center;
  gap: 18px;
  flex-wrap: wrap;
}

.layout-options {
  margin-top: 26px;
  padding-top: 22px;
  border-top: 1px dashed var(--border-strong);
  display: flex;
  flex-direction: column;
  gap: 14px;
}

.lo-row {
  display: flex;
  align-items: center;
  gap: 26px;
  flex-wrap: wrap;
}

.lo-field {
  display: flex;
  align-items: center;
  gap: 12px;
}

.lo-label {
  font-size: 13px;
  font-weight: 700;
  color: var(--text-soft);
}

.seg {
  display: inline-flex;
  padding: 3px;
  gap: 2px;
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 9px;
}

.seg-btn {
  border: none;
  background: transparent;
  padding: 6px 18px;
  border-radius: 7px;
  font-size: 13px;
  font-weight: 600;
  color: var(--text-soft);
  cursor: pointer;
  transition: background 0.16s var(--ease-out), color 0.16s var(--ease-out);
}

.seg-btn:hover {
  color: var(--text);
}

.seg-btn.on {
  background: var(--primary);
  color: #fff;
}

.cell-input {
  width: 86px;
  padding: 7px 12px;
  border: 1.5px solid var(--border-strong);
  border-radius: 9px;
  font-size: 13.5px;
  font-weight: 700;
  background: var(--surface);
  text-transform: uppercase;
  outline: none;
  transition: border-color 0.16s var(--ease-out), box-shadow 0.16s var(--ease-out);
}

.cell-input:focus {
  border-color: var(--primary);
  box-shadow: 0 0 0 3px var(--primary-soft);
}

.cell-input.invalid {
  border-color: var(--danger);
  box-shadow: 0 0 0 3px var(--danger-soft);
}

.lo-error {
  font-size: 12px;
  color: var(--danger);
}

.lo-hint {
  margin: 0;
  font-size: 12px;
  color: var(--text-faint);
}

.template-section {
  margin-top: 26px;
  padding-top: 22px;
  border-top: 1px dashed var(--border-strong);
}

.ts-head {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 12px;
}

.ts-label {
  font-size: 13px;
  font-weight: 800;
  letter-spacing: 0.3px;
}

.ts-tip {
  font-size: 12px;
  color: var(--text-faint);
}

.sheet-pick {
  margin-top: 16px;
  display: flex;
  flex-direction: column;
  gap: 9px;
  align-items: flex-start;
}

.sheet-btns {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
}

.sheet-btn {
  border: 1.5px solid var(--border-strong);
  background: var(--surface);
  border-radius: 9px;
  padding: 7px 16px;
  font-size: 13px;
  font-weight: 600;
  color: var(--text-soft);
  cursor: pointer;
  max-width: 100%;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  transition: all 0.16s var(--ease-out);
}

.sheet-btn:hover:not(:disabled) {
  border-color: var(--primary);
  color: var(--primary-ink);
}

.sheet-btn.on {
  background: var(--primary);
  border-color: var(--primary);
  color: #fff;
}

.sheet-btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.btn-make {
  border: none;
  border-radius: 12px;
  padding: 13px 38px;
  font-size: 15px;
  font-weight: 700;
  color: #fff;
  background: var(--primary);
  box-shadow: 0 8px 20px -8px rgba(13, 138, 122, 0.55);
  display: inline-flex;
  align-items: center;
  gap: 10px;
  transition: transform 0.18s var(--ease-out), box-shadow 0.18s var(--ease-out),
    background 0.18s var(--ease-out), opacity 0.18s var(--ease-out);
}

.btn-make:hover:not(:disabled) {
  transform: translateY(-1px);
  background: var(--primary-strong);
  box-shadow: 0 12px 24px -8px rgba(13, 138, 122, 0.6);
}

.btn-make:active:not(:disabled) {
  transform: scale(0.98);
}

.btn-make:disabled {
  opacity: 0.45;
  cursor: not-allowed;
  box-shadow: none;
}

.run-hint {
  margin: 0;
  font-size: 13px;
  color: var(--text-faint);
}

.run-hint b {
  color: var(--primary-ink);
  font-variant-numeric: tabular-nums;
}

.spinner {
  width: 15px;
  height: 15px;
  border: 2.5px solid rgba(255, 255, 255, 0.35);
  border-top-color: #fff;
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
}

@keyframes spin {
  to {
    transform: rotate(360deg);
  }
}

.file-row {
  display: flex;
  align-items: center;
  gap: 11px;
  padding: 9px 13px;
  border: 1px solid var(--border);
  border-radius: var(--radius-s);
  background: var(--surface-2);
  font-size: 13px;
  transition: border-color 0.15s var(--ease-out);
}

.file-row:hover {
  border-color: var(--border-strong);
}

.file-glyph {
  flex-shrink: 0;
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
  font-family: var(--font-num);
  font-variant-numeric: tabular-nums;
}

.file-remove {
  border: none;
  background: transparent;
  color: var(--text-faint);
  font-size: 13px;
  padding: 2px 6px;
  border-radius: 6px;
  cursor: pointer;
  transition: background 0.15s var(--ease-out), color 0.15s var(--ease-out);
}

.file-remove:hover:not(:disabled) {
  background: var(--danger-soft);
  color: var(--danger);
}

.file-remove:disabled {
  opacity: 0.4;
  cursor: not-allowed;
}

.colophon {
  margin-top: 52px;
  text-align: center;
  font-size: 12px;
  color: var(--text-faint);
  letter-spacing: 0.3px;
}

@media (max-width: 640px) {
  .masthead {
    margin-bottom: 32px;
  }
  .workspace {
    padding-left: 46px;
  }
  .step-dot {
    left: -46px;
  }
  .step {
    padding-bottom: 36px;
  }
}
</style>
