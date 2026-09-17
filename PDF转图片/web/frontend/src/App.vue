<script setup>
import { ref, computed, watch, nextTick, onMounted, onUnmounted } from 'vue'
import ModeSelect from './components/ModeSelect.vue'
import InvoiceTypeSelect from './components/InvoiceTypeSelect.vue'
import ReceiptTypeSelect from './components/ReceiptTypeSelect.vue'
import ReceiptSubtypeSelect from './components/ReceiptSubtypeSelect.vue'
import UploadArea from './components/UploadArea.vue'
import TemplateUpload from './components/TemplateUpload.vue'
import ReportUpload from './components/ReportUpload.vue'
import ProgressPanel from './components/ProgressPanel.vue'
import ResultPanel from './components/ResultPanel.vue'
import { createTask, createReportTask, createFundTask, continueReportTask, getTask, getWorksheets } from './api'

// 当前导航视图 (首页 / 历史记录 / 帮助中心)
const currentNav = ref('home')

// 功能选择与状态
const mode = ref('receipt')
const invType = ref('1')
const receiptPerson = ref('谢莉丽')
const receiptSubtype = ref('')
const files = ref([])

const receiptModes = { '谢莉丽': '1', '赵淑华': '2', '邵梅琳': '5' }
const paymentLabel = computed(() => receiptSubtype.value === 'alipay' ? '支付宝' : receiptSubtype.value === 'huolala' ? '货拉拉' : '微信')

const effectiveMode = computed(() => {
  if (mode.value === 'receipt') {
    if (receiptPerson.value === '赵淑华' && receiptSubtype.value === 'invoice') return receiptModes[receiptPerson.value]
    if (['wechat', 'alipay', 'huolala'].includes(receiptSubtype.value)) return '6'
    if (receiptSubtype.value === 'ordinary_invoice') return '9'
    return receiptModes[receiptPerson.value] || '1'
  }
  return mode.value
})

const layoutDir = ref('v')
const startCell = ref('A1')
const templateFile = ref(null)
const sheets = ref([])
const selectedSheet = ref('')
const sheetError = ref('')

const reportFile = ref(null)
const reportProfile = ref('bundle')
const reportSheets = ref([])
const reportSheet = ref('')
const reportError = ref('')
const reportLoading = ref(false)
const reportUploadPercent = ref(0)
const reportUploadDone = ref(false)

const fundFile1 = ref(null)
const fundFile2 = ref(null)
const fundFile3 = ref(null)
const fundFile4 = ref(null)
const fundFiles = ref([])

const status = ref('idle') // idle | processing | done | error | paused
const taskId = ref('')
const current = ref(0)
const total = ref(0)
const message = ref('')
const filename = ref('')
const error = ref('')
const logs = ref([])
const activityLogs = ref([])
const logTerminal = ref(null)
const reportStep = ref(0)
const reportMaxStep = ref(0)
const elapsedSeconds = ref(0)

let pollTimer = null
let elapsedTimer = null
let elapsedStartedAt = 0
let elapsedAccumulated = 0
let activityLogId = 0
let syncedTaskLogs = new Set()
const ACTIVE_TASK_STORAGE_KEY = 'caiwu-active-task'

const submitting = computed(() => status.value === 'processing')
const startCellValid = computed(() => /^[A-Za-z]{1,3}\d{1,7}$/.test(startCell.value))

const canSubmit = computed(() => {
  if (submitting.value) return false
  if (mode.value === '4') {
    return !!(reportFile.value && reportSheet.value && !reportLoading.value)
  }
  if (mode.value === 'fund') {
    return !!fundFile3.value
  }
  return (
    files.value.length > 0 &&
    (effectiveMode.value !== '1' || startCellValid.value) &&
    (effectiveMode.value !== '1' || !templateFile.value || selectedSheet.value)
  )
})

function formatLogTime(value = new Date()) {
  return value.toLocaleTimeString('zh-CN', { hour12: false })
}

function addActivityLog(content, level = 'info') {
  activityLogs.value.push({ id: ++activityLogId, time: formatLogTime(), content, level })
}

function saveActiveTask() {
  if (!taskId.value) return
  try {
    sessionStorage.setItem(ACTIVE_TASK_STORAGE_KEY, JSON.stringify({
      taskId: taskId.value,
      mode: mode.value,
      status: status.value,
      current: current.value,
      total: total.value,
      message: message.value,
      filename: filename.value,
      error: error.value,
      logs: logs.value,
      activityLogs: activityLogs.value,
      elapsedSeconds: elapsedSeconds.value,
      reportStep: reportStep.value,
      reportMaxStep: reportMaxStep.value,
      currentNav: currentNav.value
    }))
  } catch (_) {
    // 浏览器禁用会话存储时，仍按原来的方式正常处理任务。
  }
}

function clearSavedTask() {
  try {
    sessionStorage.removeItem(ACTIVE_TASK_STORAGE_KEY)
  } catch (_) {
    // 忽略浏览器存储不可用的情况。
  }
}

function resumeElapsedTimer() {
  stopElapsedTimer(false)
  elapsedAccumulated = elapsedSeconds.value
  elapsedStartedAt = Date.now()
  elapsedTimer = setInterval(() => {
    elapsedSeconds.value = elapsedAccumulated + Math.floor((Date.now() - elapsedStartedAt) / 1000)
  }, 1000)
}

async function restoreActiveTask() {
  let saved
  try {
    saved = JSON.parse(sessionStorage.getItem(ACTIVE_TASK_STORAGE_KEY) || 'null')
  } catch (_) {
    clearSavedTask()
    return
  }
  if (!saved || typeof saved.taskId !== 'string' || !saved.taskId) return

  taskId.value = saved.taskId
  mode.value = saved.mode || mode.value
  status.value = saved.status || 'processing'
  current.value = Number(saved.current) || 0
  total.value = Number(saved.total) || 0
  message.value = saved.message || '正在恢复任务状态…'
  filename.value = saved.filename || ''
  error.value = saved.error || ''
  logs.value = Array.isArray(saved.logs) ? saved.logs : []
  activityLogs.value = Array.isArray(saved.activityLogs) ? saved.activityLogs : []
  activityLogId = activityLogs.value.reduce((maxId, item) => Math.max(maxId, Number(item.id) || 0), 0)
  syncedTaskLogs = new Set(logs.value.map((item) => `${taskId.value}:${String(item)}`))
  elapsedSeconds.value = Number(saved.elapsedSeconds) || 0
  reportStep.value = Number(saved.reportStep) || 0
  reportMaxStep.value = Number(saved.reportMaxStep) || 0
  currentNav.value = saved.currentNav || 'home'

  if (status.value === 'processing') {
    resumeElapsedTimer()
    stopPolling()
    pollTimer = setInterval(poll, 1200)
  }
  await poll()
}

watch(
  () => activityLogs.value.length,
  async () => {
    await nextTick()
    if (logTerminal.value) logTerminal.value.scrollTop = logTerminal.value.scrollHeight
  }
)

function exportActivityLogs() {
  const content = activityLogs.value.length
    ? activityLogs.value.map(item => `[${item.time}] ${item.content}`).join('\r\n')
    : '暂无处理日志'
  const url = URL.createObjectURL(new Blob([content], { type: 'text/plain;charset=utf-8' }))
  const link = document.createElement('a')
  link.href = url
  link.download = `财务工具-处理日志-${new Date().toISOString().slice(0, 10)}.txt`
  link.click()
  URL.revokeObjectURL(url)
}

// 监听模式变化
watch(mode, (val, old) => {
  if (val !== old && !submitting.value) {
    if (val === 'receipt') {
      if (!receiptPerson.value) receiptPerson.value = '谢莉丽'
    } else {
      receiptPerson.value = ''
      receiptSubtype.value = ''
    }
    files.value = []
    clearTemplate()
    clearReportFile()
    clearFundFile()
  }
})

watch(receiptPerson, (val, old) => {
  if (val && val !== old && !submitting.value) {
    files.value = []
    clearTemplate()
    receiptSubtype.value = ''
  }
})

watch(receiptSubtype, (val, old) => {
  if (val !== old && !submitting.value) {
    files.value = []
    clearTemplate()
  }
})

function clearReportFile() {
  reportFile.value = null
  reportSheets.value = []
  reportSheet.value = ''
  reportError.value = ''
  reportLoading.value = false
  reportUploadPercent.value = 0
  reportUploadDone.value = false
}

function clearFundFile() {
  fundFile1.value = null
  fundFile2.value = null
  fundFile3.value = null
  fundFile4.value = null
  fundFiles.value = []
}

function selectFundFiles(fileList) {
  const fList = Array.from(fileList || []).filter(f => /\.(xlsx|xlsm)$/i.test(f.name))
  fundFiles.value = fList
  fundFile1.value = fList.find(f => /收款审核/.test(f.name) && !/流水/.test(f.name)) || null
  fundFile2.value = fList.find(f => /服务商付款/.test(f.name) && !/流水/.test(f.name)) || null
  fundFile4.value = fList.find(f => /流水/.test(f.name)) || null
  // 审核表的名称也可能带“中信对公”；只有不属于其他三类的文件才是系统表。
  fundFile3.value = fList.find(f => /对公/.test(f.name)
    && !/收款审核|服务商付款|流水/.test(f.name)) || null
  addActivityLog(`资金组已选择 ${fList.length} 个 Excel 文件${fundFile3.value ? '，已识别中信对公表' : '，未识别中信对公表'}`, fundFile3.value ? 'success' : 'warning')
}

const fundFileSlots = computed(() => [
  ['收款审核表', fundFile1.value],
  ['服务商付款审核表', fundFile2.value],
  ['中信对公', fundFile3.value],
  ['银行账号管理流水', fundFile4.value]
])

async function submitFund() {
  if (!fundFiles.value.length || !fundFile3.value || submitting.value) return
  stopPolling()
  stopElapsedTimer(false)
  elapsedSeconds.value = 0
  elapsedAccumulated = 0
  status.value = 'processing'
  current.value = 0
  total.value = 6
  message.value = '正在上传表格…'
  error.value = ''
  logs.value = []
  syncedTaskLogs = new Set()
  filename.value = ''
  addActivityLog('开始提交资金组核对任务')
  try {
    // 传入全部原始文件，由后端按关键词识别四类表，不能因前端槽位丢掉第 5 个文件。
    const uploadFiles = fundFiles.value
      .filter(file => file instanceof File)
    if (!uploadFiles.includes(fundFile3.value)) {
      throw new Error('中信对公文件无效，请重新选择 Excel 文件')
    }
    const data = await createFundTask(uploadFiles)
    taskId.value = data.task_id
    addActivityLog(`资金组任务已创建：${data.task_id}`, 'success')
    saveActiveTask()
    pollTimer = setInterval(poll, 1200)
    poll()
  } catch (e) {
    status.value = 'error'
    error.value = e.message
    addActivityLog(`资金组任务提交失败：${e.message}`, 'error')
    stopElapsedTimer()
  }
}

async function onReportSelected(file) {
  reportFile.value = file
  reportSheets.value = []
  reportSheet.value = ''
  reportError.value = ''
  reportLoading.value = true
  reportUploadPercent.value = 0
  reportUploadDone.value = false
  try {
    reportSheets.value = await getWorksheets(file, (percent) => {
      reportUploadPercent.value = percent
      if (percent >= 100) reportUploadDone.value = true
    })
    reportSheet.value = reportSheets.value[0] || ''
  } catch (e) {
    reportError.value = e.message
  } finally {
    reportLoading.value = false
  }
}

async function submitReport() {
  if (!reportFile.value || !reportSheet.value || submitting.value) return
  stopPolling()
  stopElapsedTimer(false)
  elapsedSeconds.value = 0
  elapsedAccumulated = 0
  status.value = 'processing'
  current.value = 0
  total.value = 5
  message.value = '正在上传表格…'
  error.value = ''
  logs.value = []
  syncedTaskLogs = new Set()
  filename.value = ''
  try {
    addActivityLog('开始提交报表处理任务')
    const data = await createReportTask(reportFile.value, reportSheet.value, 'bundle')
    taskId.value = data.task_id
    addActivityLog(`报表任务已创建：${data.task_id}`, 'success')
    saveActiveTask()
    pollTimer = setInterval(poll, 1200)
    poll()
  } catch (e) {
    status.value = 'error'
    error.value = e.message
    addActivityLog(`报表任务提交失败：${e.message}`, 'error')
    stopElapsedTimer()
  }
}

async function continueReport() {
  if (!taskId.value || status.value !== 'paused') return
  status.value = 'processing'
  message.value = '正在继续处理…'
  saveActiveTask()
  try {
    addActivityLog('开始提交文件处理任务')
    await continueReportTask(taskId.value)
    pollTimer = setInterval(poll, 1200)
    poll()
  } catch (e) {
    status.value = 'error'
    error.value = e.message
  }
}

function formatSize(bytes) {
  if (bytes < 1024) return bytes + ' B'
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB'
  return (bytes / (1024 * 1024)).toFixed(2) + ' MB'
}

function addFiles(list) {
  const seen = new Set(files.value.map((f) => `${f.name}:${f.size}`))
  for (const f of list) {
    const name = f.name.toLowerCase()
    const allowed = effectiveMode.value === '2'
      ? ['.pdf', '.xlsx', '.xlsm'].some((ext) => name.endsWith(ext))
      : ['6', '7', '8'].includes(effectiveMode.value)
      ? ['.pdf', '.png', '.jpg', '.jpeg'].some((ext) => name.endsWith(ext))
      : name.endsWith('.pdf')
    if (!allowed) continue
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

function stopElapsedTimer(captureFinal = true) {
  if (captureFinal && elapsedStartedAt) {
    elapsedAccumulated += Math.floor((Date.now() - elapsedStartedAt) / 1000)
    elapsedSeconds.value = elapsedAccumulated
  }
  if (elapsedTimer) {
    clearInterval(elapsedTimer)
    elapsedTimer = null
  }
  elapsedStartedAt = 0
}

function startElapsedTimer() {
  stopElapsedTimer()
  elapsedAccumulated = 0
  elapsedStartedAt = Date.now()
  elapsedSeconds.value = 0
  elapsedTimer = setInterval(() => {
    elapsedSeconds.value = elapsedAccumulated + Math.floor((Date.now() - elapsedStartedAt) / 1000)
  }, 1000)
}

async function submit() {
  if (mode.value === '4') {
    submitReport()
    return
  }
  if (mode.value === 'fund') {
    submitFund()
    return
  }
  if (!canSubmit.value) return
  stopPolling()
  startElapsedTimer()
  status.value = 'processing'
  current.value = 0
  total.value = 0
  message.value = '正在上传文件…'
  error.value = ''
  filename.value = ''
  try {
    const data = await createTask({
      files: files.value,
      mode: effectiveMode.value,
      invType: mode.value === 'receipt' ? receiptSubtype.value : invType.value,
      layout: effectiveMode.value === '1' ? layoutDir.value : 'v',
      startCell: effectiveMode.value === '1' ? startCell.value : 'A1',
      template: effectiveMode.value === '1' ? templateFile.value : null,
      sheetName: effectiveMode.value === '1' ? selectedSheet.value : ''
    })
    taskId.value = data.task_id
    addActivityLog(`处理任务已创建：${data.task_id}`, 'success')
    saveActiveTask()
    pollTimer = setInterval(poll, 1200)
    poll()
  } catch (e) {
    status.value = 'error'
    error.value = e.message
    addActivityLog(`处理任务提交失败：${e.message}`, 'error')
    stopElapsedTimer()
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
    clearSavedTask()
    stopPolling()
    stopElapsedTimer()
    return
  }
  current.value = data.current || 0
  total.value = data.total || 0
  message.value = data.message || ''
  logs.value = data.logs || []
  for (const taskLog of logs.value) {
    const text = String(taskLog)
    const key = `${taskId.value}:${text}`
    if (!syncedTaskLogs.has(key)) {
      syncedTaskLogs.add(key)
      addActivityLog(text, text.includes('失败') || text.includes('错误') ? 'error' : 'info')
    }
  }
  reportStep.value = data.step || 0
  reportMaxStep.value = data.max_step || 0
  if ((mode.value === '4' || mode.value === 'fund') && Number.isFinite(data.elapsed_seconds)) {
    elapsedSeconds.value = data.elapsed_seconds
  }

  if (data.status === 'done') {
    filename.value = data.filename || ''
    status.value = 'done'
    addActivityLog(`任务完成：${filename.value || '结果文件已生成'}`, 'success')
    stopPolling()
    stopElapsedTimer()
  } else if (data.status === 'paused') {
    status.value = 'paused'
    stopPolling()
    stopElapsedTimer()
  } else if (data.status === 'error') {
    error.value = data.error || data.message || '处理失败'
    status.value = 'error'
    addActivityLog(`任务处理失败：${error.value}`, 'error')
    stopPolling()
    stopElapsedTimer()
  }
  saveActiveTask()
}

function reset() {
  stopPolling()
  stopElapsedTimer(false)
  clearSavedTask()
  status.value = 'idle'
  taskId.value = ''
  current.value = 0
  total.value = 0
  message.value = ''
  filename.value = ''
  error.value = ''
  logs.value = []
  elapsedSeconds.value = 0
  elapsedAccumulated = 0
  files.value = []
  clearTemplate()
  clearReportFile()
  clearFundFile()
}

onMounted(() => {
  restoreActiveTask()
})

onUnmounted(() => {
  stopPolling()
  stopElapsedTimer(false)
})
</script>

<template>
  <div class="layout-app">
    <!-- 1. 左侧竖向导航栏 -->
    <aside class="sidebar-nav">
      <!-- 顶部 Logo 品牌 -->
      <div class="nav-brand">
        <div class="brand-icon">
          <!-- 绿白折叠叶子/纸张图标 -->
          <svg viewBox="0 0 32 32" width="24" height="24" fill="none">
            <path d="M7 6a3 3 0 013-3h12a3 3 0 013 3v13.5a6.5 6.5 0 01-6.5 6.5H10a3 3 0 01-3-3V6z" fill="#008765"/>
            <path d="M12 9h8M12 14h6" stroke="#fff" stroke-width="2" stroke-linecap="round"/>
            <path d="M7 20a6 6 0 006 6h5.5A6.5 6.5 0 0112 19.5V14H7v6z" fill="#2eb88a" fill-opacity="0.9"/>
          </svg>
        </div>
        <div class="brand-text">
          <span class="brand-title">财务内部工具</span>
          <span class="brand-desc">发票识别 · 表格处理</span>
        </div>
      </div>

      <!-- 工作组别模式选择 -->
      <div class="sidebar-section">
        <div class="sidebar-section-title">工作组别</div>
        <ModeSelect v-model="mode" :disabled="submitting" />
      </div>

      <button
        type="button"
        class="sidebar-log-link"
        :class="{ 'is-active': currentNav === 'logs' }"
        @click="currentNav = 'logs'"
      >
        <svg viewBox="0 0 24 24" width="18" height="18" fill="none" aria-hidden="true">
          <path d="M7 3.5h7l3 3V20a.5.5 0 01-.5.5h-9A.5.5 0 017 20V3.5zM10 10h4M10 14h4M10 18h2" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
        <span>处理日志</span>
        <b v-if="activityLogs.length">{{ activityLogs.length }}</b>
      </button>
    </aside>

    <!-- 2. 右侧主工作区容器 -->
    <div class="main-viewport">
      <template v-if="currentNav === 'home'">
      <!-- 步骤流程指示条 (步骤 1 / 步骤 2) -->
      <div class="step-indicator-bar">
        <div class="step-badge is-active">
          <span class="step-num">1</span>
          <div class="step-info">
            <span class="step-name">选择功能</span>
            <span class="step-desc">根据要处理的内容选择模式</span>
          </div>
        </div>
        <div class="step-connector"></div>
        <div class="step-badge" :class="{ 'is-active': mode }">
          <span class="step-num">2</span>
          <div class="step-info">
            <span class="step-name">上传文件并制作</span>
            <span class="step-desc">上传文件后系统自动识别处理生成结果</span>
          </div>
        </div>
      </div>

      <!-- 核心工作区双栏布局 (左栏: 功能选择+子选项 / 右栏: 上传+制作) -->
      <div class="workspace-grid" :class="{ 'single-col': mode === '4' || mode === 'fund' || (!receiptPerson && mode === 'receipt') }">
        <!-- 左栏：模式对应的细项配置（如发票类型、人员、排版模板） -->
        <section v-if="mode === 'receipt' || mode === '3'" class="left-section-card">
          <div class="section-head">
            <h2 class="section-title">
              {{ mode === 'receipt' ? '收款组配置' : '付款组发票类型' }}
            </h2>
            <p class="section-subtitle">
              {{ mode === 'receipt' ? '选择对应人员及发票分类' : '选择当前要处理的英文发票版式' }}
            </p>
          </div>

          <!-- 子选项：收款组选人员 / 付款组选发票类型 -->
          <div class="sub-config-box">
            <div class="sub-config-head">
              <span class="sub-config-title">{{ mode === 'receipt' ? '选择经办人员' : '选择发票类型（13种）' }}</span>
            </div>

            <!-- 收款组人员选择 -->
            <ReceiptTypeSelect
              v-if="mode === 'receipt'"
              v-model="receiptPerson"
              :disabled="submitting"
            />

            <!-- 付款组发票类型选择 -->
            <InvoiceTypeSelect
              v-else-if="mode === '3'"
              v-model="invType"
              :disabled="submitting"
            />
          </div>

          <!-- 收款组人员子类型选择 -->
          <div v-if="mode === 'receipt' && (receiptPerson === '赵淑华' || receiptPerson === '邵梅琳')" class="sub-config-box">
            <div class="sub-config-head">
              <span class="sub-config-tag">发票分类</span>
              <span class="sub-config-title">选择具体发票/凭证格式</span>
            </div>
            <ReceiptSubtypeSelect
              v-model="receiptSubtype"
              :person="receiptPerson"
              :disabled="submitting"
            />
          </div>

          <!-- 收款组谢莉丽可选模板设置 -->
          <div v-if="effectiveMode === '1'" class="sub-config-box template-box">
            <div class="sub-config-head">
              <span class="sub-config-tag">模板与排版</span>
              <span class="sub-config-title">插入已有表格（可选）</span>
            </div>
            <TemplateUpload
              :disabled="submitting"
              @selected="onTemplateSelected"
              @cleared="clearTemplate"
            />
            <div v-if="templateFile && !sheetError" class="sheet-pick-group">
              <span class="opt-label">选择工作表：</span>
              <div class="sheet-chips">
                <button
                  v-for="s in sheets"
                  :key="s"
                  type="button"
                  class="sheet-chip"
                  :class="{ on: selectedSheet === s }"
                  :disabled="submitting"
                  @click="selectedSheet = s"
                >{{ s }}</button>
              </div>
            </div>
            <div class="layout-dir-row">
              <span class="opt-label">排版方向：</span>
              <div class="seg-btn-group">
                <button type="button" class="seg-btn" :class="{ on: layoutDir === 'v' }" @click="layoutDir = 'v'">纵向排布</button>
                <button type="button" class="seg-btn" :class="{ on: layoutDir === 'h' }" @click="layoutDir = 'h'">横向排布</button>
              </div>
              <span class="opt-label ml-4">起始格：</span>
              <input v-model="startCell" class="cell-input" :class="{ invalid: !startCellValid }" placeholder="A1" />
            </div>
          </div>
        </section>

        <!-- 右栏：文件上传与操作卡片 -->
        <section class="right-section-card">
          <!-- 1. 发票/普通文件上传视图 -->
          <template v-if="mode !== '4' && mode !== 'fund'">
            <div class="section-head">
              <h2 class="section-title">
                {{ effectiveMode === '2' ? '上传原始表格 / 凭证' : effectiveMode === '6' && receiptSubtype !== 'invoice' ? '上传 ' + paymentLabel + ' 凭证' : '上传 PDF 文件' }}
              </h2>
              <p class="section-subtitle">
                {{ effectiveMode === '2' ? '支持 PDF，或 Excel 表格中的支付截图' : effectiveMode === '6' && receiptSubtype !== 'invoice' ? '支持图片、PDF，也可以直接拖入文件夹' : '支持多选，一次拖入全部发票，系统将自动识别并处理' }}
              </p>
            </div>

            <!-- 上传组件 -->
            <UploadArea
              :disabled="submitting"
              :count="files.length"
              :allow-directories="['1', '5', '6', '9'].includes(effectiveMode)"
              :accept="effectiveMode === '2' ? '.pdf,.xlsx,.xlsm' : effectiveMode === '6' ? '.pdf,.png,.jpg,.jpeg' : '.pdf'"
              :file-label="effectiveMode === '2' ? 'PDF 或 Excel 表格' : effectiveMode === '6' && receiptSubtype !== 'invoice' ? paymentLabel + '凭证' : 'PDF'"
              @add="addFiles"
              @remove="removeFile"
              @clear="clearFiles"
            >
              <div v-for="(f, i) in files" :key="f.name + i" class="file-item-row">
                <svg viewBox="0 0 20 20" width="16" height="16" fill="none" class="file-icon">
                  <path d="M6 2h5l4 4v12H6V2z" stroke="var(--primary)" stroke-width="1.6" stroke-linejoin="round" />
                  <path d="M11 2v4h4" stroke="var(--primary)" stroke-width="1.6" stroke-linejoin="round" />
                </svg>
                <span class="f-name" :title="f.name">{{ f.name }}</span>
                <span class="f-size">{{ formatSize(f.size) }}</span>
                <button class="f-remove" type="button" :disabled="submitting" @click="removeFile(i)">✕</button>
              </div>
            </UploadArea>
          </template>

          <!-- 2. 报表组上传视图 -->
          <template v-else-if="mode === '4'">
            <div class="section-head">
              <h2 class="section-title">上传报表文件</h2>
              <p class="section-subtitle">拖入 Excel 表格，读取工作表后生成无应收与无业务员成本总表</p>
            </div>

            <ReportUpload :disabled="submitting" @selected="onReportSelected" @cleared="clearReportFile" />

            <div v-if="reportLoading" class="report-loading-box">
              <div class="loading-head">
                <span>{{ reportUploadDone ? '正在读取工作表…' : '正在上传表格…' }}</span>
                <strong>{{ reportUploadPercent }}%</strong>
              </div>
              <div class="loading-track">
                <div class="loading-bar" :style="{ width: reportUploadPercent + '%' }"></div>
              </div>
            </div>

            <p v-if="reportError" class="form-error">{{ reportError }}</p>

            <div v-if="reportSheets.length" class="sheet-select-box">
              <span class="opt-label">选择要处理的工作表：</span>
              <div class="sheet-chips">
                <button
                  v-for="sheet in reportSheets"
                  :key="sheet"
                  type="button"
                  class="sheet-chip"
                  :class="{ on: reportSheet === sheet }"
                  :disabled="submitting"
                  @click="reportSheet = sheet"
                >{{ sheet }}</button>
              </div>
            </div>
          </template>

          <!-- 3. 资金组表格上传视图 -->
          <template v-else-if="mode === 'fund'">
            <div class="section-head">
              <h2 class="section-title">上传资金核对表格</h2>
              <p class="section-subtitle">拖入已有表格，按文件名自动识别对公、收款审核、服务商付款与银行流水</p>
            </div>

            <div class="fund-drag-box" :class="{ ready: fundFile3 }" @dragover.prevent @drop.prevent="e => selectFundFiles(e.dataTransfer.files)">
              <div class="fund-drag-icon">
                <svg viewBox="0 0 24 24" width="28" height="28" fill="none">
                  <path d="M12 15V7M8.5 10.5L12 7l3.5 3.5" stroke="var(--primary)" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>
                  <rect x="3" y="3" width="18" height="18" rx="4" stroke="var(--border-strong)" stroke-width="1.5"/>
                </svg>
              </div>
              <div class="fund-drag-text">
                <strong>拖入资金组 Excel 文件</strong>
                <span>对公表必传，收款、付款、银行账号管理流水按需提供</span>
              </div>
              <label class="fund-btn-browse">
                点击选择文件
                <input type="file" accept=".xlsx,.xlsm" multiple hidden :disabled="submitting" @change="e => { selectFundFiles(e.target.files); e.target.value = '' }" />
              </label>
            </div>

            <div class="fund-file-chips">
              <div v-for="([label, file]) in fundFileSlots" :key="label" class="fund-chip" :class="{ found: file }">
                <span class="fund-chip-status">{{ file ? '✓' : '·' }}</span>
                <span class="fund-chip-label">{{ label }}</span>
                <span class="fund-chip-val" :title="file ? file.name : ''">{{ file ? file.name : '未识别' }}</span>
              </div>
            </div>
          </template>

          <!-- 开始制作主操作按钮 -->
          <div class="action-footer">
            <button
              class="primary-start-btn"
              type="button"
              :disabled="!canSubmit"
              @click="submit"
            >
              <span v-if="submitting" class="btn-spinner"></span>
              <span>{{ submitting ? '正在处理中…' : '开始制作' }}</span>
              <svg v-if="!submitting" viewBox="0 0 20 20" width="18" height="18" fill="none">
                <path d="M4 10h12M11 5l5 5-5 5" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
            </button>
          </div>

          <!-- 进度监控面板 -->
          <ProgressPanel
            v-if="submitting || status === 'paused'"
            :status="'processing'"
            :current="current"
            :total="total"
            :message="message"
            :logs="logs"
            :elapsed-seconds="elapsedSeconds"
            :report-profile="reportProfile"
            :report-flow="mode === '4'"
          />

          <!-- 报表组单步暂停继续按钮 -->
          <div v-if="status === 'paused'" class="paused-action-row">
            <a class="btn-step-download" :href="`/api/tasks/${taskId}/download`" :download="filename">下载第 {{ reportStep }} 步结果</a>
            <button class="btn-step-continue" type="button" @click="continueReport">继续下一步</button>
          </div>

          <!-- 完成或失败结果面板 -->
          <ResultPanel
            v-if="status === 'done' || status === 'error'"
            :status="status"
            :task-id="taskId"
            :filename="filename"
            :error="error"
            :elapsed-seconds="elapsedSeconds"
            @reset="reset"
          />
        </section>
      </div>
      </template>

      <section v-else class="log-page">
        <header class="log-page-head">
          <div>
            <p class="log-page-kicker">TASK CONSOLE</p>
            <h1>处理日志</h1>
            <p>本会话任务的实时输出与异常记录。</p>
          </div>
          <div class="log-page-actions">
            <button type="button" class="log-back-btn" @click="currentNav = 'home'">返回工作台</button>
            <button type="button" class="log-export-btn" :disabled="!activityLogs.length" @click="exportActivityLogs">导出日志</button>
          </div>
        </header>
        <div class="log-terminal" role="log" aria-live="polite">
          <div class="log-terminal-bar">
            <div class="terminal-window-controls" aria-hidden="true"><i></i><i></i><i></i></div>
            <span class="terminal-title">finance-tool / task-output</span>
            <span class="terminal-count">{{ activityLogs.length }} lines</span>
          </div>
          <div ref="logTerminal" class="log-page-list">
          <p class="terminal-banner">财务内部工具 · 会话日志已启动</p>
          <p v-if="!activityLogs.length" class="activity-log-empty">等待任务输出…</p>
          <div v-for="item in activityLogs" :key="item.id" class="log-page-item" :class="`is-${item.level}`">
            <time>[{{ item.time }}]</time>
            <span class="terminal-level">{{ item.level === 'error' ? 'ERR' : item.level === 'warning' ? 'WRN' : item.level === 'success' ? 'OK ' : 'LOG' }}</span>
            <p><span class="terminal-prompt">›</span>{{ item.content }}</p>
          </div>
          <p v-if="activityLogs.length" class="terminal-cursor"><span>›</span></p>
          </div>
        </div>
      </section>
    </div>
  </div>
</template>

<style scoped>
/* 整个应用外壳：左侧导航 230px + 右侧工作区 */
.layout-app {
  display: flex;
  min-height: 100vh;
  background-color: var(--bg);
}

/* 1. 侧边栏导航 */
.sidebar-nav {
  width: 270px;
  flex-shrink: 0;
  background: var(--surface);
  border-right: 1px solid var(--border);
  display: flex;
  flex-direction: column;
  padding: 20px 14px 18px;
  position: sticky;
  top: 0;
  height: 100vh;
  overflow-y: auto;
}

.sidebar-section {
  display: flex;
  flex-direction: column;
  gap: 8px;
  margin-bottom: 16px;
}

.sidebar-section-title {
  font-size: 11.5px;
  font-weight: 700;
  color: var(--text-faint);
  letter-spacing: 0.5px;
  padding: 0 4px;
}

.sidebar-log-link {
  display: flex;
  align-items: center;
  gap: 9px;
  width: 100%;
  margin-top: auto;
  padding: 10px 12px;
  border: 1px solid #dbe9e3;
  border-radius: var(--radius-s);
  background: #f7fbf9;
  color: var(--text-secondary);
  font-size: 12.5px;
  font-weight: 700;
  cursor: pointer;
  text-align: left;
}

.sidebar-log-link:hover,
.sidebar-log-link.is-active {
  border-color: #b8ded0;
  background: var(--primary-soft);
  color: var(--primary-ink);
}

.sidebar-log-link b {
  min-width: 18px;
  margin-left: auto;
  padding: 1px 5px;
  border-radius: 99px;
  background: var(--primary);
  color: #fff;
  font-size: 10px;
  line-height: 16px;
  text-align: center;
}

.log-page {
  max-width: 1080px;
  margin: 22px auto 0;
  padding: clamp(22px, 4vw, 42px);
}

.log-page-head {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 24px;
  padding: 0 0 22px;
}

.log-page-kicker {
  margin: 0 0 7px;
  color: var(--primary-ink);
  font-size: 11px;
  font-weight: 800;
  letter-spacing: .12em;
}

.log-page h1 {
  margin: 0;
  color: var(--text);
  font-size: clamp(24px, 3vw, 32px);
  letter-spacing: -.025em;
}

.log-page-head p:not(.log-page-kicker) {
  margin: 8px 0 0;
  color: var(--text-soft);
  font-size: 13px;
}

.log-page-actions {
  display: flex;
  flex-shrink: 0;
  gap: 8px;
}

.log-back-btn, .log-export-btn {
  padding: 9px 12px;
  border-radius: 7px;
  font-size: 12px;
  font-weight: 700;
  cursor: pointer;
}

.log-back-btn {
  border: 1px solid var(--border-strong);
  background: var(--surface);
  color: var(--text-secondary);
}

.log-export-btn {
  border: 1px solid var(--primary);
  background: var(--primary);
  color: #fff;
}

.log-export-btn:disabled { opacity: .45; cursor: not-allowed; }

.log-terminal {
  overflow: hidden;
  border: 1px solid #24383a;
  border-radius: 12px;
  background: #101b1c;
  box-shadow: 0 20px 42px -28px rgba(15, 42, 38, .72);
}

.log-terminal-bar {
  display: flex;
  align-items: center;
  min-height: 42px;
  padding: 0 14px;
  background: #19292a;
  border-bottom: 1px solid #294344;
  color: #9cb9b4;
  font-family: var(--font-num), Consolas, monospace;
  font-size: 11px;
}

.terminal-window-controls { display: flex; gap: 6px; margin-right: 12px; }
.terminal-window-controls i { width: 8px; height: 8px; border-radius: 50%; background: #52706b; }
.terminal-window-controls i:first-child { background: #d57363; }
.terminal-window-controls i:nth-child(2) { background: #c8a259; }
.terminal-window-controls i:last-child { background: #5aad82; }
.terminal-title { color: #c8dcd7; }
.terminal-count { margin-left: auto; color: #7ca29a; font-variant-numeric: tabular-nums; }

.log-page-list {
  max-height: calc(100vh - 250px);
  min-height: clamp(480px, 62vh, 720px);
  overflow: auto;
  padding: 17px 18px 20px;
  scrollbar-color: #3b5a57 #101b1c;
}

.log-page-item {
  display: grid;
  grid-template-columns: 76px 32px minmax(0, 1fr);
  gap: 10px;
  align-items: start;
  padding: 3px 0;
  color: #c3d8d2;
  font-family: var(--font-num), Consolas, monospace;
  font-size: 12px;
  line-height: 1.65;
}

.log-page-item time {
  color: #71928c;
  font-size: inherit;
  font-variant-numeric: tabular-nums;
}

.log-page-item p { margin: 0; overflow-wrap: anywhere; }
.terminal-level { color: #81aaa1; font-size: 10px; font-weight: 800; letter-spacing: .06em; }
.terminal-prompt { margin-right: 8px; color: #58c8ac; font-weight: 800; }
.log-page-item.is-success .terminal-level, .log-page-item.is-success .terminal-prompt { color: #6ed4a3; }
.log-page-item.is-warning .terminal-level, .log-page-item.is-warning .terminal-prompt { color: #e1b85d; }
.log-page-item.is-error .terminal-level, .log-page-item.is-error .terminal-prompt { color: #ed8175; }
.terminal-banner { margin: 0 0 12px; color: #79aaa0; font-family: var(--font-num), Consolas, monospace; font-size: 11px; }
.terminal-banner::before { content: '# '; color: #58c8ac; }
.activity-log-empty { margin: 18px 0; color: #76958e; font-family: var(--font-num), Consolas, monospace; }
.activity-log-empty::before { content: '› '; color: #58c8ac; }
.terminal-cursor { margin: 6px 0 0; color: #62d2b5; font-family: var(--font-num), Consolas, monospace; font-weight: 800; }
.terminal-cursor span { display: inline-block; animation: terminal-blink 1.1s step-end infinite; }
@keyframes terminal-blink { 50% { opacity: 0; } }

.workspace-grid.single-col {
  grid-template-columns: minmax(0, 1fr);
  max-width: 820px;
  margin: 0 auto;
}

.nav-brand {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 4px 6px 24px;
}

.brand-icon {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 38px;
  height: 38px;
  border-radius: 10px;
  background: var(--primary-soft);
}

.brand-text {
  display: flex;
  flex-direction: column;
}

.brand-title {
  font-size: 15px;
  font-weight: 800;
  color: var(--text);
  line-height: 1.2;
}

.brand-desc {
  font-size: 11px;
  color: var(--text-faint);
  margin-top: 2px;
}

.nav-menu {
  display: flex;
  flex-direction: column;
  gap: 4px;
}

.nav-item {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 10px 14px;
  border-radius: var(--radius-s);
  border: none;
  background: transparent;
  color: var(--text-secondary);
  font-size: 13.5px;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.16s ease;
  text-align: left;
}

.nav-item:hover {
  background: var(--surface-2);
  color: var(--text);
}

.nav-item.is-active {
  background: var(--primary-soft);
  color: var(--primary-ink);
  font-weight: 700;
}

.nav-icon {
  flex-shrink: 0;
}

/* 侧边栏底部插画 */
.sidebar-banner {
  margin-top: auto;
  display: flex;
  flex-direction: column;
  align-items: center;
  text-align: center;
  padding: 16px 12px;
  border-radius: var(--radius-m);
  background: linear-gradient(180deg, #f3faf7 0%, #eaf5f1 100%);
  border: 1px solid #d8ede4;
}

.banner-illus {
  margin-bottom: 6px;
}

.banner-title {
  font-size: 13px;
  font-weight: 700;
  color: var(--primary-ink);
}

.banner-sub {
  font-size: 11px;
  color: var(--text-soft);
  margin-top: 2px;
}

/* 2. 右侧主工作区 */
.main-viewport {
  flex: 1;
  min-width: 0;
  padding: 24px 36px 48px;
  max-width: 1440px;
  margin: 0 auto;
}

/* 顶部 Header */
.top-header {
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  margin-bottom: 20px;
}

.page-title {
  margin: 0;
  font-size: 24px;
  font-weight: 800;
  color: var(--text);
  letter-spacing: -0.01em;
}

.page-subtitle {
  margin: 4px 0 0;
  font-size: 13px;
  color: var(--text-soft);
}

.header-right {
  display: flex;
  align-items: center;
  gap: 20px;
}

.header-link {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  background: none;
  border: none;
  color: var(--text-secondary);
  font-size: 13px;
  cursor: pointer;
  padding: 6px 10px;
  border-radius: 6px;
  transition: all 0.15s ease;
}

.header-link:hover {
  background: var(--surface);
  color: var(--primary);
}

.user-profile {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 4px 10px 4px 4px;
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius-pill);
  cursor: pointer;
}

.user-avatar {
  width: 28px;
  height: 28px;
  border-radius: 50%;
  overflow: hidden;
  background: #e2f1f8;
}

.user-avatar img {
  width: 100%;
  height: 100%;
  display: block;
}

.user-name {
  font-size: 13px;
  font-weight: 600;
  color: var(--text);
}

.dropdown-arrow {
  color: var(--text-faint);
}

/* 步骤横向指示条 */
.step-indicator-bar {
  display: flex;
  align-items: center;
  margin-bottom: 24px;
  padding: 0 4px;
}

.step-badge {
  display: flex;
  align-items: center;
  gap: 10px;
  opacity: 0.55;
  transition: opacity 0.2s ease;
}

.step-badge.is-active {
  opacity: 1;
}

.step-num {
  width: 26px;
  height: 26px;
  border-radius: 50%;
  background: var(--surface);
  border: 1.5px solid var(--border-strong);
  color: var(--text-soft);
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 13px;
  font-weight: 800;
  font-family: var(--font-num);
}

.step-badge.is-active .step-num {
  background: var(--primary);
  border-color: var(--primary);
  color: #fff;
  box-shadow: 0 2px 8px var(--primary-shadow);
}

.step-info {
  display: flex;
  flex-direction: column;
}

.step-name {
  font-size: 14px;
  font-weight: 700;
  color: var(--text);
}

.step-desc {
  font-size: 11.5px;
  color: var(--text-faint);
}

.step-connector {
  flex: 1;
  height: 1px;
  background: var(--border);
  margin: 0 24px;
}

/* 核心工作区双栏卡片网格 */
.workspace-grid {
  display: grid;
  grid-template-columns: minmax(460px, 1fr) minmax(440px, 1fr);
  gap: 24px;
  align-items: start;
}

/* 卡片统一白色底、圆角与边框 */
.left-section-card,
.right-section-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: var(--radius-xl);
  padding: 24px;
  box-shadow: var(--card-shadow);
  display: flex;
  flex-direction: column;
  gap: 20px;
}

.section-head {
  display: flex;
  flex-direction: column;
  gap: 4px;
}

.section-title {
  margin: 0;
  font-size: 18px;
  font-weight: 800;
  color: var(--text);
}

.section-subtitle {
  margin: 0;
  font-size: 12.5px;
  color: var(--text-soft);
}

/* 细项配置容器 */
.sub-config-box {
  background: var(--surface-2);
  border: 1px solid var(--border);
  border-radius: var(--radius-m);
  padding: 16px;
  display: flex;
  flex-direction: column;
  gap: 12px;
}

.sub-config-head {
  display: flex;
  align-items: center;
  gap: 8px;
}

.sub-config-tag {
  font-size: 10.5px;
  font-weight: 700;
  padding: 2px 6px;
  background: var(--primary-soft);
  color: var(--primary-ink);
  border-radius: var(--radius-xs);
}

.sub-config-title {
  font-size: 13px;
  font-weight: 700;
  color: var(--text);
}

/* 模板与排版 */
.sheet-pick-group,
.layout-dir-row {
  display: flex;
  align-items: center;
  flex-wrap: wrap;
  gap: 10px;
  margin-top: 4px;
}

.opt-label {
  font-size: 12.5px;
  font-weight: 600;
  color: var(--text-secondary);
}

.sheet-chips {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
}

.sheet-chip {
  padding: 4px 10px;
  font-size: 12px;
  border-radius: 6px;
  border: 1px solid var(--border-strong);
  background: var(--surface);
  cursor: pointer;
  color: var(--text-secondary);
}

.sheet-chip.on {
  background: var(--primary);
  color: #fff;
  border-color: var(--primary);
}

.seg-btn-group {
  display: inline-flex;
  background: var(--surface);
  padding: 2px;
  border-radius: 8px;
  border: 1px solid var(--border);
}

.seg-btn {
  border: none;
  background: transparent;
  padding: 4px 10px;
  font-size: 12px;
  border-radius: 6px;
  cursor: pointer;
  color: var(--text-soft);
}

.seg-btn.on {
  background: var(--primary);
  color: #fff;
  font-weight: 600;
}

.cell-input {
  width: 60px;
  padding: 4px 8px;
  font-size: 12.5px;
  font-weight: 700;
  text-transform: uppercase;
  border: 1.5px solid var(--border-strong);
  border-radius: 6px;
  background: #fff;
}

.cell-input.invalid {
  border-color: var(--danger);
}

.ml-4 {
  margin-left: 12px;
}

/* 底部帮助 Banner */
.helper-banner {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 12px 16px;
  border-radius: var(--radius-m);
  background: #fdfaf3;
  border: 1px solid #f6e8c3;
  margin-top: auto;
}

.helper-icon {
  font-size: 18px;
}

.helper-text {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: 2px;
}

.helper-text strong {
  font-size: 12.5px;
  color: #92580c;
}

.helper-text span {
  font-size: 11.5px;
  color: #a47638;
}

.helper-link-btn {
  display: inline-flex;
  align-items: center;
  gap: 3px;
  border: none;
  background: transparent;
  color: #b45309;
  font-size: 12px;
  font-weight: 700;
  cursor: pointer;
}

/* 右栏内部 */
.file-item-row {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 8px 10px;
  background: var(--surface-2);
  border: 1px solid var(--border);
  border-radius: var(--radius-s);
  font-size: 12.5px;
}

.file-icon {
  flex-shrink: 0;
}

.f-name {
  flex: 1;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-weight: 600;
}

.f-size {
  color: var(--text-faint);
  font-size: 11.5px;
}

.f-remove {
  border: none;
  background: transparent;
  color: var(--text-faint);
  cursor: pointer;
  padding: 2px 6px;
  border-radius: 4px;
}

.f-remove:hover {
  background: var(--danger-soft);
  color: var(--danger);
}

/* 资金组拖拽框 */
.fund-drag-box {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 10px;
  padding: 28px 16px;
  border: 1.5px dashed var(--border-strong);
  border-radius: var(--radius-m);
  background: var(--surface);
  text-align: center;
}

.fund-drag-box.ready {
  border-color: var(--primary);
  background: var(--primary-soft);
}

.fund-drag-text {
  display: flex;
  flex-direction: column;
  gap: 3px;
}

.fund-drag-text strong {
  font-size: 14.5px;
  color: var(--text);
}

.fund-drag-text span {
  font-size: 12px;
  color: var(--text-soft);
}

.fund-btn-browse {
  display: inline-block;
  padding: 6px 16px;
  border: 1px solid var(--primary);
  border-radius: var(--radius-pill);
  color: var(--primary);
  font-size: 12px;
  font-weight: 600;
  cursor: pointer;
  background: #fff;
}

.fund-btn-browse:hover {
  background: var(--primary);
  color: #fff;
}

.fund-file-chips {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 8px;
}

.fund-chip {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 8px 10px;
  background: var(--surface-2);
  border: 1px solid var(--border);
  border-radius: 8px;
  font-size: 12px;
}

.fund-chip.found {
  background: var(--primary-soft);
  border-color: var(--primary-soft-border);
  color: var(--primary-ink);
}

.fund-chip-status {
  font-weight: 800;
}

.fund-chip-label {
  font-weight: 600;
}

.fund-chip-val {
  margin-left: auto;
  font-size: 11px;
  color: var(--text-faint);
  max-width: 90px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

/* 报表上传状态 */
.report-loading-box {
  padding: 12px 14px;
  background: var(--surface-2);
  border-radius: 8px;
  border: 1px solid var(--border);
}

.loading-head {
  display: flex;
  justify-content: space-between;
  font-size: 12px;
  color: var(--text-secondary);
}

.loading-track {
  height: 6px;
  background: var(--surface);
  border-radius: 99px;
  margin-top: 6px;
  overflow: hidden;
}

.loading-bar {
  height: 100%;
  background: var(--primary);
  transition: width 0.2s ease;
}

.form-error {
  color: var(--danger);
  font-size: 12px;
  margin: 0;
}

/* 文件安全卡片 */
.security-card {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 12px 16px;
  background: #f0f9f6;
  border: 1px solid #d2ede3;
  border-radius: var(--radius-m);
}

.security-icon {
  display: flex;
  align-items: center;
  justify-content: center;
  flex-shrink: 0;
}

.security-text {
  display: flex;
  flex-direction: column;
  gap: 2px;
}

.security-text strong {
  font-size: 13px;
  color: var(--primary-ink);
}

.security-text span {
  font-size: 11.5px;
  color: var(--text-soft);
  line-height: 1.4;
}

/* 开始制作大按钮 */
.action-footer {
  margin-top: 4px;
}

.primary-start-btn {
  width: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  padding: 14px 28px;
  background: var(--primary);
  color: #ffffff;
  border: none;
  border-radius: var(--radius-pill);
  font-size: 15px;
  font-weight: 700;
  cursor: pointer;
  box-shadow: 0 4px 16px var(--primary-shadow);
  transition: all 0.2s var(--ease-smooth);
}

.primary-start-btn:hover:not(:disabled) {
  background: var(--primary-hover);
  transform: translateY(-1px);
  box-shadow: 0 8px 24px var(--primary-shadow);
}

.primary-start-btn:active:not(:disabled) {
  transform: scale(0.99);
}

.primary-start-btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
  box-shadow: none;
}

.btn-spinner {
  width: 16px;
  height: 16px;
  border: 2.5px solid rgba(255, 255, 255, 0.3);
  border-top-color: #fff;
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
}

@keyframes spin {
  to { transform: rotate(360deg); }
}

.paused-action-row {
  display: flex;
  gap: 12px;
}

.btn-step-download {
  flex: 1;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 10px;
  background: var(--primary);
  color: #fff;
  border-radius: 8px;
  font-size: 13px;
  font-weight: 700;
}

.btn-step-continue {
  flex: 1;
  padding: 10px;
  border: 1px solid var(--border-strong);
  background: var(--surface);
  border-radius: 8px;
  font-size: 13px;
  font-weight: 700;
  cursor: pointer;
}

/* 响应式断点适配 */
@media (max-width: 1080px) {
  .workspace-grid {
    grid-template-columns: 1fr;
  }
}

@media (max-width: 768px) {
  .sidebar-nav {
    display: none;
  }
  .main-viewport {
    padding: 16px;
  }
}
</style>
