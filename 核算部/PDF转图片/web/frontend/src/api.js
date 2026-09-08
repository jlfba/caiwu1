async function createTask({
  files,
  mode,
  invType,
  layout = 'v',
  startCell = 'A1',
  template,
  sheetName
}) {
  const fd = new FormData()
  for (const f of files) fd.append('files', f, f.name)
  fd.append('mode', mode)
  fd.append('inv_type', invType)
  fd.append('layout', layout)
  fd.append('start_cell', startCell)
  if (template) {
    fd.append('template', template, template.name)
    fd.append('sheet_name', sheetName || '')
  }
  const res = await fetch('/api/tasks', { method: 'POST', body: fd })
  const data = await res.json().catch(() => ({}))
  if (!res.ok) throw new Error(data.detail || '创建任务失败')
  return data
}

async function getTask(taskId) {
  const res = await fetch(`/api/tasks/${taskId}`)
  const data = await res.json()
  if (!res.ok) throw new Error(data.detail || '查询任务失败')
  return data
}

async function getWorksheets(file, onProgress) {
  const fd = new FormData()
  fd.append('template', file)
  if (!onProgress) {
    const res = await fetch('/api/worksheets', { method: 'POST', body: fd })
    const data = await res.json()
    if (!res.ok) throw new Error(data.detail || '读取工作表失败')
    return data.sheets
  }
  return await new Promise((resolve, reject) => {
    const request = new XMLHttpRequest()
    request.open('POST', '/api/worksheets')
    request.upload.addEventListener('progress', (event) => {
      if (event.lengthComputable) onProgress(Math.round((event.loaded / event.total) * 100))
    })
    request.addEventListener('load', () => {
      let data = {}
      try { data = JSON.parse(request.responseText || '{}') } catch { /* handled below */ }
      if (request.status < 200 || request.status >= 300) {
        reject(new Error(data.detail || '读取工作表失败'))
        return
      }
      onProgress(100)
      resolve(data.sheets || [])
    })
    request.addEventListener('error', () => reject(new Error('上传表格失败，请检查网络连接')))
    request.addEventListener('abort', () => reject(new Error('上传表格已取消')))
    request.send(fd)
  })
}

async function createReportTask(file, sheetName, reportProfile = 'bundle') {
  const fd = new FormData()
  fd.append('file', file, file.name)
  fd.append('sheet_name', sheetName)
  fd.append('report_profile', reportProfile)
  const res = await fetch('/api/report-tasks', { method: 'POST', body: fd })
  const data = await res.json().catch(() => ({}))
  if (!res.ok) throw new Error(data.detail || '创建报表任务失败')
  return data
}

async function continueReportTask(taskId) {
  const res = await fetch(`/api/report-tasks/${taskId}/continue`, { method: 'POST' })
  const data = await res.json().catch(() => ({}))
  if (!res.ok) throw new Error(data.detail || '继续处理失败')
  return data
}

function downloadUrl(taskId) {
  return `/api/tasks/${taskId}/download`
}

export { createTask, createReportTask, continueReportTask, getTask, getWorksheets, downloadUrl }
