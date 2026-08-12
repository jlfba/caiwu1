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

async function getWorksheets(file) {
  const fd = new FormData()
  fd.append('template', file)
  const res = await fetch('/api/worksheets', { method: 'POST', body: fd })
  const data = await res.json()
  if (!res.ok) throw new Error(data.detail || '读取工作表失败')
  return data.sheets
}

function downloadUrl(taskId) {
  return `/api/tasks/${taskId}/download`
}

export { createTask, getTask, getWorksheets, downloadUrl }
