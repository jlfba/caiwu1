async function createTask({ files, mode, invType }) {
  const fd = new FormData()
  for (const f of files) fd.append('files', f, f.name)
  fd.append('mode', mode)
  fd.append('inv_type', invType)
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

function downloadUrl(taskId) {
  return `/api/tasks/${taskId}/download`
}

export { createTask, getTask, downloadUrl }
