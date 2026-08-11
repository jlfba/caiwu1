<script setup>
import { downloadUrl } from '../api'

defineProps({
  taskId: { type: String, default: '' },
  filename: { type: String, default: '' },
  error: { type: String, default: '' },
  status: { type: String, required: true }
})
defineEmits(['reset'])
</script>

<template>
  <div v-if="status === 'done'" class="result-panel done">
    <div class="result-icon">✅</div>
    <div class="result-title">处理完成！</div>
    <div class="result-sub">{{ filename || '已生成表格' }}</div>
    <div class="result-actions">
      <a class="btn primary" :href="downloadUrl(taskId)" :download="filename">
        下载 {{ filename || 'Excel 表格' }}
      </a>
      <button class="btn ghost" type="button" @click="$emit('reset')">再做一个</button>
    </div>
  </div>

  <div v-else-if="status === 'error'" class="result-panel error">
    <div class="result-icon">⚠️</div>
    <div class="result-title">处理失败</div>
    <div class="result-sub">{{ error }}</div>
    <div class="result-actions">
      <button class="btn ghost" type="button" @click="$emit('reset')">返回重试</button>
    </div>
  </div>
</template>

<style scoped>
.result-panel {
  background: var(--card);
  border: 1.5px solid var(--border);
  border-radius: var(--radius);
  padding: 28px 24px;
  text-align: center;
  box-shadow: var(--shadow-sm);
}

.result-panel.done {
  border-color: #99f6e4;
}

.result-panel.error {
  border-color: #fecaca;
}

.result-icon {
  font-size: 40px;
  line-height: 1;
  margin-bottom: 12px;
}

.result-title {
  font-size: 18px;
  font-weight: 800;
}

.result-sub {
  margin-top: 6px;
  font-size: 13px;
  color: var(--text-soft);
  word-break: break-all;
}

.result-actions {
  margin-top: 20px;
  display: flex;
  justify-content: center;
  gap: 12px;
  flex-wrap: wrap;
}

.btn {
  border: none;
  border-radius: 10px;
  padding: 11px 22px;
  font-size: 14px;
  font-weight: 700;
  text-decoration: none;
  display: inline-flex;
  align-items: center;
  transition: all 0.15s ease;
}

.btn.primary {
  background: var(--primary);
  color: #fff;
  box-shadow: 0 6px 16px -6px rgba(15, 118, 110, 0.5);
}

.btn.primary:hover {
  background: var(--primary-strong);
}

.btn.ghost {
  background: #f1f5f9;
  color: var(--text);
}

.btn.ghost:hover {
  background: #e2e8f0;
}
</style>
