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
  <div v-if="status === 'done'" class="result-panel" :class="status">
    <span class="rp-icon">
      <svg viewBox="0 0 40 40" width="40" height="40" fill="none" aria-hidden="true">
        <circle cx="20" cy="20" r="19" fill="var(--primary-soft)" />
        <path d="M13 20.5l4.8 4.5L27.5 15" stroke="var(--primary-strong)" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" />
      </svg>
    </span>
    <div class="rp-text">
      <h3 class="rp-title">处理完成</h3>
      <p class="rp-sub" :title="filename">{{ filename || '表格已生成' }}</p>
    </div>
    <div class="rp-actions">
      <a class="btn primary" :href="downloadUrl(taskId)" :download="filename">
        <svg viewBox="0 0 20 20" width="15" height="15" fill="none" aria-hidden="true">
          <path d="M10 3v9m0 0l-3.5-3.5M10 12l3.5-3.5M4.5 15.5h11" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" />
        </svg>
        下载表格
      </a>
      <button class="btn ghost" type="button" @click="$emit('reset')">继续处理</button>
    </div>
  </div>

  <div v-else-if="status === 'error'" class="result-panel" :class="status">
    <span class="rp-icon">
      <svg viewBox="0 0 40 40" width="40" height="40" fill="none" aria-hidden="true">
        <circle cx="20" cy="20" r="19" fill="var(--danger-soft)" />
        <path d="M20 12v9" stroke="var(--danger)" stroke-width="3" stroke-linecap="round" />
        <circle cx="20" cy="26" r="1.8" fill="var(--danger)" />
      </svg>
    </span>
    <div class="rp-text">
      <h3 class="rp-title">处理失败</h3>
      <p class="rp-sub">{{ error }}</p>
    </div>
    <div class="rp-actions">
      <button class="btn ghost" type="button" @click="$emit('reset')">返回重试</button>
    </div>
  </div>
</template>

<style scoped>
.result-panel {
  display: flex;
  align-items: center;
  gap: 16px;
  flex-wrap: wrap;
  padding: 22px 24px;
  border-radius: var(--radius);
  border: 1.5px solid var(--border);
  background: var(--surface);
  animation: fade 0.35s var(--ease-out);
}

.result-panel.done {
  border-color: var(--primary);
  background: linear-gradient(180deg, var(--primary-soft), var(--surface) 56%);
}

.result-panel.error {
  border-color: rgba(194, 65, 60, 0.4);
  background: linear-gradient(180deg, var(--danger-soft), var(--surface) 56%);
}

@keyframes fade {
  from {
    opacity: 0;
    transform: translateY(8px);
  }
  to {
    opacity: 1;
    transform: none;
  }
}

.rp-icon {
  flex-shrink: 0;
  display: grid;
  place-items: center;
}

.rp-text {
  flex: 1;
  min-width: 0;
}

.rp-title {
  margin: 0;
  font-size: 16px;
  font-weight: 800;
}

.rp-sub {
  margin: 4px 0 0;
  font-size: 12.5px;
  color: var(--text-soft);
  word-break: break-all;
}

.rp-actions {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
  margin-left: auto;
}

.btn {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  border: none;
  border-radius: 10px;
  padding: 10px 20px;
  font-size: 13.5px;
  font-weight: 700;
  text-decoration: none;
  cursor: pointer;
  transition: transform 0.16s var(--ease-out), background 0.16s var(--ease-out),
    box-shadow 0.16s var(--ease-out);
}

.btn:active {
  transform: scale(0.97);
}

.btn.primary {
  background: var(--primary);
  color: #fff;
  box-shadow: 0 8px 18px -8px rgba(13, 138, 122, 0.55);
}

.btn.primary:hover {
  background: var(--primary-strong);
}

.btn.ghost {
  background: var(--surface);
  color: var(--text-soft);
  border: 1.5px solid var(--border);
}

.btn.ghost:hover {
  color: var(--text);
  border-color: var(--border-strong);
}

@media (max-width: 640px) {
  .rp-actions {
    margin-left: 0;
    width: 100%;
  }
  .btn {
    flex: 1;
    justify-content: center;
  }
}
</style>
