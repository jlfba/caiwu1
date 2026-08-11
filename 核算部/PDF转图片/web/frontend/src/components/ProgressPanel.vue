<script setup>
import { computed } from 'vue'

const props = defineProps({
  status: { type: String, required: true },
  current: { type: Number, default: 0 },
  total: { type: Number, default: 0 },
  message: { type: String, default: '' }
})

const percent = computed(() => {
  if (props.status === 'done') return 100
  if (!props.total) return 0
  return Math.min(100, Math.round((props.current / props.total) * 100))
})
</script>

<template>
  <div class="progress-panel">
    <div class="pp-head">
      <span class="pp-status" :class="status">
        <span class="dot"></span>
        {{ status === 'processing' ? '正在处理' : status === 'pending' ? '排队等待' : '处理完成' }}
      </span>
      <span class="pp-percent">{{ percent }}%</span>
    </div>

    <div class="track">
      <div class="bar" :class="{ error: status === 'error' }" :style="{ width: percent + '%' }"></div>
    </div>

    <div class="pp-msg">{{ message }}</div>
  </div>
</template>

<style scoped>
.progress-panel {
  background: var(--card);
  border: 1.5px solid var(--border);
  border-radius: var(--radius);
  padding: 18px 20px;
  box-shadow: var(--shadow-sm);
}

.pp-head {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 12px;
}

.pp-status {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  font-weight: 700;
  font-size: 14px;
  color: var(--text);
}

.pp-status .dot {
  width: 9px;
  height: 9px;
  border-radius: 50%;
  background: var(--text-faint);
}

.pp-status.processing .dot {
  background: var(--accent);
  animation: pulse 1.2s ease-in-out infinite;
}

.pp-status.done .dot {
  background: var(--primary);
}

@keyframes pulse {
  0%,
  100% {
    opacity: 1;
  }
  50% {
    opacity: 0.3;
  }
}

.pp-percent {
  font-size: 15px;
  font-weight: 800;
  color: var(--primary);
  font-variant-numeric: tabular-nums;
}

.track {
  height: 10px;
  background: #e8eef5;
  border-radius: 999px;
  overflow: hidden;
}

.bar {
  height: 100%;
  background: linear-gradient(90deg, var(--primary), var(--accent));
  border-radius: 999px;
  transition: width 0.4s ease;
}

.bar.error {
  background: var(--danger);
}

.pp-msg {
  margin-top: 12px;
  font-size: 13px;
  color: var(--text-soft);
  line-height: 1.5;
}
</style>
