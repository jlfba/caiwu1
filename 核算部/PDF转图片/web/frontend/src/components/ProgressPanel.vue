<script setup>
import { computed } from 'vue'

const props = defineProps({
  status: { type: String, required: true },
  current: { type: Number, default: 0 },
  total: { type: Number, default: 0 },
  message: { type: String, default: '' },
  elapsedSeconds: { type: Number, default: -1 }
})

const percent = computed(() => {
  if (props.status === 'done') return 100
  if (!props.total) return 0
  return Math.min(100, Math.round((props.current / props.total) * 100))
})

const elapsedText = computed(() => {
  const seconds = Math.max(0, props.elapsedSeconds)
  const hours = Math.floor(seconds / 3600)
  const minutes = Math.floor((seconds % 3600) / 60)
  const rest = seconds % 60
  const parts = [minutes, rest].map((value) => String(value).padStart(2, '0'))
  if (hours) parts.unshift(String(hours))
  return parts.join(':')
})
</script>

<template>
  <div class="progress-panel">
    <div class="pp-head">
      <span class="pp-title">
        <span class="pp-label">处理进度</span>
        <span v-if="elapsedSeconds >= 0" class="pp-elapsed">耗时 {{ elapsedText }}</span>
      </span>
      <span class="pp-percent">{{ percent }}<i>%</i></span>
    </div>
    <div class="track" role="progressbar" :aria-valuenow="percent" aria-valuemin="0" aria-valuemax="100">
      <div class="bar" :style="{ width: percent + '%' }"></div>
    </div>
    <p class="pp-msg">{{ message }}</p>
  </div>
</template>

<style scoped>
.progress-panel {
  background: var(--surface);
  border: 1.5px solid var(--border);
  border-radius: var(--radius);
  padding: 20px 22px;
  animation: fade 0.3s var(--ease-out);
}

@keyframes fade {
  from {
    opacity: 0;
    transform: translateY(6px);
  }
  to {
    opacity: 1;
    transform: none;
  }
}

.pp-head {
  display: flex;
  justify-content: space-between;
  align-items: baseline;
  margin-bottom: 12px;
}

.pp-label {
  font-size: 13px;
  font-weight: 700;
  letter-spacing: 0.3px;
}

.pp-title {
  display: inline-flex;
  align-items: baseline;
  gap: 10px;
  min-width: 0;
}

.pp-elapsed {
  font-family: var(--font-num);
  font-size: 12px;
  font-weight: 650;
  color: var(--text-soft);
  font-variant-numeric: tabular-nums;
  white-space: nowrap;
}

.pp-percent {
  font-family: var(--font-num);
  font-size: 20px;
  font-weight: 800;
  color: var(--primary);
  font-variant-numeric: tabular-nums;
}

.pp-percent i {
  font-style: normal;
  font-size: 13px;
  color: var(--text-faint);
  margin-left: 1px;
}

.track {
  height: 8px;
  background: var(--surface-2);
  border: 1px solid var(--border);
  border-radius: 999px;
  overflow: hidden;
}

.bar {
  height: 100%;
  border-radius: 999px;
  background: linear-gradient(90deg, var(--primary), var(--primary-strong));
  transition: width 0.45s var(--ease-out);
  box-shadow: 0 0 0 1px rgba(13, 138, 122, 0.15);
}

.pp-msg {
  margin: 12px 0 0;
  font-size: 12.5px;
  color: var(--text-soft);
}
</style>
