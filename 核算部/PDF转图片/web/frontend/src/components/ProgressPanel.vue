<script setup>
import { computed } from 'vue'

const props = defineProps({
  status: { type: String, required: true },
  current: { type: Number, default: 0 },
  total: { type: Number, default: 0 },
  message: { type: String, default: '' },
  elapsedSeconds: { type: Number, default: -1 },
  logs: { type: Array, default: () => [] }
})

const percent = computed(() => {
  if (props.status === 'done') return 100
  if (!props.total) return 0
  return Math.min(100, Math.round((props.current / props.total) * 100))
})

const reportStages = [
  '加载工作簿',
  '复制所选工作表',
  '筛选无应收明细',
  '生成无应收明细透视表',
  '保存处理结果'
]

const isReportFlow = computed(() => props.total === reportStages.length)

function stageState(index) {
  if (props.current > index + 1) return 'done'
  if (props.current === index + 1) return 'active'
  return 'pending'
}

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
    <ol v-if="isReportFlow" class="pp-stages" aria-label="报表处理流程">
      <li v-for="(stage, index) in reportStages" :key="stage" :class="`stage-${stageState(index)}`">
        <span class="stage-dot" aria-hidden="true">{{ stageState(index) === 'done' ? '✓' : index + 1 }}</span>
        <span>{{ stage }}</span>
        <span v-if="stageState(index) === 'active'" class="stage-now">进行中</span>
        <span v-else-if="stageState(index) === 'done'" class="stage-now">已完成</span>
      </li>
    </ol>
    <div v-if="isReportFlow" class="pp-terminal" aria-label="处理日志">
      <div class="pp-terminal-head"><span>处理日志</span><span>{{ logs.length }} 条</span></div>
      <div class="pp-terminal-body">
        <p v-for="(line, index) in logs" :key="`${index}-${line}`">{{ line }}</p>
        <p v-if="!logs.length" class="pp-terminal-empty">等待后端日志…</p>
      </div>
    </div>
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

.pp-stages {
  display: grid;
  gap: 8px;
  margin: 18px 0 0;
  padding: 14px 0 0;
  border-top: 1px solid var(--border);
  list-style: none;
}

.pp-stages li {
  display: flex;
  align-items: center;
  gap: 9px;
  min-height: 28px;
  color: var(--text-faint);
  font-size: 12.5px;
}

.stage-dot {
  display: grid;
  place-items: center;
  width: 22px;
  height: 22px;
  border: 1px solid var(--border-strong);
  border-radius: 50%;
  color: var(--text-faint);
  font-family: var(--font-num);
  font-size: 11px;
  font-weight: 700;
}

.stage-active {
  color: var(--primary-ink) !important;
  font-weight: 700;
}

.stage-active .stage-dot {
  border-color: var(--primary);
  background: var(--primary);
  color: #fff;
}

.stage-done {
  color: var(--text-soft) !important;
}

.stage-done .stage-dot {
  border-color: var(--primary);
  background: var(--primary-soft);
  color: var(--primary-ink);
}

.stage-now {
  margin-left: auto;
  color: var(--text-faint);
  font-size: 11px;
  font-weight: 600;
}

.pp-terminal {
  margin-top: 16px;
  overflow: hidden;
  border: 1px solid #263842;
  border-radius: 10px;
  background: #101a20;
  color: #c8e2df;
  font-family: var(--font-num), Consolas, monospace;
}
.pp-terminal-head {
  display: flex; justify-content: space-between; padding: 9px 12px;
  border-bottom: 1px solid #263842; color: #8bcfc6; font-size: 11px; font-weight: 700;
}
.pp-terminal-body { max-height: 150px; overflow-y: auto; padding: 10px 12px; scrollbar-width: thin; }
.pp-terminal-body p { margin: 0 0 6px; color: #c8e2df; font-size: 11px; line-height: 1.45; white-space: pre-wrap; overflow-wrap: anywhere; }
.pp-terminal-body p::before { content: '> '; color: #5bd0bd; }
.pp-terminal-empty { color: #78919a !important; }
</style>
