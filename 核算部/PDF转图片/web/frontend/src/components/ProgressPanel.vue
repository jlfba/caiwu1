<script setup>
import { computed, nextTick, ref, watch } from 'vue'

const props = defineProps({
  status: { type: String, required: true },
  current: { type: Number, default: 0 },
  total: { type: Number, default: 0 },
  message: { type: String, default: '' },
  elapsedSeconds: { type: Number, default: -1 },
  logs: { type: Array, default: () => [] }
})

const terminalBody = ref(null)

watch(
  () => [props.logs.length, props.message, props.current],
  async () => {
    await nextTick()
    if (terminalBody.value) {
      terminalBody.value.scrollTop = terminalBody.value.scrollHeight
    }
  },
  { immediate: true }
)

const percent = computed(() => {
  if (props.status === 'done') return 100
  if (!props.total) return 0
  return Math.min(100, Math.round((props.current / props.total) * 100))
})

const reportStages = [
  '保留应收单价小于 1',
  '删除客户简称关键词',
  '删除业务员华南KA',
  '删除备注J开头或无应收',
  '删除配仓单号刘丹整柜',
  '删除备注免费补发',
  '删除整柜高金额',
  '新增无应收分类',
  '生成无应收明细透视表',
  '清理临时删除记录'
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
    <p v-if="!isReportFlow" class="pp-msg">{{ message }}</p>
    <div v-if="isReportFlow" class="pp-terminal" aria-label="处理日志">
      <div class="pp-terminal-head"><span>全部处理日志</span><span>{{ logs.length }} 条</span></div>
      <div ref="terminalBody" class="pp-terminal-body">
        <p class="terminal-current">[当前状态] {{ message || '等待处理状态…' }}</p>
        <p v-for="(stage, index) in reportStages" :key="`stage-${stage}`" :class="`terminal-stage stage-${stageState(index)}`">
          [步骤 {{ index + 1 }}/5] {{ stage }} — {{ stageState(index) === 'done' ? '已完成' : stageState(index) === 'active' ? '进行中' : '待处理' }}
        </p>
        <p class="terminal-divider">---------------- 实时记录 ----------------</p>
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
.pp-terminal-body { max-height: 360px; min-height: 240px; overflow-y: auto; padding: 12px 14px; scrollbar-width: thin; }
.pp-terminal-body p { margin: 0 0 6px; color: #c8e2df; font-size: 11px; line-height: 1.45; white-space: pre-wrap; overflow-wrap: anywhere; }
.pp-terminal-body p::before { content: '> '; color: #5bd0bd; }
.pp-terminal-empty { color: #78919a !important; }
.pp-terminal-body .terminal-current { color: #8fe3d7; font-weight: 700; }
.pp-terminal-body .terminal-stage { color: #78919a; }
.pp-terminal-body .terminal-stage.stage-active { color: #8fe3d7 !important; }
.pp-terminal-body .terminal-stage.stage-done { color: #b6ccc9 !important; }
.pp-terminal-body .terminal-divider { color: #4f6b72; }
.pp-terminal-body .terminal-divider::before { content: ''; }
</style>
