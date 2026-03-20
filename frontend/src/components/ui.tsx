import {
  CircleAlert,
  CircleCheckBig,
  CircleDashed,
  Clock3,
  FolderOpen,
  LoaderCircle,
} from 'lucide-react'
import type { ReactNode } from 'react'
import type { JobRecord, ValidationIssue, ValidationSummary } from '../types'

/* ────────── Status Badges ────────── */

export function JobStatusBadge({ status }: { status: JobRecord['status'] }) {
  const icon =
    status === 'success' ? (
      <CircleCheckBig size={14} />
    ) : status === 'failed' ? (
      <CircleAlert size={14} />
    ) : status === 'running' ? (
      <LoaderCircle size={14} className="spin" />
    ) : (
      <Clock3 size={14} />
    )
  return (
    <span className={`status-badge ${status}`}>
      {icon}
      {statusLabel(status)}
    </span>
  )
}

function statusLabel(status: JobRecord['status']) {
  switch (status) {
    case 'pending':
      return 'Wartend'
    case 'running':
      return 'Läuft'
    case 'success':
      return 'Erfolg'
    case 'failed':
      return 'Fehler'
    default:
      return status
  }
}

/* ────────── Cards ────────── */

export function MetricCard(props: {
  icon: ReactNode
  label: string
  value: string
  tone?: 'success' | 'neutral' | 'info' | 'warning'
}) {
  return (
    <div className={`metric-card ${props.tone ?? 'neutral'}`}>
      <div className="metric-icon">{props.icon}</div>
      <span className="metric-label">{props.label}</span>
      <strong className="metric-value">{props.value}</strong>
    </div>
  )
}

export function StatBox(props: {
  label: string
  value: string
  icon: ReactNode
  tone?: 'default' | 'critical'
}) {
  return (
    <div className={`stat-box ${props.tone === 'critical' ? 'critical' : ''}`}>
      <div className="metric-icon">{props.icon}</div>
      <span>{props.label}</span>
      <strong>{props.value}</strong>
    </div>
  )
}

export function InfoTile(props: { title: string; value: string }) {
  return (
    <div className="info-tile">
      <span className="text-muted">{props.title}</span>
      <strong>{props.value}</strong>
    </div>
  )
}

export function FeatureTile(props: {
  icon: ReactNode
  title: string
  text: string
}) {
  return (
    <div className="feature-tile">
      <div className="metric-icon">{props.icon}</div>
      <strong>{props.title}</strong>
      <p className="text-muted">{props.text}</p>
    </div>
  )
}

/* ────────── Checks & Validation ────────── */

export function CheckItem(props: { ok: boolean; text: string }) {
  return (
    <div className={`check-item ${props.ok ? 'ok' : 'off'}`}>
      {props.ok ? <CircleCheckBig size={16} /> : <CircleDashed size={16} />}
      <span>{props.text}</span>
    </div>
  )
}

export function IssueCard(props: { issue: ValidationIssue }) {
  return (
    <div className={`issue-card ${props.issue.level}`}>
      <strong>{props.issue.code}</strong>
      <p>{props.issue.message}</p>
    </div>
  )
}

export function StatusSummary(props: {
  validation: ValidationSummary | null
}) {
  if (!props.validation) {
    return (
      <div className="empty-state compact">
        Noch keine Validierung vorhanden.
      </div>
    )
  }

  return (
    <div className="issue-stack">
      {props.validation.issues.length === 0 ? (
        <div className="empty-state compact">Keine Hinweise vorhanden.</div>
      ) : (
        props.validation.issues.map((issue) => (
          <IssueCard
            key={`${issue.code}-${issue.message}`}
            issue={issue}
          />
        ))
      )}
    </div>
  )
}

/* ────────── Form Controls ────────── */

export function Field(props: {
  label: string
  value: string
  onChange: (v: string) => void
  placeholder?: string
}) {
  return (
    <label className="form-field">
      <span className="field-label">{props.label}</span>
      <input
        className="text-input"
        value={props.value}
        onChange={(e) => props.onChange(e.target.value)}
        placeholder={props.placeholder}
      />
    </label>
  )
}

export function SelectField(props: {
  label: string
  value: string
  onChange: (v: string) => void
  options: Array<{ value: string; label: string }>
}) {
  return (
    <label className="form-field">
      <span className="field-label">{props.label}</span>
      <select
        className="text-input"
        value={props.value}
        onChange={(e) => props.onChange(e.target.value)}
      >
        {props.options.map((o) => (
          <option key={o.value} value={o.value}>
            {o.label}
          </option>
        ))}
      </select>
    </label>
  )
}

export function ToggleChip(props: {
  label: string
  checked: boolean
  onChange: (v: boolean) => void
}) {
  return (
    <label className={`toggle-chip ${props.checked ? 'checked' : ''}`}>
      <input
        type="checkbox"
        checked={props.checked}
        onChange={(e) => props.onChange(e.target.checked)}
      />
      <span>{props.label}</span>
    </label>
  )
}

/* ────────── Misc ────────── */

export function InfoPill(props: { label: string; value: string }) {
  return (
    <div className="info-pill">
      <span className="text-muted">{props.label}</span>
      <strong>{props.value}</strong>
    </div>
  )
}

export function OpenFolderButton(props: {
  label: string
  path: string
  onOpen: (path: string) => void
}) {
  return (
    <button className="btn-ghost btn-sm" onClick={() => props.onOpen(props.path)}>
      <FolderOpen size={14} />
      {props.label}
    </button>
  )
}

/* ────────── Step Indicator ────────── */

export function StepBar(props: {
  steps: string[]
  current: number
}) {
  return (
    <div className="step-bar">
      {props.steps.map((label, idx) => {
        const done = idx < props.current
        const active = idx === props.current
        return (
          <div
            key={label}
            className={`step-item${done ? ' done' : ''}${active ? ' active' : ''}`}
          >
            <div className="step-circle">
              {done ? <CircleCheckBig size={16} /> : idx + 1}
            </div>
            <span className="step-label">{label}</span>
          </div>
        )
      })}
    </div>
  )
}

/* ────────── Helpers ────────── */

export function trimPath(path?: string) {
  if (!path) return 'Nicht gesetzt'
  if (path.length <= 52) return path
  return `${path.slice(0, 22)} … ${path.slice(-24)}`
}

export function parentPath(path: string) {
  if (!path) return path
  const norm = path.replace(/[\\/]+$/, '')
  const idx = Math.max(norm.lastIndexOf('\\'), norm.lastIndexOf('/'))
  return idx <= 0 ? norm : norm.slice(0, idx)
}

export function formatDateTime(value?: string | null) {
  if (!value) return '—'
  return new Intl.DateTimeFormat('de-DE', {
    dateStyle: 'medium',
    timeStyle: 'short',
  }).format(new Date(value))
}

export function formatBytes(bytes: number) {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
}
