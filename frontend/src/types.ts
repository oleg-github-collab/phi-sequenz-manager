export type RunMode = 'headless' | 'visible'
export type JobStatus = 'pending' | 'running' | 'success' | 'failed'
export type ValidationLevel = 'error' | 'warning' | 'info'
export type FileKind = 'file' | 'directory'

export interface NamedProfile {
  name: string
  source_workbook_path: string
  bas_folder_path: string
  dat_file_path: string
  output_dir: string
  macro_name: string
  preferred_mode: RunMode
}

export interface WatchSettings {
  enabled: boolean
  watch_folder: string
  debounce_ms: number
  auto_run_mode: RunMode
  pattern: string
}

export interface AppSettings {
  source_workbook_path: string
  bas_folder_path: string
  dat_file_path: string
  output_dir: string
  generated_workbook_path: string
  macro_name: string
  preferred_mode: RunMode
  auto_rebuild_workbook: boolean
  profiles: NamedProfile[]
  watch: WatchSettings
}

export interface ValidationIssue {
  level: ValidationLevel
  code: string
  message: string
}

export interface ValidationStats {
  row_count: number
  header_count: number
  k_sample_count: number
  p_sample_count: number
  duplicate_sample_count: number
  phi50_candidates: number
}

export interface ValidationSummary {
  ok: boolean
  workbook_exists: boolean
  bas_modules_found: boolean
  dat_exists: boolean
  headers: string[]
  issues: ValidationIssue[]
  stats: ValidationStats
}

export interface RunArtifact {
  label: string
  path: string
}

export interface JobPreview {
  sheet_name: string
  headers: string[]
  rows: string[][]
}

export interface RunRequest {
  source_workbook_path: string
  bas_folder_path: string
  dat_file_path: string
  output_dir: string
  macro_name: string
  mode: RunMode
  force_rebuild_workbook: boolean
}

export interface JobRecord {
  id: string
  status: JobStatus
  title: string
  created_at: string
  started_at: string | null
  finished_at: string | null
  request: RunRequest
  validation: ValidationSummary
  message: string
  artifacts: RunArtifact[]
  preview: JobPreview | null
  step_log: string[]
}

export interface DashboardResponse {
  settings: AppSettings
  latest_validation: ValidationSummary | null
  jobs: JobRecord[]
  active_watch: boolean
}

export interface LogEntry {
  timestamp: string
  event_type: string
  message: string
}

export interface FileSystemEntry {
  name: string
  path: string
  kind: FileKind
  size_bytes: number
  modified_at: string | null
}

export interface DirectoryListing {
  requested_path: string
  exists: boolean
  entries: FileSystemEntry[]
}

export interface ApiEnvelope<T> {
  ok: boolean
  data: T
  validation?: ValidationSummary
  error?: string
}
