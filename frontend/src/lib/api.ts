import type {
  ApiEnvelope,
  AppSettings,
  DashboardResponse,
  DirectoryListing,
  JobRecord,
  LogEntry,
  RunRequest,
  ValidationSummary,
  WatchSettings,
} from '../types'

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(path, {
    headers: {
      'Content-Type': 'application/json',
      ...(init?.headers ?? {}),
    },
    ...init,
  })

  const payload = (await response.json()) as ApiEnvelope<T>
  if (!response.ok || !payload.ok) {
    throw new Error(payload.error || 'Die Anfrage konnte nicht verarbeitet werden.')
  }

  return payload.data
}

export async function fetchDashboard() {
  return request<DashboardResponse>('/api/dashboard')
}

export async function fetchValidation() {
  return request<ValidationSummary>('/api/validate')
}

export async function validateRequest(runRequest: RunRequest) {
  return request<ValidationSummary>('/api/validate', {
    method: 'POST',
    body: JSON.stringify(runRequest),
  })
}

export async function saveSettings(settings: AppSettings) {
  const response = await fetch('/api/settings', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(settings),
  })

  const payload = (await response.json()) as ApiEnvelope<AppSettings>
  if (!response.ok || !payload.ok) {
    throw new Error(payload.error || 'Einstellungen konnten nicht gespeichert werden.')
  }

  return {
    settings: payload.data,
    validation: payload.validation ?? null,
  }
}

export async function runPipeline(runRequest: RunRequest) {
  return request<JobRecord>('/api/run', {
    method: 'POST',
    body: JSON.stringify(runRequest),
  })
}

export async function fetchJobs() {
  return request<JobRecord[]>('/api/jobs')
}

export async function fetchLogs() {
  return request<LogEntry[]>('/api/logs')
}

export async function fetchFiles(path?: string) {
  const suffix = path ? `?path=${encodeURIComponent(path)}` : ''
  return request<DirectoryListing>(`/api/files${suffix}`)
}

export async function openFolder(path: string) {
  await request<unknown>('/api/open-folder', {
    method: 'POST',
    body: JSON.stringify({ path }),
  })
}

export async function startWatch(watch: WatchSettings) {
  await request<unknown>('/api/watch/start', {
    method: 'POST',
    body: JSON.stringify({
      watch_folder: watch.watch_folder,
      pattern: watch.pattern,
      debounce_ms: watch.debounce_ms,
      mode: watch.auto_run_mode,
    }),
  })
}

export async function stopWatch() {
  await request<unknown>('/api/watch/stop', {
    method: 'POST',
  })
}
