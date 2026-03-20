import {
  Activity,
  ChevronRight,
  FileText,
  Home,
  LoaderCircle,
  Play,
  RefreshCw,
  ScrollText,
  Settings,
  X,
} from 'lucide-react'
import { startTransition, useEffect, useEffectEvent, useState } from 'react'
import { NavLink, Route, Routes, useNavigate } from 'react-router-dom'
import { SceneCanvas } from './components/SceneCanvas'
import {
  fetchDashboard,
  fetchFiles,
  fetchLogs,
  fetchValidation,
  openFolder,
  runPipeline,
  saveSettings,
  startWatch,
  stopWatch,
} from './lib/api'
import { HomePage } from './pages/HomePage'
import { SetupPage } from './pages/SetupPage'
import { RunPage } from './pages/RunPage'
import { ResultsPage } from './pages/ResultsPage'
import { LogsPage } from './pages/LogsPage'
import type {
  AppSettings,
  DashboardResponse,
  DirectoryListing,
  LogEntry,
  RunMode,
  RunRequest,
  ValidationSummary,
  WatchSettings,
} from './types'

const navigation = [
  { to: '/', label: 'Startseite', icon: Home },
  { to: '/einrichten', label: 'Einrichtung', icon: Settings },
  { to: '/ausfuehren', label: 'Ausführung', icon: Play },
  { to: '/ergebnisse', label: 'Ergebnisse', icon: Activity },
  { to: '/protokoll', label: 'Protokoll', icon: ScrollText },
] as const

type ToastTone = 'info' | 'success' | 'warning' | 'error'

interface Toast {
  tone: ToastTone
  text: string
}

export default function App() {
  const navigate = useNavigate()
  const [dashboard, setDashboard] = useState<DashboardResponse | null>(null)
  const [logs, setLogs] = useState<LogEntry[]>([])
  const [files, setFiles] = useState<DirectoryListing | null>(null)
  const [validation, setValidation] = useState<ValidationSummary | null>(null)
  const [processing, setProcessing] = useState<string | null>(null)
  const [toast, setToast] = useState<Toast | null>(null)
  const [fileBrowserPath, setFileBrowserPath] = useState('')
  const [selectedJobId, setSelectedJobId] = useState<string | null>(null)
  const [initialLoading, setInitialLoading] = useState(true)

  const latestJob = dashboard?.jobs[0] ?? null
  const selectedJob =
    dashboard?.jobs.find((j) => j.id === selectedJobId) ?? latestJob ?? null

  const refreshAll = useEffectEvent(async () => {
    try {
      const [dash, logData] = await Promise.all([
        fetchDashboard(),
        fetchLogs(),
      ])
      const nextPath = fileBrowserPath || dash.settings.output_dir
      const fileData = await fetchFiles(nextPath)

      startTransition(() => {
        setDashboard(dash)
        setLogs(logData)
        setFiles(fileData)
        setValidation(dash.latest_validation)
        if (!fileBrowserPath) setFileBrowserPath(nextPath)
        if (!selectedJobId && dash.jobs[0]) setSelectedJobId(dash.jobs[0].id)
      })
    } catch (err) {
      setToast({
        tone: 'error',
        text:
          err instanceof Error
            ? err.message
            : 'Verbindung zum Backend fehlgeschlagen.',
      })
    } finally {
      setInitialLoading(false)
    }
  })

  useEffect(() => {
    void refreshAll()
    const id = window.setInterval(() => void refreshAll(), 8000)
    return () => window.clearInterval(id)
  }, [])

  const runWithMode = async (mode: RunMode) => {
    if (!dashboard) return
    const req: RunRequest = {
      source_workbook_path: dashboard.settings.source_workbook_path,
      bas_folder_path: dashboard.settings.bas_folder_path,
      dat_file_path: dashboard.settings.dat_file_path,
      output_dir: dashboard.settings.output_dir,
      macro_name: dashboard.settings.macro_name,
      mode,
      force_rebuild_workbook: dashboard.settings.auto_rebuild_workbook,
    }
    setProcessing(
      mode === 'headless'
        ? 'Headless-Lauf wird ausgeführt …'
        : 'Sichtbarer Lauf wird ausgeführt …',
    )
    try {
      const job = await runPipeline(req)
      startTransition(() => setSelectedJobId(job.id))
      await refreshAll()
      setToast({
        tone: job.status === 'success' ? 'success' : 'warning',
        text: job.message,
      })
      if (job.status === 'success') navigate('/ergebnisse')
    } catch (err) {
      setToast({
        tone: 'error',
        text:
          err instanceof Error
            ? err.message
            : 'Der Lauf konnte nicht gestartet werden.',
      })
    } finally {
      setProcessing(null)
    }
  }

  const triggerValidation = async () => {
    setProcessing('Validierung läuft …')
    try {
      const v = await fetchValidation()
      startTransition(() => setValidation(v))
      setToast({
        tone: v.ok ? 'success' : 'warning',
        text: v.ok
          ? 'Alle Prüfungen bestanden.'
          : 'Es wurden Hinweise oder Fehler gefunden.',
      })
    } catch (err) {
      setToast({
        tone: 'error',
        text:
          err instanceof Error
            ? err.message
            : 'Validierung fehlgeschlagen.',
      })
    } finally {
      setProcessing(null)
    }
  }

  const handleSave = async (settings: AppSettings) => {
    setProcessing('Einstellungen werden gespeichert …')
    try {
      const res = await saveSettings(settings)
      startTransition(() => {
        setValidation(res.validation)
        if (dashboard) setDashboard({ ...dashboard, settings: res.settings })
        setFileBrowserPath(res.settings.output_dir)
      })
      await refreshAll()
      setToast({ tone: 'success', text: 'Einstellungen gespeichert.' })
    } catch (err) {
      setToast({
        tone: 'error',
        text:
          err instanceof Error
            ? err.message
            : 'Speichern fehlgeschlagen.',
      })
    } finally {
      setProcessing(null)
    }
  }

  const handleOpen = async (path: string) => {
    try {
      await openFolder(path)
    } catch {
      setToast({ tone: 'error', text: 'Pfad konnte nicht geöffnet werden.' })
    }
  }

  const handleToggleWatch = async (ws: WatchSettings, enable: boolean) => {
    setProcessing(enable ? 'Watchdog wird gestartet …' : 'Watchdog wird gestoppt …')
    try {
      if (enable) await startWatch(ws)
      else await stopWatch()
      await refreshAll()
      setToast({
        tone: 'success',
        text: enable
          ? 'Ordnerüberwachung aktiviert.'
          : 'Ordnerüberwachung gestoppt.',
      })
    } catch (err) {
      setToast({
        tone: 'error',
        text: err instanceof Error ? err.message : 'Watchdog-Aktion fehlgeschlagen.',
      })
    } finally {
      setProcessing(null)
    }
  }

  const refreshFileBrowser = async (path: string) => {
    try {
      const listing = await fetchFiles(path)
      startTransition(() => {
        setFiles(listing)
        setFileBrowserPath(path)
      })
    } catch {
      setToast({ tone: 'error', text: 'Dateiansicht konnte nicht geladen werden.' })
    }
  }

  /* ─── Splash / initial loading with UnicornStudio ─── */
  if (initialLoading) {
    return (
      <div className="splash-screen">
        <div className="splash-scene">
          <SceneCanvas />
        </div>
        <div className="splash-content animate-fade-in">
          <h1>PHI Sequenz-Manager</h1>
          <p>Verbindung wird hergestellt …</p>
          <div className="splash-loader">
            <LoaderCircle size={28} className="spin" />
          </div>
        </div>
      </div>
    )
  }

  return (
    <div className="app-shell">
      {/* ─── Sidebar ─── */}
      <aside className="sidebar">
        <div className="sidebar-brand">
          <div className="brand-mark">PHI</div>
          <div>
            <div className="brand-name">Sequenz-Manager</div>
            <div className="brand-sub">Hybrid Runtime</div>
          </div>
        </div>

        <nav className="sidebar-nav">
          {navigation.map((entry) => {
            const Icon = entry.icon
            return (
              <NavLink
                key={entry.to}
                to={entry.to}
                end={entry.to === '/'}
                className={({ isActive }) =>
                  `nav-item${isActive ? ' active' : ''}`
                }
              >
                <Icon size={18} />
                <span>{entry.label}</span>
                <ChevronRight size={14} className="nav-arrow" />
              </NavLink>
            )
          })}
        </nav>

        <div className="sidebar-footer">
          <button className="btn-ghost btn-sm" onClick={() => void refreshAll()}>
            <RefreshCw size={14} />
            Aktualisieren
          </button>
          <div className="sidebar-status">
            {processing ? (
              <span className="status-badge running">
                <LoaderCircle size={12} className="spin" />
                Aktiv
              </span>
            ) : (
              <span className="status-badge success">Bereit</span>
            )}
          </div>
        </div>
      </aside>

      {/* ─── Main Content ─── */}
      <main className="main-panel">
        {/* Toast notification */}
        {toast && (
          <div className={`toast toast-${toast.tone} animate-slide-down`}>
            <span>{toast.text}</span>
            <button className="toast-close" onClick={() => setToast(null)}>
              <X size={14} />
            </button>
          </div>
        )}

        <div className="page-content">
          <Routes>
            <Route
              path="/"
              element={
                <HomePage
                  dashboard={dashboard}
                  validation={validation}
                  latestJob={latestJob}
                  onNavigate={(path) => navigate(path)}
                />
              }
            />
            <Route
              path="/einrichten"
              element={
                <SetupPage
                  settings={dashboard?.settings ?? null}
                  validation={validation}
                  onSave={handleSave}
                  onOpenPath={handleOpen}
                  onValidate={() => void triggerValidation()}
                  onToggleWatch={handleToggleWatch}
                />
              }
            />
            <Route
              path="/ausfuehren"
              element={
                <RunPage
                  dashboard={dashboard}
                  validation={validation}
                  onValidate={() => void triggerValidation()}
                  onRunHeadless={() => void runWithMode('headless')}
                  onRunVisible={() => void runWithMode('visible')}
                  onToggleWatch={handleToggleWatch}
                  onOpenPath={handleOpen}
                />
              }
            />
            <Route
              path="/ergebnisse"
              element={
                <ResultsPage
                  jobs={dashboard?.jobs ?? []}
                  selectedJob={selectedJob}
                  selectedJobId={selectedJobId}
                  onSelectJob={setSelectedJobId}
                  files={files}
                  fileBrowserPath={fileBrowserPath}
                  onChangeBrowserPath={setFileBrowserPath}
                  onRefreshFiles={(p) => void refreshFileBrowser(p)}
                  onOpenPath={handleOpen}
                />
              }
            />
            <Route path="/protokoll" element={<LogsPage logs={logs} />} />
          </Routes>
        </div>
      </main>

      {/* ─── Processing Overlay with UnicornStudio ─── */}
      {processing && (
        <div className="processing-overlay animate-fade-in">
          <div className="processing-scene">
            <SceneCanvas />
          </div>
          <div className="processing-card animate-slide-up">
            <LoaderCircle size={36} className="spin processing-spinner" />
            <h2>Verarbeitung läuft</h2>
            <p className="processing-detail">{processing}</p>
            <div className="progress-bar">
              <div className="progress-fill shimmer" />
            </div>
            <p className="text-muted text-sm">
              Bitte warten Sie, bis der Vorgang abgeschlossen ist.
            </p>
          </div>
        </div>
      )}
    </div>
  )
}
