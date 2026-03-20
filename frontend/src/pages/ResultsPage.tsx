import { Clock3, FolderOpen, RefreshCw, ScanSearch } from 'lucide-react'
import {
  formatBytes,
  formatDateTime,
  JobStatusBadge,
  parentPath,
} from '../components/ui'
import type { DirectoryListing, JobRecord } from '../types'

export function ResultsPage(props: {
  jobs: JobRecord[]
  selectedJob: JobRecord | null
  selectedJobId: string | null
  onSelectJob: (id: string) => void
  files: DirectoryListing | null
  fileBrowserPath: string
  onChangeBrowserPath: (p: string) => void
  onRefreshFiles: (p: string) => void
  onOpenPath: (p: string) => void
}) {
  const job = props.selectedJob

  return (
    <div className="page-grid animate-fade-in">
      {/* Job List + Preview in 2-column layout */}
      <div className="results-layout">
        {/* Left: Job list */}
        <section className="glass-card">
          <div className="card-header">
            <div>
              <span className="eyebrow">Laufhistorie</span>
              <h2>Abgeschlossene Jobs</h2>
            </div>
            <span className="status-badge neutral">
              <Clock3 size={14} />
              {props.jobs.length} Einträge
            </span>
          </div>
          <div className="job-list">
            {props.jobs.length === 0 ? (
              <div className="empty-state compact">
                Noch keine Jobs vorhanden.
              </div>
            ) : (
              props.jobs.map((j) => (
                <button
                  key={j.id}
                  className={`job-card${props.selectedJobId === j.id ? ' selected' : ''}`}
                  onClick={() => props.onSelectJob(j.id)}
                >
                  <div className="job-card-head">
                    <strong>{j.title}</strong>
                    <JobStatusBadge status={j.status} />
                  </div>
                  <p className="text-muted text-sm">{j.message}</p>
                  <span className="text-muted text-xs">
                    {formatDateTime(j.finished_at || j.created_at)}
                  </span>
                </button>
              ))
            )}
          </div>
        </section>

        {/* Right: Preview */}
        <section className="glass-card">
          <div className="card-header">
            <div>
              <span className="eyebrow">Vorschau</span>
              <h2>Ergebnistabelle</h2>
            </div>
            {job && <JobStatusBadge status={job.status} />}
          </div>

          {job?.preview ? (
            <>
              <div className="preview-meta">
                <strong>{job.preview.sheet_name}</strong>
                <span className="text-muted">{job.message}</span>
              </div>
              <div className="table-wrap">
                <table className="data-table">
                  <thead>
                    <tr>
                      {job.preview.headers.map((h) => (
                        <th key={h}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {job.preview.rows.map((row, ri) => (
                      <tr key={`${job.id}-${ri}`}>
                        {row.map((cell, ci) => (
                          <td key={`${ri}-${ci}`}>{cell}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="btn-row" style={{ marginTop: '1rem' }}>
                {job.artifacts.map((a) => (
                  <button
                    key={`${a.label}-${a.path}`}
                    className="btn-ghost btn-sm"
                    onClick={() => props.onOpenPath(parentPath(a.path))}
                  >
                    <FolderOpen size={14} />
                    {a.label}
                  </button>
                ))}
              </div>
            </>
          ) : (
            <div className="empty-state">
              Wählen Sie links einen Job aus, um die Vorschau zu sehen.
            </div>
          )}
        </section>
      </div>

      {/* File Browser */}
      <section className="glass-card">
        <div className="card-header">
          <div>
            <span className="eyebrow">Dateibrowser</span>
            <h2>Artefakte und Ausgabe</h2>
          </div>
          <button
            className="btn-ghost btn-sm"
            onClick={() => props.onRefreshFiles(props.fileBrowserPath)}
          >
            <RefreshCw size={14} />
            Aktualisieren
          </button>
        </div>

        <div className="browser-bar">
          <input
            className="text-input"
            value={props.fileBrowserPath}
            onChange={(e) => props.onChangeBrowserPath(e.target.value)}
            placeholder="Pfad eingeben …"
          />
          <button
            className="btn-secondary btn-sm"
            onClick={() => props.onRefreshFiles(props.fileBrowserPath)}
          >
            <ScanSearch size={14} />
            Laden
          </button>
          <button
            className="btn-ghost btn-sm"
            onClick={() => props.onOpenPath(props.fileBrowserPath)}
          >
            <FolderOpen size={14} />
            Öffnen
          </button>
        </div>

        <div className="file-list">
          {props.files?.entries.length ? (
            props.files.entries.map((entry) => (
              <div key={entry.path} className="file-row">
                <div>
                  <strong>{entry.name}</strong>
                  <span className="text-muted text-sm">
                    {entry.kind === 'directory'
                      ? 'Ordner'
                      : formatBytes(entry.size_bytes)}
                  </span>
                </div>
                <div className="btn-row">
                  <span
                    className={`status-badge ${entry.kind === 'directory' ? 'info' : 'neutral'}`}
                  >
                    {entry.kind === 'directory' ? 'Ordner' : 'Datei'}
                  </span>
                  <button
                    className="btn-ghost btn-sm"
                    onClick={() =>
                      props.onOpenPath(
                        entry.kind === 'directory'
                          ? entry.path
                          : parentPath(entry.path),
                      )
                    }
                  >
                    <FolderOpen size={14} />
                    Öffnen
                  </button>
                </div>
              </div>
            ))
          ) : (
            <div className="empty-state compact">
              Keine Dateien für diesen Pfad gefunden.
            </div>
          )}
        </div>
      </section>
    </div>
  )
}
