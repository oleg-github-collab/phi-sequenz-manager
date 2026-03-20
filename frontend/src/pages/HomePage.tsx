import {
  ArrowRight,
  BellRing,
  Boxes,
  CircleCheckBig,
  Cpu,
  FileText,
  Gauge,
  Zap,
} from 'lucide-react'
import { CheckItem, formatDateTime, InfoPill, JobStatusBadge, MetricCard, trimPath } from '../components/ui'
import type { DashboardResponse, JobRecord, ValidationSummary } from '../types'

export function HomePage(props: {
  dashboard: DashboardResponse | null
  validation: ValidationSummary | null
  latestJob: JobRecord | null
  onNavigate: (path: string) => void
}) {
  const { dashboard, validation, latestJob } = props

  return (
    <div className="page-grid animate-fade-in">
      {/* Hero */}
      <section className="hero-section">
        <div className="hero-content">
          <span className="eyebrow">Willkommen</span>
          <h1 className="hero-title">
            PHI Sequenz-Manager
          </h1>
          <p className="hero-desc">
            Ihr Leitstand für die automatisierte Erzeugung von Messsequenzen.
            Konfigurieren Sie Quelldateien, prüfen Sie die Daten und starten
            Sie den Lauf — alles in einer Oberfläche.
          </p>

          <div className="hero-actions">
            <button
              className="btn-primary btn-lg"
              onClick={() => props.onNavigate('/einrichten')}
            >
              <Zap size={18} />
              Jetzt einrichten
              <ArrowRight size={16} />
            </button>
            <button
              className="btn-secondary"
              onClick={() => props.onNavigate('/ausfuehren')}
            >
              <Cpu size={18} />
              Direkt zur Ausführung
            </button>
          </div>
        </div>

        <div className="hero-steps">
          <div className="step-preview">
            <div className="step-preview-item">
              <div className="step-num">1</div>
              <div>
                <strong>Einrichten</strong>
                <p>Quelldateien und Pfade konfigurieren</p>
              </div>
            </div>
            <div className="step-preview-item">
              <div className="step-num">2</div>
              <div>
                <strong>Prüfen</strong>
                <p>Validierung der Konfiguration</p>
              </div>
            </div>
            <div className="step-preview-item">
              <div className="step-num">3</div>
              <div>
                <strong>Ausführen</strong>
                <p>Sequenz automatisch erzeugen</p>
              </div>
            </div>
            <div className="step-preview-item">
              <div className="step-num">4</div>
              <div>
                <strong>Ergebnisse</strong>
                <p>Tabelle und Artefakte einsehen</p>
              </div>
            </div>
          </div>
        </div>
      </section>

      {/* Quick Status */}
      <section className="glass-card">
        <div className="card-header">
          <div>
            <span className="eyebrow">Systemstatus</span>
            <h2>Schnellübersicht</h2>
          </div>
        </div>
        <div className="grid-3">
          <MetricCard
            icon={<CircleCheckBig size={18} />}
            label="Watchdog"
            value={dashboard?.active_watch ? 'Aktiv' : 'Aus'}
            tone={dashboard?.active_watch ? 'success' : 'neutral'}
          />
          <MetricCard
            icon={<Boxes size={18} />}
            label="Jobs"
            value={String(dashboard?.jobs.length ?? 0)}
            tone="info"
          />
          <MetricCard
            icon={<BellRing size={18} />}
            label="Prüfung"
            value={
              validation
                ? validation.ok
                  ? 'Bestanden'
                  : 'Fehler'
                : 'Offen'
            }
            tone={validation?.ok ? 'success' : validation ? 'warning' : 'neutral'}
          />
        </div>
      </section>

      {/* Current Config Summary */}
      <section className="glass-card">
        <div className="card-header">
          <div>
            <span className="eyebrow">Aktuelle Konfiguration</span>
            <h2>Pfade und Einstellungen</h2>
          </div>
          <button
            className="btn-ghost btn-sm"
            onClick={() => props.onNavigate('/einrichten')}
          >
            <FileText size={14} />
            Bearbeiten
          </button>
        </div>
        <div className="grid-2">
          <InfoPill
            label="Excel-Quelldatei"
            value={trimPath(dashboard?.settings.source_workbook_path)}
          />
          <InfoPill
            label="DAT-Datei"
            value={trimPath(dashboard?.settings.dat_file_path)}
          />
          <InfoPill
            label="Makro"
            value={dashboard?.settings.macro_name || 'Nicht gesetzt'}
          />
          <InfoPill
            label="Modus"
            value={
              dashboard?.settings.preferred_mode === 'visible'
                ? 'Sichtbar'
                : 'Headless'
            }
          />
        </div>
      </section>

      {/* Validation Checks */}
      {validation && (
        <section className="glass-card">
          <div className="card-header">
            <div>
              <span className="eyebrow">Vorprüfung</span>
              <h2>Bereitschaftsstatus</h2>
            </div>
            <span
              className={`status-badge ${validation.ok ? 'success' : 'warning'}`}
            >
              {validation.ok ? 'Bereit' : 'Prüfen'}
            </span>
          </div>
          <div className="grid-2">
            <CheckItem
              ok={validation.workbook_exists}
              text="Excel-Quelldatei vorhanden"
            />
            <CheckItem
              ok={validation.bas_modules_found}
              text="Alle BAS-Module gefunden"
            />
            <CheckItem ok={validation.dat_exists} text="DAT-Datei vorhanden" />
            <CheckItem
              ok={validation.ok}
              text="Konfiguration ohne kritische Fehler"
            />
          </div>
        </section>
      )}

      {/* Latest Job */}
      {latestJob && (
        <section className="glass-card">
          <div className="card-header">
            <div>
              <span className="eyebrow">Letzter Lauf</span>
              <h2>{latestJob.title}</h2>
            </div>
            <JobStatusBadge status={latestJob.status} />
          </div>
          <p className="text-muted">{latestJob.message}</p>
          <div className="card-footer">
            <span className="text-muted text-sm">
              {formatDateTime(latestJob.finished_at || latestJob.created_at)}
            </span>
            <button
              className="btn-ghost btn-sm"
              onClick={() => props.onNavigate('/ergebnisse')}
            >
              <Gauge size={14} />
              Ergebnisse ansehen
            </button>
          </div>
        </section>
      )}
    </div>
  )
}
