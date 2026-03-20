import {
  Bot,
  CircleAlert,
  Eye,
  FolderOpen,
  ListChecks,
  Play,
  RadioTower,
  ScanSearch,
} from 'lucide-react'
import {
  CheckItem,
  InfoTile,
  IssueCard,
  StepBar,
  StatBox,
} from '../components/ui'
import type { DashboardResponse, ValidationSummary, WatchSettings } from '../types'

const RUN_STEPS = ['Validieren', 'Modus wählen', 'Starten']

export function RunPage(props: {
  dashboard: DashboardResponse | null
  validation: ValidationSummary | null
  onValidate: () => void
  onRunHeadless: () => void
  onRunVisible: () => void
  onToggleWatch: (ws: WatchSettings, enable: boolean) => Promise<void>
  onOpenPath: (path: string) => void
}) {
  const { dashboard, validation } = props
  const issues = validation?.issues ?? []
  const errors = issues.filter((i) => i.level === 'error').length
  const warnings = issues.filter((i) => i.level === 'warning').length
  const ready = validation?.ok ?? false

  const currentStep = !validation ? 0 : ready ? 2 : 1

  return (
    <div className="page-grid animate-fade-in">
      <StepBar steps={RUN_STEPS} current={currentStep} />

      {/* Validation Section */}
      <section className="glass-card animate-slide-up">
        <div className="card-header">
          <div>
            <span className="eyebrow">Schritt 1</span>
            <h2>Konfiguration validieren</h2>
          </div>
          <button className="btn-secondary btn-sm" onClick={props.onValidate}>
            <ScanSearch size={16} />
            Jetzt prüfen
          </button>
        </div>

        {validation ? (
          <>
            <div className="grid-4">
              <StatBox
                label="Datensätze"
                value={String(validation.stats.row_count)}
                icon={<ListChecks size={18} />}
              />
              <StatBox
                label="Warnungen"
                value={String(warnings)}
                icon={<CircleAlert size={18} />}
              />
              <StatBox
                label="Fehler"
                value={String(errors)}
                icon={<CircleAlert size={18} />}
                tone={errors > 0 ? 'critical' : 'default'}
              />
              <StatBox
                label="phI-50"
                value={String(validation.stats.phi50_candidates)}
                icon={<Bot size={18} />}
              />
            </div>

            <div className="grid-2">
              <CheckItem
                ok={validation.workbook_exists}
                text="Excel-Quelldatei gefunden"
              />
              <CheckItem
                ok={validation.bas_modules_found}
                text="Alle BAS-Module vorhanden"
              />
              <CheckItem
                ok={validation.dat_exists}
                text="DAT-Datei vorhanden"
              />
              <CheckItem
                ok={validation.ok}
                text="Keine kritischen Fehler"
              />
            </div>

            {issues.length > 0 && (
              <div className="issue-stack">
                {issues.slice(0, 6).map((issue) => (
                  <IssueCard
                    key={`${issue.code}-${issue.message}`}
                    issue={issue}
                  />
                ))}
              </div>
            )}
          </>
        ) : (
          <div className="empty-state">
            Klicken Sie auf „Jetzt prüfen", um die Validierung zu starten.
          </div>
        )}
      </section>

      {/* Run Controls */}
      <section className="glass-card animate-slide-up">
        <div className="card-header">
          <div>
            <span className="eyebrow">Schritt 2 & 3</span>
            <h2>Lauf starten</h2>
          </div>
          <span className={`status-badge ${ready ? 'success' : 'warning'}`}>
            {ready ? 'Bereit' : 'Prüfung erforderlich'}
          </span>
        </div>

        <p className="text-muted card-desc">
          Wählen Sie den gewünschten Ausführungsmodus. Im Headless-Modus
          läuft Excel unsichtbar im Hintergrund. Im sichtbaren Modus
          können Sie den Ablauf direkt in Excel verfolgen.
        </p>

        <div className="run-mode-grid">
          <button
            className="mode-card"
            onClick={props.onRunHeadless}
            disabled={!ready}
          >
            <Play size={32} />
            <strong>Headless starten</strong>
            <p>
              Excel läuft unsichtbar. Ideal für die automatisierte
              Verarbeitung ohne Unterbrechung.
            </p>
          </button>
          <button
            className="mode-card"
            onClick={props.onRunVisible}
            disabled={!ready}
          >
            <Eye size={32} />
            <strong>Sichtbar starten</strong>
            <p>
              Excel wird geöffnet und der Ablauf ist live sichtbar.
              Ideal zur Kontrolle und Fehlersuche.
            </p>
          </button>
        </div>

        {!ready && (
          <p className="text-muted text-sm" style={{ textAlign: 'center' }}>
            Die Ausführung ist erst möglich, wenn die Validierung ohne
            kritische Fehler bestanden wurde.
          </p>
        )}
      </section>

      {/* Watchdog */}
      <section className="glass-card">
        <div className="card-header">
          <div>
            <span className="eyebrow">Automatik</span>
            <h2>Ordnerüberwachung (Watchdog)</h2>
          </div>
          <span
            className={`status-badge ${dashboard?.active_watch ? 'success' : 'neutral'}`}
          >
            <RadioTower size={14} />
            {dashboard?.active_watch ? 'Aktiv' : 'Gestoppt'}
          </span>
        </div>

        <div className="grid-2">
          <InfoTile
            title="Watch-Ordner"
            value={dashboard?.settings.watch.watch_folder || '—'}
          />
          <InfoTile
            title="Muster"
            value={dashboard?.settings.watch.pattern || '*.dat'}
          />
          <InfoTile
            title="Verzögerung"
            value={`${dashboard?.settings.watch.debounce_ms ?? 0} ms`}
          />
          <InfoTile
            title="Modus"
            value={
              dashboard?.settings.watch.auto_run_mode === 'visible'
                ? 'Sichtbar'
                : 'Headless'
            }
          />
        </div>

        <div className="card-footer">
          <button
            className="btn-secondary"
            onClick={() =>
              dashboard &&
              props.onToggleWatch(
                dashboard.settings.watch,
                !dashboard.active_watch,
              )
            }
          >
            <RadioTower size={16} />
            {dashboard?.active_watch
              ? 'Watchdog stoppen'
              : 'Watchdog starten'}
          </button>
          <div className="btn-row">
            <button
              className="btn-ghost btn-sm"
              onClick={() =>
                dashboard && props.onOpenPath(dashboard.settings.output_dir)
              }
            >
              <FolderOpen size={14} />
              Ausgabeordner
            </button>
          </div>
        </div>
      </section>
    </div>
  )
}
