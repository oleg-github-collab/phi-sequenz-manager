import {
  CircleCheckBig,
  RadioTower,
  ScanSearch,
  Settings2,
} from 'lucide-react'
import { FeatureTile, formatDateTime } from '../components/ui'
import type { LogEntry } from '../types'

export function LogsPage(props: { logs: LogEntry[] }) {
  return (
    <div className="page-grid animate-fade-in">
      {/* Event Log */}
      <section className="glass-card">
        <div className="card-header">
          <div>
            <span className="eyebrow">App-Ereignisse</span>
            <h2>Protokoll</h2>
          </div>
          <span className="status-badge neutral">{props.logs.length} Zeilen</span>
        </div>

        <div className="table-wrap">
          <table className="data-table">
            <thead>
              <tr>
                <th>Zeitpunkt</th>
                <th>Typ</th>
                <th>Nachricht</th>
              </tr>
            </thead>
            <tbody>
              {props.logs.length === 0 ? (
                <tr>
                  <td colSpan={3} className="text-muted" style={{ textAlign: 'center' }}>
                    Noch keine Ereignisse vorhanden.
                  </td>
                </tr>
              ) : (
                props.logs.slice(0, 120).map((e) => (
                  <tr key={`${e.timestamp}-${e.event_type}-${e.message}`}>
                    <td>{formatDateTime(e.timestamp)}</td>
                    <td>
                      <span className="status-badge neutral">{e.event_type}</span>
                    </td>
                    <td>{e.message}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </section>

      {/* Feature Info */}
      <section className="glass-card">
        <div className="card-header">
          <div>
            <span className="eyebrow">Betriebshinweise</span>
            <h2>Was diese Runtime auszeichnet</h2>
          </div>
        </div>

        <div className="grid-2">
          <FeatureTile
            icon={<CircleCheckBig size={18} />}
            title="Dialogfreie Automatisierung"
            text="VBA-Module werden für Headless-Läufe gepatcht, damit MsgBox-Dialoge keine unsichtbaren Excel-Sitzungen blockieren."
          />
          <FeatureTile
            icon={<Settings2 size={18} />}
            title="Reproduzierbarer Build"
            text="Das Runtime-Workbook wird nur dann neu gebaut, wenn sich Quelldatei, BAS-Stand oder DAT-Pfad geändert haben."
          />
          <FeatureTile
            icon={<ScanSearch size={18} />}
            title="Maschinenlesbarer Report"
            text="Jeder Lauf erzeugt einen JSON-Report mit Artefakten, Status und tabellarischer Vorschau."
          />
          <FeatureTile
            icon={<RadioTower size={18} />}
            title="Watchdog für neue Dateien"
            text="Neue DAT-Dateien im Inbox-Ordner werden automatisch erkannt und in einen Lauf überführt."
          />
        </div>
      </section>
    </div>
  )
}
