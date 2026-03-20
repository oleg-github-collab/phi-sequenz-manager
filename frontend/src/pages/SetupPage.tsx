import {
  ArrowLeft,
  ArrowRight,
  FolderOpen,
  Save,
  Upload,
} from 'lucide-react'
import { useEffect, useRef, useState } from 'react'
import type { ChangeEvent } from 'react'
import {
  Field,
  SelectField,
  StepBar,
  StatusSummary,
  ToggleChip,
} from '../components/ui'
import type {
  AppSettings,
  NamedProfile,
  RunMode,
  ValidationSummary,
  WatchSettings,
} from '../types'

const WIZARD_STEPS = ['Quelldateien', 'Messdaten', 'Optionen', 'Prüfung']

export function SetupPage(props: {
  settings: AppSettings | null
  validation: ValidationSummary | null
  onSave: (s: AppSettings) => void
  onOpenPath: (p: string) => void
  onValidate: () => void
  onToggleWatch: (ws: WatchSettings, enable: boolean) => Promise<void>
}) {
  const [step, setStep] = useState(0)
  const [draft, setDraft] = useState<AppSettings | null>(props.settings)
  const importRef = useRef<HTMLInputElement | null>(null)

  useEffect(() => {
    setDraft(props.settings)
  }, [props.settings])

  if (!draft) {
    return (
      <div className="empty-state animate-fade-in">
        Einstellungen werden geladen …
      </div>
    )
  }

  const up = <K extends keyof AppSettings>(k: K, v: AppSettings[K]) =>
    setDraft({ ...draft, [k]: v })

  const upWatch = <K extends keyof WatchSettings>(k: K, v: WatchSettings[K]) =>
    setDraft({ ...draft, watch: { ...draft.watch, [k]: v } })

  const applyProfile = (p: NamedProfile) =>
    setDraft({
      ...draft,
      source_workbook_path: p.source_workbook_path,
      bas_folder_path: p.bas_folder_path,
      dat_file_path: p.dat_file_path,
      output_dir: p.output_dir,
      macro_name: p.macro_name,
      preferred_mode: p.preferred_mode,
    })

  const saveAsProfile = () => {
    const name = window.prompt(
      'Profilname für die aktuelle Konfiguration:',
      'Standard',
    )
    if (!name) return
    const profile: NamedProfile = {
      name,
      source_workbook_path: draft.source_workbook_path,
      bas_folder_path: draft.bas_folder_path,
      dat_file_path: draft.dat_file_path,
      output_dir: draft.output_dir,
      macro_name: draft.macro_name,
      preferred_mode: draft.preferred_mode,
    }
    const rest = draft.profiles.filter((e) => e.name !== name)
    setDraft({
      ...draft,
      profiles: [profile, ...rest].sort((a, b) =>
        a.name.localeCompare(b.name, 'de'),
      ),
    })
  }

  const exportJSON = () => {
    const blob = new Blob([JSON.stringify(draft, null, 2)], {
      type: 'application/json',
    })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = 'phi-sequenz-profil.json'
    a.click()
    URL.revokeObjectURL(url)
  }

  const importJSON = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    setDraft(JSON.parse(await file.text()) as AppSettings)
    e.target.value = ''
  }

  const next = () => setStep((s) => Math.min(s + 1, WIZARD_STEPS.length - 1))
  const back = () => setStep((s) => Math.max(s - 1, 0))

  return (
    <div className="page-grid animate-fade-in">
      {/* Step indicator */}
      <StepBar steps={WIZARD_STEPS} current={step} />

      {/* Step 0: Quelldateien */}
      {step === 0 && (
        <section className="glass-card animate-slide-up">
          <div className="card-header">
            <div>
              <span className="eyebrow">Schritt 1 von 4</span>
              <h2>Quelldateien festlegen</h2>
            </div>
          </div>
          <p className="text-muted card-desc">
            Wählen Sie die Excel-Quelldatei und den Ordner mit den
            exportierten BAS-Modulen. Diese Dateien bilden die Grundlage
            für den automatisierten Workbook-Build.
          </p>
          <div className="form-stack">
            <Field
              label="Excel-Quelldatei (.xlsm)"
              value={draft.source_workbook_path}
              onChange={(v) => up('source_workbook_path', v)}
              placeholder="C:\Pfad\zur\Quelldatei.xlsm"
            />
            <Field
              label="BAS-Modulordner"
              value={draft.bas_folder_path}
              onChange={(v) => up('bas_folder_path', v)}
              placeholder="C:\Pfad\zu\den\BAS-Modulen"
            />
          </div>
          <div className="card-footer">
            <div className="btn-row">
              <button
                className="btn-ghost btn-sm"
                onClick={() => props.onOpenPath(draft.source_workbook_path)}
              >
                <FolderOpen size={14} />
                Workbook öffnen
              </button>
              <button
                className="btn-ghost btn-sm"
                onClick={() => props.onOpenPath(draft.bas_folder_path)}
              >
                <FolderOpen size={14} />
                BAS-Ordner öffnen
              </button>
            </div>
            <button className="btn-primary" onClick={next}>
              Weiter
              <ArrowRight size={16} />
            </button>
          </div>
        </section>
      )}

      {/* Step 1: Messdaten */}
      {step === 1 && (
        <section className="glass-card animate-slide-up">
          <div className="card-header">
            <div>
              <span className="eyebrow">Schritt 2 von 4</span>
              <h2>Messdaten und Ausgabe</h2>
            </div>
          </div>
          <p className="text-muted card-desc">
            Geben Sie den Pfad zur DAT-Datei mit den Messdaten an und
            bestimmen Sie, wohin die erzeugten Artefakte geschrieben werden.
          </p>
          <div className="form-stack">
            <Field
              label="DAT-Datei"
              value={draft.dat_file_path}
              onChange={(v) => up('dat_file_path', v)}
              placeholder="C:\dialims\messsequenzPHI.dat"
            />
            <Field
              label="Ausgabeordner"
              value={draft.output_dir}
              onChange={(v) => up('output_dir', v)}
              placeholder="C:\Pfad\zur\Ausgabe"
            />
            <Field
              label="Runtime-Workbook (automatisch erzeugt)"
              value={draft.generated_workbook_path}
              onChange={(v) => up('generated_workbook_path', v)}
              placeholder="runtime\generated\phi_hybrid_runtime.xlsm"
            />
          </div>
          <div className="card-footer">
            <button className="btn-ghost" onClick={back}>
              <ArrowLeft size={16} />
              Zurück
            </button>
            <button className="btn-primary" onClick={next}>
              Weiter
              <ArrowRight size={16} />
            </button>
          </div>
        </section>
      )}

      {/* Step 2: Optionen */}
      {step === 2 && (
        <section className="glass-card animate-slide-up">
          <div className="card-header">
            <div>
              <span className="eyebrow">Schritt 3 von 4</span>
              <h2>Optionen und Watchdog</h2>
            </div>
          </div>
          <p className="text-muted card-desc">
            Passen Sie den Makronamen, den Ausführungsmodus und die
            automatische Ordnerüberwachung an.
          </p>
          <div className="form-stack">
            <div className="grid-2">
              <Field
                label="Makroname"
                value={draft.macro_name}
                onChange={(v) => up('macro_name', v)}
              />
              <SelectField
                label="Bevorzugter Modus"
                value={draft.preferred_mode}
                onChange={(v) => up('preferred_mode', v as RunMode)}
                options={[
                  { value: 'headless', label: 'Headless (unsichtbar)' },
                  { value: 'visible', label: 'Sichtbar (Excel-Fenster)' },
                ]}
              />
            </div>
            <div className="toggle-row">
              <ToggleChip
                label="Workbook automatisch neu bauen"
                checked={draft.auto_rebuild_workbook}
                onChange={(v) => up('auto_rebuild_workbook', v)}
              />
              <ToggleChip
                label="Watchdog beim Speichern aktivieren"
                checked={draft.watch.enabled}
                onChange={(v) => upWatch('enabled', v)}
              />
            </div>
            {draft.watch.enabled && (
              <div className="grid-2">
                <Field
                  label="Watch-Ordner"
                  value={draft.watch.watch_folder}
                  onChange={(v) => upWatch('watch_folder', v)}
                />
                <Field
                  label="Dateimuster"
                  value={draft.watch.pattern}
                  onChange={(v) => upWatch('pattern', v)}
                />
                <Field
                  label="Verzögerung (ms)"
                  value={String(draft.watch.debounce_ms)}
                  onChange={(v) =>
                    upWatch('debounce_ms', Number.parseInt(v || '0', 10) || 0)
                  }
                />
                <SelectField
                  label="Auto-Lauf Modus"
                  value={draft.watch.auto_run_mode}
                  onChange={(v) => upWatch('auto_run_mode', v as RunMode)}
                  options={[
                    { value: 'headless', label: 'Headless' },
                    { value: 'visible', label: 'Sichtbar' },
                  ]}
                />
              </div>
            )}
          </div>
          <div className="card-footer">
            <button className="btn-ghost" onClick={back}>
              <ArrowLeft size={16} />
              Zurück
            </button>
            <button className="btn-primary" onClick={next}>
              Weiter
              <ArrowRight size={16} />
            </button>
          </div>
        </section>
      )}

      {/* Step 3: Prüfung & Speichern */}
      {step === 3 && (
        <section className="glass-card animate-slide-up">
          <div className="card-header">
            <div>
              <span className="eyebrow">Schritt 4 von 4</span>
              <h2>Prüfung und Speichern</h2>
            </div>
            <button
              className="btn-ghost btn-sm"
              onClick={props.onValidate}
            >
              Erneut prüfen
            </button>
          </div>
          <p className="text-muted card-desc">
            Überprüfen Sie die Zusammenfassung und speichern Sie die
            Konfiguration. Die Validierung prüft, ob alle Dateien
            vorhanden und korrekt sind.
          </p>
          <StatusSummary validation={props.validation} />
          <div className="card-footer">
            <button className="btn-ghost" onClick={back}>
              <ArrowLeft size={16} />
              Zurück
            </button>
            <button
              className="btn-primary btn-lg"
              onClick={() => props.onSave(draft)}
            >
              <Save size={18} />
              Speichern und übernehmen
            </button>
          </div>
        </section>
      )}

      {/* Profiles (always visible) */}
      <section className="glass-card">
        <div className="card-header">
          <div>
            <span className="eyebrow">Profile</span>
            <h2>Schnellwahl und Import/Export</h2>
          </div>
          <div className="btn-row">
            <button className="btn-ghost btn-sm" onClick={saveAsProfile}>
              <Save size={14} />
              Als Profil sichern
            </button>
            <button className="btn-ghost btn-sm" onClick={exportJSON}>
              <Upload size={14} />
              Exportieren
            </button>
            <button
              className="btn-ghost btn-sm"
              onClick={() => importRef.current?.click()}
            >
              <Upload size={14} />
              Importieren
            </button>
            <input
              ref={importRef}
              type="file"
              accept="application/json"
              hidden
              onChange={(e) => void importJSON(e)}
            />
          </div>
        </div>
        <div className="profile-list">
          {draft.profiles.length === 0 ? (
            <div className="empty-state compact">
              Noch keine gespeicherten Profile.
            </div>
          ) : (
            draft.profiles.map((p) => (
              <div key={p.name} className="profile-card">
                <div>
                  <strong>{p.name}</strong>
                  <p className="text-muted text-sm">
                    {p.source_workbook_path}
                  </p>
                </div>
                <div className="btn-row">
                  <button
                    className="btn-ghost btn-sm"
                    onClick={() => applyProfile(p)}
                  >
                    Anwenden
                  </button>
                  <button
                    className="btn-ghost btn-sm"
                    onClick={() =>
                      setDraft({
                        ...draft,
                        profiles: draft.profiles.filter(
                          (e) => e.name !== p.name,
                        ),
                      })
                    }
                  >
                    Entfernen
                  </button>
                </div>
              </div>
            ))
          )}
        </div>
      </section>
    </div>
  )
}
