# Hybrid Approach

## Zielbild

Dieser Ordner enthaelt eine alternative, technisch orientierte Laufumgebung fuer die vorhandenen PHI-Makros.
Die Idee ist bewusst hybrid:

- Excel/VBA bleibt der fachliche Zielpunkt fuer `Daten`, `Needle1` und das bestehende Workbook.
- PowerShell uebernimmt Build, Workbook-Rebuild, Headless-Ausfuehrung, Artefakt-Export und Smoke-Tests.
- Rust stellt die lokale API, den Watchdog, die Validierung, die Web-Auslieferung und die technische Laufsteuerung bereit.
- React/Vite liefert eine moderne deutsche Bedienoberflaeche mit Dashboard, Dateien & Profilen, Ergebnistabelle und separater Log-Seite.

Damit entsteht keine lose Sammlung aus Skripten, sondern eine reproduzierbare Runtime fuer Bediener, Entwickler und Power-User.

## Warum dieser Ansatz deutlich robuster ist

### 1. Reproduzierbarer Workbook-Build

Das Runtime-Workbook wird nicht manuell gepflegt, sondern aus dem Quell-Workbook und den exportierten BAS-Modulen neu aufgebaut.
Die PowerShell-Pipeline importiert die Module gepatcht, kompiliert das VBA-Projekt und schreibt ein Manifest.
Ein Rebuild erfolgt automatisch, wenn sich eines dieser Dinge geaendert hat:

- Quell-Workbook
- BAS-Module
- verwendeter DAT-Pfad

Das reduziert stille Abweichungen zwischen Entwicklungsstand und Laufzeit erheblich.

### 2. Echte Headless-Faehigkeit

Die VBA-Module enthalten Benutzer-Dialoge per `MsgBox`.
Fuer unsichtbare Excel-Laeufe ist das gefaehrlich, weil ein unsichtbarer Dialog den kompletten Prozess blockieren kann.

Deshalb wird beim Runtime-Build ein zusaetzliches Automationsmodul importiert und die kritischen Module werden fuer Automationslaeufe gepatcht:

- Dialoge werden unterdrueckt
- statt dessen wird ins Log geschrieben
- Excel kann damit unsichtbar und ohne Haenger durchlaufen

### 3. Klare technische Rueckmeldung

Jeder Lauf erzeugt einen JSON-Report mit:

- Laufstatus
- erzeugten Artefakten
- Ergebnisvorschau der `Needle1`-Tabelle
- exportierten CSV-Dateien fuer `Needle1`, `Daten` und `_Execution_Log`

Die Weboberflaeche muss also nicht raten, sondern kann saubere Resultate anzeigen.

### 4. Lokale API statt VBA-Magie im Browser

Die Weboberflaeche spricht nicht direkt mit Excel, sondern mit dem Rust-Dienst.
Dadurch bleiben diese Aufgaben an einer robusten Stelle gebuendelt:

- Validierung von Pfaden und DAT-Dateien
- Starten von PowerShell-Pipelines
- Job-Historie
- Log-Ausgabe
- Watchdog fuer neue DAT-Dateien
- Auslieferung der gebauten Weboberflaeche

### 5. Lokales, gepatchtes Unicorn-SDK ohne Branding

Die Lade-/Visualisierungsflaeche nutzt `unicornstudio-react`, aber nicht das CDN-SDK.
Stattdessen wird eine lokal gepatchte Datei verwendet:

- Quelle: `artifacts/unicornStudio.umd.js`
- Ziel: `frontend/public/vendor/unicornStudio.patched.js`

Der Branding-Block fuer das Unicorn-Logo wird lokal entfernt, sodass die Szene weiterhin laeuft, aber kein externes SDK-Branding eingeblendet wird.

## Was hier bereits implementiert ist

### Backend in Rust

Pfad: `backend/`

Funktionen:

- REST-API fuer Dashboard, Settings, Validierung, Jobs, Logs und Dateibrowser
- Watchdog fuer `*.dat`
- Speicherung von `settings.json`, `jobs.json` und `app-events.jsonl`
- Auslieferung des gebauten Frontends aus `runtime/web`
- CLI-Subcommands:
  - `serve`
  - `validate`
  - `run`

### PowerShell Runtime

Pfad: `scripts/`

Wichtige Skripte:

- `run-excel-pipeline.ps1`
  - fuehrt die komplette Excel/VBA-Kette aus
- `build-workbook.ps1`
  - erzeugt ein runtime-faehiges Workbook
- `common.ps1`
  - COM-, Export- und Build-Helfer
- `patch-unicorn-sdk.ps1`
  - erzeugt das lokale gepatchte Unicorn-SDK
- `start-hybrid.ps1`
  - baut Frontend und Backend und startet den Server
- `smoke-test.ps1`
  - testet Frontend-Build, Backend-Build, Excel-Pipeline und API
- `package-local-release.ps1`
  - erstellt ein lokales Release-Zip
- `run-release.ps1`
  - startet ein vorhandenes Binary direkt als lokalen Server

### Weboberflaeche

Pfad: `frontend/`

Bereiche:

- `Leitstand`
  - Schnellstart headless oder sichtbar
  - Validierung
  - Watchdog-Status
  - letzter Lauf
- `Dateien & Profile`
  - Pfade
  - Makro
  - Auto-Rebuild
  - Watchdog-Einstellungen
  - Profilimport und -export
- `Ergebnisse`
  - Job-Historie
  - Ergebnistabelle aus `Needle1`
  - Dateibrowser fuer Artefakte
- `Logs`
  - App-Ereignisse
  - technische Betriebsnotizen

## Schnellstart lokal

### 1. Voller Start der lokalen Oberflaeche

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\start-hybrid.ps1 -OpenBrowser
```

Danach ist die Oberflaeche unter `http://127.0.0.1:8765` erreichbar.

### 2. Nur Backend bauen

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-backend.ps1
```

### 3. Workbook- und Excel-Pipeline direkt testen

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\run-excel-pipeline.ps1 `
  -SourceWorkbook "C:\Pfad\zum\Workbook.xlsm" `
  -BasFolder "C:\Pfad\zu\den\bas" `
  -DatPath "C:\Pfad\zur\datei.dat" `
  -OutputDir "C:\Pfad\zu\output" `
  -GeneratedWorkbook "C:\Pfad\zu\runtime.xlsm" `
  -MacroName "MakrosAusfuehren" `
  -ReportPath "C:\Pfad\zu\report.json" `
  -ForceRebuild
```

### 4. End-to-End Smoke-Test

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\smoke-test.ps1
```

Dieser Test prueft:

- lokalen Unicorn-Patch
- Frontend-Build
- Rust-Backend-Build
- Excel/VBA-Lauf gegen eine Beispiel-DAT
- API-Start
- `health`, `dashboard` und `validate`

## Release und Verteilung

### Lokales Release-Zip

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\package-local-release.ps1
```

Ergebnis:

- Ordner: `runtime/releases/hybrid-approach-local-win11/`
- Zip: `runtime/releases/hybrid-approach-local-win11.zip`

### GitHub Actions

Workflow:

- `.github/workflows/windows-release.yml`

Der Workflow:

- patcht das lokale Unicorn-SDK
- baut das React-Frontend
- baut das Rust-Backend als Windows-Release
- legt ein gebrauchsfertiges Bundle als ZIP ab
- haengt das ZIP bei Tag-Releases automatisch an eine GitHub Release

Wichtig:

- Die Workflow-Datei ist vorbereitet.
- Ein echter Download einer GitHub-Release in diesen Ordner ist erst moeglich, wenn dieses Projekt in einem GitHub-Repository liegt und dort ein Release erzeugt wurde.

## Validierte Punkte in diesem Ordner

Zum Stand dieser Implementierung wurden erfolgreich geprueft:

- Excel/VBA-Pipeline gegen `tests/fixtures/sample_phi.dat`
- Rebuild des Runtime-Workbooks
- VBA-Compile waehrend des Workbook-Builds
- Export von `Needle1`, `Daten` und `_Execution_Log`
- Rust-Backend-Build lokal mit GNU-Toolchain
- Frontend-Build mit Vite
- API-Smoke-Test mit `health`, `dashboard` und `validate`

## Wichtige Voraussetzungen

- Microsoft Excel muss lokal installiert sein.
- In Excel muss der Zugriff auf das VBA-Projektmodell erlaubt sein:
  - `Trust access to the VBA project object model`
- Die Originaldateien muessen weiterhin vorhanden sein:
  - Quell-Workbook im uebergeordneten Ordner
  - exportierte BAS-Module im uebergeordneten Ordner

## Technische Grenzen

- GitHub-Release-Downloads koennen hier lokal nicht "fertig" geliefert werden, solange kein echtes Remote-Repository existiert.
- Die Runtime ist bewusst auf Windows ausgerichtet, weil Excel COM Teil des Kerns ist.
- Der fachliche Sequenzalgorithmus bleibt im VBA-Workbook; Rust ist hier vor allem Steuerungs-, API- und Validierungs-Layer.

## Empfohlene naechste Ausbaustufen

- Fachliche Kernlogik schrittweise aus VBA nach Rust auslagern
- Vergleich gegen Goldstandard-Dateien fuer echte Regressionspruefungen
- signierte Release-Pakete fuer kontrollierte Verteilung
- automatische Inbox-Verarbeitung als Windows-Aufgabe oder Dienst
- projektweites Repo mit echter GitHub Release Pipeline
