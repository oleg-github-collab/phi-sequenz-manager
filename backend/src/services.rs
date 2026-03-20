use crate::models::{
    AppSettings, DashboardResponse, DirectoryListing, FileKind, FileSystemEntry, JobPreview, JobRecord,
    JobStatus, LogEntry, NamedProfile, OpenFolderRequest, RunArtifact, RunMode, RunRequest,
    ValidationIssue, ValidationLevel, ValidationSummary, WatchRequest, WatchSettings,
};
use anyhow::{bail, Context, Result};
use chrono::{DateTime, Utc};
use csv::ReaderBuilder;
use notify::{Config as NotifyConfig, Event, EventKind, RecommendedWatcher, RecursiveMode, Watcher};
use regex::Regex;
use serde::{Deserialize, Serialize};
use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};
use std::fs::{self, OpenOptions};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use tokio::process::Command;
use tokio::sync::{mpsc, Mutex, RwLock};
use tokio::task::JoinHandle;
use tokio::time::{sleep, Duration, Instant};
use tracing::{error, warn};
use uuid::Uuid;

const REQUIRED_BAS_FILES: [&str; 6] = [
    "Modul1_Konstanten.bas",
    "Modul2_Logger.bas",
    "Modul3_Hilfsfunktionen.bas",
    "Modul4_ImportCSV.bas",
    "Modul5_SequenzSchreiben.bas",
    "Modul6_Hauptsteuerung.bas",
];

#[derive(Debug, Clone)]
pub struct AppPaths {
    pub root: PathBuf,
    pub runtime_output: PathBuf,
    pub runtime_generated: PathBuf,
    pub runtime_logs: PathBuf,
    pub runtime_inbox: PathBuf,
    pub settings_file: PathBuf,
    pub jobs_file: PathBuf,
    pub scripts_dir: PathBuf,
    pub web_dir: PathBuf,
}

impl AppPaths {
    pub fn new(root: PathBuf) -> Result<Self> {
        let runtime = root.join("runtime");
        let runtime_output = runtime.join("output");
        let runtime_generated = runtime.join("generated");
        let runtime_logs = runtime.join("logs");
        let runtime_inbox = runtime.join("inbox");
        let settings_file = runtime.join("settings.json");
        let jobs_file = runtime.join("jobs.json");
        let scripts_dir = root.join("scripts");
        let web_dir = runtime.join("web");

        for dir in [
            &runtime,
            &runtime_output,
            &runtime_generated,
            &runtime_logs,
            &runtime_inbox,
            &web_dir,
        ] {
            fs::create_dir_all(dir)?;
        }

        Ok(Self {
            root,
            runtime_output,
            runtime_generated,
            runtime_logs,
            runtime_inbox,
            settings_file,
            jobs_file,
            scripts_dir,
            web_dir,
        })
    }
}

struct WatchController {
    _watcher: RecommendedWatcher,
    task: JoinHandle<()>,
}

pub struct AppState {
    pub paths: AppPaths,
    pub settings: RwLock<AppSettings>,
    pub jobs: RwLock<Vec<JobRecord>>,
    pub latest_validation: RwLock<Option<ValidationSummary>>,
    watch_controller: Mutex<Option<WatchController>>,
}

impl AppState {
    pub async fn load(paths: AppPaths) -> Result<Arc<Self>> {
        let settings = load_or_default_settings(&paths).context("settings konnten nicht geladen werden")?;
        let jobs = load_or_default_jobs(&paths).context("jobs konnten nicht geladen werden")?;

        Ok(Arc::new(Self {
            paths,
            settings: RwLock::new(settings),
            jobs: RwLock::new(jobs),
            latest_validation: RwLock::new(None),
            watch_controller: Mutex::new(None),
        }))
    }

    pub async fn dashboard(self: &Arc<Self>) -> DashboardResponse {
        DashboardResponse {
            settings: self.settings.read().await.clone(),
            latest_validation: self.latest_validation.read().await.clone(),
            jobs: self.jobs.read().await.clone(),
            active_watch: self.watch_controller.lock().await.is_some(),
        }
    }

    pub async fn current_settings(&self) -> AppSettings {
        self.settings.read().await.clone()
    }

    pub async fn jobs_snapshot(&self) -> Vec<JobRecord> {
        self.jobs.read().await.clone()
    }

    pub async fn save_settings(&self, settings: AppSettings) -> Result<AppSettings> {
        if settings.output_dir.trim().is_empty() {
            bail!("Das Ausgabeverzeichnis darf nicht leer sein.");
        }

        fs::create_dir_all(&settings.output_dir)?;
        if !settings.generated_workbook_path.trim().is_empty() {
            if let Some(parent) = Path::new(&settings.generated_workbook_path).parent() {
                fs::create_dir_all(parent)?;
            }
        }
        if !settings.watch.watch_folder.trim().is_empty() {
            fs::create_dir_all(&settings.watch.watch_folder)?;
        }

        write_json_pretty(&self.paths.settings_file, &settings)?;
        *self.settings.write().await = settings.clone();
        self.append_event("settings_saved", "Einstellungen wurden gespeichert.")
            .await?;
        Ok(settings)
    }

    pub async fn validate_with_settings(&self) -> Result<ValidationSummary> {
        let settings = self.settings.read().await.clone();
        let request = RunRequest {
            source_workbook_path: settings.source_workbook_path.clone(),
            bas_folder_path: settings.bas_folder_path.clone(),
            dat_file_path: settings.dat_file_path.clone(),
            output_dir: settings.output_dir.clone(),
            macro_name: settings.macro_name.clone(),
            mode: settings.preferred_mode.clone(),
            force_rebuild_workbook: settings.auto_rebuild_workbook,
        };

        self.validate_request(&request).await
    }

    pub async fn validate_request(&self, request: &RunRequest) -> Result<ValidationSummary> {
        let mut summary = ValidationSummary::default();
        let workbook = PathBuf::from(&request.source_workbook_path);
        let bas_folder = PathBuf::from(&request.bas_folder_path);
        let dat_path = PathBuf::from(&request.dat_file_path);

        summary.workbook_exists = workbook.exists();
        summary.dat_exists = dat_path.exists();
        summary.bas_modules_found = REQUIRED_BAS_FILES
            .iter()
            .all(|file_name| bas_folder.join(file_name).exists());

        if !summary.workbook_exists {
            summary.issues.push(issue(
                ValidationLevel::Error,
                "workbook_missing",
                "Die Excel-Quelldatei wurde nicht gefunden.",
            ));
        }

        if !summary.bas_modules_found {
            summary.issues.push(issue(
                ValidationLevel::Error,
                "bas_missing",
                "Mindestens ein exportiertes BAS-Modul fehlt.",
            ));
        }

        if request.dat_file_path.trim().is_empty() {
            summary.issues.push(issue(
                ValidationLevel::Error,
                "dat_empty",
                "Es wurde keine DAT-Datei angegeben.",
            ));
        } else if !summary.dat_exists {
            summary.issues.push(issue(
                ValidationLevel::Error,
                "dat_missing",
                "Die angegebene DAT-Datei wurde nicht gefunden.",
            ));
        }

        if summary.dat_exists {
            self.populate_dat_validation(&dat_path, &mut summary)?;
        }

        summary.ok = !summary
            .issues
            .iter()
            .any(|item| item.level == ValidationLevel::Error);

        *self.latest_validation.write().await = Some(summary.clone());
        Ok(summary)
    }

    pub async fn run_job(self: &Arc<Self>, request: RunRequest) -> Result<JobRecord> {
        let validation = self.validate_request(&request).await?;
        let mut job = JobRecord {
            id: Uuid::new_v4(),
            status: JobStatus::Pending,
            title: format!("Hybrid-Lauf {}", Utc::now().format("%Y-%m-%d %H:%M:%S")),
            created_at: Utc::now(),
            started_at: None,
            finished_at: None,
            request: request.clone(),
            validation: validation.clone(),
            message: String::new(),
            artifacts: Vec::new(),
            preview: None,
            step_log: vec!["Validierung abgeschlossen.".to_string()],
        };

        if !validation.ok {
            job.status = JobStatus::Failed;
            job.finished_at = Some(Utc::now());
            job.message = "Die Validierung hat kritische Fehler gefunden.".to_string();
            self.upsert_job(job.clone()).await?;
            return Ok(job);
        }

        fs::create_dir_all(&request.output_dir)?;
        job.status = JobStatus::Running;
        job.started_at = Some(Utc::now());
        self.upsert_job(job.clone()).await?;

        let generated_workbook_path = self.resolve_generated_workbook_path().await;
        let report_path = PathBuf::from(&request.output_dir).join(format!("job-{}-report.json", job.id));
        let script_path = self.paths.scripts_dir.join("run-excel-pipeline.ps1");

        if !script_path.exists() {
            bail!("Das PowerShell-Skript '{}' wurde nicht gefunden.", script_path.display());
        }

        let mut command = Command::new("powershell.exe");
        command
            .arg("-NoProfile")
            .arg("-ExecutionPolicy")
            .arg("Bypass")
            .arg("-File")
            .arg(&script_path)
            .arg("-SourceWorkbook")
            .arg(&request.source_workbook_path)
            .arg("-BasFolder")
            .arg(&request.bas_folder_path)
            .arg("-DatPath")
            .arg(&request.dat_file_path)
            .arg("-OutputDir")
            .arg(&request.output_dir)
            .arg("-GeneratedWorkbook")
            .arg(&generated_workbook_path)
            .arg("-MacroName")
            .arg(&request.macro_name)
            .arg("-ReportPath")
            .arg(&report_path);

        if request.mode == RunMode::Visible {
            command.arg("-Visible");
        }

        if request.force_rebuild_workbook {
            command.arg("-ForceRebuild");
        }

        let output = command
            .output()
            .await
            .context("Die PowerShell-Pipeline konnte nicht gestartet werden.")?;

        let stdout = String::from_utf8_lossy(&output.stdout).trim().to_string();
        let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();

        if !stdout.is_empty() {
            job.step_log.push(stdout);
        }
        if !stderr.is_empty() {
            job.step_log.push(stderr.clone());
        }

        let mut report: Option<ScriptReport> = None;
        if report_path.exists() {
            report = Some(read_json(&report_path).context("Pipeline-Report konnte nicht gelesen werden.")?);
        }

        if let Some(report) = report {
            job.artifacts = report
                .artifacts
                .into_iter()
                .map(|artifact| RunArtifact {
                    label: artifact.label,
                    path: artifact.path,
                })
                .collect();
            job.preview = report.preview.map(|preview| JobPreview {
                sheet_name: preview.sheet_name,
                headers: preview.headers,
                rows: preview.rows,
            });
            job.step_log.extend(report.messages);

            if output.status.success() && report.success {
                job.status = JobStatus::Success;
                job.message = "Hybrid-Lauf erfolgreich abgeschlossen.".to_string();
            } else {
                job.status = JobStatus::Failed;
                job.message = if stderr.is_empty() {
                    "Die Excel-Pipeline ist fehlgeschlagen.".to_string()
                } else {
                    format!("Die Excel-Pipeline ist fehlgeschlagen: {stderr}")
                };
            }
        } else if output.status.success() {
            job.status = JobStatus::Failed;
            job.message = "Der Lauf war erfolgreich, aber der technische Report fehlt.".to_string();
            job.step_log
                .push("Warnung: Die Pipeline hat keinen Report geschrieben.".to_string());
        } else {
            job.status = JobStatus::Failed;
            job.message = if stderr.is_empty() {
                "Die Excel-Pipeline ist fehlgeschlagen.".to_string()
            } else {
                format!("Die Excel-Pipeline ist fehlgeschlagen: {stderr}")
            };
        }

        job.finished_at = Some(Utc::now());
        self.upsert_job(job.clone()).await?;
        self.append_event(
            "job_finished",
            &format!("Job {} wurde mit Status {:?} beendet.", job.id, job.status),
        )
        .await?;

        Ok(job)
    }

    pub async fn open_folder(&self, request: OpenFolderRequest) -> Result<()> {
        let path = PathBuf::from(&request.path);
        if !path.exists() {
            bail!("Pfad existiert nicht: {}", path.display());
        }

        Command::new("explorer.exe")
            .arg(&path)
            .spawn()
            .context("Der Explorer konnte nicht gestartet werden.")?;
        Ok(())
    }

    pub async fn list_directory(&self, path: Option<String>) -> Result<DirectoryListing> {
        let requested_path = path
            .filter(|value| !value.trim().is_empty())
            .unwrap_or_else(|| self.paths.runtime_output.to_string_lossy().into_owned());
        let resolved = PathBuf::from(&requested_path);

        if !resolved.exists() {
            return Ok(DirectoryListing {
                requested_path,
                exists: false,
                entries: Vec::new(),
            });
        }

        let mut entries = Vec::new();
        for entry in fs::read_dir(&resolved)? {
            let entry = entry?;
            let path = entry.path();
            let metadata = entry.metadata()?;
            let kind = if metadata.is_dir() {
                FileKind::Directory
            } else {
                FileKind::File
            };
            let modified_at = metadata
                .modified()
                .ok()
                .map(DateTime::<Utc>::from);

            entries.push(FileSystemEntry {
                name: entry.file_name().to_string_lossy().into_owned(),
                path: path.to_string_lossy().into_owned(),
                kind,
                size_bytes: if metadata.is_file() { metadata.len() } else { 0 },
                modified_at,
            });
        }

        entries.sort_by(|left, right| match (&left.kind, &right.kind) {
            (FileKind::Directory, FileKind::File) => Ordering::Less,
            (FileKind::File, FileKind::Directory) => Ordering::Greater,
            _ => left.name.to_lowercase().cmp(&right.name.to_lowercase()),
        });

        Ok(DirectoryListing {
            requested_path,
            exists: true,
            entries,
        })
    }

    pub async fn read_logs(&self) -> Result<Vec<LogEntry>> {
        let path = self.paths.runtime_logs.join("app-events.jsonl");
        if !path.exists() {
            return Ok(Vec::new());
        }

        let content = fs::read_to_string(path)?;
        let mut entries = Vec::new();
        for line in content.lines() {
            let line = line.trim();
            if line.is_empty() {
                continue;
            }

            match serde_json::from_str::<LogEntry>(line) {
                Ok(entry) => entries.push(entry),
                Err(error) => {
                    warn!("Logzeile konnte nicht gelesen werden: {error}");
                }
            }
        }

        entries.sort_by(|left, right| right.timestamp.cmp(&left.timestamp));
        Ok(entries)
    }

    pub async fn start_watch(self: &Arc<Self>, request: WatchRequest) -> Result<()> {
        self.stop_watch().await?;

        let watch_path = PathBuf::from(&request.watch_folder);
        if !watch_path.exists() {
            bail!("Der Watch-Ordner existiert nicht: {}", watch_path.display());
        }

        let (tx, mut rx) = mpsc::unbounded_channel::<notify::Result<Event>>();
        let mut watcher = RecommendedWatcher::new(
            move |result| {
                let _ = tx.send(result);
            },
            NotifyConfig::default(),
        )?;
        watcher.watch(&watch_path, RecursiveMode::NonRecursive)?;

        let state = Arc::clone(self);
        let pattern = request.pattern.clone();
        let mode = request.mode.clone();
        let debounce = Duration::from_millis(request.debounce_ms.max(250));

        let task = tokio::spawn(async move {
            let mut last_trigger = Instant::now() - debounce;
            while let Some(result) = rx.recv().await {
                match result {
                    Ok(event) => {
                        if !matches!(event.kind, EventKind::Create(_) | EventKind::Modify(_)) {
                            continue;
                        }

                        for changed_path in event.paths {
                            if !matches_watch_pattern(&changed_path, &pattern) {
                                continue;
                            }

                            if last_trigger.elapsed() < debounce {
                                continue;
                            }

                            last_trigger = Instant::now();
                            sleep(Duration::from_millis(350)).await;

                            let settings = state.settings.read().await.clone();
                            let run_request = RunRequest {
                                source_workbook_path: settings.source_workbook_path,
                                bas_folder_path: settings.bas_folder_path,
                                dat_file_path: changed_path.to_string_lossy().into_owned(),
                                output_dir: settings.output_dir,
                                macro_name: settings.macro_name,
                                mode: mode.clone(),
                                force_rebuild_workbook: settings.auto_rebuild_workbook,
                            };

                            if let Err(error) = state
                                .append_event(
                                    "watch_triggered",
                                    &format!("Watch-Trigger für '{}'.", changed_path.display()),
                                )
                                .await
                            {
                                warn!("Watch-Log konnte nicht geschrieben werden: {error:#}");
                            }

                            if let Err(error) = state.run_job(run_request).await {
                                error!("Watcher-Lauf fehlgeschlagen: {error:#}");
                            }
                        }
                    }
                    Err(error) => warn!("Fehler beim Watch-Ereignis: {error}"),
                }
            }
        });

        *self.watch_controller.lock().await = Some(WatchController {
            _watcher: watcher,
            task,
        });

        self.append_event("watch_started", "Ordnerbeobachtung wurde aktiviert.")
            .await?;
        Ok(())
    }

    pub async fn stop_watch(&self) -> Result<()> {
        if let Some(controller) = self.watch_controller.lock().await.take() {
            controller.task.abort();
            self.append_event("watch_stopped", "Ordnerbeobachtung wurde gestoppt.")
                .await?;
        }
        Ok(())
    }

    async fn resolve_generated_workbook_path(&self) -> String {
        let settings = self.settings.read().await.clone();
        if settings.generated_workbook_path.trim().is_empty() {
            self.paths
                .runtime_generated
                .join("phi_hybrid_runtime.xlsm")
                .to_string_lossy()
                .into_owned()
        } else {
            settings.generated_workbook_path
        }
    }

    fn populate_dat_validation(&self, dat_path: &Path, summary: &mut ValidationSummary) -> Result<()> {
        let mut reader = ReaderBuilder::new()
            .delimiter(b';')
            .flexible(true)
            .from_path(dat_path)
            .with_context(|| format!("DAT-Datei konnte nicht gelesen werden: {}", dat_path.display()))?;

        let headers = reader
            .headers()
            .context("Header konnten nicht gelesen werden.")?
            .iter()
            .map(|value| value.trim().to_string())
            .collect::<Vec<_>>();

        summary.headers = headers.clone();
        summary.stats.header_count = headers.len();

        let matrix_index = headers.iter().position(|value| value.eq_ignore_ascii_case("Matrix"));
        let probe_index = headers.iter().position(|value| value.eq_ignore_ascii_case("Probe")).or(Some(1));
        let analyse_index = headers.iter().position(|value| value.eq_ignore_ascii_case("Analyse"));

        if matrix_index.is_none() {
            summary.issues.push(issue(
                ValidationLevel::Error,
                "missing_matrix",
                "Die Pflichtspalte 'Matrix' fehlt in der DAT-Datei.",
            ));
        }

        if probe_index.is_none() {
            summary.issues.push(issue(
                ValidationLevel::Error,
                "missing_probe",
                "Die Probenspalte konnte nicht bestimmt werden.",
            ));
        }

        if analyse_index.is_none() {
            summary.issues.push(issue(
                ValidationLevel::Warning,
                "missing_analyse",
                "Die Spalte 'Analyse' fehlt. Kommentare können dadurch unvollständig sein.",
            ));
        }

        let sample_regex = Regex::new(r"^[KP]\d{7}$").expect("sample regex");
        let mut duplicates = HashMap::<String, usize>::new();
        let mut seen_matrices = HashSet::<String>::new();

        for record in reader.records() {
            let record = record?;
            summary.stats.row_count += 1;

            if let Some(index) = probe_index {
                let sample = record.get(index).unwrap_or("").trim().to_string();
                if sample_regex.is_match(&sample) {
                    if sample.starts_with('K') {
                        summary.stats.k_sample_count += 1;
                    }
                    if sample.starts_with('P') {
                        summary.stats.p_sample_count += 1;
                    }
                    *duplicates.entry(sample).or_insert(0) += 1;
                }
            }

            if let Some(index) = matrix_index {
                let matrix = record.get(index).unwrap_or("").trim().to_string();
                if !matrix.is_empty() {
                    seen_matrices.insert(matrix.to_lowercase());
                }
                if matrix.eq_ignore_ascii_case("phI-50") {
                    summary.stats.phi50_candidates += 1;
                }
            }
        }

        summary.stats.duplicate_sample_count = duplicates.values().filter(|count| **count > 1).count();

        if summary.stats.row_count == 0 {
            summary.issues.push(issue(
                ValidationLevel::Error,
                "no_rows",
                "Die DAT-Datei enthält keine Datenzeilen.",
            ));
        }

        if summary.stats.k_sample_count == 0 {
            summary.issues.push(issue(
                ValidationLevel::Warning,
                "no_control_samples",
                "Es wurden keine Kontrollproben im Format KYYxxxxx erkannt.",
            ));
        }

        if summary.stats.p_sample_count == 0 {
            summary.issues.push(issue(
                ValidationLevel::Warning,
                "no_patient_samples",
                "Es wurden keine Proben im Format PYYxxxxx erkannt.",
            ));
        }

        if summary.stats.phi50_candidates == 0 {
            summary.issues.push(issue(
                ValidationLevel::Warning,
                "phi50_missing",
                "Es wurde keine phI-50-Kontrollprobe erkannt.",
            ));
        }

        if summary.stats.duplicate_sample_count > 0 {
            summary.issues.push(issue(
                ValidationLevel::Info,
                "duplicate_samples",
                "Doppelte Probennummern wurden erkannt. Das kann für Mehrfachanalysen korrekt sein.",
            ));
        }

        if !seen_matrices.iter().any(|item| item.contains("w")) {
            summary.issues.push(issue(
                ValidationLevel::Info,
                "no_water_matrix",
                "Es wurden keine Wasser-Matrizen erkannt.",
            ));
        }

        Ok(())
    }

    async fn upsert_job(&self, job: JobRecord) -> Result<()> {
        let mut jobs = self.jobs.write().await;
        if let Some(existing) = jobs.iter_mut().find(|item| item.id == job.id) {
            *existing = job.clone();
        } else {
            jobs.insert(0, job.clone());
        }
        jobs.truncate(30);
        write_json_pretty(&self.paths.jobs_file, &*jobs)?;
        Ok(())
    }

    async fn append_event(&self, event_type: &str, message: &str) -> Result<()> {
        let entry = LogEntry {
            timestamp: Utc::now(),
            event_type: event_type.to_string(),
            message: message.to_string(),
        };

        let path = self.paths.runtime_logs.join("app-events.jsonl");
        let mut handle = OpenOptions::new().create(true).append(true).open(path)?;
        writeln!(handle, "{}", serde_json::to_string(&entry)?)?;
        Ok(())
    }
}

#[derive(Debug, Deserialize)]
struct ScriptArtifact {
    label: String,
    path: String,
}

#[derive(Debug, Deserialize)]
struct ScriptPreview {
    sheet_name: String,
    headers: Vec<String>,
    rows: Vec<Vec<String>>,
}

#[derive(Debug, Deserialize)]
struct ScriptReport {
    success: bool,
    messages: Vec<String>,
    artifacts: Vec<ScriptArtifact>,
    preview: Option<ScriptPreview>,
}

fn issue(level: ValidationLevel, code: &str, message: &str) -> ValidationIssue {
    ValidationIssue {
        level,
        code: code.to_string(),
        message: message.to_string(),
    }
}

fn matches_watch_pattern(path: &Path, pattern: &str) -> bool {
    let file_name = match path.file_name().and_then(|value| value.to_str()) {
        Some(value) => value.to_ascii_lowercase(),
        None => return false,
    };

    let pattern = pattern.trim().to_ascii_lowercase();
    if pattern.is_empty() || pattern == "*.dat" {
        return file_name.ends_with(".dat");
    }

    if pattern.contains('*') || pattern.contains('?') {
        let regex_pattern = format!(
            "^{}$",
            regex::escape(&pattern)
                .replace("\\*", ".*")
                .replace("\\?", ".")
        );

        if let Ok(regex) = Regex::new(&regex_pattern) {
            return regex.is_match(&file_name);
        }
    }

    if let Some(suffix) = pattern.strip_prefix("*.") {
        return file_name.ends_with(&format!(".{suffix}"));
    }

    file_name.contains(&pattern)
}

fn write_json_pretty<T: Serialize>(path: &Path, value: &T) -> Result<()> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }
    fs::write(path, serde_json::to_vec_pretty(value)?)?;
    Ok(())
}

fn read_json<T: for<'de> Deserialize<'de>>(path: &Path) -> Result<T> {
    let content = fs::read_to_string(path)?;
    Ok(serde_json::from_str(content.trim_start_matches('\u{feff}'))?)
}

fn load_or_default_settings(paths: &AppPaths) -> Result<AppSettings> {
    if paths.settings_file.exists() {
        return read_json(&paths.settings_file);
    }

    let parent = paths.root.parent().unwrap_or(paths.root.as_path());
    let default_workbook = parent.join("Messsequenzübertragung_PHI.xlsm");
    let default_bas_folder = parent.to_path_buf();
    let default_output_dir = paths.runtime_output.to_string_lossy().into_owned();
    let default_generated = paths
        .runtime_generated
        .join("phi_hybrid_runtime.xlsm")
        .to_string_lossy()
        .into_owned();

    let default_profile = NamedProfile {
        name: "Standard".to_string(),
        source_workbook_path: default_workbook.to_string_lossy().into_owned(),
        bas_folder_path: default_bas_folder.to_string_lossy().into_owned(),
        dat_file_path: String::new(),
        output_dir: default_output_dir.clone(),
        macro_name: "MakrosAusfuehren".to_string(),
        preferred_mode: RunMode::Headless,
    };

    let settings = AppSettings {
        source_workbook_path: default_profile.source_workbook_path.clone(),
        bas_folder_path: default_profile.bas_folder_path.clone(),
        dat_file_path: String::new(),
        output_dir: default_output_dir,
        generated_workbook_path: default_generated,
        macro_name: "MakrosAusfuehren".to_string(),
        preferred_mode: RunMode::Headless,
        auto_rebuild_workbook: true,
        profiles: vec![default_profile],
        watch: WatchSettings {
            enabled: false,
            watch_folder: paths.runtime_inbox.to_string_lossy().into_owned(),
            debounce_ms: 1500,
            auto_run_mode: RunMode::Headless,
            pattern: "*.dat".to_string(),
        },
    };

    write_json_pretty(&paths.settings_file, &settings)?;
    Ok(settings)
}

fn load_or_default_jobs(paths: &AppPaths) -> Result<Vec<JobRecord>> {
    if paths.jobs_file.exists() {
        return read_json(&paths.jobs_file);
    }

    write_json_pretty(&paths.jobs_file, &Vec::<JobRecord>::new())?;
    Ok(Vec::new())
}
