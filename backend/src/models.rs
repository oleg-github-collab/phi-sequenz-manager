use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};
use uuid::Uuid;

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum RunMode {
    Headless,
    Visible,
}

impl Default for RunMode {
    fn default() -> Self {
        Self::Headless
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NamedProfile {
    pub name: String,
    pub source_workbook_path: String,
    pub bas_folder_path: String,
    pub dat_file_path: String,
    pub output_dir: String,
    pub macro_name: String,
    pub preferred_mode: RunMode,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WatchSettings {
    pub enabled: bool,
    pub watch_folder: String,
    pub debounce_ms: u64,
    pub auto_run_mode: RunMode,
    pub pattern: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AppSettings {
    pub source_workbook_path: String,
    pub bas_folder_path: String,
    pub dat_file_path: String,
    pub output_dir: String,
    pub generated_workbook_path: String,
    pub macro_name: String,
    pub preferred_mode: RunMode,
    pub auto_rebuild_workbook: bool,
    pub profiles: Vec<NamedProfile>,
    pub watch: WatchSettings,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum ValidationLevel {
    Error,
    Warning,
    Info,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ValidationIssue {
    pub level: ValidationLevel,
    pub code: String,
    pub message: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ValidationStats {
    pub row_count: usize,
    pub header_count: usize,
    pub k_sample_count: usize,
    pub p_sample_count: usize,
    pub duplicate_sample_count: usize,
    pub phi50_candidates: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ValidationSummary {
    pub ok: bool,
    pub workbook_exists: bool,
    pub bas_modules_found: bool,
    pub dat_exists: bool,
    pub headers: Vec<String>,
    pub issues: Vec<ValidationIssue>,
    pub stats: ValidationStats,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RunRequest {
    pub source_workbook_path: String,
    pub bas_folder_path: String,
    pub dat_file_path: String,
    pub output_dir: String,
    pub macro_name: String,
    pub mode: RunMode,
    pub force_rebuild_workbook: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RunArtifact {
    pub label: String,
    pub path: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct JobPreview {
    pub sheet_name: String,
    pub headers: Vec<String>,
    pub rows: Vec<Vec<String>>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum JobStatus {
    Pending,
    Running,
    Success,
    Failed,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct JobRecord {
    pub id: Uuid,
    pub status: JobStatus,
    pub title: String,
    pub created_at: DateTime<Utc>,
    pub started_at: Option<DateTime<Utc>>,
    pub finished_at: Option<DateTime<Utc>>,
    pub request: RunRequest,
    pub validation: ValidationSummary,
    pub message: String,
    pub artifacts: Vec<RunArtifact>,
    pub preview: Option<JobPreview>,
    pub step_log: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DashboardResponse {
    pub settings: AppSettings,
    pub latest_validation: Option<ValidationSummary>,
    pub jobs: Vec<JobRecord>,
    pub active_watch: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OpenFolderRequest {
    pub path: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WatchRequest {
    pub watch_folder: String,
    pub pattern: String,
    pub debounce_ms: u64,
    pub mode: RunMode,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LogEntry {
    pub timestamp: DateTime<Utc>,
    pub event_type: String,
    pub message: String,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum FileKind {
    File,
    Directory,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FileSystemEntry {
    pub name: String,
    pub path: String,
    pub kind: FileKind,
    pub size_bytes: u64,
    pub modified_at: Option<DateTime<Utc>>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DirectoryListing {
    pub requested_path: String,
    pub exists: bool,
    pub entries: Vec<FileSystemEntry>,
}
