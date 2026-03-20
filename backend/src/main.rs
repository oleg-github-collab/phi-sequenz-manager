mod models;
mod routes;
mod services;

use anyhow::{bail, Result};
use clap::{Parser, Subcommand, ValueEnum};
use models::RunMode;
use services::{AppPaths, AppState};
use std::net::SocketAddr;
use std::path::PathBuf;
use tracing_subscriber::EnvFilter;

#[derive(Debug, Parser)]
#[command(
    name = "hybrid-approach",
    version,
    about = "Lokaler Hybrid-Dienst für Excel/VBA, PowerShell und Rust"
)]
struct Cli {
    #[command(subcommand)]
    command: Option<Commands>,
}

#[derive(Debug, Subcommand)]
enum Commands {
    Serve {
        #[arg(long, default_value = ".")]
        root: PathBuf,
        #[arg(long, default_value = "127.0.0.1")]
        host: String,
        #[arg(long, default_value_t = 8765)]
        port: u16,
    },
    Validate {
        #[arg(long, default_value = ".")]
        root: PathBuf,
    },
    Run {
        #[arg(long, default_value = ".")]
        root: PathBuf,
        #[arg(long)]
        dat_file: Option<PathBuf>,
        #[arg(long, value_enum, default_value_t = CliRunMode::Headless)]
        mode: CliRunMode,
        #[arg(long, default_value_t = false)]
        force_rebuild: bool,
    },
}

#[derive(Debug, Clone, Copy, ValueEnum)]
enum CliRunMode {
    Headless,
    Visible,
}

impl From<CliRunMode> for RunMode {
    fn from(value: CliRunMode) -> Self {
        match value {
            CliRunMode::Headless => RunMode::Headless,
            CliRunMode::Visible => RunMode::Visible,
        }
    }
}

#[tokio::main]
async fn main() -> Result<()> {
    init_tracing();

    let cli = Cli::parse();
    match cli.command.unwrap_or(Commands::Serve {
        root: PathBuf::from("."),
        host: "127.0.0.1".to_string(),
        port: 8765,
    }) {
        Commands::Serve { root, host, port } => serve(root, host, port).await,
        Commands::Validate { root } => validate(root).await,
        Commands::Run {
            root,
            dat_file,
            mode,
            force_rebuild,
        } => run_once(root, dat_file, mode, force_rebuild).await,
    }
}

async fn serve(root: PathBuf, host: String, port: u16) -> Result<()> {
    let root = canonicalize_root(root)?;
    let state = build_state(root).await?;

    let settings = state.current_settings().await;
    if settings.watch.enabled {
        state
            .start_watch(models::WatchRequest {
                watch_folder: settings.watch.watch_folder.clone(),
                pattern: settings.watch.pattern.clone(),
                debounce_ms: settings.watch.debounce_ms,
                mode: settings.watch.auto_run_mode.clone(),
            })
            .await?;
    }

    let address: SocketAddr = format!("{host}:{port}").parse()?;
    let listener = tokio::net::TcpListener::bind(address).await?;
    tracing::info!("Hybrid-Backend lauscht auf http://{address}");
    axum::serve(listener, routes::router(state)).await?;
    Ok(())
}

async fn validate(root: PathBuf) -> Result<()> {
    let root = canonicalize_root(root)?;
    let state = build_state(root).await?;
    let validation = state.validate_with_settings().await?;
    println!("{}", serde_json::to_string_pretty(&validation)?);
    if !validation.ok {
        bail!("Die Validierung hat Fehler gefunden.");
    }
    Ok(())
}

async fn run_once(
    root: PathBuf,
    dat_file: Option<PathBuf>,
    mode: CliRunMode,
    force_rebuild: bool,
) -> Result<()> {
    let root = canonicalize_root(root)?;
    let state = build_state(root).await?;
    let settings = state.current_settings().await;

    let request = models::RunRequest {
        source_workbook_path: settings.source_workbook_path,
        bas_folder_path: settings.bas_folder_path,
        dat_file_path: dat_file
            .unwrap_or_else(|| PathBuf::from(settings.dat_file_path))
            .to_string_lossy()
            .into_owned(),
        output_dir: settings.output_dir,
        macro_name: settings.macro_name,
        mode: mode.into(),
        force_rebuild_workbook: force_rebuild || settings.auto_rebuild_workbook,
    };

    let job = state.run_job(request).await?;
    println!("{}", serde_json::to_string_pretty(&job)?);
    if job.status != models::JobStatus::Success {
        bail!("Der Lauf wurde nicht erfolgreich abgeschlossen.");
    }
    Ok(())
}

async fn build_state(root: PathBuf) -> Result<std::sync::Arc<AppState>> {
    let paths = AppPaths::new(root)?;
    AppState::load(paths).await
}

fn canonicalize_root(root: PathBuf) -> Result<PathBuf> {
    let resolved = if root.is_absolute() {
        root
    } else {
        std::env::current_dir()?.join(root)
    };
    Ok(resolved)
}

fn init_tracing() {
    let filter = EnvFilter::try_from_default_env()
        .unwrap_or_else(|_| EnvFilter::new("info,tower_http=info,hybrid_approach=info"));

    tracing_subscriber::fmt()
        .with_env_filter(filter)
        .with_target(false)
        .compact()
        .init();
}
