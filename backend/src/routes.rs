use crate::models::{AppSettings, OpenFolderRequest, RunRequest, WatchRequest};
use crate::services::AppState;
use anyhow::Error;
use axum::extract::{Query, State};
use axum::http::{header, StatusCode, Uri};
use axum::response::{Html, IntoResponse, Response};
use axum::routing::{get, post};
use axum::{Json, Router};
use serde::Deserialize;
use serde_json::json;
use std::path::PathBuf;
use std::sync::Arc;
use tower_http::cors::{Any, CorsLayer};
use tower_http::trace::TraceLayer;

#[derive(Debug)]
struct ApiError(Error);

impl<E> From<E> for ApiError
where
    E: Into<Error>,
{
    fn from(error: E) -> Self {
        Self(error.into())
    }
}

impl IntoResponse for ApiError {
    fn into_response(self) -> Response {
        let body = Json(json!({
            "ok": false,
            "error": self.0.to_string()
        }));
        (StatusCode::INTERNAL_SERVER_ERROR, body).into_response()
    }
}

#[derive(Debug, Deserialize)]
struct FilesQuery {
    path: Option<String>,
}

pub fn router(state: Arc<AppState>) -> Router {
    let api_router = Router::new()
        .route("/health", get(health))
        .route("/dashboard", get(dashboard))
        .route("/settings", get(get_settings).put(save_settings))
        .route("/validate", get(validate_current).post(validate_request))
        .route("/run", post(run_job))
        .route("/jobs", get(get_jobs))
        .route("/logs", get(get_logs))
        .route("/files", get(get_files))
        .route("/open-folder", post(open_folder))
        .route("/watch/start", post(start_watch))
        .route("/watch/stop", post(stop_watch));

    Router::new()
        .nest("/api", api_router)
        .fallback(get(serve_frontend))
        .layer(
            CorsLayer::new()
                .allow_origin(Any)
                .allow_methods(Any)
                .allow_headers(Any),
        )
        .layer(TraceLayer::new_for_http())
        .with_state(state)
}

async fn health() -> Json<serde_json::Value> {
    Json(json!({
        "ok": true,
        "service": "hybrid-approach-backend"
    }))
}

async fn dashboard(State(state): State<Arc<AppState>>) -> Result<Json<serde_json::Value>, ApiError> {
    Ok(Json(json!({
        "ok": true,
        "data": state.dashboard().await
    })))
}

async fn get_settings(State(state): State<Arc<AppState>>) -> Result<Json<serde_json::Value>, ApiError> {
    Ok(Json(json!({
        "ok": true,
        "data": state.current_settings().await
    })))
}

async fn save_settings(
    State(state): State<Arc<AppState>>,
    Json(settings): Json<AppSettings>,
) -> Result<Json<serde_json::Value>, ApiError> {
    let saved = state.save_settings(settings).await?;
    if saved.watch.enabled {
        state
            .start_watch(WatchRequest {
                watch_folder: saved.watch.watch_folder.clone(),
                pattern: saved.watch.pattern.clone(),
                debounce_ms: saved.watch.debounce_ms,
                mode: saved.watch.auto_run_mode.clone(),
            })
            .await?;
    } else {
        state.stop_watch().await?;
    }

    let validation = state.validate_with_settings().await?;
    Ok(Json(json!({
        "ok": true,
        "data": saved,
        "validation": validation
    })))
}

async fn validate_current(
    State(state): State<Arc<AppState>>,
) -> Result<Json<serde_json::Value>, ApiError> {
    Ok(Json(json!({
        "ok": true,
        "data": state.validate_with_settings().await?
    })))
}

async fn validate_request(
    State(state): State<Arc<AppState>>,
    Json(request): Json<RunRequest>,
) -> Result<Json<serde_json::Value>, ApiError> {
    Ok(Json(json!({
        "ok": true,
        "data": state.validate_request(&request).await?
    })))
}

async fn run_job(
    State(state): State<Arc<AppState>>,
    Json(request): Json<RunRequest>,
) -> Result<Json<serde_json::Value>, ApiError> {
    Ok(Json(json!({
        "ok": true,
        "data": state.run_job(request).await?
    })))
}

async fn get_jobs(State(state): State<Arc<AppState>>) -> Result<Json<serde_json::Value>, ApiError> {
    Ok(Json(json!({
        "ok": true,
        "data": state.jobs_snapshot().await
    })))
}

async fn get_logs(State(state): State<Arc<AppState>>) -> Result<Json<serde_json::Value>, ApiError> {
    Ok(Json(json!({
        "ok": true,
        "data": state.read_logs().await?
    })))
}

async fn get_files(
    State(state): State<Arc<AppState>>,
    Query(query): Query<FilesQuery>,
) -> Result<Json<serde_json::Value>, ApiError> {
    Ok(Json(json!({
        "ok": true,
        "data": state.list_directory(query.path).await?
    })))
}

async fn open_folder(
    State(state): State<Arc<AppState>>,
    Json(request): Json<OpenFolderRequest>,
) -> Result<Json<serde_json::Value>, ApiError> {
    state.open_folder(request).await?;
    Ok(Json(json!({ "ok": true })))
}

async fn start_watch(
    State(state): State<Arc<AppState>>,
    Json(request): Json<WatchRequest>,
) -> Result<Json<serde_json::Value>, ApiError> {
    state.start_watch(request).await?;
    Ok(Json(json!({ "ok": true })))
}

async fn stop_watch(State(state): State<Arc<AppState>>) -> Result<Json<serde_json::Value>, ApiError> {
    state.stop_watch().await?;
    Ok(Json(json!({ "ok": true })))
}

async fn serve_frontend(
    State(state): State<Arc<AppState>>,
    uri: Uri,
) -> Result<Response, ApiError> {
    let requested = uri.path().trim_start_matches('/');
    if requested.contains("..") {
        return Ok((StatusCode::FORBIDDEN, "Unzulässiger Pfad.").into_response());
    }

    let web_root = state.paths.web_dir.clone();
    let mut candidate = if requested.is_empty() {
        web_root.join("index.html")
    } else {
        web_root.join(PathBuf::from(requested))
    };

    if candidate.is_dir() {
        candidate = candidate.join("index.html");
    }

    let fallback = web_root.join("index.html");
    let file_path = if candidate.exists() && candidate.is_file() {
        candidate
    } else if fallback.exists() {
        fallback
    } else {
        let html = r#"
<!doctype html>
<html lang="de">
  <head>
    <meta charset="utf-8" />
    <title>Hybrid Approach</title>
    <style>
      body { font-family: "Segoe UI", sans-serif; background: #08111f; color: #f3f6fb; margin: 0; padding: 48px; }
      .card { max-width: 760px; padding: 32px; border-radius: 24px; background: linear-gradient(160deg, rgba(22,38,60,.95), rgba(10,18,34,.98)); box-shadow: 0 24px 80px rgba(0,0,0,.35); }
      h1 { margin-top: 0; font-size: 32px; }
      p { line-height: 1.6; color: #d5deeb; }
      code { background: rgba(255,255,255,.08); padding: 2px 6px; border-radius: 6px; }
    </style>
  </head>
  <body>
    <div class="card">
      <h1>Frontend-Build fehlt</h1>
      <p>Der Backend-Dienst läuft, aber im Ordner <code>runtime/web</code> wurde noch kein gebautes Frontend gefunden.</p>
      <p>Starten Sie das Projekt über <code>scripts\\start-hybrid.ps1</code>, damit die React-Oberfläche gebaut und automatisch ausgeliefert wird.</p>
    </div>
  </body>
</html>
"#;
        return Ok(Html(html).into_response());
    };

    let bytes = tokio::fs::read(&file_path).await?;
    let mime = mime_guess::from_path(&file_path).first_or_octet_stream();
    let response = (
        [(header::CONTENT_TYPE, mime.as_ref())],
        bytes,
    )
        .into_response();
    Ok(response)
}
