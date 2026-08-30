use serde::{Deserialize, Serialize};
use std::path::PathBuf;
use std::process::Child;

/// Internal session state — not serialised; lives in AppState.
pub struct SessionHandle {
    pub session_id: String,
    pub label: String,
    pub started_at: String,
    /// The spawned terminal process.
    pub process: Child,
    /// Temp .ps1 path; cleaned up on terminate.
    pub script_path: PathBuf,
}

/// Serialisable projection returned to the frontend by list_sessions.
#[derive(Debug, Serialize, Deserialize)]
pub struct SessionInfo {
    pub session_id: String,
    pub label: String,
    pub started_at: String,
    pub alive: bool,
}

/// Inputs for launch_session (§11 command surface).
#[derive(Debug, Deserialize)]
pub struct LaunchRequest {
    pub upn: Option<String>,
    pub services: Vec<String>,
    pub spo_url: Option<String>,
    pub pnp_url: Option<String>,
    pub graph_scopes: Vec<String>,
}

/// Return value of launch_session.
#[derive(Debug, Serialize)]
pub struct LaunchResult {
    pub session_id: String,
    pub status: String,
}

/// Return value of terminate_all_sessions.
#[derive(Debug, Serialize)]
pub struct TerminateAllResult {
    pub ok: bool,
    pub count: usize,
}
