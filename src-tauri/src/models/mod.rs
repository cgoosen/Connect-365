use std::collections::HashMap;
use std::sync::Mutex;

pub mod module;
pub mod session;
pub mod settings;

pub use module::{ModuleCatalogEntry, ModuleStatus, PrereqResult};
pub use session::{LaunchRequest, LaunchResult, SessionHandle, SessionInfo, TerminateAllResult};
pub use settings::Settings;

/// Shared mutable state managed by Tauri (Send + Sync via Mutex).
pub struct AppState {
    /// Keyed by session_id; holds the live process handle.
    pub sessions: Mutex<HashMap<String, SessionHandle>>,
}

impl Default for AppState {
    fn default() -> Self {
        Self {
            sessions: Mutex::new(HashMap::new()),
        }
    }
}
