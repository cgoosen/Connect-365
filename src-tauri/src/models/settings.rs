use serde::{Deserialize, Serialize};

/// User-convenience config persisted to the Tauri app data dir (§9).
/// Nothing sensitive — no credentials, no tokens.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Settings {
    /// Service IDs that were checked on last use.
    pub last_selected_services: Vec<String>,
    pub last_upn: Option<String>,
    pub last_spo_url: Option<String>,
    pub last_pnp_url: Option<String>,
    /// Opt-in scope-list auto-update toggle (default OFF per §7).
    pub auto_update_scopes: bool,
    /// Whether the window minimises to the system tray after Connect (§5).
    pub minimise_to_tray_on_connect: bool,
}

impl Default for Settings {
    fn default() -> Self {
        Self {
            last_selected_services: vec![],
            last_upn: None,
            last_spo_url: None,
            last_pnp_url: None,
            auto_update_scopes: false,
            minimise_to_tray_on_connect: true,
        }
    }
}
