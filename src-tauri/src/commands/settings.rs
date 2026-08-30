use crate::models::Settings;
use tauri::Manager;
use tauri_plugin_opener::OpenerExt;

// Only PSGallery and the official PowerShell download page are allowed (§11).
const ALLOWED_URL_PREFIXES: &[&str] = &[
    "https://www.powershellgallery.com/packages/",
    "https://learn.microsoft.com/en-us/powershell/scripting/install/install-powershell",
];

/// Reads settings.json from the Tauri app data dir; returns defaults if absent (§9).
#[tauri::command]
pub fn get_settings(app: tauri::AppHandle) -> Result<Settings, String> {
    let path = settings_path(&app)?;
    if !path.exists() {
        return Ok(Settings::default());
    }
    let json = std::fs::read_to_string(&path).map_err(|e| format!("Read settings failed: {e}"))?;
    serde_json::from_str(&json).map_err(|e| format!("Malformed settings.json: {e}"))
}

/// Persists settings to the Tauri app data dir (§9). No secrets are written.
#[tauri::command]
pub fn set_settings(app: tauri::AppHandle, settings: Settings) -> Result<(), String> {
    let path = settings_path(&app)?;
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)
            .map_err(|e| format!("Could not create app data dir: {e}"))?;
    }
    let json =
        serde_json::to_string_pretty(&settings).map_err(|e| format!("Serialise error: {e}"))?;
    std::fs::write(&path, json).map_err(|e| format!("Write settings failed: {e}"))
}

/// Opens a PSGallery URL in the default browser.
/// Validates the URL against the known allowlist before opening (§11).
#[tauri::command]
pub async fn open_install_link(app: tauri::AppHandle, url: String) -> Result<(), String> {
    let allowed = ALLOWED_URL_PREFIXES
        .iter()
        .any(|prefix| url.starts_with(prefix));

    if !allowed {
        return Err(format!(
            "URL not in allowed list: {url}"
        ));
    }

    app.opener()
        .open_url(&url, None::<String>)
        .map_err(|e| format!("Failed to open URL: {e}"))
}

fn settings_path(app: &tauri::AppHandle) -> Result<std::path::PathBuf, String> {
    app.path()
        .app_data_dir()
        .map(|d| d.join("settings.json"))
        .map_err(|e| format!("App data dir unavailable: {e}"))
}
