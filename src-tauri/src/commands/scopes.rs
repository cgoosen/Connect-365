use serde::{Deserialize, Serialize};
use tauri::Manager;

// Vendored at compile time; the app works fully offline from this copy (§7).
const VENDORED_SCOPES: &str = include_str!("../../../src/assets/graph-scopes.json");

/// One scope entry from graph-scopes.json (§7 schema).
#[derive(Debug, Serialize, Deserialize)]
pub struct ScopeItem {
    pub value: String,
    #[serde(rename = "type")]
    pub scope_type: String,
    pub description: String,
}

/// One of the five curated quick-connect profiles (§7).
#[derive(Debug, Serialize)]
pub struct ScopeProfile {
    pub id: String,
    pub label: String,
    pub scopes: Vec<String>,
}

/// Result of check_scope_list_update.
#[derive(Debug, Serialize)]
pub struct ScopeUpdateResult {
    pub updated: bool,
    pub new_version: Option<String>,
}

// ── Tauri commands ─────────────────────────────────────────────────────────────

/// Returns the five curated quick-connect profiles (§7 table).
#[tauri::command]
pub fn get_scope_profiles() -> Vec<ScopeProfile> {
    vec![
        ScopeProfile {
            id: "user-admin".into(),
            label: "User Administration".into(),
            scopes: vec![
                "User.ReadWrite.All".into(),
                "Group.ReadWrite.All".into(),
                "Directory.Read.All".into(),
            ],
        },
        ScopeProfile {
            id: "identity-access".into(),
            label: "Identity & Access".into(),
            scopes: vec![
                "Policy.ReadWrite.ConditionalAccess".into(),
                "Policy.Read.All".into(),
                "RoleManagement.ReadWrite.Directory".into(),
            ],
        },
        ScopeProfile {
            id: "device-intune".into(),
            label: "Device & Intune".into(),
            scopes: vec![
                "DeviceManagementManagedDevices.ReadWrite.All".into(),
                "DeviceManagementConfiguration.ReadWrite.All".into(),
            ],
        },
        ScopeProfile {
            id: "applications".into(),
            label: "Applications".into(),
            scopes: vec![
                "Application.ReadWrite.All".into(),
                "AppRoleAssignment.ReadWrite.All".into(),
            ],
        },
        ScopeProfile {
            id: "reports-audit".into(),
            label: "Reports & Audit".into(),
            scopes: vec![
                "AuditLog.Read.All".into(),
                "Reports.Read.All".into(),
                "Directory.Read.All".into(),
            ],
        },
    ]
}

/// Returns the full scope list. Prefers a cached updated copy in the app data
/// dir (written by check_scope_list_update) over the vendored compile-time copy.
#[tauri::command]
pub fn get_all_scopes(app: tauri::AppHandle) -> Result<Vec<ScopeItem>, String> {
    let json_str = load_scope_json(&app);
    parse_scope_items(&json_str)
}

/// Non-blocking background check against the project GitHub raw URL (§7).
/// Validates schema before adopting; the update takes effect on the next launch.
#[tauri::command]
pub async fn check_scope_list_update(
    app: tauri::AppHandle,
) -> Result<ScopeUpdateResult, String> {
    check_update_inner(app).await
}

// ── Internal helpers ───────────────────────────────────────────────────────────

fn load_scope_json(app: &tauri::AppHandle) -> String {
    // Try cached updated copy first
    if let Ok(data_dir) = app.path().app_data_dir() {
        let cached = data_dir.join("graph-scopes.json");
        if let Ok(contents) = std::fs::read_to_string(&cached) {
            return contents;
        }
    }
    VENDORED_SCOPES.to_string()
}

fn parse_scope_items(json_str: &str) -> Result<Vec<ScopeItem>, String> {
    let root: serde_json::Value =
        serde_json::from_str(json_str).map_err(|e| format!("Invalid scopes JSON: {e}"))?;

    let arr = root["scopes"]
        .as_array()
        .ok_or("scopes.json missing 'scopes' array")?;

    arr.iter()
        .map(|v| serde_json::from_value(v.clone()).map_err(|e| format!("Malformed scope entry: {e}")))
        .collect()
}

fn scope_version(json_str: &str) -> i64 {
    serde_json::from_str::<serde_json::Value>(json_str)
        .ok()
        .and_then(|v| v["version"].as_i64())
        .unwrap_or(0)
}

fn validate_remote_schema(remote: &serde_json::Value) -> Result<(), String> {
    let scopes = remote["scopes"]
        .as_array()
        .ok_or("Remote scopes.json missing 'scopes' array")?;

    if scopes.is_empty() {
        return Err("Remote scopes.json has empty 'scopes' array".into());
    }

    for scope in scopes {
        if scope["value"].as_str().is_none() || scope["type"].as_str().is_none() {
            return Err("Remote scopes.json has a malformed scope entry (missing value or type)".into());
        }
    }

    Ok(())
}

// Compiled only when scope-update feature is active (default ON, see Cargo.toml §13)
#[cfg(feature = "scope-update")]
async fn check_update_inner(app: tauri::AppHandle) -> Result<ScopeUpdateResult, String> {
    // Hardcoded URL — not user-configurable (§7 supply-chain note)
    const REMOTE_URL: &str =
        "https://raw.githubusercontent.com/cgoosen/Connect-365/main/src/assets/graph-scopes.json";

    let current_json = load_scope_json(&app);
    let current_version = scope_version(&current_json);

    let client = reqwest::Client::builder()
        .timeout(std::time::Duration::from_secs(10))
        .user_agent("Connect-365")
        .build()
        .map_err(|e| format!("HTTP client error: {e}"))?;

    let response = client
        .get(REMOTE_URL)
        .send()
        .await
        .map_err(|e| format!("Fetch error: {e}"))?;

    let remote: serde_json::Value = response
        .json()
        .await
        .map_err(|e| format!("Invalid JSON from remote: {e}"))?;

    let remote_version = remote["version"].as_i64().unwrap_or(0);

    if remote_version <= current_version {
        return Ok(ScopeUpdateResult {
            updated: false,
            new_version: None,
        });
    }

    // Validate before adopting (§7: inert data but must match expected schema)
    validate_remote_schema(&remote)?;

    let data_dir = app
        .path()
        .app_data_dir()
        .map_err(|e| format!("App data dir error: {e}"))?;

    std::fs::create_dir_all(&data_dir).map_err(|e| format!("Could not create app data dir: {e}"))?;

    let cache_path = data_dir.join("graph-scopes.json");
    let serialised =
        serde_json::to_string(&remote).map_err(|e| format!("Serialise error: {e}"))?;

    std::fs::write(&cache_path, serialised)
        .map_err(|e| format!("Failed to write cached scopes: {e}"))?;

    Ok(ScopeUpdateResult {
        updated: true,
        new_version: Some(remote_version.to_string()),
    })
}

#[cfg(not(feature = "scope-update"))]
async fn check_update_inner(_app: tauri::AppHandle) -> Result<ScopeUpdateResult, String> {
    Ok(ScopeUpdateResult {
        updated: false,
        new_version: None,
    })
}
