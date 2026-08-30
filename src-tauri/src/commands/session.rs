use crate::models::{
    AppState, LaunchRequest, LaunchResult, ModuleCatalogEntry, SessionHandle, SessionInfo,
    TerminateAllResult,
};
use std::collections::HashSet;
use std::process::Command;
use tauri::State;

// ── Module catalogue (§4) ──────────────────────────────────────────────────────

pub fn catalogue() -> Vec<ModuleCatalogEntry> {
    vec![
        ModuleCatalogEntry {
            id: "entra".into(),
            display_name: "Entra ID".into(),
            module_name: "Microsoft.Entra".into(),
            connect_cmdlet: "Connect-Entra".into(),
            gallery_url: "https://www.powershellgallery.com/packages/Microsoft.Entra".into(),
            requires_spo_url: false,
            requires_pnp_url: false,
            // Unlike Microsoft.Graph, Microsoft.Entra's submodules
            // (Microsoft.Entra.Users, .Groups, .Applications, etc.) don't
            // reliably autoload from just Connect-Entra — the module's own
            // docs/community guidance say to import it explicitly. Importing
            // only .Authentication left every other Entra cmdlet unavailable.
            import_module_name: None,
        },
        ModuleCatalogEntry {
            id: "graph".into(),
            display_name: "Microsoft Graph".into(),
            module_name: "Microsoft.Graph".into(),
            connect_cmdlet: "Connect-MgGraph -Scopes @(...)".into(),
            gallery_url: "https://www.powershellgallery.com/packages/Microsoft.Graph".into(),
            requires_spo_url: false,
            requires_pnp_url: false,
            // Importing the full Microsoft.Graph meta-package eagerly loads
            // ~30 submodules and can take minutes; Connect-MgGraph only
            // needs Microsoft.Graph.Authentication. Other Mg* cmdlets
            // autoload their own submodule on first use.
            import_module_name: Some("Microsoft.Graph.Authentication".into()),
        },
        ModuleCatalogEntry {
            id: "exchange".into(),
            display_name: "Exchange Online".into(),
            module_name: "ExchangeOnlineManagement".into(),
            connect_cmdlet: "Connect-ExchangeOnline".into(),
            gallery_url: "https://www.powershellgallery.com/packages/ExchangeOnlineManagement".into(),
            requires_spo_url: false,
            requires_pnp_url: false,
            import_module_name: None,
        },
        ModuleCatalogEntry {
            id: "compliance".into(),
            display_name: "Security & Compliance".into(),
            // Shares the module with Exchange Online (§4 note)
            module_name: "ExchangeOnlineManagement".into(),
            connect_cmdlet: "Connect-IPPSSession".into(),
            gallery_url: "https://www.powershellgallery.com/packages/ExchangeOnlineManagement".into(),
            requires_spo_url: false,
            requires_pnp_url: false,
            import_module_name: None,
        },
        ModuleCatalogEntry {
            id: "teams".into(),
            display_name: "Microsoft Teams".into(),
            module_name: "MicrosoftTeams".into(),
            connect_cmdlet: "Connect-MicrosoftTeams".into(),
            gallery_url: "https://www.powershellgallery.com/packages/MicrosoftTeams".into(),
            requires_spo_url: false,
            requires_pnp_url: false,
            import_module_name: None,
        },
        ModuleCatalogEntry {
            id: "sharepoint".into(),
            display_name: "SharePoint Online".into(),
            module_name: "Microsoft.Online.SharePoint.PowerShell".into(),
            connect_cmdlet: "Connect-SPOService -Url <admin URL>".into(),
            gallery_url: "https://www.powershellgallery.com/packages/Microsoft.Online.SharePoint.PowerShell".into(),
            requires_spo_url: true,
            requires_pnp_url: false,
            import_module_name: None,
        },
        ModuleCatalogEntry {
            id: "aip".into(),
            display_name: "Information Protection".into(),
            module_name: "AIPService".into(),
            connect_cmdlet: "Connect-AIPService".into(),
            gallery_url: "https://www.powershellgallery.com/packages/AIPService".into(),
            requires_spo_url: false,
            requires_pnp_url: false,
            import_module_name: None,
        },
        ModuleCatalogEntry {
            id: "powerplatform-admin".into(),
            display_name: "Power Platform (Admin)".into(),
            module_name: "Microsoft.PowerApps.Administration.PowerShell".into(),
            connect_cmdlet: "Add-PowerAppsAccount".into(),
            gallery_url: "https://www.powershellgallery.com/packages/Microsoft.PowerApps.Administration.PowerShell".into(),
            requires_spo_url: false,
            requires_pnp_url: false,
            import_module_name: None,
        },
        ModuleCatalogEntry {
            id: "powerplatform-maker".into(),
            display_name: "Power Platform (Maker)".into(),
            module_name: "Microsoft.PowerApps.PowerShell".into(),
            connect_cmdlet: "Add-PowerAppsAccount".into(),
            gallery_url: "https://www.powershellgallery.com/packages/Microsoft.PowerApps.PowerShell".into(),
            requires_spo_url: false,
            requires_pnp_url: false,
            import_module_name: None,
        },
        ModuleCatalogEntry {
            id: "powerbi".into(),
            display_name: "Power BI".into(),
            module_name: "MicrosoftPowerBIMgmt".into(),
            connect_cmdlet: "Connect-PowerBIServiceAccount".into(),
            gallery_url: "https://www.powershellgallery.com/packages/MicrosoftPowerBIMgmt".into(),
            requires_spo_url: false,
            requires_pnp_url: false,
            import_module_name: None,
        },
        ModuleCatalogEntry {
            id: "pnp".into(),
            display_name: "PnP PowerShell".into(),
            module_name: "PnP.PowerShell".into(),
            connect_cmdlet: "Connect-PnPOnline -Url <url> -Interactive".into(),
            gallery_url: "https://www.powershellgallery.com/packages/PnP.PowerShell".into(),
            requires_spo_url: false,
            requires_pnp_url: true,
            import_module_name: None,
        },
    ]
}

// ── Tauri commands ─────────────────────────────────────────────────────────────

#[tauri::command]
pub fn get_modules() -> Vec<ModuleCatalogEntry> {
    catalogue()
}

#[tauri::command]
pub async fn launch_session(
    state: State<'_, AppState>,
    request: LaunchRequest,
) -> Result<LaunchResult, String> {
    let cat = catalogue();
    let script_content = build_connection_script(&request, &cat);

    let session_id = uuid::Uuid::new_v4().to_string();
    let script_path = std::env::temp_dir().join(format!("connect365-{}.ps1", session_id));

    std::fs::write(&script_path, &script_content)
        .map_err(|e| format!("Failed to write session script: {e}"))?;

    let process = spawn_terminal(&script_path)?;

    let started_at = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs()
        .to_string();

    let label = build_session_label(&request, &cat);

    let handle = SessionHandle {
        session_id: session_id.clone(),
        label,
        started_at,
        process,
        script_path,
    };

    state.sessions.lock().unwrap().insert(session_id.clone(), handle);

    Ok(LaunchResult {
        session_id,
        status: "launched".into(),
    })
}

#[tauri::command]
pub fn list_sessions(state: State<'_, AppState>) -> Vec<SessionInfo> {
    let mut sessions = state.sessions.lock().unwrap();
    sessions
        .values_mut()
        .map(|handle| {
            let alive = match handle.process.try_wait() {
                Ok(None) => true, // process still running (Linux/Windows terminal)
                // Process exited (e.g. macOS osascript): use script file as proxy
                Ok(Some(_)) | Err(_) => handle.script_path.exists(),
            };
            SessionInfo {
                session_id: handle.session_id.clone(),
                label: handle.label.clone(),
                started_at: handle.started_at.clone(),
                alive,
            }
        })
        .collect()
}

#[tauri::command]
pub fn terminate_session(
    state: State<'_, AppState>,
    session_id: String,
) -> Result<(), String> {
    let mut sessions = state.sessions.lock().unwrap();
    let mut handle = sessions
        .remove(&session_id)
        .ok_or_else(|| format!("Session not found: {session_id}"))?;

    let _ = handle.process.kill();
    let _ = handle.process.wait(); // reap zombie on Unix
    let _ = std::fs::remove_file(&handle.script_path);
    Ok(())
}

#[tauri::command]
pub fn terminate_all_sessions(
    state: State<'_, AppState>,
) -> Result<TerminateAllResult, String> {
    let count = terminate_all_internal(&state);
    Ok(TerminateAllResult { ok: true, count })
}

/// Kills every tracked session process and removes its temp script.
/// Shared by `terminate_all_sessions`, the tray menu, and `quit_app`.
pub(crate) fn terminate_all_internal(state: &AppState) -> usize {
    let mut sessions = state.sessions.lock().unwrap();
    let count = sessions.len();
    for (_, mut handle) in sessions.drain() {
        let _ = handle.process.kill();
        let _ = handle.process.wait();
        let _ = std::fs::remove_file(&handle.script_path);
    }
    count
}

/// Terminates all tracked sessions and exits the app (§5 "Exit"). The caller
/// (JS Exit nav item) is responsible for confirming with the user first when
/// sessions are active — this command just executes the decision.
#[tauri::command]
pub fn quit_app(state: State<'_, AppState>, app: tauri::AppHandle) -> Result<(), String> {
    terminate_all_internal(&state);
    app.exit(0);
    Ok(())
}

// ── Script builder ─────────────────────────────────────────────────────────────

fn build_connection_script(request: &LaunchRequest, cat: &[ModuleCatalogEntry]) -> String {
    // Import + connect per service, interleaved in selection order — NOT all
    // imports followed by all connects. Some modules take a while to import;
    // front-loading every import meant the very first auth prompt couldn't
    // appear until every selected module had finished loading. Interleaving
    // lets each service prompt for auth as soon as its own module is ready.
    let mut steps: Vec<String> = Vec::new();
    let mut seen_imports: HashSet<String> = HashSet::new();
    let mut powerplatform_authed = false;

    for service_id in &request.services {
        let Some(entry) = cat.iter().find(|e| e.id == *service_id) else {
            continue;
        };

        let import_name = entry
            .import_module_name
            .as_deref()
            .unwrap_or(entry.module_name.as_str());

        if seen_imports.insert(import_name.to_string()) {
            // First-time module loads can take 30-60s (PowerShell building its
            // command analysis cache) with zero output from Import-Module itself
            // — without this the terminal looks frozen. Echo before, not after.
            steps.push(format!(
                "Write-Host \"→ Loading module '{import_name}'…\" -ForegroundColor Yellow"
            ));
            steps.push(format!("Import-Module '{}' -ErrorAction Stop", import_name));
        }

        steps.push(format!(
            "Write-Host \"→ Connecting: {}…\" -ForegroundColor Yellow",
            entry.display_name
        ));
        steps.push(build_connect_line(service_id, request, &mut powerplatform_authed));
    }

    let service_labels: Vec<&str> = request
        .services
        .iter()
        .filter_map(|id| cat.iter().find(|e| e.id == *id))
        .map(|e| e.display_name.as_str())
        .collect();

    let upn_display = request
        .upn
        .as_deref()
        .map(|u| format!(" as {u}"))
        .unwrap_or_default();

    format!(
        "# Connect-365 session script\n\
         # Services: {services}\n\
         # DO NOT EDIT — deleted when session is terminated via Connect-365\n\n\
         $ErrorActionPreference = 'Stop'\n\n\
         Write-Host \"Connect-365: Starting session…\" -ForegroundColor Cyan\n\
         Write-Host \"First-time module loads can take 30-60s while PowerShell builds its command cache — this is normal, later launches are faster.\" -ForegroundColor DarkGray\n\n\
         {steps}\n\n\
         Write-Host \"`nConnect-365: Session ready{upn_display}.\" -ForegroundColor Green\n\
         Write-Host \"Connected: {services}\" -ForegroundColor Cyan\n",
        services = service_labels.join(", "),
        steps = steps.join("\n"),
        upn_display = upn_display,
    )
}

fn build_connect_line(
    service_id: &str,
    req: &LaunchRequest,
    powerplatform_authed: &mut bool,
) -> String {
    let upn = req.upn.as_deref().unwrap_or("");

    match service_id {
        // Connect-Entra is an alias for Connect-MgGraph and shares its
        // parameter sets — there is no UPN/account-hint parameter for
        // interactive sign-in (unlike Connect-ExchangeOnline/-MicrosoftTeams).
        // The account is chosen in the browser sign-in prompt instead.
        "entra" => "Connect-Entra".into(),
        "graph" => {
            if req.graph_scopes.is_empty() {
                "Connect-MgGraph".into()
            } else {
                let scopes = req
                    .graph_scopes
                    .iter()
                    .map(|s| format!("'{}'", ps_escape(s)))
                    .collect::<Vec<_>>()
                    .join(", ");
                format!("Connect-MgGraph -Scopes @({scopes})")
            }
        }
        "exchange" => {
            if upn.is_empty() {
                "Connect-ExchangeOnline".into()
            } else {
                format!("Connect-ExchangeOnline -UserPrincipalName '{}'", ps_escape(upn))
            }
        }
        "compliance" => {
            if upn.is_empty() {
                "Connect-IPPSSession".into()
            } else {
                format!("Connect-IPPSSession -UserPrincipalName '{}'", ps_escape(upn))
            }
        }
        "teams" => {
            if upn.is_empty() {
                "Connect-MicrosoftTeams".into()
            } else {
                format!("Connect-MicrosoftTeams -AccountId '{}'", ps_escape(upn))
            }
        }
        "sharepoint" => {
            let url = req.spo_url.as_deref().unwrap_or("");
            format!("Connect-SPOService -Url '{}'", ps_escape(url))
        }
        "aip" => "Connect-AIPService".into(),
        "powerplatform-admin" | "powerplatform-maker" => {
            // Both modules authenticate with Add-PowerAppsAccount; run it once (§4 note)
            if *powerplatform_authed {
                "# Power Platform: already authenticated above".into()
            } else {
                *powerplatform_authed = true;
                "Add-PowerAppsAccount".into()
            }
        }
        "powerbi" => "Connect-PowerBIServiceAccount".into(),
        "pnp" => {
            let url = req
                .pnp_url
                .as_deref()
                .or(req.spo_url.as_deref())
                .unwrap_or("");
            if url.is_empty() {
                "Connect-PnPOnline -Interactive".into()
            } else {
                format!("Connect-PnPOnline -Url '{}' -Interactive", ps_escape(url))
            }
        }
        other => format!("# Unknown service id '{other}' — skipped"),
    }
}

/// Escape single-quotes for PowerShell single-quoted strings.
fn ps_escape(s: &str) -> String {
    s.replace('\'', "''")
}

fn build_session_label(request: &LaunchRequest, cat: &[ModuleCatalogEntry]) -> String {
    let services: Vec<&str> = request
        .services
        .iter()
        .filter_map(|id| cat.iter().find(|e| e.id == *id))
        .map(|e| e.display_name.as_str())
        .collect();

    match &request.upn {
        Some(upn) => format!("{} ({})", services.join(", "), upn),
        None => services.join(", "),
    }
}

// ── Terminal launcher ──────────────────────────────────────────────────────────

pub(crate) fn spawn_terminal(script_path: &std::path::Path) -> Result<std::process::Child, String> {
    let path_str = script_path.to_string_lossy().into_owned();
    spawn_terminal_inner(&path_str)
}

#[cfg(target_os = "windows")]
fn spawn_terminal_inner(path_str: &str) -> Result<std::process::Child, String> {
    // Prefer Windows Terminal; fall back to a new cmd console
    let wt = Command::new("wt.exe")
        .args(["--window", "new", "pwsh.exe", "-NoExit", "-File", path_str])
        .spawn();

    if let Ok(child) = wt {
        return Ok(child);
    }

    // cmd /c start opens a new console window; cmd itself exits quickly but the
    // console window (conhost) stays open. v1 tracks this cmd child.
    Command::new("cmd.exe")
        .args([
            "/c",
            "start",
            "Connect-365 Session",
            "pwsh.exe",
            "-NoExit",
            "-File",
            path_str,
        ])
        .spawn()
        .map_err(|e| format!("Failed to launch terminal on Windows: {e}"))
}

#[cfg(target_os = "macos")]
fn spawn_terminal_inner(path_str: &str) -> Result<std::process::Child, String> {
    // Escape path for AppleScript double-quoted string
    let escaped = path_str.replace('\\', "\\\\").replace('"', "\\\"");

    // Prefer iTerm2 if installed
    let iterm_paths = ["/Applications/iTerm2.app", "/Applications/iTerm.app"];
    for iterm_path in &iterm_paths {
        if std::path::Path::new(iterm_path).exists() {
            let app = if iterm_path.contains("iTerm2") {
                "iTerm2"
            } else {
                "iTerm"
            };
            let script = format!(
                "tell application \"{app}\"\n\
                 \tactivate\n\
                 \tcreate window with default profile\n\
                 \ttell current session of current window\n\
                 \t\twrite text \"pwsh -NoExit -File \\\"{escaped}\\\"\"\n\
                 \tend tell\n\
                 end tell"
            );
            if let Ok(child) = Command::new("osascript").arg("-e").arg(&script).spawn() {
                return Ok(child);
            }
        }
    }

    // Fall back to Terminal.app
    let script = format!(
        "tell application \"Terminal\"\n\
         \tactivate\n\
         \tdo script \"pwsh -NoExit -File \\\"{escaped}\\\"\"\n\
         end tell"
    );
    Command::new("osascript")
        .arg("-e")
        .arg(&script)
        .spawn()
        .map_err(|e| format!("Failed to launch Terminal.app: {e}"))
}

#[cfg(not(any(target_os = "windows", target_os = "macos")))]
fn spawn_terminal_inner(path_str: &str) -> Result<std::process::Child, String> {
    // Try common Linux terminal emulators in order of preference
    let attempts: &[(&str, &[&str])] = &[
        ("gnome-terminal", &["--", "pwsh", "-NoExit", "-File"]),
        ("konsole", &["-e", "pwsh", "-NoExit", "-File"]),
        ("xfce4-terminal", &["--", "pwsh", "-NoExit", "-File"]),
        ("mate-terminal", &["--", "pwsh", "-NoExit", "-File"]),
        ("xterm", &["-e", "pwsh", "-NoExit", "-File"]),
    ];

    for (term, prefix_args) in attempts {
        let mut args: Vec<&str> = prefix_args.to_vec();
        args.push(path_str);
        if let Ok(child) = Command::new(term).args(&args).spawn() {
            return Ok(child);
        }
    }

    Err("No supported terminal emulator found. Install gnome-terminal, konsole, or xterm.".into())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn graph_and_exchange_interleave_import_then_connect() {
        let cat = catalogue();
        let req = LaunchRequest {
            upn: Some("admin@contoso.onmicrosoft.com".into()),
            services: vec!["exchange".into(), "graph".into()],
            spo_url: None,
            pnp_url: None,
            graph_scopes: vec!["User.Read.All".into()],
        };
        let script = build_connection_script(&req, &cat);
        println!("{script}");

        // Exchange's own import+connect must both appear before Graph's,
        // proving imports are no longer front-loaded as a single block.
        let exo_import = script.find("Import-Module 'ExchangeOnlineManagement'").unwrap();
        let exo_connect = script.find("Connect-ExchangeOnline").unwrap();
        let graph_import = script.find("Import-Module 'Microsoft.Graph.Authentication'").unwrap();
        let graph_connect = script.find("Connect-MgGraph").unwrap();

        assert!(exo_import < exo_connect);
        assert!(exo_connect < graph_import, "Graph's import should come after Exchange's own connect, not before it");
        assert!(graph_import < graph_connect);

        // Full meta-package must never be imported.
        assert!(!script.contains("Import-Module 'Microsoft.Graph'\n"));
    }
}
