use crate::commands::session::catalogue;
use crate::models::{ModuleStatus, PrereqResult};
use std::process::Command;
use std::sync::OnceLock;

/// Resolves the pwsh executable to invoke.
///
/// GUI apps launched via Finder/Dock (macOS) or a `.desktop` launcher (Linux)
/// get a minimal PATH from the OS session, not the user's shell-configured
/// PATH — so Homebrew's `/opt/homebrew/bin`, the official PowerShell
/// installer's `/usr/local/bin` symlink, etc. are invisible to `pwsh` even
/// though the exact same binary resolves fine when the app is run from a
/// terminal (`cargo tauri dev`). Fall back to resolving through the user's
/// real login shell, then to common known install locations, before giving
/// up and reporting pwsh as missing.
fn resolve_pwsh() -> &'static str {
    static RESOLVED: OnceLock<String> = OnceLock::new();
    RESOLVED.get_or_init(|| {
        // Fast path: already resolvable as-is (dev builds, Windows, most Linux setups)
        if Command::new("pwsh").arg("--version").output().is_ok() {
            return "pwsh".to_string();
        }

        // Fall back to the user's login shell PATH (macOS/Linux GUI launch)
        #[cfg(not(target_os = "windows"))]
        {
            let shell = std::env::var("SHELL").unwrap_or_else(|_| "/bin/zsh".into());
            if let Ok(out) = Command::new(&shell).args(["-lc", "command -v pwsh"]).output() {
                if out.status.success() {
                    let path = String::from_utf8_lossy(&out.stdout).trim().to_string();
                    if !path.is_empty() {
                        return path;
                    }
                }
            }
        }

        // Fall back to common known install locations
        #[cfg(not(target_os = "windows"))]
        let candidates: &[&str] = &[
            "/opt/homebrew/bin/pwsh",              // Homebrew, Apple Silicon
            "/usr/local/bin/pwsh",                 // Homebrew (Intel) / official installer symlink
            "/usr/local/microsoft/powershell/7/pwsh", // official macOS .pkg installer
            "/usr/bin/pwsh",                       // Linux package managers
            "/snap/bin/pwsh",                      // Linux snap
        ];
        #[cfg(target_os = "windows")]
        let candidates: &[&str] = &[];

        for candidate in candidates {
            if std::path::Path::new(candidate).exists() {
                return candidate.to_string();
            }
        }

        "pwsh".to_string()
    })
}

/// Checks pwsh availability and runs Get-Module -ListAvailable for each catalogue module
/// in a single pwsh invocation, returning installed status and version.
#[tauri::command]
pub async fn check_prerequisites() -> Result<PrereqResult, String> {
    let pwsh_bin = resolve_pwsh();

    // ── 1. Detect pwsh ────────────────────────────────────────────────────────
    let (pwsh_installed, pwsh_version) = match Command::new(pwsh_bin).arg("--version").output() {
        Ok(out) if out.status.success() => {
            let v = String::from_utf8_lossy(&out.stdout).trim().to_string();
            (true, Some(v))
        }
        _ => (false, None),
    };

    if !pwsh_installed {
        let modules = catalogue()
            .into_iter()
            .map(|m| ModuleStatus {
                id: m.id,
                module_name: m.module_name,
                installed: false,
                version: None,
            })
            .collect();
        return Ok(PrereqResult {
            pwsh_installed,
            pwsh_version,
            modules,
        });
    }

    // ── 2. Check modules — single pwsh invocation, deduped by module name ────
    let cat = catalogue();
    let unique_names: Vec<String> = {
        let mut seen = std::collections::HashSet::new();
        cat.iter()
            .filter(|e| seen.insert(e.module_name.clone()))
            .map(|e| e.module_name.clone())
            .collect()
    };

    // Build a comma-separated PS array of quoted module names
    let names_ps = unique_names
        .iter()
        .map(|n| format!("'{}'", n.replace('\'', "''")))
        .collect::<Vec<_>>()
        .join(",");

    // ForEach-Object returns {Name, Installed, Version} for each module.
    // Sort descending so the highest installed version is selected.
    let ps_expr = format!(
        "@({names}) | ForEach-Object {{ \
            $mod = Get-Module -ListAvailable -Name $_ | Sort-Object Version -Descending | Select-Object -First 1; \
            [PSCustomObject]@{{Name=$_; Installed=($null -ne $mod); Version=if ($null -ne $mod) {{$mod.Version.ToString()}} else {{$null}}}} \
        }} | ConvertTo-Json -Compress",
        names = names_ps
    );

    let output = Command::new(pwsh_bin)
        .args(["-NoProfile", "-NonInteractive", "-Command", &ps_expr])
        .output()
        .map_err(|e| format!("pwsh check failed: {e}"))?;

    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        return Err(format!("Prereq check script failed: {stderr}"));
    }

    let stdout = String::from_utf8_lossy(&output.stdout);
    let raw = stdout.trim();

    let json: serde_json::Value = serde_json::from_str(raw)
        .map_err(|e| format!("Could not parse prereq output: {e}\nRaw: {raw}"))?;

    // PowerShell may return a single object (not array) when only one module is checked
    let results: Vec<&serde_json::Value> = match &json {
        serde_json::Value::Array(arr) => arr.iter().collect(),
        obj => vec![obj],
    };

    let installed_map: std::collections::HashMap<&str, (bool, Option<String>)> = results
        .iter()
        .filter_map(|obj| {
            let name = obj["Name"].as_str()?;
            let installed = obj["Installed"].as_bool().unwrap_or(false);
            let version = obj["Version"].as_str().map(|s| s.to_string());
            Some((name, (installed, version)))
        })
        .collect();

    // Expand back to full catalogue (one row per entry, including shared modules like
    // ExchangeOnlineManagement which backs both Exchange Online and Compliance)
    let modules: Vec<ModuleStatus> = cat
        .into_iter()
        .map(|m| {
            let (installed, version) = installed_map
                .get(m.module_name.as_str())
                .cloned()
                .unwrap_or((false, None));
            ModuleStatus {
                id: m.id,
                module_name: m.module_name,
                installed,
                version,
            }
        })
        .collect();

    Ok(PrereqResult {
        pwsh_installed,
        pwsh_version,
        modules,
    })
}

/// Opens a terminal running Install-Module for the given module, scoped to CurrentUser.
/// Ensures the NuGet provider is present first to avoid interactive prompts.
#[tauri::command]
pub async fn install_module(module_name: String) -> Result<(), String> {
    let safe_name = module_name.replace('\'', "''");

    let script = format!(
        "# Connect-365 — Install / Update module\n\
         $ErrorActionPreference = 'Stop'\n\
         Write-Host \"`nConnect-365: Installing / updating '{name}'...`n\" -ForegroundColor Cyan\n\
         \n\
         # Ensure NuGet provider is present so Install-Module doesn't prompt interactively\n\
         if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue | Where-Object {{$_.Version -ge '2.8.5.201'}})) {{\n\
             Write-Host 'Installing NuGet provider...' -ForegroundColor Yellow\n\
             Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force | Out-Null\n\
         }}\n\
         \n\
         Install-Module '{name}' -Scope CurrentUser -Force -AllowClobber\n\
         \n\
         $installed = Get-Module -ListAvailable -Name '{name}' | Sort-Object Version -Descending | Select-Object -First 1\n\
         if ($installed) {{\n\
             Write-Host \"`nDone — $($installed.Name) v$($installed.Version) is installed.\" -ForegroundColor Green\n\
         }} else {{\n\
             Write-Host \"`nInstall completed.\" -ForegroundColor Green\n\
         }}\n\
         \n\
         Write-Host 'Press Enter to close this window.'\n\
         Read-Host\n",
        name = safe_name
    );

    let script_path = std::env::temp_dir()
        .join(format!("connect365-install-{}.ps1", uuid::Uuid::new_v4()));

    std::fs::write(&script_path, &script)
        .map_err(|e| format!("Failed to write install script: {e}"))?;

    // Reuse the same terminal launcher as session launches
    let child = super::session::spawn_terminal(&script_path)?;

    // Detach — we don't track install processes; the terminal manages its own lifetime
    drop(child);

    Ok(())
}
