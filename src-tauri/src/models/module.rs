use serde::{Deserialize, Serialize};

/// One entry in the static module catalogue (§4).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ModuleCatalogEntry {
    /// Stable identifier used in LaunchRequest.services[] and stored prefs.
    pub id: String,
    pub display_name: String,
    /// PowerShell module name passed to Get-Module / Import-Module.
    pub module_name: String,
    /// Module actually Import-Module'd in the launch script, if different
    /// from `module_name`. Used for umbrella meta-packages (Microsoft.Graph,
    /// Microsoft.Entra) where importing the full package eagerly loads
    /// dozens of submodules and can take minutes; importing just the
    /// `.Authentication` submodule is fast and sufficient for Connect-*
    /// (other submodules autoload lazily as their cmdlets are used).
    /// `check_prerequisites` still checks `module_name` (the umbrella
    /// package), since that's what users install from PSGallery.
    #[serde(default)]
    pub import_module_name: Option<String>,
    /// Connect cmdlet template; placeholders: {upn}, {spo_url}, {pnp_url}, {scopes}.
    pub connect_cmdlet: String,
    pub gallery_url: String,
    /// Whether the user must supply an SPO admin URL when this service is selected.
    pub requires_spo_url: bool,
    /// Whether the user may supply a PnP site URL (optional, not required).
    pub requires_pnp_url: bool,
}

/// Prereq-check result for one module.
#[derive(Debug, Serialize, Deserialize)]
pub struct ModuleStatus {
    pub id: String,
    pub module_name: String,
    pub installed: bool,
    pub version: Option<String>,
}

/// Full result returned by check_prerequisites.
#[derive(Debug, Serialize, Deserialize)]
pub struct PrereqResult {
    pub pwsh_installed: bool,
    pub pwsh_version: Option<String>,
    pub modules: Vec<ModuleStatus>,
}
