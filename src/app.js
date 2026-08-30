/**
 * Connect-365 frontend — communicates with the Rust backend via Tauri invoke.
 * No bundler; plain ES2020 module-style (no import/export — Tauri loads this as a classic script).
 */

// Lazy wrapper — window.__TAURI__ is injected by Tauri after page load
const invoke = (cmd, args) => window.__TAURI__.core.invoke(cmd, args);

// ── State ─────────────────────────────────────────────────────────────────────

const state = {
  modules: [],            // ModuleCatalogEntry[]
  selectedServices: new Set(),
  scopeProfiles: [],       // ScopeProfile[]
  activeProfileIds: new Set(),
  selectedScopes: new Set(),
  settings: null,          // Settings
};

// ── Boot ──────────────────────────────────────────────────────────────────────

document.addEventListener("DOMContentLoaded", async () => {
  initTheme();
  initTabs();

  // ── Static button wiring (no inline onclick — blocked by CSP) ──────────────
  document.getElementById("connect-form")
    ?.addEventListener("submit", e => { e.preventDefault(); connectServices(); });

  document.getElementById("settings-form")
    ?.addEventListener("submit", e => { e.preventDefault(); saveSettings(); });

  document.getElementById("btn-refresh-prereqs")
    ?.addEventListener("click", () => checkPrerequisites());

  document.getElementById("scope-search")
    ?.addEventListener("input", e => filterScopes(e.target.value));

  document.getElementById("btn-terminate-all")
    ?.addEventListener("click", () => terminateAllSessions());

  document.getElementById("btn-exit")
    ?.addEventListener("click", e => { e.preventDefault(); exitApp(); });

  // Event delegation for prereq table — handles Install/Update and PSGallery buttons
  document.getElementById("prereq-body")
    ?.addEventListener("click", e => {
      const btn = e.target.closest("button[data-action]");
      if (!btn) return;
      if (btn.dataset.action === "install")   installModule(btn.dataset.moduleName);
      if (btn.dataset.action === "open-link") openLink(btn.dataset.url);
    });

  // Event delegation for sessions list — per-session terminate
  document.getElementById("sessions-list")
    ?.addEventListener("click", e => {
      const btn = e.target.closest("button[data-terminate]");
      if (btn) terminateSession(btn.dataset.terminate);
    });

  await Promise.all([
    loadModules(),
    loadScopeProfiles(),
    loadSettings(),
  ]);

  // Opt-in, non-blocking (§7) — never delayed startup, adopted on next launch
  if (state.settings?.auto_update_scopes) {
    invoke("check_scope_list_update").catch(() => {});
  }

  window.__TAURI__.app.getVersion()
    .then(v => { document.getElementById("about-version").textContent = v; })
    .catch(() => { document.getElementById("about-version").textContent = "2.0.0-Preview1"; });

  // Fit after initial data (service checkboxes + last-used selections are now rendered)
  requestAnimationFrame(() => requestAnimationFrame(fitWindowToContent));

  checkPrerequisites();
  pollSessions();
});

// ── Theme ─────────────────────────────────────────────────────────────────────

const THEME_KEY = "c365-theme";

function initTheme() {
  const html = document.documentElement;
  const saved = localStorage.getItem(THEME_KEY);
  if (saved === "light" || saved === "dark") {
    html.setAttribute("data-theme", saved);
  }

  const themeSwitch = document.getElementById("theme-switch");
  themeSwitch.checked = html.getAttribute("data-theme") !== "light";
  themeSwitch.addEventListener("change", () => {
    const theme = themeSwitch.checked ? "dark" : "light";
    html.setAttribute("data-theme", theme);
    localStorage.setItem(THEME_KEY, theme);
  });
}

// ── Window auto-sizing ────────────────────────────────────────────────────────

async function fitWindowToContent() {
  try {
    const { getCurrentWindow, LogicalSize } = window.__TAURI__.window;
    const win = getCurrentWindow();

    // Measure natural content height of the active tab panel and sidebar
    const activePanel = document.querySelector(".tab-panel:not([hidden])");
    const sidebar     = document.querySelector(".app-sidebar");
    if (!activePanel || !sidebar) return;

    const contentH = Math.max(activePanel.scrollHeight, sidebar.scrollHeight);

    // Add a small buffer for the native title bar (≈28 px macOS, ≈30 px Windows)
    const CHROME = 30;
    const targetH = contentH + CHROME;

    // Preserve the current inner width
    const inner = await win.innerSize();
    const dpr   = window.devicePixelRatio || 1;
    const currentW = Math.round(inner.width / dpr);

    const minH = 500;
    const maxH = Math.floor(screen.availHeight * 0.92);
    const clampedH = Math.max(minH, Math.min(targetH, maxH));

    await win.setSize(new LogicalSize(currentW, clampedH));
  } catch (_) {
    // Falls back to the fixed size in tauri.conf.json when window API is unavailable
  }
}

// ── Tab switching ─────────────────────────────────────────────────────────────

function initTabs() {
  document.querySelectorAll(".nav-link[data-tab]").forEach(link => {
    link.addEventListener("click", e => {
      e.preventDefault();
      const tabId = link.dataset.tab;
      document.querySelectorAll(".nav-link").forEach(l => {
        l.classList.remove("active");
        l.removeAttribute("aria-current");
      });
      document.querySelectorAll(".tab-panel").forEach(p => { p.classList.remove("active"); p.hidden = true; });
      link.classList.add("active");
      link.setAttribute("aria-current", "page");
      const panel = document.getElementById(`tab-${tabId}`);
      panel.classList.add("active");
      panel.hidden = false;
      // Resize after the new tab's content is laid out
      requestAnimationFrame(() => requestAnimationFrame(fitWindowToContent));
    });
  });

  // Lazy-load advanced scopes when the details element is opened
  document.querySelectorAll("details").forEach(d => {
    d.addEventListener("toggle", () => {
      if (d.open && d.querySelector("#scope-advanced-list")) loadAllScopes();
    });
  });
}

// ── Module catalogue ──────────────────────────────────────────────────────────

async function loadModules() {
  try {
    state.modules = await invoke("get_modules");
    renderServiceCheckboxes();
  } catch (err) {
    console.error("get_modules failed", err);
  }
}

function renderServiceCheckboxes() {
  const container = document.getElementById("service-checkboxes");
  container.innerHTML = "";

  state.modules.forEach(mod => {
    const label = document.createElement("label");
    const input = document.createElement("input");
    input.type = "checkbox";
    input.name = "service";
    input.value = mod.id;
    label.appendChild(input);
    label.append(` ${mod.display_name}`);
    input.addEventListener("change", () => onServiceToggle());
    container.appendChild(label);
  });

  // Restore last-used selections from settings
  if (state.settings?.last_selected_services?.length) {
    state.settings.last_selected_services.forEach(id => {
      const cb = container.querySelector(`input[value="${CSS.escape(id)}"]`);
      if (cb) cb.checked = true;
    });
    onServiceToggle();
  }
}

function onServiceToggle() {
  state.selectedServices.clear();
  document.querySelectorAll('input[name="service"]:checked').forEach(cb => {
    state.selectedServices.add(cb.value);
  });

  // Conditional fields
  document.getElementById("spo-url-row").hidden = !state.selectedServices.has("sharepoint");
  document.getElementById("pnp-url-row").hidden  = !state.selectedServices.has("pnp");
  document.getElementById("graph-scope-picker").hidden = !state.selectedServices.has("graph");
}

// ── Graph scope picker ────────────────────────────────────────────────────────

async function loadScopeProfiles() {
  try {
    state.scopeProfiles = await invoke("get_scope_profiles");
    renderScopeProfiles();
  } catch (err) {
    console.error("get_scope_profiles failed", err);
  }
}

function renderScopeProfiles() {
  const container = document.getElementById("scope-profiles");
  container.innerHTML = "";

  state.scopeProfiles.forEach(profile => {
    const chip = document.createElement("span");
    chip.className = "scope-chip";
    chip.setAttribute("role", "button");
    chip.setAttribute("tabindex", "0");
    chip.setAttribute("aria-pressed", "false");
    chip.dataset.profileId = profile.id;
    chip.title = profile.scopes.join(", ");
    chip.textContent = profile.label;
    container.appendChild(chip);
  });

  container.addEventListener("click", e => {
    const chip = e.target.closest(".scope-chip");
    if (chip) toggleProfileChip(chip);
  });
  container.addEventListener("keydown", e => {
    const chip = e.target.closest(".scope-chip");
    if (chip && (e.key === "Enter" || e.key === " ")) {
      e.preventDefault();
      toggleProfileChip(chip);
    }
  });
}

function toggleProfileChip(chip) {
  const pressed = chip.getAttribute("aria-pressed") === "true";
  const next = !pressed;
  chip.setAttribute("aria-pressed", String(next));
  if (next) state.activeProfileIds.add(chip.dataset.profileId);
  else state.activeProfileIds.delete(chip.dataset.profileId);
  resolveScopes();
}

async function loadAllScopes() {
  const container = document.getElementById("scope-advanced-list");
  if (container.dataset.loaded) return;
  try {
    const scopes = await invoke("get_all_scopes");
    renderAdvancedScopes(scopes);
    container.dataset.loaded = "1";
  } catch (err) {
    container.textContent = "Failed to load scope list.";
  }
}

function renderAdvancedScopes(scopes) {
  const container = document.getElementById("scope-advanced-list");
  container.innerHTML = "";
  scopes.forEach(scope => {
    const label = document.createElement("label");
    label.dataset.value = scope.value.toLowerCase();

    const input = document.createElement("input");
    input.type = "checkbox";
    input.className = "scope-cb";
    input.value = scope.value;

    const span = document.createElement("span");
    span.className = "scope-value";
    span.textContent = ` ${scope.value}`;

    const small = document.createElement("small");
    small.className = "scope-desc";
    small.textContent = scope.description;

    label.append(input, span, small);
    input.addEventListener("change", resolveScopes);
    container.appendChild(label);
  });

  // Reflect any scopes already selected via active profile chips
  resolveScopes();
}

function filterScopes(query) {
  const q = query.toLowerCase();
  document.querySelectorAll("#scope-advanced-list label").forEach(label => {
    label.hidden = q.length > 0 && !label.dataset.value.includes(q);
  });
}

function resolveScopes() {
  const set = new Set();
  state.activeProfileIds.forEach(id => {
    const profile = state.scopeProfiles.find(p => p.id === id);
    profile?.scopes.forEach(s => set.add(s));
  });
  document.querySelectorAll(".scope-cb:checked").forEach(cb => set.add(cb.value));
  state.selectedScopes = set;

  // Visually reflect profile-contributed scopes as checked
  document.querySelectorAll(".scope-cb").forEach(cb => {
    if (set.has(cb.value)) cb.checked = true;
  });

  const box = document.getElementById("resolved-scopes");
  const list = document.getElementById("resolved-scope-list");
  const sorted = [...set].sort();
  box.hidden = sorted.length === 0;
  list.textContent = sorted.join("  ");
}

// ── Connect ───────────────────────────────────────────────────────────────────

async function connectServices() {
  const upn      = document.getElementById("upn").value.trim() || null;
  const spo_url  = document.getElementById("spo-url")?.value.trim() || null;
  const pnp_url  = document.getElementById("pnp-url")?.value.trim() || null;
  const services = [...state.selectedServices];
  const graph_scopes = [...state.selectedScopes];

  if (services.length === 0) {
    showConnectStatus("Select at least one service.", "error");
    return;
  }
  if (state.selectedServices.has("sharepoint") && !spo_url) {
    showConnectStatus("SharePoint Online requires an admin URL.", "error");
    return;
  }

  const btn = document.getElementById("btn-connect");
  btn.disabled = true;
  btn.setAttribute("aria-busy", "true");
  showConnectStatus("Launching session…", "");

  try {
    await invoke("launch_session", { request: { upn, services, spo_url, pnp_url, graph_scopes } });
    hideConnectStatus();
    pollSessions();

    if (state.settings?.minimise_to_tray_on_connect) {
      try {
        await window.__TAURI__.window.getCurrentWindow().hide();
      } catch (_) { /* window API unavailable — ignore */ }
    }
  } catch (err) {
    showConnectStatus(`Error: ${err}`, "error");
  } finally {
    btn.disabled = false;
    btn.removeAttribute("aria-busy");
  }
}

function showConnectStatus(msg, type) {
  const el = document.getElementById("connect-status");
  el.textContent = msg;
  el.className = `connect-status ${type}`;
  el.hidden = false;
}

function hideConnectStatus() {
  const el = document.getElementById("connect-status");
  el.hidden = true;
  el.textContent = "";
}

// ── Prerequisites ─────────────────────────────────────────────────────────────

async function checkPrerequisites() {
  document.getElementById("pwsh-message").textContent = "Checking…";
  document.getElementById("pwsh-status").setAttribute("aria-busy", "true");
  const tbody = document.getElementById("prereq-body");
  tbody.innerHTML = "";

  try {
    const result = await invoke("check_prerequisites");
    renderPrereqResult(result);
    requestAnimationFrame(() => requestAnimationFrame(fitWindowToContent));
  } catch (err) {
    document.getElementById("pwsh-message").textContent = `Check failed: ${err}`;
  } finally {
    document.getElementById("pwsh-status").removeAttribute("aria-busy");
  }
}

function statusPill(installed) {
  const span = document.createElement("span");
  span.className = `status-pill ${installed ? "status-ok" : "status-missing"}`;
  span.textContent = installed ? "Installed" : "Missing";
  return span;
}

function renderPrereqResult(result) {
  const pwshEl = document.getElementById("pwsh-message");
  pwshEl.innerHTML = "";

  if (result.pwsh_installed) {
    pwshEl.appendChild(statusPill(true));
    pwshEl.append(` ${result.pwsh_version ?? "version unknown"}`);
  } else {
    pwshEl.appendChild(statusPill(false));
    pwshEl.append(" — ");
    const dl = document.createElement("a");
    dl.href = "#";
    dl.textContent = "Download PowerShell 7";
    dl.addEventListener("click", e => { e.preventDefault(); openLink("https://learn.microsoft.com/en-us/powershell/scripting/install/install-powershell?view=powershell-7.6"); });
    pwshEl.appendChild(dl);
  }

  const tbody = document.getElementById("prereq-body");
  tbody.innerHTML = "";

  result.modules.forEach(mod => {
    const cat         = state.modules.find(m => m.id === mod.id) ?? {};
    const displayName = cat.display_name ?? mod.module_name;
    const galleryUrl  = cat.gallery_url  ?? "";

    const row = document.createElement("tr");

    // Module name
    const tdName = document.createElement("td");
    tdName.textContent = displayName;

    // Status
    const tdStatus = document.createElement("td");
    tdStatus.appendChild(statusPill(mod.installed));

    // Version
    const tdVersion = document.createElement("td");
    if (mod.version) {
      const code = document.createElement("code");
      code.textContent = mod.version;
      tdVersion.appendChild(code);
    } else {
      tdVersion.textContent = "—";
    }

    // Action — Install or Update button
    const tdAction = document.createElement("td");
    const actionBtn = document.createElement("button");
    actionBtn.className = "outline secondary prereq-action-btn";
    actionBtn.textContent = mod.installed ? "Update" : "Install";
    actionBtn.title = mod.installed ? "Update to latest version" : "Install module";
    actionBtn.dataset.action = "install";
    actionBtn.dataset.moduleName = mod.module_name;
    tdAction.appendChild(actionBtn);

    // Link — PSGallery button
    const tdLink = document.createElement("td");
    if (galleryUrl) {
      const linkBtn = document.createElement("button");
      linkBtn.className = "outline secondary prereq-action-btn";
      linkBtn.textContent = "PSGallery";
      linkBtn.title = "View on PowerShell Gallery";
      linkBtn.dataset.action = "open-link";
      linkBtn.dataset.url = galleryUrl;
      tdLink.appendChild(linkBtn);
    }

    row.append(tdName, tdStatus, tdVersion, tdAction, tdLink);
    tbody.appendChild(row);
  });
}

async function installModule(moduleName) {
  try {
    // Tauri v2: multi-word Rust param `module_name` → camelCase `moduleName` in JS
    await invoke("install_module", { moduleName });
  } catch (err) {
    alert(`Failed to launch installer: ${err}`);
  }
}

async function openLink(url) {
  await invoke("open_install_link", { url });
}

// ── Settings ──────────────────────────────────────────────────────────────────

async function loadSettings() {
  try {
    state.settings = await invoke("get_settings");
    applySettingsToForm();
  } catch (err) {
    console.error("get_settings failed", err);
    state.settings = {};
  }
}

function applySettingsToForm() {
  if (!state.settings) return;
  const s = state.settings;
  document.getElementById("upn").value = s.last_upn ?? "";
  document.getElementById("spo-url").value = s.last_spo_url ?? "";
  document.getElementById("pnp-url").value = s.last_pnp_url ?? "";
  document.getElementById("setting-tray").checked = s.minimise_to_tray_on_connect ?? true;
  document.getElementById("setting-autoupdate").checked = s.auto_update_scopes ?? false;
}

async function saveSettings() {
  const s = {
    last_selected_services: [...state.selectedServices],
    last_upn:  document.getElementById("upn").value.trim() || null,
    last_spo_url: document.getElementById("spo-url").value.trim() || null,
    last_pnp_url: document.getElementById("pnp-url").value.trim() || null,
    minimise_to_tray_on_connect: document.getElementById("setting-tray").checked,
    auto_update_scopes: document.getElementById("setting-autoupdate").checked,
  };
  try {
    await invoke("set_settings", { settings: s });
    state.settings = s;
  } catch (err) {
    alert(`Failed to save settings: ${err}`);
  }
}

// ── Sessions ──────────────────────────────────────────────────────────────────

let lastSessionsSnapshot = "";

async function pollSessions() {
  try {
    const sessions = await invoke("list_sessions");
    renderSessions(sessions);
  } catch (_) {/* ignore during shutdown */}
  setTimeout(pollSessions, 5000);
}

function renderSessions(sessions) {
  const alive = sessions.filter(s => s.alive);

  // Skip re-render if nothing changed (avoids flicker on the 5s poll)
  const snapshot = JSON.stringify(alive.map(s => [s.session_id, s.label]));
  if (snapshot === lastSessionsSnapshot) return;
  lastSessionsSnapshot = snapshot;

  const badge = document.getElementById("sessions-count-badge");
  badge.hidden = alive.length === 0;
  badge.textContent = String(alive.length);

  document.getElementById("sessions-toolbar").hidden = alive.length === 0;

  const list = document.getElementById("sessions-list");
  list.innerHTML = "";

  if (alive.length === 0) {
    const empty = document.createElement("div");
    empty.className = "sessions-empty";
    empty.innerHTML = `<i class="fa-solid fa-terminal"></i>`;
    const p = document.createElement("p");
    p.textContent = "No active sessions.";
    const small = document.createElement("small");
    small.textContent = "Connect to a service to start a session.";
    p.appendChild(document.createElement("br"));
    p.appendChild(small);
    empty.appendChild(p);
    list.appendChild(empty);
    return;
  }

  alive.forEach(session => {
    const card = document.createElement("article");
    card.className = "session-card";

    const meta = document.createElement("div");
    meta.className = "session-meta";

    const label = document.createElement("span");
    label.className = "session-label";
    label.textContent = session.label || "Connect-365 session";

    const time = document.createElement("span");
    time.className = "session-time";
    const started = Number(session.started_at);
    time.textContent = Number.isFinite(started)
      ? `Started ${new Date(started * 1000).toLocaleString()}`
      : "";

    meta.append(label, time);

    const btn = document.createElement("button");
    btn.className = "outline";
    btn.textContent = "Terminate";
    btn.dataset.terminate = session.session_id;

    card.append(meta, btn);
    list.appendChild(card);
  });
}

async function terminateSession(sessionId) {
  try {
    await invoke("terminate_session", { sessionId });
    lastSessionsSnapshot = ""; // force re-render
    pollSessions();
  } catch (err) {
    alert(`Failed to terminate session: ${err}`);
  }
}

async function terminateAllSessions() {
  try {
    await invoke("terminate_all_sessions");
    lastSessionsSnapshot = "";
    pollSessions();
  } catch (err) {
    alert(`Failed to terminate sessions: ${err}`);
  }
}

// ── Exit ──────────────────────────────────────────────────────────────────────

async function exitApp() {
  try {
    const sessions = await invoke("list_sessions");
    const aliveCount = sessions.filter(s => s.alive).length;
    if (aliveCount > 0) {
      const label = aliveCount === 1 ? "session" : "sessions";
      const ok = confirm(`Terminating will close ${aliveCount} active ${label}. Exit Connect-365?`);
      if (!ok) return;
    }
    await invoke("quit_app");
  } catch (err) {
    console.error("Exit failed", err);
  }
}
