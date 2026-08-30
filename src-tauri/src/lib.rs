mod commands;
mod models;

use models::AppState;
use tauri::Manager;

pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .manage(AppState::default())
        .invoke_handler(tauri::generate_handler![
            commands::prerequisites::check_prerequisites,
            commands::prerequisites::install_module,
            commands::session::get_modules,
            commands::session::launch_session,
            commands::session::list_sessions,
            commands::session::terminate_session,
            commands::session::terminate_all_sessions,
            commands::session::quit_app,
            commands::scopes::get_scope_profiles,
            commands::scopes::get_all_scopes,
            commands::scopes::check_scope_list_update,
            commands::settings::get_settings,
            commands::settings::set_settings,
            commands::settings::open_install_link,
        ])
        .setup(|app| {
            build_tray(app)?;
            Ok(())
        })
        // Hide to tray on close rather than quitting (§5 "app remains in tray")
        .on_window_event(|window, event| {
            if let tauri::WindowEvent::CloseRequested { api, .. } = event {
                api.prevent_close();
                let _ = window.hide();
            }
        })
        .run(tauri::generate_context!())
        .expect("error running Connect-365");
}

fn build_tray(app: &tauri::App) -> tauri::Result<()> {
    use tauri::menu::{MenuBuilder, MenuItemBuilder};
    use tauri::tray::{MouseButton, MouseButtonState, TrayIconBuilder, TrayIconEvent};

    let show = MenuItemBuilder::with_id("show", "Show Window").build(app)?;
    let sep = tauri::menu::PredefinedMenuItem::separator(app)?;
    let terminate_all = MenuItemBuilder::with_id("terminate_all", "Terminate All Sessions").build(app)?;
    let sep2 = tauri::menu::PredefinedMenuItem::separator(app)?;
    let quit = MenuItemBuilder::with_id("quit", "Quit Connect-365").build(app)?;

    let menu = MenuBuilder::new(app)
        .item(&show)
        .item(&sep)
        .item(&terminate_all)
        .item(&sep2)
        .item(&quit)
        .build()?;

    TrayIconBuilder::new()
        .icon(app.default_window_icon().unwrap().clone())
        .menu(&menu)
        .show_menu_on_left_click(false)
        .tooltip("Connect-365")
        .on_menu_event(|app, event| match event.id.as_ref() {
            "show" => {
                if let Some(win) = app.get_webview_window("main") {
                    let _ = win.show();
                    let _ = win.set_focus();
                }
            }
            "terminate_all" => {
                let state = app.state::<AppState>();
                commands::session::terminate_all_internal(&state);
            }
            "quit" => {
                // Tray Quit has no UI to confirm through; still clean up tracked
                // sessions/temp scripts before exiting (§5).
                let state = app.state::<AppState>();
                commands::session::terminate_all_internal(&state);
                app.exit(0);
            }
            _ => {}
        })
        .on_tray_icon_event(|tray, event| {
            // Left-click toggles the window
            if let TrayIconEvent::Click {
                button: MouseButton::Left,
                button_state: MouseButtonState::Up,
                ..
            } = event
            {
                let app = tray.app_handle();
                if let Some(win) = app.get_webview_window("main") {
                    if win.is_visible().unwrap_or(false) {
                        let _ = win.hide();
                    } else {
                        let _ = win.show();
                        let _ = win.set_focus();
                    }
                }
            }
        })
        .build(app)?;

    Ok(())
}
