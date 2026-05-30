use std::process::Command;
use std::os::windows::process::CommandExt;
use tauri::{
    menu::{Menu, MenuItem},
    tray::{MouseButton, MouseButtonState, TrayIconBuilder, TrayIconEvent},
    Manager,
};

#[derive(serde::Serialize, serde::Deserialize, Debug)]
struct OutlookStore {
    #[serde(rename = "displayName")]
    display_name: String,
    #[serde(rename = "storeId")]
    store_id: Option<String>,
    #[serde(rename = "filePath")]
    file_path: Option<String>,
    #[serde(rename = "fileSize")]
    file_size: Option<String>,
    #[serde(rename = "exchangeStoreType")]
    exchange_store_type: Option<serde_json::Value>,
    #[serde(rename = "storeType")]
    store_type: Option<String>,
}

#[derive(serde::Serialize, serde::Deserialize, Debug)]
struct PSOutput {
    #[serde(rename = "type")]
    output_type: String,
    stores: Vec<OutlookStore>,
}

// Prevents additional console window on Windows when executing child processes
const CREATE_NO_WINDOW: u32 = 0x08000000;

// Comando para buscar buzones de tipo OST ejecutando el script PowerShell
#[tauri::command]
fn find_outlook_stores(app_handle: tauri::AppHandle) -> Result<Vec<OutlookStore>, String> {
    // Intentar obtener la ruta del recurso empaquetado en Tauri
    let mut resource_path = app_handle
        .path()
        .resource_dir()
        .ok()
        .map(|p| p.join("scripts-ps1").join("outlook-list-stores.ps1"));

    // Fallback para desarrollo (debug): si el archivo no existe en el directorio de recursos empaquetados,
    // buscamos directamente en la carpeta del código fuente de src-tauri.
    if resource_path.is_none() || !resource_path.as_ref().unwrap().exists() {
        let dev_path = std::env::current_dir()
            .map(|p| p.join("scripts-ps1").join("outlook-list-stores.ps1"));
        if let Ok(p) = dev_path {
            if p.exists() {
                resource_path = Some(p);
            }
        }
    }

    let final_path = resource_path
        .ok_or_else(|| "No se pudo encontrar el script 'outlook-list-stores.ps1' en recursos ni en origen.".to_string())?;

    // Limpiar el prefijo UNC \\?\ que Rust añade automáticamente en Windows para rutas largas, 
    // ya que cmd/PowerShell no lo reconocen correctamente en la mayoría de sus argumentos como -File.
    let mut script_str = final_path.to_string_lossy().into_owned();
    if script_str.starts_with("\\\\?\\") {
        script_str = script_str.replacen("\\\\?\\", "", 1);
    }

    // Comando de PowerShell: ejecuta de forma no interactiva, con bypass de ejecución, filtrando por OST
    let output = Command::new("powershell")
        .creation_flags(CREATE_NO_WINDOW)
        .arg("-NoProfile")
        .arg("-NonInteractive")
        .arg("-ExecutionPolicy")
        .arg("Bypass")
        .arg("-File")
        .arg(&script_str)
        .arg("-StoreType")
        .arg("OST")
        .arg("-Json")
        .output()
        .map_err(|e| format!("No se pudo ejecutar PowerShell: {}", e))?;

    if !output.status.success() {
        let err_msg = String::from_utf8_lossy(&output.stderr);
        return Err(format!("Error del script PowerShell: {}", err_msg));
    }

    let stdout_str = String::from_utf8_lossy(&output.stdout);
    
    // El script en caso de error puede retornar un JSON {"type": "error", "message": "..."}
    // o bien el listado de stores {"type": "stores", "stores": [...]}
    if stdout_str.trim().is_empty() {
        return Ok(Vec::new());
    }

    let parsed: serde_json::Value = serde_json::from_str(&stdout_str)
        .map_err(|e| format!("Error al parsear JSON: {}, raw: {}", e, stdout_str))?;

    if let Some(err_type) = parsed.get("type") {
        if err_type == "error" {
            let msg = parsed.get("message")
                .and_then(|m| m.as_str())
                .unwrap_or("Error desconocido en script");
            return Err(msg.to_string());
        }
    }

    let data: PSOutput = serde_json::from_value(parsed)
        .map_err(|e| format!("Error de correspondencia del JSON: {}", e))?;

    // Filtrar aquellas que tengan filePath válido (no vacío ni nulo)
    let filtered_stores = data.stores.into_iter()
        .filter(|s| s.file_path.is_some() && !s.file_path.as_ref().unwrap().trim().is_empty())
        .collect();

    Ok(filtered_stores)
}

// Learn more about Tauri commands at https://tauri.app/develop/calling-rust/
#[tauri::command]
fn greet(name: &str) -> String {
    format!("Hello, {}! You've been greeted from Rust!", name)
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![greet, find_outlook_stores])
        .setup(|app| {
            // Obtener la ventana principal
            let main_window = app.get_webview_window("main").unwrap();

            // Configurar el evento de cierre de la ventana principal para ocultarla en lugar de cerrarla
            let window_clone = main_window.clone();
            main_window.on_window_event(move |event| {
                if let tauri::WindowEvent::CloseRequested { api, .. } = event {
                    api.prevent_close();
                    let _ = window_clone.hide();
                }
            });

            // Crear los ítems del menú de la bandeja de sistema
            let show_item = MenuItem::with_id(app, "show", "Mostrar", true, None::<&str>)?;
            let settings_item = MenuItem::with_id(app, "settings", "Configuración", true, None::<&str>)?;
            let quit_item = MenuItem::with_id(app, "quit", "Salir", true, None::<&str>)?;

            // Crear el menú y añadir los ítems
            let tray_menu = Menu::with_items(app, &[&show_item, &settings_item, &quit_item])?;

            // Construir el TrayIcon de Tauri v2
            let _tray = TrayIconBuilder::new()
                .icon(app.default_window_icon().unwrap().clone())
                .menu(&tray_menu)
                .on_menu_event(|app, event| match event.id.as_ref() {
                    "show" | "settings" => {
                        if let Some(window) = app.get_webview_window("main") {
                            let _ = window.show();
                            let _ = window.set_focus();
                        }
                    }
                    "quit" => {
                        app.exit(0);
                    }
                    _ => {}
                })
                .on_tray_icon_event(|tray, event| {
                    if let TrayIconEvent::Click {
                        button: MouseButton::Left,
                        button_state: MouseButtonState::Up,
                        ..
                    } = event
                    {
                        let app = tray.app_handle();
                        if let Some(window) = app.get_webview_window("main") {
                            if window.is_visible().unwrap_or(false) {
                                let _ = window.hide();
                            } else {
                                let _ = window.show();
                                let _ = window.set_focus();
                            }
                        }
                    }
                })
                .build(app)?;

            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
