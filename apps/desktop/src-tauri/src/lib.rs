use std::io::{BufRead, BufReader};
use std::os::windows::process::CommandExt;
use std::process::{Command, Stdio};
use tauri::{
    menu::{Menu, MenuItem},
    tray::{MouseButton, MouseButtonState, TrayIconBuilder, TrayIconEvent},
    Emitter, Manager,
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

// Resuelve la ruta absoluta de un script .ps1, buscando primero en los recursos
// empaquetados de Tauri y, como fallback de desarrollo, en la carpeta de origen.
fn resolve_script_path(app_handle: &tauri::AppHandle, script_name: &str) -> Result<String, String> {
    let mut resource_path = app_handle
        .path()
        .resource_dir()
        .ok()
        .map(|p| p.join("scripts-ps1").join(script_name));

    if resource_path.is_none() || !resource_path.as_ref().unwrap().exists() {
        let dev_path = std::env::current_dir().map(|p| p.join("scripts-ps1").join(script_name));
        if let Ok(p) = dev_path {
            if p.exists() {
                resource_path = Some(p);
            }
        }
    }

    let final_path = resource_path.ok_or_else(|| {
        format!(
            "No se pudo encontrar el script '{}' en recursos ni en origen.",
            script_name
        )
    })?;

    // Limpiar el prefijo UNC \\?\ que Rust añade automáticamente en Windows para rutas largas,
    // ya que cmd/PowerShell no lo reconocen correctamente como argumento -File.
    let mut script_str = final_path.to_string_lossy().into_owned();
    if script_str.starts_with("\\\\?\\") {
        script_str = script_str.replacen("\\\\?\\", "", 1);
    }
    Ok(script_str)
}

// Comando para buscar buzones de tipo OST ejecutando el script PowerShell
#[tauri::command]
fn find_outlook_stores(app_handle: tauri::AppHandle) -> Result<Vec<OutlookStore>, String> {
    let script_str = resolve_script_path(&app_handle, "outlook-list-stores.ps1")?;

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
            let msg = parsed
                .get("message")
                .and_then(|m| m.as_str())
                .unwrap_or("Error desconocido en script");
            return Err(msg.to_string());
        }
    }

    let data: PSOutput = serde_json::from_value(parsed)
        .map_err(|e| format!("Error de correspondencia del JSON: {}", e))?;

    // Filtrar aquellas que tengan filePath válido (no vacío ni nulo)
    let filtered_stores = data
        .stores
        .into_iter()
        .filter(|s| s.file_path.is_some() && !s.file_path.as_ref().unwrap().trim().is_empty())
        .collect();

    Ok(filtered_stores)
}

// Comando para escanear un buzón (StoreId) ejecutando outlook-scan-pst.ps1 con streaming.
// Cada línea JSON emitida por el script se reenvía a la UI mediante eventos Tauri:
//   - "scan://event"    -> payload crudo de cada evento (scanMeta, scanProgress, log, etc.)
//   - "scan://complete" -> payload final (summary) cuando el escaneo termina correctamente
//   - "scan://error"    -> mensaje de error
#[tauri::command]
fn scan_mailbox(
    app_handle: tauri::AppHandle,
    window: tauri::Window,
    store_id: String,
) -> Result<(), String> {
    if store_id.trim().is_empty() {
        return Err("StoreId vacío. Seleccione un buzón válido.".to_string());
    }

    let script_str = resolve_script_path(&app_handle, "outlook-scan-pst.ps1")?;

    let mut child = Command::new("powershell")
        .creation_flags(CREATE_NO_WINDOW)
        .arg("-NoProfile")
        .arg("-NonInteractive")
        .arg("-ExecutionPolicy")
        .arg("Bypass")
        .arg("-File")
        .arg(&script_str)
        .arg("-StoreId")
        .arg(&store_id)
        .arg("-IncludeSize")
        .arg("-Summary")
        .arg("-Json")
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .map_err(|e| format!("No se pudo ejecutar PowerShell: {}", e))?;

    let stdout = child
        .stdout
        .take()
        .ok_or_else(|| "No se pudo capturar la salida del script.".to_string())?;

    let reader = BufReader::new(stdout);

    for line in reader.lines() {
        let line = match line {
            Ok(l) => l,
            Err(_) => continue,
        };
        let trimmed = line.trim();
        if trimmed.is_empty() {
            continue;
        }

        // Cada línea debería ser un objeto JSON. Si no parsea, lo ignoramos.
        let parsed: serde_json::Value = match serde_json::from_str(trimmed) {
            Ok(v) => v,
            Err(_) => continue,
        };

        let event_type = parsed
            .get("type")
            .and_then(|t| t.as_str())
            .unwrap_or("")
            .to_string();

        match event_type.as_str() {
            "error" => {
                let msg = parsed
                    .get("message")
                    .and_then(|m| m.as_str())
                    .unwrap_or("Error desconocido en el escaneo")
                    .to_string();
                let _ = window.emit("scan://error", &msg);
            }
            "summary" => {
                let _ = window.emit("scan://complete", &parsed);
            }
            _ => {
                let _ = window.emit("scan://event", &parsed);
            }
        }
    }

    let status = child
        .wait()
        .map_err(|e| format!("Error esperando al proceso: {}", e))?;

    if !status.success() {
        let mut err_msg = String::new();
        if let Some(mut stderr) = child.stderr.take() {
            use std::io::Read;
            let _ = stderr.read_to_string(&mut err_msg);
        }
        if err_msg.trim().is_empty() {
            err_msg = format!("El script finalizó con código de error {:?}.", status.code());
        }
        let _ = window.emit("scan://error", &err_msg);
        return Err(err_msg);
    }

    Ok(())
}

// Learn more about Tauri commands at https://tauri.app/develop/calling-rust/
#[tauri::command]
fn greet(name: &str) -> String {
    format!("Hello, {}! You've been greeted from Rust!", name)
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_store::Builder::new().build())
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_store::Builder::default().build())
        .invoke_handler(tauri::generate_handler![
            greet,
            find_outlook_stores,
            scan_mailbox
        ])
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
            let settings_item =
                MenuItem::with_id(app, "settings", "Configuración", true, None::<&str>)?;
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
