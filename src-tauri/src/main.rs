// Desktop shell, modelled on HOP's Tauri layer: the web UI owns the document engine,
// and the shell provides native file dialogs and file I/O.
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::path::PathBuf;
use std::sync::atomic::{AtomicUsize, Ordering};
use tauri::Manager;
use tauri_plugin_dialog::DialogExt;

#[derive(serde::Serialize)]
struct OpenedFile {
    name: String,
    path: String,
    bytes: Vec<u8>,
}

fn file_name(p: &PathBuf) -> String {
    p.file_name().map(|s| s.to_string_lossy().into_owned()).unwrap_or_default()
}

#[tauri::command]
async fn open_document(app: tauri::AppHandle) -> Result<Option<OpenedFile>, String> {
    let picked = app
        .dialog()
        .file()
        .add_filter("Word 문서", &["docx"])
        .blocking_pick_file();
    let Some(fp) = picked else { return Ok(None) };
    let path = fp.into_path().map_err(|e| e.to_string())?;
    let bytes = std::fs::read(&path).map_err(|e| e.to_string())?;
    Ok(Some(OpenedFile { name: file_name(&path), path: path.to_string_lossy().into_owned(), bytes }))
}

#[tauri::command]
fn read_document(path: String) -> Result<OpenedFile, String> {
    let p = PathBuf::from(&path);
    let bytes = std::fs::read(&p).map_err(|e| e.to_string())?;
    Ok(OpenedFile { name: file_name(&p), path, bytes })
}

/// Write the document. With no path (or save-as), ask for one first. Returns the saved path.
#[tauri::command]
async fn save_document(app: tauri::AppHandle, path: Option<String>, suggested: String, bytes: Vec<u8>) -> Result<Option<String>, String> {
    let target = match path {
        Some(p) => PathBuf::from(p),
        None => {
            let picked = app
                .dialog()
                .file()
                .add_filter("Word 문서", &["docx"])
                .set_file_name(&suggested)
                .blocking_save_file();
            match picked {
                Some(fp) => fp.into_path().map_err(|e| e.to_string())?,
                None => return Ok(None),
            }
        }
    };
    // Write to a sibling temp file first so a failed write never truncates the original.
    let tmp = target.with_extension("docx.tmp");
    std::fs::write(&tmp, &bytes).map_err(|e| e.to_string())?;
    std::fs::rename(&tmp, &target).map_err(|e| e.to_string())?;
    Ok(Some(target.to_string_lossy().into_owned()))
}

/// Files passed on the command line (e.g. "Open With" on Windows/Linux). Only the first window opens them.
#[tauri::command]
fn launch_paths(window: tauri::WebviewWindow) -> Vec<String> {
    if window.label() != "main" {
        return Vec::new();
    }
    std::env::args().skip(1).filter(|a| a.to_lowercase().ends_with(".docx")).collect()
}

// ---- windows (View ▸ Switch Windows / New Window)

static NEXT_WINDOW: AtomicUsize = AtomicUsize::new(1);

#[derive(serde::Serialize)]
struct WindowInfo {
    label: String,
    title: String,
    focused: bool,
}

/// Every open document window with its title (the title carries the file name and the ● dirty mark).
#[tauri::command]
fn list_windows(app: tauri::AppHandle) -> Vec<WindowInfo> {
    let mut out: Vec<WindowInfo> = app
        .webview_windows()
        .into_iter()
        .map(|(label, w)| WindowInfo { title: w.title().unwrap_or_default(), focused: w.is_focused().unwrap_or(false), label })
        .collect();
    out.sort_by(|a, b| a.label.cmp(&b.label));
    out
}

#[tauri::command]
fn focus_window(app: tauri::AppHandle, label: String) -> Result<(), String> {
    let w = app.get_webview_window(&label).ok_or("창을 찾을 수 없습니다")?;
    let _ = w.unminimize();
    w.show().map_err(|e| e.to_string())?;
    w.set_focus().map_err(|e| e.to_string())
}

/// Open a new editor window, optionally with a document. Async: building windows from a
/// synchronous command deadlocks on Windows.
#[tauri::command]
async fn new_window(app: tauri::AppHandle, path: Option<String>) -> Result<String, String> {
    let label = format!("doc-{}", NEXT_WINDOW.fetch_add(1, Ordering::SeqCst));
    let open = serde_json::to_string(&path).map_err(|e| e.to_string())?;
    tauri::WebviewWindowBuilder::new(&app, &label, tauri::WebviewUrl::App("index.html".into()))
        .title("Oword")
        .inner_size(1300.0, 900.0)
        .min_inner_size(800.0, 600.0)
        .disable_drag_drop_handler()
        .initialization_script(&format!("window.__OWORD_OPEN__ = {open};"))
        .build()
        .map_err(|e| e.to_string())?;
    Ok(label)
}

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![open_document, read_document, save_document, launch_paths, list_windows, focus_window, new_window])
        .run(tauri::generate_context!())
        .expect("error while running Oword");
}
