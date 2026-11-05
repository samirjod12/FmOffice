#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use futures_util::StreamExt;
use serde::Serialize;
use std::{fs, io::Cursor, path::PathBuf};
use tauri::{AppHandle, Emitter};

/// Set at build time by the GitHub workflow: FMOFFICE_REPO=owner/repo
const REPO: &str = match option_env!("FMOFFICE_REPO") {
    Some(r) => r,
    None => "OWNER/FmOffice",
};

struct AppInfo {
    id: &'static str,
    name: &'static str,
    asset: &'static str,
    exe: &'static str,
}
const APPS: &[AppInfo] = &[AppInfo {
    id: "sheets",
    name: "FmOffice Sheets",
    asset: "FmOffice-Sheets-windows-x64.zip",
    exe: "fmoffice-sheets.exe",
}];

#[derive(Serialize, Clone)]
struct Progress {
    app: String,
    stage: String,
    percent: f64,
}

fn install_root() -> PathBuf {
    let base = std::env::var("LOCALAPPDATA")
        .or_else(|_| std::env::var("HOME").map(|h| format!("{h}/.local/share")))
        .unwrap_or_else(|_| ".".into());
    PathBuf::from(base).join("FmOffice")
}

fn emit(app: &AppHandle, id: &str, stage: &str, percent: f64) {
    let _ = app.emit("progress", Progress { app: id.into(), stage: stage.into(), percent });
}

#[tauri::command]
fn get_info() -> serde_json::Value {
    let root = install_root();
    let installed: Vec<&str> = APPS.iter().filter(|a| root.join(a.id).join(a.exe).exists()).map(|a| a.id).collect();
    serde_json::json!({
        "repo": REPO,
        "installDir": root.to_string_lossy(),
        "installed": installed,
        "uninstall": std::env::args().any(|a| a == "--uninstall"),
    })
}

#[tauri::command]
async fn install(app: AppHandle, ids: Vec<String>) -> Result<(), String> {
    let root = install_root();
    fs::create_dir_all(&root).map_err(|e| e.to_string())?;
    let client = reqwest::Client::builder().user_agent("FmOffice-Setup").build().map_err(|e| e.to_string())?;
    for id in ids {
        let info = APPS.iter().find(|a| a.id == id).ok_or("Unknown app")?;
        let url = format!("https://github.com/{REPO}/releases/latest/download/{}", info.asset);
        emit(&app, &id, "Downloading", 0.0);
        let resp = client.get(&url).send().await.map_err(|e| format!("Download failed: {e}"))?;
        if !resp.status().is_success() {
            return Err(format!("Download failed ({}). Has a release been published at github.com/{REPO}/releases ?", resp.status()));
        }
        let total = resp.content_length().unwrap_or(0) as f64;
        let mut got = 0f64;
        let mut buf: Vec<u8> = Vec::new();
        let mut stream = resp.bytes_stream();
        while let Some(chunk) = stream.next().await {
            let chunk = chunk.map_err(|e| e.to_string())?;
            got += chunk.len() as f64;
            buf.extend_from_slice(&chunk);
            if total > 0.0 { emit(&app, &id, "Downloading", got / total * 90.0); }
        }
        emit(&app, &id, "Installing", 92.0);
        let dest = root.join(info.id);
        if dest.exists() { let _ = fs::remove_dir_all(&dest); }
        fs::create_dir_all(&dest).map_err(|e| e.to_string())?;
        let mut zip = zip::ZipArchive::new(Cursor::new(buf)).map_err(|e| e.to_string())?;
        zip.extract(&dest).map_err(|e| e.to_string())?;
        // keep a copy of this setup so the user can uninstall later
        if let Ok(me) = std::env::current_exe() { let _ = fs::copy(me, root.join("FmOffice-Setup.exe")); }
        emit(&app, &id, "Creating shortcuts", 97.0);
        make_shortcuts(info, &dest);
        register_uninstall(&root);
        emit(&app, &id, "Done", 100.0);
    }
    Ok(())
}

#[tauri::command]
fn launch(id: String) -> Result<(), String> {
    let info = APPS.iter().find(|a| a.id == id).ok_or("Unknown app")?;
    std::process::Command::new(install_root().join(info.id).join(info.exe)).spawn().map_err(|e| e.to_string())?;
    Ok(())
}

#[tauri::command]
fn uninstall() -> Result<(), String> {
    #[cfg(windows)]
    {
        let ps = r#"
$start = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'
Remove-Item -Force -ErrorAction SilentlyContinue (Join-Path $start 'FmOffice Sheets.lnk')
Remove-Item -Force -ErrorAction SilentlyContinue (Join-Path ([Environment]::GetFolderPath('Desktop')) 'FmOffice Sheets.lnk')
Remove-Item -Force -ErrorAction SilentlyContinue 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\FmOffice' -Recurse
"#;
        run_ps(ps);
    }
    let root = install_root();
    for a in APPS { let _ = fs::remove_dir_all(root.join(a.id)); }
    // the running setup copy can't delete itself on Windows; remove the rest, leave the exe
    Ok(())
}

#[cfg(windows)]
fn run_ps(script: &str) {
    use std::os::windows::process::CommandExt;
    let _ = std::process::Command::new("powershell")
        .args(["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script])
        .creation_flags(0x08000000) // CREATE_NO_WINDOW
        .status();
}

fn make_shortcuts(info: &AppInfo, dir: &PathBuf) {
    #[cfg(windows)]
    {
        let exe = dir.join(info.exe);
        let ps = format!(
            r#"$w = New-Object -ComObject WScript.Shell
foreach ($p in @((Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs'), [Environment]::GetFolderPath('Desktop'))) {{
  $s = $w.CreateShortcut((Join-Path $p '{name}.lnk')); $s.TargetPath = '{exe}'; $s.WorkingDirectory = '{dir}'; $s.Save()
}}"#,
            name = info.name, exe = exe.display(), dir = dir.display()
        );
        run_ps(&ps);
    }
    #[cfg(not(windows))]
    { let _ = (info, dir); }
}

fn register_uninstall(root: &PathBuf) {
    #[cfg(windows)]
    {
        let setup = root.join("FmOffice-Setup.exe");
        let ps = format!(
            r#"$k = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\FmOffice'
New-Item -Path $k -Force | Out-Null
Set-ItemProperty $k DisplayName 'FmOffice'
Set-ItemProperty $k Publisher 'FmOffice (fan-made)'
Set-ItemProperty $k UninstallString '"{s}" --uninstall'
Set-ItemProperty $k InstallLocation '{r}'"#,
            s = setup.display(), r = root.display()
        );
        run_ps(&ps);
    }
    #[cfg(not(windows))]
    { let _ = root; }
}

fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![get_info, install, launch, uninstall])
        .run(tauri::generate_context!())
        .expect("error while running FmOffice Setup");
}
