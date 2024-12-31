#[cfg(not(debug_assertions))]
#[macro_use]
extern crate windows_service;

use std::collections::HashMap;
use std::error::Error;
use std::ffi::OsStr;
use std::fs::{read_dir, read_to_string};
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::mpsc;

use notify::{
    recommended_watcher, Event, EventKind, RecursiveMode, Result as NotifyResult, Watcher,
};
use serde::Deserialize;
use toml::from_str;
use win_toast_notify::{Duration as ToastDuration, WinToastNotify};

#[cfg(not(debug_assertions))]
use std::env::var;

#[cfg(not(debug_assertions))]
use std::ffi::OsString;

#[cfg(not(debug_assertions))]
use std::time::Duration;

#[cfg(not(debug_assertions))]
use windows_service::service::{
    ServiceControl, ServiceControlAccept, ServiceExitCode, ServiceState, ServiceStatus, ServiceType,
};

#[cfg(not(debug_assertions))]
use windows_service::service_control_handler::{self, ServiceControlHandlerResult};

#[cfg(not(debug_assertions))]
use windows_service::service_dispatcher;

#[cfg(debug_assertions)]
fn print_debug(message: &str) {
    println!("[DEBUG]: {}", message);
}

#[cfg(not(debug_assertions))]
fn print_debug(_: &str) {}

#[derive(Deserialize)]
struct Settings {
    listened_directory: String,
    filename_prefix: String,
    hidden_filename_prefix: String,
    ext_name: String,
    script_directory: String,
    script_filename: String,
    env_name: String,
}

#[derive(Deserialize)]
struct PathConfig {
    settings: Settings,
}

fn load_config(file_path: &str) -> Result<PathConfig, Box<dyn Error>> {
    let content = read_to_string(file_path)?;
    let config = from_str::<PathConfig>(&content)?;

    Ok(config)
}

fn generate_tiangan_map() -> HashMap<String, usize> {
    let tiangan = vec!["甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸"];

    tiangan
        .into_iter()
        .enumerate()
        .map(|(i, v)| (v.to_string(), i))
        .collect()
}

fn get_tiangan_from_filename(
    filename: &str,
    filename_prefix: &str,
    tiangan_order: &HashMap<String, usize>,
) -> Option<usize> {
    if let Some(pos) = filename.strip_prefix(filename_prefix) {
        tiangan_order.get(pos).cloned()
    } else {
        None
    }
}

fn get_filename_with_largest_tiangan(
    folder_path: &str,
    filename_prefix: &str,
    hidden_filename_prefix: &str,
    ext_name: &str,
    tiangan_order: &HashMap<String, usize>,
    is_hidden_file: bool,
) -> Option<PathBuf> {
    read_dir(folder_path)
        .ok()?
        .filter_map(|entry| entry.ok())
        .filter(|entry| {
            entry
                .path()
                .extension()
                .map(|cur_ext| cur_ext == ext_name)
                .unwrap_or(false)
        })
        .filter_map(|entry| {
            if let Some(version) = get_tiangan_from_filename(
                &entry.path().file_stem()?.to_string_lossy(),
                filename_prefix,
                tiangan_order,
            ) {
                Some((version, entry.path()))
            } else {
                None
            }
        })
        .max_by_key(|(version, _)| *version)
        .map(|(_, path)| {
            if !is_hidden_file {
                return path;
            }

            let new_filename = path
                .file_name()
                .unwrap_or(OsStr::new(""))
                .to_string_lossy()
                .to_string();

            if new_filename.starts_with(filename_prefix) {
                path.with_file_name(new_filename.replace(filename_prefix, hidden_filename_prefix))
            } else {
                path.with_file_name("")
            }
        })
}

fn is_expected_file(
    event: &Event,
    folder_path: &str,
    filename_prefix: &str,
    hidden_filename_prefix: &str,
    ext_name: &str,
    tiangan_order: &HashMap<String, usize>,
    is_hidden_file: bool,
) -> bool {
    if let Some(expected_hidden_filename) = get_filename_with_largest_tiangan(
        folder_path,
        filename_prefix,
        hidden_filename_prefix,
        ext_name,
        tiangan_order,
        is_hidden_file,
    ) {
        event
            .paths
            .iter()
            .any(|path| path == &expected_hidden_filename)
    } else {
        false
    }
}

fn is_same_file(event: &Event, expected_filename: &str) -> bool {
    get_filename_from_event(event).map_or(false, |filename| filename == expected_filename)
}

fn run_script(directory: &str, filename: &str, env_name: &str) -> bool {
    if !Path::new(directory).exists() {
        return false;
    }

    if !Path::new(directory).join(filename).exists() {
        return false;
    }

    print_debug(&format!("Running {}", filename));
    match Command::new("cmd")
        .arg("/C")
        .arg(format!(
            "conda activate {} && python {} -m SheetWizard",
            env_name, filename
        ))
        .current_dir(directory)
        .status()
    {
        Ok(exit_status) => {
            if exit_status.success() {
                print_debug("Executed script successfully");

                true
            } else {
                print_debug(&format!(
                    "Executed script failed with exit code: {}",
                    exit_status.code().unwrap_or(-1)
                ));

                false
            }
        }
        Err(_) => false,
    }
}

fn show_notification(title: &str, message: &str) {
    WinToastNotify::new()
        .set_title(title)
        .set_messages(vec![message])
        .set_duration(ToastDuration::Short)
        .show()
        .expect("Failed to show toast notification")
}

fn get_filename_from_event(event: &Event) -> Option<String> {
    event.paths.iter().find_map(|path| {
        path.file_name()
            .map(|name| name.to_string_lossy().to_string())
    })
}

fn run_watcher(
    config_path: &str,
    tx: mpsc::Sender<NotifyResult<Event>>,
    rx: &mpsc::Receiver<NotifyResult<Event>>,
) -> Result<(), Box<dyn Error>> {
    let path_config = load_config(config_path)?;
    let mut watcher = recommended_watcher(tx)?;
    let tiangan_order = generate_tiangan_map();
    let mut is_expected_hidden_file_opened = false;
    let mut is_expected_file_modified = false;
    let mut cur_expected_hidden_filename = "".to_string();

    watcher
        .watch(
            Path::new(&path_config.settings.listened_directory),
            RecursiveMode::Recursive,
        )
        .unwrap_or(());

    for res in rx {
        match res {
            Ok(event) => match event.kind {
                EventKind::Create(_) => {
                    if is_expected_file(
                        &event,
                        &path_config.settings.listened_directory,
                        &path_config.settings.filename_prefix,
                        &path_config.settings.hidden_filename_prefix,
                        &path_config.settings.ext_name,
                        &tiangan_order,
                        true,
                    ) {
                        cur_expected_hidden_filename =
                            get_filename_from_event(&event).unwrap_or("".to_string());
                        is_expected_hidden_file_opened = true;
                        is_expected_file_modified = false;
                        print_debug(&format!("{} opened", cur_expected_hidden_filename));
                    }
                }
                EventKind::Modify(_) => {
                    if is_expected_file(
                        &event,
                        &path_config.settings.listened_directory,
                        &path_config.settings.filename_prefix,
                        &path_config.settings.hidden_filename_prefix,
                        &path_config.settings.ext_name,
                        &tiangan_order,
                        false,
                    ) {
                        is_expected_file_modified = true;
                    }
                }
                EventKind::Remove(_) => {
                    if is_expected_hidden_file_opened
                        && is_expected_file_modified
                        && is_same_file(&event, &cur_expected_hidden_filename)
                    {
                        cur_expected_hidden_filename = "".to_string();
                        is_expected_hidden_file_opened = false;
                        is_expected_file_modified = false;

                        let success = run_script(
                            &path_config.settings.script_directory,
                            &path_config.settings.script_filename,
                            &path_config.settings.env_name,
                        );

                        print_debug(&format!("{} closed", cur_expected_hidden_filename));

                        if success {
                            show_notification("Sheet Wizard", "Processed successfully.");
                        } else {
                            show_notification(
                                "Sheet Wizard",
                                "Processing failed, the file may not have changed.",
                            );
                        }
                    }
                }
                EventKind::Access(_) => {}
                _ => {
                    break;
                }
            },
            Err(e) => {
                print_debug(&format!("Error occurred in watcher: {:?}", e));
            }
        }
    }

    Ok(())
}

#[cfg(debug_assertions)]
fn run_debug_service() -> Result<(), Box<dyn Error>> {
    let (tx, rx) = mpsc::channel::<NotifyResult<Event>>();

    run_watcher("./path.toml", tx, &rx)
}

#[cfg(not(debug_assertions))]
fn run_service() -> Result<(), Box<dyn Error>> {
    let (tx, rx) = mpsc::channel::<NotifyResult<Event>>();
    let tx_clone = tx.clone();
    let status_handle = service_control_handler::register(
        "SheetWizard",
        move |control_event| -> ServiceControlHandlerResult {
            match control_event {
                ServiceControl::Stop => {
                    if let Err(_) = tx_clone.send(Ok(Event::new(EventKind::Other))) {}

                    ServiceControlHandlerResult::NoError
                }
                ServiceControl::Interrogate => ServiceControlHandlerResult::NoError,
                _ => ServiceControlHandlerResult::NotImplemented,
            }
        },
    )?;

    status_handle.set_service_status(ServiceStatus {
        service_type: ServiceType::OWN_PROCESS,
        current_state: ServiceState::Running,
        controls_accepted: ServiceControlAccept::STOP,
        exit_code: ServiceExitCode::Win32(0),
        checkpoint: 0,
        wait_hint: Duration::default(),
        process_id: None,
    })?;

    let path_config_directory = var("SW_TOML_PATH").unwrap_or("./".to_string());

    run_watcher(&format!("{}\\path.toml", path_config_directory), tx, &rx)?;

    status_handle.set_service_status(ServiceStatus {
        service_type: ServiceType::OWN_PROCESS,
        current_state: ServiceState::Stopped,
        controls_accepted: ServiceControlAccept::empty(),
        exit_code: ServiceExitCode::Win32(0),
        checkpoint: 0,
        wait_hint: Duration::default(),
        process_id: None,
    })?;

    Ok(())
}

#[cfg(not(debug_assertions))]
fn run_service_entry(_: Vec<OsString>) {
    if let Err(_) = run_service() {}
}
#[cfg(not(debug_assertions))]
define_windows_service!(ffi_service_main, run_service_entry);

#[cfg(not(debug_assertions))]
fn main() -> Result<(), Box<dyn Error>> {
    service_dispatcher::start("SheetWizard", ffi_service_main)?;

    Ok(())
}

#[cfg(debug_assertions)]
fn main() -> Result<(), Box<dyn Error>> {
    run_debug_service()?;

    Ok(())
}
