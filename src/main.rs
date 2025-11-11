#![windows_subsystem = "windows"]

use anyhow::{Context, Result, anyhow};
use clap::{CommandFactory, Parser};
use give_back_to_ceasar_or_god::*;
use serde::Deserialize;
use std::ffi::OsString;
use std::path::*;

use windows::{
    Win32::System::Com::{COINIT_MULTITHREADED, CoInitializeEx, CoUninitialize},
    Win32::Storage::EnhancedStorage::*, 
    Win32::System::Variant::*,
    Win32::UI::WindowsAndMessaging::*,
};
use windows_registry::*;

#[derive(Parser, Debug)]
#[command(
    // name = "Schrödinger's Office Launcher",
    version = "1.0",
    author = "X.B.G",
    about = "A utility to auto open office documents with WPS or MS Office based on file properties"
)]
struct Cli {
    /// 注册文件关联
    #[arg(short, long)]
    registry: bool,

    /// 打开文件路径
    #[arg(short, long, value_name = "FILE")]
    open: Option<PathBuf>,
}
#[derive(Deserialize)]
struct ExePaths {
    wps: String,
    word: String,
    powerpoint: String,
    excel: String,
}

struct ComInitializer;

impl ComInitializer {
    fn new() -> Result<Self> {
        unsafe {
            let result = CoInitializeEx(None, COINIT_MULTITHREADED);
            if result.is_err() {
                return Err(anyhow!("Failed to initialize COM"));
            }
        }
        Ok(ComInitializer)
    }
}

impl Drop for ComInitializer {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

fn get_program_name_from_meta(file_path: &Path) -> Result<String> {
    let store = get_file_property_store(file_path)?; //.context("failure to get property store.")?;
    unsafe {
        let pkey_value = store.GetValue(&PKEY_ApplicationName);
        match pkey_value {
            Ok(prop_var) => {
                let value = match prop_var.vt() {
                    VT_LPWSTR => prop_var.to_string(),
                    _ => "<Unsupported Type>".to_string(),
                };
                return Ok(value);
            }
            Err(e) => {
                return Err(anyhow!("Failed to get documnet property value: {}", e));
            }
        }
    }
}

fn register_file_associations(exe_paths: &ExePaths) -> Result<()> {
    let app_path = std::env::current_exe()?;
    let app_path_str = app_path
        .to_str()
        .ok_or_else(|| anyhow::anyhow!("无法获取可执行文件路径"))?;

    let extensions = vec![".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"];

    for ext in extensions {
        let prog_id = format!(
            "Schrödinger's {}",
            String::from(ext).to_uppercase().strip_prefix('.').unwrap()
        );
        let description = match ext {
            ".doc" => "Microsoft Word 97-2003 Document",
            ".docx" => "Microsoft Word Document",
            ".xls" => "Microsoft Excel 97-2003 Spreadsheet",
            ".xlsx" => "Microsoft Excel Spreadsheet",
            ".ppt" => "Microsoft PowerPoint 97-2003 Presentation",
            ".pptx" => "Microsoft PowerPoint Presentation",
            _ => ext,
        };

        // 设置文件扩展名关联
        set_file_association(ext, &prog_id, &description, app_path_str, exe_paths)
            .map_err(|e| anyhow::anyhow!("Failed to set file association for {}: {}", ext, e))?;
    }

    Ok(())
}

fn set_file_association(
    ext: &str,
    prog_id: &str,
    description: &str,
    app_path_str: &str,
    _exe_paths: &ExePaths,
) -> Result<()> {
    // 创建程序ID注册表项
    let prog_id_key = CURRENT_USER
        .create(&format!("Software\\Classes\\{}", prog_id))
        .map_err(|e| anyhow!("Failed to create ProgID key: {}", e))?;

    // 设置程序描述
    prog_id_key
        .set_string("", description)
        .map_err(|e| anyhow!("Failed to set ProgID description: {}", e))?;

    // 创建shell\open\command路径
    let command_key = CURRENT_USER
        .create(&format!(
            "Software\\Classes\\{}\\shell\\open\\command",
            prog_id
        ))
        .map_err(|e| anyhow!("Failed to create command key: {}", e))?;

    // 设置打开命令，使用引号包围路径以处理空格
    let command = format!("\"{}\" --open \"%1\"", app_path_str);
    command_key
        .set_string("", &command)
        .map_err(|e| anyhow!("Failed to set command: {}", e))?;

    // 设置文件扩展名关联
    let ext_key = CURRENT_USER
        .create(&format!("Software\\Classes\\{}", ext))
        .map_err(|e| anyhow!("Failed to create extension key: {}", e))?;

    // 将扩展名指向程序ID
    ext_key
        .set_string("", &prog_id)
        .map_err(|e| anyhow!("Failed to set extension association: {}", e))?;

    // 清除UserChoice键和OpenWithList键，防止系统优先使用UserChoice中的设置
    let _ = CURRENT_USER
        .create("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts")
        .and_then(|file_exts_key| {
            let _ = file_exts_key.create(ext).and_then(|ext_key: Key| {
                ext_key.remove_tree("UserChoice").ok();
                ext_key.remove_tree("OpenWithList").ok();
                ext_key.remove_tree("OpenWithProgids").ok();
                Ok(ext_key)
            });
            Ok(file_exts_key)
        })
        .ok();

    // 设置OpenWithList ,设置此程序为当前用户下的唯一打开项，让系统不用弹出打开程序选择框
    let _open_with_list_key = CURRENT_USER
        .create(&format!(
            "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\{}\\OpenWithList",
            ext
        ))
        .and_then(|key| {
            key.set_string(
                "a",
                PathBuf::from(app_path_str)
                    .file_name()
                    .unwrap()
                    .to_str()
                    .unwrap(),
            )?;
            key.set_string("MRUList", "a")?;
            Ok(key)
        })
        .map_err(|e| anyhow!("Failed to create OpenWithList key: {}", e))?;

    // 设置默认图标
    let icon = match ext {
        ".doc" | ".docx" => {
            let icon_path = PathBuf::from(&app_path_str)
                .parent()
                .unwrap()
                .join("assets")
                .join("wordicon.exe,13");
            icon_path.to_string_lossy().to_string()
        }
        ".xls" | ".xlsx" => {
            let icon_path = PathBuf::from(&app_path_str)
                .parent()
                .unwrap()
                .join("assets")
                .join("xlicons.exe,1");
            icon_path.to_string_lossy().to_string()
        }
        ".ppt" | ".pptx" => {
            let icon_path = PathBuf::from(&app_path_str)
                .parent()
                .unwrap()
                .join("assets")
                .join("pptico.exe,10");
            icon_path.to_string_lossy().to_string()
        }
        _ => String::new(),
    };
    let default_icon_key = CURRENT_USER
        .create(&format!("Software\\Classes\\{}\\DefaultIcon", prog_id))
        .map_err(|e| anyhow!("Failed to create DefaultIcon key: {}", e))?;
    default_icon_key
        .set_string("", &format!("{}", icon))
        .map_err(|e| anyhow!("Failed to set default icon: {}", e))?;

    Ok(())
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    // 如果没有提供任何参数，显示帮助信息
    if !cli.registry && cli.open.is_none() {
        let message = Cli::command()
            .disable_help_flag(true)
            .disable_version_flag(true)
            .no_binary_name(true)
            .render_help()
            .to_string();
        message_box("HELP", &message, MB_OK);
        return Ok(());
    }

    let kso_path = get_kso_path()?;
    let mso_root = get_mso_path()?;
    let exe_paths = &ExePaths {
        wps: kso_path,
        word: format!("{}WINWORD.EXE", &mso_root.clone()),
        powerpoint: format!("{}POWERPNT.EXE", &mso_root.clone()),
        excel: format!("{}EXCEL.EXE", &mso_root.clone()),
    };

    if cli.registry {
        // 注册文件关联
        register_file_associations(exe_paths).context("Failed to register file associations")?;
        message_box("OK", "注册成功", MB_OK);
        return Ok(());
    }

    if cli.open.is_some() {
        if let Some(file_path) = cli.open.as_deref() {
            if !file_path.exists() {
                message_box("错误", "目标文件不存在", MB_ICONWARNING);
                anyhow::bail!("File does not exist: {}", file_path.to_str().unwrap());
            }

            let _com = ComInitializer::new()?;

            let pkey_application_name = get_program_name_from_meta(&file_path); //.context("Failed to get 'PKEY_ApplicationName'")?;
            if pkey_application_name.is_err() {
                open_with_default_app(file_path);
                return Ok(());
            };

            let program_name = pkey_application_name.unwrap();
            let ext: OsString = file_path.extension().unwrap().to_ascii_lowercase();

            let mut res: Result<(), anyhow::Error> = Err(anyhow::anyhow!("Unknown file type"));

            if is_wps_pattern(&program_name) {
                let exe_path = &exe_paths.wps;
                println!("Launching WPS Office for file: {}", &exe_path);
                match ext.to_str() {
                    Some("doc") | Some("docx") => {
                        res = launch_process(&exe_path, "/wps", file_path.to_str().unwrap())
                            .context("Failed to launch WPS Writer process");
                    }
                    Some("ppt") | Some("pptx") => {
                        res = launch_process(&exe_path, "/wpp", file_path.to_str().unwrap())
                            .context("Failed to launch WPS Presentation process");
                    }
                    Some("xls") | Some("xlsx") => {
                        res = launch_process(&exe_path, "/et", file_path.to_str().unwrap())
                            .context("Failed to launch WPS Spreadsheets process");
                    }
                    _ => {}
                };
            } else if is_ms_pattern(&program_name) {
                match ext.to_str() {
                    Some("doc") | Some("docx") => {
                        let exe_path = &exe_paths.word;
                        res = launch_process(&exe_path, "", file_path.to_str().unwrap())
                            .context("Failed to launch WinWord.exe process");
                    }
                    Some("ppt") | Some("pptx") => {
                        let exe_path = &exe_paths.powerpoint;
                        res = launch_process(&exe_path, "", file_path.to_str().unwrap())
                            .context("Failed to launch PowerPnt.exe process");
                    }
                    Some("xls") | Some("xlsx") => {
                        let exe_path = &exe_paths.excel;
                        res = launch_process(&exe_path, "", file_path.to_str().unwrap())
                            .context("Failed to launch Excel.exe process");
                    }
                    _ => {}
                }
            };
            if res.is_err() {
                //未知类型，使用系统默认打开方式
                open_with_default_app(file_path);
            }
            return Ok(());
        }
    }
    Ok(())
}
