#![windows_subsystem = "windows"]
use anyhow::{Context, Result, anyhow};
use regex::Regex;
use serde::Deserialize;
use windows::Win32::UI::WindowsAndMessaging::MB_ICONWARNING;
use windows::Win32::UI::WindowsAndMessaging::MESSAGEBOX_STYLE;
use std::ffi::OsString;
use std::fs;
use std::path::*;
use toml;
use windows::{
    Win32::Foundation::*,
    Win32::Storage::EnhancedStorage::*,
    Win32::System::Com::*,
    Win32::System::Threading::*,
    Win32::System::Variant::*,
    Win32::UI::Shell::PropertiesSystem::*,
    Win32::UI::Shell::ShellExecuteW,
    Win32::UI::WindowsAndMessaging::*,
    core::*,
};
use windows_registry::*;
use windows_strings::PCWSTR;
use clap::{Parser, CommandFactory};

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct  Cli {
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

#[derive(Deserialize)]
struct Config {
    paths: ExePaths,
}

fn message_box(title: &str, message: &str,utype:MESSAGEBOX_STYLE) { 
    let title_wide: Vec<u16> = title.encode_utf16().chain(std::iter::once(0)).collect();
    let message_wide: Vec<u16> = message.encode_utf16().chain(std::iter::once(0)).collect();
    unsafe {
        MessageBoxW(
            None,
            PCWSTR(message_wide.as_ptr()),
            PCWSTR(title_wide.as_ptr()),
            utype,
        );
    }
}
fn read_config(config_path: &str) -> anyhow::Result<Config> {
    let content = fs::read_to_string(config_path).map_err(|e| {
        message_box("错误", "M$_or_WP$ 配置文件读取失败",MB_ICONWARNING);
        anyhow::anyhow!("Failed to read config: {}", e)
    })?;
    let cfg: Config = toml::from_str(&content).map_err(|e| {
        message_box("错误", "M$_or_WP$ 配置文件解析失败",MB_ICONWARNING);
        anyhow::anyhow!("Failed to parse config: {}", e)
    })?;
    Ok(cfg)
}

fn is_wps_pattern(prog: &str) -> bool {
    let re = Regex::new(r"(?i)wps").unwrap();
    re.is_match(prog)
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

fn get_file_property_store(file_path: &Path) -> Result<IPropertyStore> {
    if !file_path.exists(){
        return Err(anyhow!("{:?}不存在!",file_path));
    }
    let file_path_wide: Vec<u16> = file_path
        .to_str()
        .unwrap().replace("/", "\\")
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect();
    let path = PCWSTR(file_path_wide.as_ptr());

    unsafe {
        SHGetPropertyStoreFromParsingName(
            path,
            None,
            GPS_DEFAULT, // &IPropertyStore::IID
        )
        .map_err(|e| anyhow!("in fn get_file_property_store: {}", e))
    }
}

fn get_program_name(file_path: &Path) -> Result<String> {
    let store = get_file_property_store(file_path)?;//.context("failure to get property store.")?;
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
    //
}

fn launch_process(exe_path: &str, options: &str ,file_arg: &str) -> Result<()> {
    let exe_wide: Vec<u16> = format!("\"{}\" {} \"{}\"", exe_path, options,file_arg)
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect();

    let mut si: STARTUPINFOW = unsafe { std::mem::zeroed() };
    si.cb = std::mem::size_of::<STARTUPINFOW>() as u32;

    let mut pi: PROCESS_INFORMATION = unsafe { std::mem::zeroed() };
    let success = unsafe {
        windows::Win32::System::Threading::CreateProcessW(
            None,
            Some(PWSTR(exe_wide.as_ptr() as *mut _)),
            None,
            None,
            FALSE.into(),
            CREATE_NEW_CONSOLE | CREATE_NO_WINDOW,
            None,
            None,
            &mut si,
            &mut pi,
        )
    };

    if success.is_ok() {
        unsafe {
            let _ = CloseHandle(pi.hProcess);
            let _ = CloseHandle(pi.hThread);
        }
        Ok(())
    } else {
        anyhow::bail!("CreateProcessW failed");
    }
}

fn get_current_exe_dir() -> Result<PathBuf, anyhow::Error> {
    let exe_path = std::env::current_exe()?;
    let dir = exe_path
        .parent()
        .ok_or_else(|| anyhow::anyhow!("无法获取可执行文件目录"))?
        .to_path_buf();
    Ok(dir)
}

fn register_file_associations(exe_paths:&ExePaths) -> Result<()> {
    let app_path = std::env::current_exe()?;
    let app_path_str = app_path.to_str().ok_or_else(|| anyhow::anyhow!("无法获取可执行文件路径"))?;

    let extensions = vec![".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"];

    for ext in extensions {
        let prog_id = format!("Schrödinger's {}", String::from(ext).to_uppercase().strip_prefix('.').unwrap());
        let description = match ext{
            ".doc" => "Microsoft Word 97-2003 Document",
            ".docx" => "Microsoft Word Document",
            ".xls" => "Microsoft Excel 97-2003 Spreadsheet",
            ".xlsx" => "Microsoft Excel Spreadsheet",
            ".ppt" => "Microsoft PowerPoint 97-2003 Presentation",
            ".pptx" => "Microsoft PowerPoint Presentation",
            _ => ext,
        };

        // 设置文件扩展名关联
        set_file_association(ext, &prog_id, &description, app_path_str,exe_paths)
            .map_err(|e| anyhow::anyhow!("Failed to set file association for {}: {}", ext, e))?;
    }

    Ok(())
}

fn set_file_association(ext: &str, prog_id: &str, description: &str, app_path_str: &str, exe_paths:&ExePaths) -> Result<()>  {
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
        .create(&format!("Software\\Classes\\{}\\shell\\open\\command", prog_id))
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
            let _ = file_exts_key.create(ext).and_then(|ext_key: Key|{
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
        .create(&format!("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\{}\\OpenWithList", ext))
        .and_then(|key| {
            key.set_string("a", PathBuf::from(app_path_str).file_name().unwrap().to_str().unwrap())?;
            key.set_string("MRUList", "a")?;
            Ok(key)
        })
        .map_err(|e| anyhow!("Failed to create OpenWithList key: {}", e))?;

    // 设置默认图标
    let icon = match ext {
        ".doc" | ".docx" => {
            let icon_path = PathBuf::from(&exe_paths.word).parent().unwrap().join("wordicon.exe,13");
            icon_path.to_string_lossy().to_string()
        },
        ".xls" | ".xlsx" => {
            let icon_path = PathBuf::from(&exe_paths.excel).parent().unwrap().join("xlicons.exe,1");
            icon_path.to_string_lossy().to_string()
        },
        ".ppt" | ".pptx" => {
            let icon_path = PathBuf::from(&exe_paths.powerpoint).parent().unwrap().join("pptico.exe,10");
            icon_path.to_string_lossy().to_string()
        },
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
        let message = Cli::command().render_help().to_string();
        message_box("HELP", &message, MB_OK);
        return Ok(());
    }

    let cfg_path = format!("{}\\config.toml", &get_current_exe_dir()?.to_str().unwrap());
    let cfg = read_config(&cfg_path)?;
    if cli.registry {
        // 注册文件关联
        register_file_associations(&cfg.paths).context("Failed to register file associations")?;
        message_box("OK", "注册成功",MB_OK);
        return Ok(());
    }

    if let Some(file_path) = cli.open.as_deref() {
        if !file_path.exists() {
                message_box("错误", "目标文件不存在",MB_ICONWARNING);
                anyhow::bail!("File does not exist: {}", file_path.to_str().unwrap());
            }        

        let _com = ComInitializer::new()?;

        let program_name = get_program_name(&file_path)?; //.context("Failed to get 'PKEY_ApplicationName'")?;

        let ext: OsString = file_path.extension().unwrap().to_ascii_lowercase();

        if is_wps_pattern(&program_name) {
            let exe_path = cfg.paths.wps;
            println!("Launching WPS Office for file: {}", &exe_path);
            match ext.to_str() {
                Some("doc") | Some("docx") => {
                    launch_process(&exe_path,"/wps", file_path.to_str().unwrap())
                        .context("Failed to launch WPS Writer process")?;
                }
                Some("ppt") | Some("pptx") => {
                    launch_process(&exe_path,"/wpp", file_path.to_str().unwrap())
                        .context("Failed to launch WPS Presentation process")?;
                }
                Some("xls") | Some("xlsx") => {
                    launch_process(&exe_path, "/et", file_path.to_str().unwrap())
                        .context("Failed to launch WPS Spreadsheets process")?;
                }
                _ => { //未知类型，使用系统默认打开方式
                    unsafe { ShellExecuteW(
                        None,
                        PCWSTR("open".encode_utf16().chain(std::iter::once(0)).collect::<Vec<u16>>().as_ptr()),
                        PCWSTR(file_path.to_str().unwrap().encode_utf16().chain(std::iter::once(0)).collect::<Vec<u16>>().as_ptr()),
                        PCWSTR::null(),
                        PCWSTR::null(),
                        SW_SHOWNORMAL,
                    ) };
                }
            };

        } else {
            match ext.to_str() {
                Some("doc") | Some("docx") => {
                    let exe_path = cfg.paths.word;
                    launch_process(&exe_path, "",file_path.to_str().unwrap())
                        .context("Failed to launch WinWord.exe process")?;
                }
                Some("ppt") | Some("pptx") => {
                    let exe_path = cfg.paths.powerpoint;
                    launch_process(&exe_path, "",file_path.to_str().unwrap())
                        .context("Failed to launch PowerPnt.exe process")?;
                }
                Some("xls") | Some("xlsx") => {
                    let exe_path = cfg.paths.excel;
                    launch_process(&exe_path, "",file_path.to_str().unwrap())
                        .context("Failed to launch Excel.exe process")?;
                }
                _ => { //未知类型，使用系统默认打开方式
                    unsafe { ShellExecuteW(
                        None,
                        PCWSTR("open".encode_utf16().chain(std::iter::once(0)).collect::<Vec<u16>>().as_ptr()),
                        PCWSTR(file_path.to_str().unwrap().encode_utf16().chain(std::iter::once(0)).collect::<Vec<u16>>().as_ptr()),
                        PCWSTR::null(),
                        PCWSTR::null(),
                         SW_HIDE //SW_SHOWNORMAL,
                    ) };
                }
            };
        }  
        return Ok(());      
    }
    Ok(())
}
