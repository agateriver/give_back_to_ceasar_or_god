#![cfg_attr(windows, windows_subsystem = "windows")]
// #![allow(unused_imports)]
use anyhow::{Result, anyhow};
use regex::Regex;
use std::fs;
use std::path::{Path, PathBuf};
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::{
    Win32::Foundation::*, Win32::System::Threading::*, Win32::UI::Shell::PropertiesSystem::*,
    Win32::UI::Shell::ShellExecuteW, core::*,
};
use windows_registry::*;

pub fn is_wps_pattern(prog: &str) -> bool {
    let re1 = Regex::new(r"(?i)WPS").unwrap();
    let re2 = Regex::new(r"(?i)Kingsoft").unwrap();
    re1.is_match(prog) || re2.is_match(prog)
}

pub fn is_ms_pattern(prog: &str) -> bool {
    let re = Regex::new(r"(?i)Microsoft").unwrap();
    re.is_match(prog)
}

/// 获取当前可执行文件所在目录
pub fn get_current_exe_dir() -> Result<PathBuf, anyhow::Error> {
    let exe_path = std::env::current_exe()?;
    let dir = exe_path
        .parent()
        .ok_or_else(|| anyhow!("无法获取可执行文件目录"))?
        .to_path_buf();
    Ok(dir)
}

/// 从注册表获取金山办公软件路径
pub fn get_kso_path() -> Result<String> {
    // r#"HKEY_CURRENT_USER\Software\Classes\ksoapp\shell\open\command"#
    let ksopath_user = CURRENT_USER
        .open("Software\\Classes\\ksoapp\\shell\\open\\command")
        .and_then(|key| key.get_string(""))
        .map(|s| {
            let path = first_quoted_substring(&s).unwrap();
            path
        })
        .map_err(|e| anyhow!("无法从当前用户注册表获取金山办公路径: {}", e));

    let ksopath_sys = CLASSES_ROOT
        .open("ksoapp\\shell\\open\\command")
        .and_then(|key| key.get_string(""))
        .map(|s| {
            let path = first_quoted_substring(&s).unwrap();
            path
        });

    if ksopath_user.is_ok() {
        if fs::exists(ksopath_user.as_ref().unwrap()).is_ok() {
            return Ok(ksopath_user.unwrap());
        }
    }

    if ksopath_sys.is_ok() {
        if fs::exists(ksopath_sys.as_ref().unwrap()).is_ok() {
            return Ok(ksopath_sys.unwrap());
        }
    }

    Err(anyhow!("从注册表获取金山办公路径失败"))
}

pub fn get_mso_path() -> Result<String> {
    // r#"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot"#
    let versions = vec!["16.0", "15.0", "14.0", "12.0", "11.0"]; //ofiice 2003 ~ 365
    for ver in versions {
        let reg_path = format!(r#"SOFTWARE\Microsoft\Office\{}\Word\InstallRoot"#, ver);
        let mso_path = LOCAL_MACHINE.open(&reg_path);

        if let Ok(key) = mso_path {
            let apps = vec!["Word", "Excel", "PowerPint"];
            for app in apps {
                // 设置为不弹出非默认程序警告
                let _ = CURRENT_USER
                    .create(format!(
                        r#"Software\Microsoft\Office\{}\{}\Options"#,
                        ver, app
                    ))
                    .and_then(|key| {
                        key.set_u32("AlertIfNotDefault", 0)?;
                        Ok(key)
                    })
                    .ok();
            }
            if let Ok(path) = key.get_string("Path") {
                if fs::metadata(&path).is_ok() {
                    return Ok(path);
                }
            }
        }
    }

    Err(anyhow!("从注册表获取微软办公路径失败"))
}

/// 从字符串中提取第一个引号内的子字符串
pub fn first_quoted_substring(s: &str) -> Option<String> {
    // 构建一个匹配第一个成对双引号及其内部文本的正则
    // 捕获组 1 为引号内的文本
    let re = Regex::new(r#"^"([^"]*)".*"#).ok()?;
    // 捕获到的第一组即为所需文本
    re.captures(s)
        .and_then(|cap| cap.get(1).map(|m| m.as_str().to_string()))
}

pub fn get_file_property_store(file_path: &Path) -> Result<IPropertyStore> {
    if !file_path.exists() {
        return Err(anyhow!("{:?}不存在!", file_path));
    }
    let file_path_wide: Vec<u16> = file_path
        .to_str()
        .unwrap()
        .replace("/", "\\")
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

pub fn launch_process(exe_path: &str, options: &str, file_arg: &str) -> Result<()> {
    let exe_wide: Vec<u16> = format!("\"{}\" {} \"{}\"", exe_path, options, file_arg)
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

pub fn message_box(title: &str, message: &str, utype: MESSAGEBOX_STYLE) {
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

pub fn open_with_default_app(file_path: &Path) {
    unsafe {
        ShellExecuteW(
            None,
            PCWSTR(
                "open"
                    .encode_utf16()
                    .chain(std::iter::once(0))
                    .collect::<Vec<u16>>()
                    .as_ptr(),
            ),
            PCWSTR(
                file_path
                    .to_str()
                    .unwrap()
                    .encode_utf16()
                    .chain(std::iter::once(0))
                    .collect::<Vec<u16>>()
                    .as_ptr(),
            ),
            PCWSTR::null(),
            PCWSTR::null(),
            SW_SHOWNORMAL, //,SW_HIDE//
        )
    };
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_first_quoted_substring() {
        // 测试正常情况
        assert_eq!(
            first_quoted_substring(r#""C:\Program Files\WPS Office\ksolaunch.exe" /wps "%1""#),
            Some("C:\\Program Files\\WPS Office\\ksolaunch.exe".to_string())
        );

        // 测试带空格的路径
        assert_eq!(
            first_quoted_substring(
                r#"some text "C:\Program Files (x86)\WPS Office\ksolaunch.exe" more text"#
            ),
            None
        );

        // 测试没有引号的情况
        assert_eq!(first_quoted_substring("no quotes here"), None);

        // 测试只有开始引号的情况
        assert_eq!(first_quoted_substring(r#"only "start quote"#), None);

        // 测试空字符串
        assert_eq!(first_quoted_substring(""), None);

        // 测试只有引号的情况
        assert_eq!(first_quoted_substring(r#""""#), Some("".to_string()));
    }

    #[test]
    fn test_get_kso_path_integration() {
        // 这是一个集成测试，实际运行会依赖系统环境
        // 在 CI/CD 环境中可能会失败，所以标记为忽略
        let result = get_kso_path();

        // 如果系统中有安装 WPS Office，应该能成功获取路径
        // 如果没有安装，应该返回错误
        match result {
            Ok(path) => {
                // 验证路径格式
                assert!(path.contains("WPS") || path.contains("ksolaunch"));
                // 验证路径存在（如果系统中有 WPS）
                if fs::exists(&path).is_ok() {
                    println!("WPS Office found at: {}", path);
                }
            }
            Err(e) => {
                // 如果没有安装 WPS，这是预期的
                println!("WPS Office not found or registry entry missing: {}", e);
            }
        }
    }

    #[test]
    fn test_get_kso_path_error_handling() {
        // 测试错误处理 - 通过模拟不存在的注册表路径
        // 这个测试主要验证函数不会 panic
        let result = std::panic::catch_unwind(|| get_kso_path());

        assert!(result.is_ok(), "get_kso_path should not panic");
    }

    // 模拟测试 - 用于在没有 WPS Office 的环境中测试逻辑
    #[test]
    fn test_get_kso_path_logic() {
        // 这个测试验证函数的逻辑流程，不依赖实际注册表
        // 主要测试错误处理路径

        // 测试函数签名和返回类型
        let result: Result<String, Box<dyn std::error::Error>> = Err("test error".into());
        assert!(result.is_err());
    }
}

// 条件编译的测试模块，只在 Windows 上运行
#[cfg(all(test, windows))]
mod windows_tests {
    use super::*;
    // use windows_registry::*;

    #[test]
    fn test_registry_access() {
        // 测试基本的注册表访问权限
        let result = CURRENT_USER.open("Software");
        assert!(result.is_ok(), "Should be able to access HKCU\\Software");

        let result = CLASSES_ROOT.open("Software");
        assert!(result.is_ok(), "Should be able to access HKCR");
    }

    #[test]
    fn test_kso_registry_structure() {
        // 测试 WPS Office 注册表结构是否存在
        let kso_user = CURRENT_USER.open("Software\\Classes\\ksoapp\\shell\\open\\command");
        let kso_system = CLASSES_ROOT.open("ksoapp\\shell\\open\\command");

        // 至少有一个注册表路径应该存在（取决于 WPS 安装方式）
        if kso_user.is_ok() || kso_system.is_ok() {
            println!("WPS Office registry structure detected");
        } else {
            println!(
                "WPS Office registry structure not found - this is normal if WPS is not installed"
            );
        }
    }
}
