use std::fs;
use fs_extra;
use anyhow::Result;

fn main() -> Result<()> {
    let mut res = winres::WindowsResource::new();
    res.set_icon("assets/app.ico");
    res.set(
        "FileDescription",
        "自动调用 Microsoft 或 WPS Office 组件打开文档",
    );
    res.set("ProductName", "M$ or WP$ ?");
    res.compile().unwrap();

    let  options = fs_extra::dir::CopyOptions::new().overwrite(true);

    if cfg!(debug_assertions) {
        let _= fs::remove_dir_all(std::path::Path::new("target/debug/assets"));
        let _= fs::create_dir(std::path::Path::new("target/debug/assets"));
        fs_extra::copy_items(
            &vec!["assets/register.cmd","README.md"],
            "target/debug/",
            &options,
        )?;
        fs_extra::copy_items(
            &vec!["assets/PPTICO.exe","assets/WORDICON.exe","assets/XLICONS.exe"],
            "target/debug/assets/",
            &options,
        )?; 
    } else {
        let _= fs::remove_dir_all(std::path::Path::new("target/debug/assets"));
        let _= fs::create_dir(std::path::Path::new("target/release/assets"));
        fs_extra::copy_items(
            &vec!["assets/register.cmd","README.md"],
            "target/release/",
            &options,
        )?;
        fs_extra::copy_items(
            &vec!["assets/PPTICO.exe","assets/WORDICON.exe","assets/XLICONS.exe"],
            "target/release/assets/",
            &options,
        )?;       
    }
    Ok(())
}
