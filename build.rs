use chrono::Local;
use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};
use zip::write::FileOptions;

fn create_release_zip() -> Result<(), Box<dyn std::error::Error>> {
    let target_dir = Path::new("target/release");
    let date = format!(
        "target/release/Schrödinger's Office_v{}.zip",
        Local::now().format("%Y%m%d")
    );
    let zip_path = Path::new(date.as_str());

    // 确保目标目录存在
    if !target_dir.exists() {
        return Err("Release directory does not exist".into());
    }

    // 创建 zip 文件
    let file = fs::File::create(zip_path)?;
    let mut zip = zip::ZipWriter::new(file);
    let options: FileOptions<()> =
        FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

    // 要包含在 zip 中的文件列表
    let files_to_include = vec![
        "target/release/give_back_to_ceasar_or_god.exe",
        "assets/WORDICON.EXE",
        "assets/PPTICO.EXE",
        "assets/XLICONS.EXE",
        "assets/register.cmd",
        "README.md",
        "test/test_ms.docx",
        "test/test_ms.pptx",
        "test/test_ms.xlsx",
        "test/test_kso.doc",
        "test/test_kso.ppt",
        "test/test_kso.xls",
    ];

    for file_name in files_to_include {
        let file_path = PathBuf::from(file_name);
        if file_path.starts_with("target")||file_path.ends_with("register.cmd") {
            if file_path.exists() {
                let file_data = fs::read(&file_path)?;
                zip.start_file(file_path.file_name().unwrap().to_string_lossy(), options)?;
                zip.write_all(&file_data)?;
                println!("Added {} to zip", file_name);
            } else {
                println!("Warning: {} not found", file_name);
            }
        }else{
            if file_path.exists() {
                let file_data = fs::read(&file_path)?;
                zip.start_file(file_path.to_string_lossy(), options)?;
                zip.write_all(&file_data)?;
                println!("Added {} to zip", file_name);
            } else {
                println!("Warning: {} not found", file_name);
            }
        }
    }
    zip.finish()?;
    println!("Successfully created release zip: {}", zip_path.display());

    Ok(())
}

fn main() {
    let mut res = winres::WindowsResource::new();
    res.set_icon("assets/app.ico");
    res.set(
        "FileDescription",
        "自动调用 Microsoft 或 WPS Office 组件打开文档",
    );
    res.set("ProductName", "M$ or WP$ ?");
    res.compile().unwrap();

    if cfg!(debug_assertions) {
        copy_to_output::copy_to_output("assets/", "debug").unwrap();
        copy_to_output::copy_to_output("test/", "debug").unwrap();
        copy_to_output::copy_to_output("assets/register.cmd", "debug").unwrap();
        let _ = fs::remove_file("target/debug/assets/register.cmd");
        let _ = fs::remove_file("target/debug/assets/app.ico");
        copy_to_output::copy_to_output("README.md", "debug").unwrap();
    } else {
        copy_to_output::copy_to_output("assets/", "release").unwrap();
        copy_to_output::copy_to_output("test/", "release").unwrap();
        copy_to_output::copy_to_output("assets/register.cmd", "release").unwrap();
        copy_to_output::copy_to_output("README.md", "debug").unwrap();
        let _ = fs::remove_file("target/release/assets/register.cmd");
        let _ = fs::remove_file("target/release/assets/app.ico");

        // 在发布模式下创建 zip 文件
        create_release_zip().unwrap_or_else(|e| {
            println!("Warning: Failed to create release zip: {}", e);
        });
    }
}
