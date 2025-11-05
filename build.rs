use std::fs;

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
        copy_to_output::copy_to_output("README.md", "debug").unwrap();
        let _ = fs::remove_file("target/debug/assets/register.cmd");
        let _ = fs::remove_file("target/debug/assets/app.ico");
    } else {
        copy_to_output::copy_to_output("assets/", "release").unwrap();
        copy_to_output::copy_to_output("test/", "release").unwrap();
        copy_to_output::copy_to_output("README.md", "debug").unwrap();
        let _ = fs::remove_file("target/release/assets/register.cmd");
        let _ = fs::remove_file("target/release/assets/app.ico");
    }
}
