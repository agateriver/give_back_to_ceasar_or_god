#[cfg(windows)]
fn main() {
    let mut res = winres::WindowsResource::new();
    res.set_icon("assets/app.ico");
    res.set("FileDescription", "M$ or WP$ ?");
    res.set("ProductName", "M$ or WP$ ?");
    res.compile().unwrap();
    // if cfg!(debug_assertions) {
    //     copy_to_output::copy_to_output("src/config.toml", "debug").unwrap();
    // } else {
    //     copy_to_output::copy_to_output("src/config.toml", "release").unwrap();
    // }

}
