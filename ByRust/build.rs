fn main() {
    if std::env::var("CARGO_CFG_TARGET_OS").as_deref() == Ok("windows")
        && std::env::var_os("CARGO_FEATURE_CLI").is_some()
    {
        let mut resource = winresource::WindowsResource::new();
        resource
            .set("ProductName", "Excel 模板汇总工具 Rust")
            .set("FileDescription", "Excel 模板汇总工具 Rust")
            .set("InternalName", "excel-summary-rust")
            .set("OriginalFilename", "excel-summary-rust.exe")
            .set("FileVersion", env!("CARGO_PKG_VERSION"))
            .set("ProductVersion", env!("CARGO_PKG_VERSION"));
        resource
            .compile()
            .expect("compile Windows version resource");
    }
}
