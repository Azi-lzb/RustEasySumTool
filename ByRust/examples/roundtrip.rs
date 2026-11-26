fn main() -> Result<(), Box<dyn std::error::Error>> {
    let arguments: Vec<_> = std::env::args_os().collect();
    let workbook = umya_spreadsheet::reader::xlsx::read(std::path::Path::new(&arguments[1]))?;
    umya_spreadsheet::writer::xlsx::write(&workbook, std::path::Path::new(&arguments[2]))?;
    Ok(())
}
