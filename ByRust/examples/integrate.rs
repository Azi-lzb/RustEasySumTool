use excel_summary::{AggregationRequest, RunControl, aggregate};
use std::sync::atomic::AtomicBool;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let paths = std::env::args_os().skip(1).collect::<Vec<_>>();
    if paths.len() < 4 {
        eprintln!("用法：integrate 模板.xlsx 结果.xlsx config.xlsx 源文件.xlsx [更多源文件.xlsx]");
        std::process::exit(2);
    }
    let request = AggregationRequest {
        template_path: (&paths[0]).into(),
        output_path: (&paths[1]).into(),
        config_path: Some((&paths[2]).into()),
        source_paths: paths[3..].iter().map(Into::into).collect(),
    };
    let cancel = AtomicBool::new(false);
    let progress = |done: usize, total: usize, message: &str| {
        eprintln!("{done}/{total} {message}");
    };
    let result = aggregate(
        &request,
        &RunControl {
            cancel: Some(&cancel),
            progress: Some(&progress),
        },
    )?;
    println!("{}", serde_json::to_string_pretty(&result)?);
    Ok(())
}
