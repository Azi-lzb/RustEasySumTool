use clap::Parser;
use excel_summary::{AggregationRequest, RunControl, aggregate, default_output_path};
use std::{
    io::{self, Write},
    path::PathBuf,
    sync::{
        Arc,
        atomic::{AtomicBool, Ordering},
    },
};

#[derive(Parser)]
#[command(
    name = "Excel 模板汇总工具 Rust",
    version,
    about = "按模板中的汇总指令，合并多份Excel工作簿"
)]
struct Arguments {
    #[arg(long)]
    template: Option<PathBuf>,
    #[arg(long,num_args=0..)]
    sources: Vec<PathBuf>,
    #[arg(long)]
    output: Option<PathBuf>,
    #[arg(long)]
    config: Option<PathBuf>,
    /// Emit a machine-readable result/error on stdout; progress goes to stderr.
    #[arg(long)]
    json: bool,
    #[arg(long, hide = true)]
    check_dialogs: bool,
}
fn execute(request: AggregationRequest, json: bool, cancel: &AtomicBool) -> bool {
    cancel.store(false, Ordering::Relaxed);
    let progress = |_: usize, _: usize, message: &str| {
        eprintln!("{message}");
    };
    let result = aggregate(
        &request,
        &RunControl {
            cancel: Some(cancel),
            progress: Some(&progress),
        },
    );
    if json {
        let payload = match &result {
            Ok(result) => serde_json::json!({"ok":true,"result":result}),
            Err(error) => serde_json::json!({"ok":false,"error":error}),
        };
        println!(
            "{}",
            serde_json::to_string(&payload).expect("serializable aggregation result")
        );
    } else {
        match &result {
            Ok(result) => {
                println!("已生成：{}", result.output_path.display());
                println!(
                    "来源文件 {} 份，汇总指令 {} 个。",
                    result.source_count, result.rule_count
                );
            }
            Err(error) => {
                eprintln!("{}", error.message);
                for issue in error
                    .issues
                    .iter()
                    .filter(|issue| issue.severity == excel_summary::model::Severity::Error)
                {
                    eprintln!(
                        "[{}] 模板 {}!{} 源 {} {}!{}：{}",
                        issue.code,
                        issue.template_sheet,
                        issue.template_cell,
                        issue.source_file,
                        issue.source_sheet,
                        issue.source_cell,
                        issue.message
                    );
                }
            }
        }
    }
    result.is_ok()
}
fn choose(config: Option<PathBuf>) -> Option<AggregationRequest> {
    let directory = std::env::current_exe()
        .ok()
        .and_then(|path| path.parent().map(|parent| parent.to_path_buf()))
        .unwrap_or_default();
    let template = rfd::FileDialog::new()
        .set_title("选择汇总模板")
        .set_directory(&directory)
        .add_filter("Excel工作簿", &["xlsx"])
        .pick_file()?;
    println!("模板：{}", template.display());
    let sources = rfd::FileDialog::new()
        .set_title("选择来源工作簿（可多选）")
        .set_directory(template.parent().unwrap_or(&directory))
        .add_filter("Excel工作簿", &["xlsx"])
        .pick_files()?;
    println!("已选择{}份来源文件。", sources.len());
    let output_directory = rfd::FileDialog::new()
        .set_title("选择结果保存目录")
        .set_directory(&directory)
        .pick_folder()?;
    Some(AggregationRequest {
        output_path: default_output_path(&template, &output_directory),
        template_path: template,
        source_paths: sources,
        config_path: config,
    })
}
fn main() -> std::process::ExitCode {
    let args = Arguments::parse();
    let cancel = Arc::new(AtomicBool::new(false));
    let signal = Arc::clone(&cancel);
    if let Err(error) = ctrlc::set_handler(move || {
        signal.store(true, Ordering::Relaxed);
    }) {
        eprintln!("无法设置取消处理：{error}");
    }
    if args.check_dialogs {
        let _ = rfd::FileDialog::new()
            .set_title("文件对话框检查")
            .add_filter("Excel工作簿", &["xlsx"]);
        println!("本机文件对话框入口构建成功。");
        return std::process::ExitCode::SUCCESS;
    }
    if args.template.is_some() || !args.sources.is_empty() || args.output.is_some() {
        let (Some(template), Some(output)) = (args.template, args.output) else {
            eprintln!("批处理需要同时提供--template、--sources和--output。");
            return std::process::ExitCode::from(2);
        };
        return std::process::ExitCode::from(
            if execute(
                AggregationRequest {
                    template_path: template,
                    source_paths: args.sources,
                    output_path: output,
                    config_path: args.config,
                },
                args.json,
                &cancel,
            ) {
                0
            } else {
                1
            },
        );
    }
    loop {
        println!(
            "\nExcel 模板汇总工具 Rust v{}\n1. 按模板公式汇总\n2. 退出",
            env!("CARGO_PKG_VERSION")
        );
        let _ = io::stdout().flush();
        let mut selection = String::new();
        if io::stdin().read_line(&mut selection).unwrap_or(0) == 0 {
            break;
        }
        match selection.trim() {
            "1" => {
                if let Some(request) = choose(args.config.clone()) {
                    execute(request, false, &cancel);
                } else {
                    println!("已取消选择。");
                }
            }
            "2" => break,
            _ => println!("请输入1或2。"),
        }
    }
    std::process::ExitCode::SUCCESS
}
