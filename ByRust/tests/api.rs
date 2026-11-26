use excel_summary::{
    AggregationRequest, Aggregator, RunControl, aggregate,
    functions::{CalculationResult, Calculator, FunctionSpec, Notice, ResultSource},
    model::Severity,
    numbers::Number,
};
use std::{
    path::{Path, PathBuf},
    sync::atomic::{AtomicBool, Ordering},
};

fn workbook(path: &Path, sheets: &[(&str, &[(&str, &str)])]) {
    let mut workbook = umya_spreadsheet::new_file();
    workbook.sheet_mut(0).unwrap().set_name(sheets[0].0);
    for (index, (name, cells)) in sheets.iter().enumerate() {
        if index > 0 {
            workbook.new_sheet(*name).unwrap();
        }
        for (coordinate, value) in *cells {
            workbook
                .sheet_by_name_mut(name)
                .unwrap()
                .cell_mut(*coordinate)
                .set_value_string(*value);
        }
    }
    umya_spreadsheet::writer::xlsx::write(&workbook, path).unwrap();
}
fn fixture() -> (tempfile::TempDir, AggregationRequest) {
    let directory = tempfile::tempdir().unwrap();
    let path = |name: &str| directory.path().join(name);
    workbook(&path("template.xlsx"), &[("表", &[("K4", "FUN_1()")])]);
    workbook(&path("source.xlsx"), &[("表", &[("K4", "12")])]);
    workbook(
        &path("config.xlsx"),
        &[(
            "函数映射",
            &[
                ("A1", "函数别名"),
                ("B1", "映射函数"),
                ("A2", "FUN_1"),
                ("B2", "SUM"),
            ],
        )],
    );
    let request = AggregationRequest {
        template_path: path("template.xlsx"),
        source_paths: vec![path("source.xlsx")],
        output_path: path("output.xlsx"),
        config_path: Some(path("config.xlsx")),
    };
    (directory, request)
}
fn number(path: &Path, cell: &str) -> f64 {
    umya_spreadsheet::reader::xlsx::read(path)
        .unwrap()
        .sheet_by_name("表")
        .unwrap()
        .value_number(cell)
        .unwrap()
}

#[test]
fn cancellation_before_read_and_before_publication() {
    for stage in [0, 1] {
        let (_directory, request) = fixture();
        let flag = AtomicBool::new(stage == 0);
        let progress = |done: usize, _: usize, _: &str| {
            if done == stage {
                flag.store(true, Ordering::Relaxed);
            }
        };
        let error = aggregate(
            &request,
            &RunControl {
                cancel: Some(&flag),
                progress: Some(&progress),
            },
        )
        .unwrap_err();
        assert!(error.cancelled);
        assert!(!request.output_path.exists());
        assert!(
            !request
                .output_path
                .parent()
                .unwrap()
                .read_dir()
                .unwrap()
                .any(|entry| entry
                    .unwrap()
                    .file_name()
                    .to_string_lossy()
                    .starts_with(".汇总"))
        );
    }
}
#[test]
fn configuration_snapshot_changes_only_next_batch() {
    let (_directory, mut request) = fixture();
    let second = request.template_path.with_file_name("second.xlsx");
    workbook(&second, &[("表", &[("K4", "18")])]);
    request.source_paths.push(second);
    let config = request.config_path.clone().unwrap();
    let progress = |done: usize, _: usize, _: &str| {
        if done == 0 {
            workbook(&config, &[("函数映射", &[("A1", "FUN_1"), ("B1", "AVG")])]);
        }
    };
    let first = aggregate(
        &request,
        &RunControl {
            progress: Some(&progress),
            cancel: None,
        },
    )
    .unwrap();
    assert_eq!(first.function_bindings["FUN_1"], "SUM");
    assert_eq!(number(&request.output_path, "K4"), 30.0);
    request.output_path.set_file_name("second-output.xlsx");
    let second = aggregate(&request, &RunControl::default()).unwrap();
    assert_eq!(second.function_bindings["FUN_1"], "AVG");
    assert_eq!(number(&request.output_path, "K4"), 15.0);
}
#[test]
fn appearing_output_is_never_overwritten() {
    let (_directory, request) = fixture();
    let progress = |done: usize, _: usize, _: &str| {
        if done == 1 {
            std::fs::write(&request.output_path, b"existing-user-content").unwrap();
        }
    };
    assert!(
        aggregate(
            &request,
            &RunControl {
                progress: Some(&progress),
                cancel: None
            }
        )
        .is_err()
    );
    assert_eq!(
        std::fs::read(&request.output_path).unwrap(),
        b"existing-user-content"
    );
}
#[test]
fn changed_template_or_source_prevents_publication() {
    for template in [false, true] {
        let (_directory, request) = fixture();
        let changed = if template {
            &request.template_path
        } else {
            &request.source_paths[0]
        };
        let progress = |done: usize, _: usize, _: &str| {
            if done == 1 {
                use std::io::Write;
                std::fs::OpenOptions::new()
                    .append(true)
                    .open(changed)
                    .unwrap()
                    .write_all(b"changed")
                    .unwrap();
            }
        };
        let error = aggregate(
            &request,
            &RunControl {
                progress: Some(&progress),
                cancel: None,
            },
        )
        .unwrap_err();
        assert!(error.message.contains("变化"));
        assert!(!request.output_path.exists());
    }
}
#[derive(Default)]
struct Custom {
    total: bigdecimal::BigDecimal,
}
impl Calculator for Custom {
    fn add(&mut self, values: &[Number]) -> Vec<Notice> {
        self.total += &values[0].value;
        Vec::new()
    }
    fn finish(&self, _: usize) -> CalculationResult {
        CalculationResult::new(&self.total * 2)
    }
}
#[test]
fn custom_algorithm_needs_only_registry_and_signature() {
    let (_directory, request) = fixture();
    workbook(
        &request.template_path,
        &[("表", &[("K4", "FUN_17()"), ("K5", "DOUBLE(K4)")])],
    );
    workbook(
        request.config_path.as_ref().unwrap(),
        &[("函数映射", &[("A1", "FUN_17"), ("B1", "DOUBLE")])],
    );
    let mut aggregator = Aggregator::default();
    aggregator
        .registry
        .register(FunctionSpec::new("DOUBLE", &["数值"], || {
            Box::new(Custom::default())
        }));
    let result = aggregator
        .aggregate(&request, &RunControl::default())
        .unwrap();
    assert_eq!(result.function_counts["DOUBLE"], 2);
    assert_eq!(number(&request.output_path, "K4"), 24.0);
    assert_eq!(number(&request.output_path, "K5"), 24.0);
}
struct FinalError;
impl Calculator for FinalError {
    fn add(&mut self, _: &[Number]) -> Vec<Notice> {
        Vec::new()
    }
    fn finish(&self, _: usize) -> CalculationResult {
        CalculationResult {
            notices: vec![Notice::new(
                Severity::Error,
                "custom_final_error",
                "最终计算失败",
                None,
            )],
            ..CalculationResult::default()
        }
    }
}
#[test]
fn final_error_notice_blocks_output() {
    let (_directory, request) = fixture();
    let mut aggregator = Aggregator::default();
    aggregator
        .registry
        .register(FunctionSpec::new("SUM", &["数值"], || {
            Box::new(FinalError)
        }));
    let error = aggregator
        .aggregate(&request, &RunControl::default())
        .unwrap_err();
    assert!(
        error
            .issues
            .iter()
            .any(|issue| issue.code == "custom_final_error")
    );
    assert!(!request.output_path.exists());
}
struct BadSource;
impl Calculator for BadSource {
    fn add(&mut self, _: &[Number]) -> Vec<Notice> {
        Vec::new()
    }
    fn finish(&self, _: usize) -> CalculationResult {
        CalculationResult {
            sources: vec![ResultSource {
                source_index: 999,
                parameter_index: 0,
            }],
            ..CalculationResult::default()
        }
    }
}
#[test]
fn invalid_custom_result_sources_are_errors() {
    let (_directory, request) = fixture();
    let mut aggregator = Aggregator::default();
    aggregator
        .registry
        .register(FunctionSpec::new("SUM", &["数值"], || {
            Box::new(BadSource)
        }));
    let error = aggregator
        .aggregate(&request, &RunControl::default())
        .unwrap_err();
    assert_eq!(error.issues[0].code, "invalid_result_source");
    assert!(!request.output_path.exists());
}
struct LongNotice;
impl Calculator for LongNotice {
    fn add(&mut self, _: &[Number]) -> Vec<Notice> {
        Vec::new()
    }
    fn finish(&self, _: usize) -> CalculationResult {
        CalculationResult {
            notices: vec![Notice::new(
                Severity::Warning,
                "long_notice",
                &"𠮷测试".repeat(20000),
                None,
            )],
            ..CalculationResult::default()
        }
    }
}
#[test]
fn long_utf16_log_fields_are_split_losslessly() {
    let (_directory, request) = fixture();
    let mut aggregator = Aggregator::default();
    aggregator
        .registry
        .register(FunctionSpec::new("SUM", &["数值"], || {
            Box::new(LongNotice)
        }));
    aggregator
        .aggregate(&request, &RunControl::default())
        .unwrap();
    let book = umya_spreadsheet::reader::xlsx::read(&request.output_path).unwrap();
    let log = book.sheet_by_name("运行日志").unwrap();
    let mut text = String::new();
    for row in 1..=log.highest_row() {
        let content = log.value((6, row));
        if content.starts_with('𠮷') || content.starts_with('测') || content.starts_with('试') {
            assert!(content.encode_utf16().count() <= 32767);
            text.push_str(&content);
        }
    }
    assert_eq!(text, "𠮷测试".repeat(20000));
}
#[test]
fn full_existing_log_fails_without_touching_history() {
    let (_directory, request) = fixture();
    let mut book = umya_spreadsheet::reader::xlsx::read(&request.template_path).unwrap();
    book.new_sheet("运行日志")
        .unwrap()
        .cell_mut("A1048576")
        .set_value_string("历史");
    umya_spreadsheet::writer::xlsx::write(&book, &request.template_path).unwrap();
    let before = std::fs::read(&request.template_path).unwrap();
    assert!(
        aggregate(&request, &RunControl::default())
            .unwrap_err()
            .message
            .contains("剩余行数")
    );
    assert_eq!(std::fs::read(&request.template_path).unwrap(), before);
    assert!(!request.output_path.exists());
}
#[test]
fn hardlink_deduplication_preserves_real_file_count() {
    let (_directory, mut request) = fixture();
    let alias = request.source_paths[0].with_file_name("hardlink.xlsx");
    std::fs::hard_link(&request.source_paths[0], &alias).unwrap();
    request.source_paths.push(alias);
    let result = aggregate(&request, &RunControl::default()).unwrap();
    assert_eq!(result.source_count, 1);
    assert_eq!(number(&request.output_path, "K4"), 12.0);
}
#[test]
fn incomplete_configuration_and_alias_chains_fail() {
    for cells in [
        &[("A1", "FUN_1")][..],
        &[("A1", "FUN_1"), ("B1", "FUN_2")][..],
        &[("A1", "SUM"), ("B1", "AVG")][..],
        &[
            ("A1", "FUN_1"),
            ("B1", "SUM"),
            ("A2", "fun_1"),
            ("B2", "AVG"),
        ][..],
    ] {
        let (_directory, request) = fixture();
        workbook(
            request.config_path.as_ref().unwrap(),
            &[("函数映射", cells)],
        );
        assert!(aggregate(&request, &RunControl::default()).is_err());
        assert!(!request.output_path.exists());
    }
}
#[test]
fn default_output_is_unique_and_keeps_input_extension() {
    let one = excel_summary::default_output_path(Path::new("模板.xlsx"), Path::new("dist"));
    assert_eq!(one.parent(), Some(PathBuf::from("dist").as_path()));
    assert!(
        one.file_name()
            .unwrap()
            .to_string_lossy()
            .starts_with("模板_汇总_")
    );
    assert_eq!(one.extension().unwrap(), "xlsx");
}

#[test]
fn distinct_unicode_paths_are_not_casefolded_into_duplicates() {
    let (_directory, mut request) = fixture();
    let first = request.template_path.with_file_name("Straße.xlsx");
    let second = request.template_path.with_file_name("Strasse.xlsx");
    workbook(&first, &[("表", &[("K4", "10")])]);
    workbook(&second, &[("表", &[("K4", "20")])]);
    request.source_paths = vec![first, second];
    let result = aggregate(&request, &RunControl::default()).unwrap();
    assert_eq!(result.source_count, 2);
    assert_eq!(number(&request.output_path, "K4"), 30.0);
}
