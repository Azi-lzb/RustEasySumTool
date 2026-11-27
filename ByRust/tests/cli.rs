#![cfg(feature = "cli")]

use std::{path::Path, process::Command};

fn workbook(path: &Path, instructions: bool, zero_total_weight: bool) {
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.sheet_mut(0).unwrap();
    sheet.set_name("业务明细");
    if instructions {
        sheet.cell_mut("D7").set_value_string("WAVG(A7,B7)");
        sheet.cell_mut("F7").set_value_string("WAVG(A7,C7)");
    } else {
        sheet
            .cell_mut("B7")
            .set_value_number(if zero_total_weight { -1 } else { 1 });
        sheet.cell_mut("C7").set_value_number(2);
    }
    umya_spreadsheet::writer::xlsx::write(&book, path).unwrap();
}

#[test]
fn successful_cli_keeps_warnings_in_run_log() {
    let directory = tempfile::tempdir().unwrap();
    let template = directory.path().join("template.xlsx");
    let source = directory.path().join("source.xlsx");
    let output = directory.path().join("result.xlsx");
    workbook(&template, true, false);
    workbook(&source, false, false);
    let completed = Command::new(env!("CARGO_BIN_EXE_excel-summary-rust"))
        .arg("--template")
        .arg(&template)
        .arg("--sources")
        .arg(&source)
        .arg("--output")
        .arg(&output)
        .output()
        .unwrap();
    assert!(
        completed.status.success(),
        "{}",
        String::from_utf8_lossy(&completed.stderr)
    );
    let terminal = format!(
        "{}{}",
        String::from_utf8_lossy(&completed.stdout),
        String::from_utf8_lossy(&completed.stderr)
    );
    assert!(!terminal.contains("利率空白"));
    assert!(!terminal.contains("权重仍进入分母"));
    assert!(!terminal.contains("提醒"));
    assert!(!terminal.contains("业务明细!D7"));
    assert!(!terminal.contains("业务明细!F7"));
    let result = umya_spreadsheet::reader::xlsx::read(&output).unwrap();
    let log = result.sheet_by_name("运行日志").unwrap();
    let warnings = (1..=log.highest_row())
        .filter(|row| log.value((3, *row)) == "提醒")
        .collect::<Vec<_>>();
    assert_eq!(warnings.len(), 2);
    for (row, coordinate) in warnings.iter().zip(["D7", "F7"]) {
        assert!(log.value((4, *row)).contains(coordinate));
        assert!(log.value((6, *row)).contains("利率空白"));
        assert!(log.value((6, *row)).contains("权重仍进入分母"));
    }
}

#[test]
fn failed_cli_prints_errors_without_warning_dump() {
    let directory = tempfile::tempdir().unwrap();
    let template = directory.path().join("template.xlsx");
    let first = directory.path().join("first.xlsx");
    let second = directory.path().join("second.xlsx");
    let output = directory.path().join("result.xlsx");
    workbook(&template, true, false);
    workbook(&first, false, false);
    workbook(&second, false, true);
    let completed = Command::new(env!("CARGO_BIN_EXE_excel-summary-rust"))
        .arg("--template")
        .arg(&template)
        .arg("--sources")
        .arg(&first)
        .arg(&second)
        .arg("--output")
        .arg(&output)
        .output()
        .unwrap();
    assert!(!completed.status.success());
    let terminal = format!(
        "{}{}",
        String::from_utf8_lossy(&completed.stdout),
        String::from_utf8_lossy(&completed.stderr)
    );
    assert!(terminal.contains("zero_weight"));
    assert!(!terminal.contains("利率空白"));
    assert!(!terminal.contains("负权重"));
    assert!(!terminal.contains("blank_rate_with_weight"));
    assert!(!output.exists());
}
