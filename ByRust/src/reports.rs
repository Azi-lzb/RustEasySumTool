use crate::{
    functions::FunctionRegistry,
    model::{AggregationResult, Result, Severity},
    package::failure,
    provenance::{split_log_text, utf16_length},
};
use std::path::{Path, PathBuf};
use umya_spreadsheet::Workbook;

pub(crate) fn append_run_log(
    workbook: &mut Workbook,
    template: &Path,
    sources: &[PathBuf],
    stamp: &str,
    result: &AggregationResult,
    registry: &FunctionRegistry,
) -> Result<u32> {
    let row = |level: &str, label: &str, file: String, content: String| {
        [stamp.into(), level.into(), label.into(), file, content]
    };
    let counts = result
        .function_counts
        .iter()
        .map(|(name, count)| format!("{name} {count} 个单元格"))
        .collect::<Vec<_>>()
        .join("；");
    let mut rows=vec![
        row("信息","程序版本",String::new(),format!("{} Rust",env!("CARGO_PKG_VERSION"))),
        row("信息","模板",template.file_name().unwrap_or_default().to_string_lossy().into(),template.display().to_string()),
        row("信息","输出文件",result.output_path.file_name().unwrap_or_default().to_string_lossy().into(),result.output_path.display().to_string()),
        row("信息","配置文件",String::new(),result.config_path.as_ref().map_or_else(||"默认配置不存在，可使用直接函数及兼容旧别名。".into(),|path|path.display().to_string())),
        row("信息","汇总统计",String::new(),format!("源文件 {} 个；参与源表 {} 张；输出模板页 {} 张；{counts}。",result.source_count,result.source_sheet_count,result.template_sheet_count)),
        row("信息","计算耗时",String::new(),format!("{:.3} 秒（不含结果保存）",result.duration_seconds)),
        row("信息","计算约定",String::new(),"SUM/AVG/WAVG/SUMPRODUCT/RATIO空白按0；AVG分母为有效源文件总数；空白利率的权重仍进入加权分母；WAVG区域按位置配对且总权重为0报错；SUMPRODUCT先逐文件相乘再求和；RATIO为分子合计除以分母合计，分母合计0报错；MAX/MIN忽略空白、真实0有效，全空0不代表真实极值；COUNT空白不计、真实0计数；输出仅为结果值。".into()),
    ];
    for (name, target) in &result.function_bindings {
        let signature = registry
            .get(target)
            .ok_or_else(|| failure("日志算法未注册"))?
            .signature(target);
        rows.push(row("信息","本批函数绑定",name.clone(),format!("{name} -> {target}；算法签名：{signature}。本批使用固定快照，配置修改从下一批生效。")));
    }
    if result.function_bindings.is_empty() {
        rows.push(row(
            "信息",
            "本批函数绑定",
            String::new(),
            "本批未使用别名；直接函数名称的固定意义不受配置改变。".into(),
        ));
    }
    for source in sources {
        rows.push(row(
            "信息",
            "源文件",
            source
                .file_name()
                .unwrap_or_default()
                .to_string_lossy()
                .into(),
            source.display().to_string(),
        ));
    }
    for issue in &result.issues {
        let mut positions = Vec::new();
        if !issue.template_sheet.is_empty() {
            positions.push(format!(
                "模板 {}!{}",
                issue.template_sheet, issue.template_cell
            ));
        }
        if !issue.source_sheet.is_empty() {
            positions.push(format!("源 {}!{}", issue.source_sheet, issue.source_cell));
        }
        let label = if positions.is_empty() {
            issue.code.clone()
        } else {
            positions.join("；")
        };
        rows.push(row(
            match issue.severity {
                Severity::Info => "信息",
                Severity::Warning => "提醒",
                Severity::Error => "错误",
            },
            &label,
            issue.source_file.clone(),
            issue.message.clone(),
        ));
    }
    if result.issues.is_empty() {
        rows.push(row(
            "信息",
            "检查结果",
            String::new(),
            "本次未发现需要提醒的情况。".into(),
        ));
    }
    let mut expanded = Vec::new();
    for row in rows {
        let parts = row.each_ref().map(|value| split_log_text(value));
        let count = parts.iter().map(Vec::len).max().unwrap_or(1);
        for index in 0..count {
            let mut line: [String; 5] =
                std::array::from_fn(|column| parts[column].get(index).cloned().unwrap_or_default());
            if index > 0 {
                line[0] = row[0].clone();
                line[1] = row[1].clone();
                if parts[2].len() == 1 {
                    let label = format!("{}（续 {}/{count}）", row[2], index + 1);
                    line[2] = if utf16_length(&label) <= 32767 {
                        label
                    } else {
                        row[2].clone()
                    };
                }
            }
            expanded.push(line);
        }
    }
    if workbook.sheet_by_name("运行日志").is_err() {
        workbook.new_sheet("运行日志").map_err(failure)?;
    }
    let sheet = workbook.sheet_by_name_mut("运行日志").map_err(failure)?;
    let has_content = !sheet.cells().is_empty() || !sheet.merge_cells().is_empty();
    let highest = sheet
        .merge_cells()
        .iter()
        .filter_map(|range| range.range().rsplit(':').next().map(str::to_owned))
        .filter_map(|coordinate| {
            coordinate
                .trim_start_matches(|character: char| {
                    character.is_ascii_alphabetic() || character == '$'
                })
                .replace('$', "")
                .parse::<u32>()
                .ok()
        })
        .max()
        .unwrap_or(0)
        .max(sheet.highest_row());
    let start = if has_content { highest + 3 } else { 2 };
    let header = start + 2;
    if u64::from(header) + expanded.len() as u64 > 1_048_576 {
        return Err(failure(
            "运行日志没有足够的剩余行数，请使用不含历史日志的模板。",
        ));
    }
    sheet.cell_mut((2, start)).set_value_string(if has_content {
        "本次运行日志"
    } else {
        "运行日志"
    });
    for (index, value) in ["时间", "级别", "项目／位置", "文件", "内容"]
        .iter()
        .enumerate()
    {
        sheet
            .cell_mut((index as u32 + 2, header))
            .set_value_string(*value);
    }
    for (index, line) in expanded.iter().enumerate() {
        let row = header + 1 + index as u32;
        for (column, value) in line.iter().enumerate() {
            sheet
                .cell_mut((column as u32 + 2, row))
                .set_value_string(value);
        }
        let height = line
            .iter()
            .map(|text| text.lines().count() + text.chars().count() / 55)
            .max()
            .unwrap_or(1) as f64
            * 16.0
            + 8.0;
        sheet
            .row_dimension_mut(row)
            .set_height(height.clamp(24.0, 409.5));
    }
    if !has_content {
        for (column, width) in [
            ("A", 3.0),
            ("B", 22.0),
            ("C", 10.0),
            ("D", 36.0),
            ("E", 36.0),
            ("F", 64.0),
        ] {
            sheet.column_dimension_mut(column).set_width(width);
        }
    }
    Ok(start)
}
