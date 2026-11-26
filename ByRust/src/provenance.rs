use crate::{
    functions::CalculationResult,
    model::{Issue, Result, Severity},
    package::failure,
    rules::TemplateRule,
};
use std::path::PathBuf;
use umya_spreadsheet::{Comment, CommentText, Worksheet};

pub fn utf16_length(text: &str) -> usize {
    text.encode_utf16().count()
}
pub fn split_log_text(text: &str) -> Vec<String> {
    let mut result = Vec::new();
    let mut part = String::new();
    let mut length = 0;
    for character in text.chars() {
        let width = character.len_utf16();
        if length + width > 32767 {
            result.push(std::mem::take(&mut part));
            length = 0;
        }
        part.push(character);
        length += width;
    }
    if !part.is_empty() || result.is_empty() {
        result.push(part);
    }
    result
}
fn abbreviate(text: &str, limit: usize) -> String {
    if utf16_length(text) <= limit {
        return text.into();
    }
    let suffix = "…（完整内容见运行日志）";
    let mut left = limit.saturating_sub(utf16_length(suffix));
    let prefix = text
        .chars()
        .take_while(|character| {
            if character.len_utf16() > left {
                false
            } else {
                left -= character.len_utf16();
                true
            }
        })
        .collect::<String>();
    format!("{prefix}{suffix}")
}
fn comment_text(comment: &Comment) -> String {
    comment
        .text()
        .rich_text()
        .map(|value| value.text().into_owned())
        .or_else(|| comment.text().text().map(|value| value.value().into()))
        .unwrap_or_default()
}
pub(crate) fn append(
    sheet: &mut Worksheet,
    rule: &TemplateRule,
    calculated: &CalculationResult,
    paths: &[PathBuf],
) -> Result<(Vec<Issue>, bool)> {
    if calculated.sources.is_empty() && calculated.source_note.is_empty() {
        return Ok((Vec::new(), false));
    }
    let algorithm = if rule.function_name != rule.operation {
        format!("{}（原函数/别名：{}）", rule.operation, rule.function_name)
    } else {
        rule.operation.clone()
    };
    let total = calculated.sources.len();
    let mut issues = Vec::new();
    let mut preview = Vec::new();
    for (ordinal, source) in calculated.sources.iter().enumerate() {
        let invalid = || {
            crate::model::AggregationError::with_issues(
                "计算函数返回了无效的来源索引。",
                vec![
                    Issue::new(
                        Severity::Error,
                        "invalid_result_source",
                        "来源文件索引或参数索引超出范围。",
                    )
                    .at_template(&rule.sheet, &rule.coordinate),
                ],
            )
        };
        let path = paths.get(source.source_index).ok_or_else(invalid)?;
        let reference = rule
            .references
            .get(source.parameter_index)
            .ok_or_else(invalid)?;
        let mut issue = Issue::new(
            Severity::Info,
            "result_source",
            format!(
                "{algorithm} 结果来源 {}/{total}（源文件序号 {}，参数序号 {}）。",
                ordinal + 1,
                source.source_index + 1,
                source.parameter_index + 1
            ),
        )
        .at_template(&rule.sheet, &rule.coordinate);
        issue.source_file = path.display().to_string();
        issue.source_sheet = reference.sheet.clone();
        issue.source_cell = reference.coordinate.clone();
        issues.push(issue);
        if ordinal < 10 {
            preview.push((source.source_index, path, reference));
        }
    }
    if !calculated.source_note.is_empty()
        && !calculated
            .notices
            .iter()
            .any(|notice| notice.message == calculated.source_note)
    {
        issues.push(
            Issue::new(
                Severity::Info,
                "result_source_note",
                &calculated.source_note,
            )
            .at_template(&rule.sheet, &rule.coordinate),
        );
    }
    let summary = if total > 1 {
        format!("并列共{total}个，完整来源见运行日志")
    } else if total == 1 {
        "来源共1个，完整来源见运行日志".into()
    } else {
        "无有效数值，完整说明见运行日志".into()
    };
    let header = "—— 汇总结果来源 ——";
    let mut detailed = vec![header.into(), format!("算法：{algorithm}")];
    let mut compact = vec![
        header.into(),
        format!("算法：{}", abbreviate(&algorithm, 160)),
    ];
    for (ordinal, (index, path, reference)) in preview.iter().enumerate() {
        let name = path.file_name().unwrap_or_default().to_string_lossy();
        detailed.extend([
            format!(
                "来源 {}/{total}（源文件序号 {}）：{name}",
                ordinal + 1,
                index + 1
            ),
            format!("表格地址：{}", reference.label()),
            format!("完整路径：{}", path.display()),
        ]);
        compact.extend([
            format!(
                "来源 {}/{total}（源文件序号 {}）：{}",
                ordinal + 1,
                index + 1,
                abbreviate(&name, 160)
            ),
            format!("表格地址：{}", abbreviate(&reference.label(), 160)),
            "完整路径：批注空间有限，见运行日志".into(),
        ]);
    }
    if total == 0 {
        detailed.push("无有效数值，未记录来源文件。".into());
    }
    if !calculated.source_note.is_empty() {
        detailed.push(format!("说明：{}", calculated.source_note));
        compact.push(format!(
            "说明：{}",
            abbreviate(&calculated.source_note, 512)
        ));
    }
    if total > preview.len() {
        detailed.push(summary.clone());
    }
    compact.push(summary.clone());
    let candidates = [
        detailed.join("\n"),
        compact.join("\n"),
        format!("{header}\n算法：{}\n{summary}", abbreviate(&algorithm, 160)),
        format!("{header}\n{summary}"),
        format!(
            "——\n{}",
            if total > 0 {
                "完整来源见运行日志"
            } else {
                "无有效数值，完整说明见运行日志"
            }
        ),
    ];
    let existing = sheet
        .comments()
        .iter()
        .find(|comment| comment.coordinate().get_coordinate() == rule.coordinate)
        .cloned();
    let old = existing.as_ref().map(comment_text).unwrap_or_default();
    let prefix = if old.is_empty() {
        old
    } else {
        format!("{old}\n\n")
    };
    let Some(section) = candidates
        .iter()
        .find(|section| utf16_length(&prefix) + utf16_length(section) <= 32767)
    else {
        issues.push(
            Issue::new(
                Severity::Warning,
                "source_comment_full",
                "原批注剩余空间不足，未改动原批注；完整来源见运行日志。",
            )
            .at_template(&rule.sheet, &rule.coordinate),
        );
        return Ok((issues, false));
    };
    let text = prefix + section;
    let height = (text.lines().count() as f64 * 18.0 + 24.0).clamp(120.0, 900.0);
    let mut comment = existing.unwrap_or_else(|| {
        let mut comment = Comment::default();
        comment
            .new_comment(rule.coordinate.as_str())
            .set_author("汇总工具");
        comment
    });
    let mut body = CommentText::default();
    body.set_text_string(text);
    comment.set_text(body);
    let style = comment.shape().style();
    let mut widths = (640.0f64, height);
    for component in style.split(';') {
        if let Some((name, value)) = component.split_once(':') {
            let value = value.trim();
            let pixels = if let Some(value) = value.strip_suffix("pt") {
                value.parse::<f64>().unwrap_or(0.0) / 0.75
            } else {
                value.trim_end_matches("px").parse::<f64>().unwrap_or(0.0)
            };
            if name.trim() == "width" {
                widths.0 = widths.0.max(pixels);
            }
            if name.trim() == "height" {
                widths.1 = widths.1.max(pixels);
            }
        }
    }
    let retained = style
        .split(';')
        .filter(|component| {
            !component.trim_start().starts_with("width:")
                && !component.trim_start().starts_with("height:")
        })
        .collect::<Vec<_>>()
        .join(";");
    comment.shape_mut().set_style(format!(
        "{retained};width:{}pt;height:{}pt",
        widths.0 * 0.75,
        widths.1 * 0.75
    ));
    let position = sheet
        .comments()
        .iter()
        .position(|comment| comment.coordinate().get_coordinate() == rule.coordinate);
    if let Some(position) = position {
        *sheet
            .comments_mut()
            .get_mut(position)
            .ok_or_else(|| failure("批注位置失效"))? = comment;
    } else {
        sheet.add_comments(comment);
    }
    Ok((issues, true))
}
