use crate::{
    configuration::load_configuration,
    functions::{Calculator, FunctionRegistry, Notice},
    model::*,
    numbers::excel_number,
    package::*,
    preserve, provenance,
    reader::{SourceWorkbook, fold},
    reports,
    rules::{Reference, TemplateRule, is_alias, parse_rule},
};
use chrono::Local;
use std::{
    collections::{BTreeMap, BTreeSet},
    path::{Path, PathBuf},
    time::{Instant, SystemTime},
};

#[derive(Clone, Default)]
pub struct Aggregator {
    pub registry: FunctionRegistry,
}
pub fn default_output_path(template: &Path, directory: &Path) -> PathBuf {
    directory.join(format!(
        "{}_汇总_{}.xlsx",
        template.file_stem().unwrap_or_default().to_string_lossy(),
        Local::now().format("%Y%m%d_%H%M%S_%6f")
    ))
}
fn absolute(path: &Path) -> Result<PathBuf> {
    let expanded = if let Ok(home) = std::env::var("USERPROFILE") {
        if path == Path::new("~") {
            PathBuf::from(home)
        } else if path.to_string_lossy().starts_with("~/")
            || path.to_string_lossy().starts_with("~\\")
        {
            PathBuf::from(home).join(&path.to_string_lossy()[2..])
        } else {
            path.into()
        }
    } else {
        path.into()
    };
    if expanded.exists() {
        dunce::canonicalize(expanded).map_err(failure)
    } else {
        std::path::absolute(expanded).map_err(failure)
    }
}
fn same(left: &Path, right: &Path) -> bool {
    path_identity(left) == path_identity(right)
        || same_file::is_same_file(left, right).unwrap_or(false)
}
fn path_identity(path: &Path) -> String {
    if cfg!(windows) {
        path.to_string_lossy().to_lowercase()
    } else {
        path.to_string_lossy().into_owned()
    }
}
fn xlsx(path: &Path) -> bool {
    path.extension()
        .is_some_and(|extension| extension.eq_ignore_ascii_case("xlsx"))
}
fn fingerprint(path: &Path) -> Result<(u64, SystemTime)> {
    let metadata = std::fs::metadata(path).map_err(failure)?;
    Ok((metadata.len(), metadata.modified().map_err(failure)?))
}
pub(crate) fn template_safety(package: &Package) -> Result<()> {
    for (name, bytes) in package {
        let lower = name.to_lowercase();
        let drawing = lower.starts_with("xl/drawings/") && lower.ends_with(".xml");
        let empty = drawing
            && std::str::from_utf8(bytes)
                .ok()
                .and_then(|text| document(text).ok())
                .is_some_and(|doc| {
                    doc.root_element().has_tag_name((
                        "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                        "wsDr",
                    )) && !doc.root_element().children().any(|node| node.is_element())
                });
        if [
            "xl/media/",
            "xl/charts/",
            "xl/activex/",
            "xl/ctrlprops/",
            "_xmlsignatures/",
        ]
        .iter()
        .any(|prefix| lower.starts_with(prefix))
            || lower == "xl/vbaproject.bin"
            || (drawing && !empty)
        {
            return Err(failure(
                "模板包含图片、图表、绘图、控件、宏或签名。当前版本仅支持单元格表样，为避免另存时丢失这些对象，本次未处理。请使用不含这些对象的模板。",
            ));
        }
    }
    Ok(())
}
fn read_rules(
    workbook: &umya_spreadsheet::Workbook,
    bindings: &BTreeMap<String, String>,
    registry: &FunctionRegistry,
) -> Result<Vec<TemplateRule>> {
    let mut rules = Vec::new();
    let mut issues = Vec::new();
    for sheet in workbook.sheet_collection() {
        for cell in sheet.cells_sorted() {
            let coordinate = cell.coordinate().get_coordinate();
            let value = cell.value().into_owned();
            if cell.cell_value().is_formula() {
                issues.push(
                    Issue::new(
                        Severity::Error,
                        "template_formula_not_supported",
                        format!(
                            "模板含带等号的普通Excel公式“={}”。请使用不带等号的工具指令。",
                            cell.cell_value().formula()
                        ),
                    )
                    .at_template(sheet.name(), &coordinate),
                );
                continue;
            }
            if cell.cell_value().data_type() == "e" {
                issues.push(
                    Issue::new(
                        Severity::Error,
                        "template_excel_error",
                        format!("模板含有Excel错误值：{value}"),
                    )
                    .at_template(sheet.name(), &coordinate),
                );
                continue;
            }
            if sheet.name() == "运行日志" {
                continue;
            }
            match parse_rule(&value, sheet.name(), &coordinate, bindings, registry) {
                Ok(Some(rule)) => rules.push(rule),
                Ok(None) => (),
                Err(message) => issues.push(
                    Issue::new(Severity::Error, "invalid_rule", message)
                        .at_template(sheet.name(), &coordinate),
                ),
            }
        }
    }
    if !issues.is_empty() {
        return Err(AggregationError::with_issues(
            "模板规则校验失败，请修正指定位置后重试。",
            issues,
        ));
    }
    if rules.is_empty() {
        return Err(failure(
            "模板中没有汇总指令，请使用不带等号的直接算法名称或FUN_编号别名。",
        ));
    }
    Ok(rules)
}
fn notice_issue(
    notice: Notice,
    rule: &TemplateRule,
    source: Option<&Path>,
    references: &[Reference],
) -> Result<Issue> {
    let reference = notice
        .parameter_index
        .map(|index| {
            references.get(index).ok_or_else(|| {
                failure(format!(
                    "计算函数{}返回了无效的提醒参数位置。",
                    rule.operation
                ))
            })
        })
        .transpose()?;
    let mut issue = Issue::new(notice.severity, &notice.code, notice.message)
        .at_template(&rule.sheet, &rule.coordinate);
    if let Some(source) = source {
        issue.source_file = source.display().to_string();
        if let Some(reference) = reference {
            issue.source_sheet = reference.sheet.clone();
            issue.source_cell = reference.coordinate.clone();
        }
    }
    Ok(issue)
}
fn error_count(issues: &[Issue]) -> usize {
    issues
        .iter()
        .filter(|issue| issue.severity == Severity::Error)
        .count()
}
impl Aggregator {
    pub fn aggregate(
        &self,
        request: &AggregationRequest,
        control: &RunControl<'_>,
    ) -> Result<AggregationResult> {
        let started = Instant::now();
        let stamp = Local::now().format("%Y-%m-%d %H:%M:%S").to_string();
        let template = absolute(&request.template_path)?;
        let output = absolute(&request.output_path)?;
        let mut issues = Vec::new();
        if !xlsx(&template) || !template.is_file() {
            return Err(failure("请选择实际存在的.xlsx模板文件。"));
        }
        if !xlsx(&output) {
            return Err(failure("结果文件扩展名必须为.xlsx。"));
        }
        if same(&template, &output) {
            return Err(failure("输出路径不能与模板相同，模板不会被覆盖。"));
        }
        if output.exists() {
            return Err(failure("输出文件已经存在，请换一个文件名。"));
        }
        if output
            .parent()
            .is_some_and(|parent| parent.exists() && !parent.is_dir())
        {
            return Err(failure("输出目录不是一个文件夹。"));
        }
        let mut sources: Vec<PathBuf> = Vec::new();
        let mut seen = BTreeSet::new();
        for candidate in &request.source_paths {
            let path = absolute(candidate)?;
            let identity = path_identity(&path);
            if !seen.insert(identity) || sources.iter().any(|existing| same(existing, &path)) {
                let mut issue = Issue::new(
                    Severity::Info,
                    "duplicate_source",
                    "重复选择的同一源文件已去重，只计算一次。",
                );
                issue.source_file = path.display().to_string();
                issues.push(issue);
                continue;
            }
            if same(&path, &template) {
                return Err(failure("模板不能同时作为待汇总源文件。"));
            }
            if same(&path, &output) {
                return Err(failure("输出路径不能与源文件相同。"));
            }
            if path
                .file_name()
                .is_some_and(|name| name.to_string_lossy().starts_with("~$"))
            {
                let mut issue = Issue::new(Severity::Info, "lock_file", "Excel临时锁文件已忽略。");
                issue.source_file = path.display().to_string();
                issues.push(issue);
                continue;
            }
            if !xlsx(&path) || !path.is_file() {
                return Err(failure(format!(
                    "源文件不存在或不是.xlsx：{}",
                    path.display()
                )));
            }
            sources.push(path);
        }
        if sources.is_empty() {
            return Err(failure("请至少选择一份有效的.xlsx源工作簿。"));
        }
        control.check()?;
        let config = load_configuration(request.config_path.as_deref(), &self.registry)?;
        let template_state = fingerprint(&template)?;
        let original = read_package(&template)?;
        template_safety(&original)?;
        let mut workbook = umya_spreadsheet::reader::xlsx::read(&template)
            .map_err(|error| failure(format!("模板打开失败：{error}")))?;
        let rules = read_rules(&workbook, &config.bindings, &self.registry)?;
        let template_sheet_count = workbook
            .sheet_collection()
            .iter()
            .filter(|sheet| sheet.name() != "运行日志")
            .count();
        let definitions = rules
            .iter()
            .map(|rule| self.registry.get(&rule.operation).unwrap())
            .collect::<Vec<_>>();
        let mut calculators: Vec<Box<dyn Calculator>> =
            definitions.iter().map(|spec| (spec.factory)()).collect();
        let mut source_states = BTreeMap::new();
        let mut source_sheet_count = 0;
        control.report(
            0,
            sources.len() + 1,
            &format!(
                "已识别{}个汇总指令，开始读取{}份源文件。",
                rules.len(),
                sources.len()
            ),
        );
        for (index, source) in sources.iter().enumerate() {
            control.check()?;
            let mut seen_errors = BTreeSet::new();
            let read = (|| -> Result<()> {
                source_states.insert(source.clone(), fingerprint(source)?);
                let mut reader = SourceWorkbook::open(source)?;
                for ((rule, specification), calculator) in
                    rules.iter().zip(&definitions).zip(&mut calculators)
                {
                    control.check()?;
                    let mut consume = |references: Vec<Reference>| -> Result<bool> {
                        control.check()?;
                        let values = references
                            .iter()
                            .map(|reference| reader.get(reference))
                            .collect::<std::result::Result<Vec<_>, _>>();
                        match values {
                            Ok(values) => {
                                for notice in calculator.add(&values) {
                                    issues.push(notice_issue(
                                        notice,
                                        rule,
                                        Some(source),
                                        &references,
                                    )?);
                                }
                                Ok(true)
                            }
                            Err(error) => {
                                if seen_errors.insert((
                                    error.code.clone(),
                                    fold(&error.sheet),
                                    error.cell.clone(),
                                )) {
                                    let mut issue =
                                        Issue::new(Severity::Error, &error.code, error.message)
                                            .at_template(&rule.sheet, &rule.coordinate);
                                    issue.source_file = source.display().to_string();
                                    issue.source_sheet = error.sheet;
                                    issue.source_cell = error.cell;
                                    issues.push(issue);
                                }
                                Ok(false)
                            }
                        }
                    };
                    if specification.variadic {
                        'parameters: for reference in &rule.references {
                            for cell in reference.cells() {
                                if !consume(vec![cell])? {
                                    break 'parameters;
                                }
                            }
                        }
                    } else {
                        let mut iterators = rule
                            .references
                            .iter()
                            .map(Reference::cells)
                            .collect::<Vec<_>>();
                        loop {
                            let row = iterators
                                .iter_mut()
                                .map(Iterator::next)
                                .collect::<Option<Vec<_>>>();
                            let Some(row) = row else {
                                break;
                            };
                            if !consume(row)? {
                                break;
                            }
                        }
                    }
                }
                source_sheet_count += reader.read_sheet_count();
                if fingerprint(source)? != source_states[source] {
                    let mut issue = Issue::new(
                        Severity::Error,
                        "source_changed",
                        "读取期间源文件发生变化，请关闭编辑后重试。",
                    );
                    issue.source_file = source.display().to_string();
                    issues.push(issue);
                }
                Ok(())
            })();
            if let Err(error) = read {
                if error.cancelled {
                    return Err(error);
                }
                let mut issue = Issue::new(
                    Severity::Error,
                    "open_failed",
                    format!("源文件处理失败：{error}"),
                );
                issue.source_file = source.display().to_string();
                issues.push(issue);
            }
            control.report(
                index + 1,
                sources.len() + 1,
                &format!(
                    "已检查{}（{}/{}）。",
                    source.file_name().unwrap_or_default().to_string_lossy(),
                    index + 1,
                    sources.len()
                ),
            );
        }
        if error_count(&issues) > 0 {
            return Err(AggregationError::with_issues(
                format!("发现{}项错误，未生成汇总结果。", error_count(&issues)),
                issues,
            ));
        }
        let mut values = BTreeMap::new();
        let mut comment_sheets = BTreeSet::new();
        for (rule, calculator) in rules.iter().zip(calculators) {
            control.check()?;
            let calculated = calculator.finish(sources.len());
            for notice in calculated.notices.clone() {
                issues.push(notice_issue(notice, rule, None, &rule.references)?);
            }
            let value = excel_number(&calculated.value).map_err(|message| {
                AggregationError::with_issues(
                    "计算结果超出Excel数值范围。",
                    vec![
                        Issue::new(Severity::Error, "numeric_range", message)
                            .at_template(&rule.sheet, &rule.coordinate),
                    ],
                )
            })?;
            let sheet = workbook.sheet_by_name_mut(&rule.sheet).map_err(failure)?;
            sheet
                .cell_mut(rule.coordinate.as_str())
                .set_value_number(value);
            values.insert((rule.sheet.clone(), rule.coordinate.clone()), value);
            let (notices, changed) = provenance::append(sheet, rule, &calculated, &sources)?;
            issues.extend(notices);
            if changed {
                comment_sheets.insert(rule.sheet.clone());
            }
        }
        if error_count(&issues) > 0 {
            return Err(AggregationError::with_issues(
                format!(
                    "最终计算发现{}项错误，未生成汇总结果。",
                    error_count(&issues)
                ),
                issues,
            ));
        }
        control.check()?;
        let mut counts = BTreeMap::new();
        let mut bindings = BTreeMap::new();
        for rule in &rules {
            *counts.entry(rule.operation.clone()).or_insert(0) += 1;
            if is_alias(&rule.function_name) {
                bindings.insert(rule.function_name.clone(), rule.operation.clone());
            }
        }
        let mut result = AggregationResult {
            output_path: output.clone(),
            source_count: sources.len(),
            source_sheet_count,
            template_sheet_count,
            rule_count: rules.len(),
            sum_rule_count: *counts.get("SUM").unwrap_or(&0),
            avg_rule_count: *counts.get("AVG").unwrap_or(&0),
            wavg_rule_count: *counts.get("WAVG").unwrap_or(&0),
            function_counts: counts,
            function_bindings: bindings,
            config_path: config.path,
            issues,
            duration_seconds: started.elapsed().as_secs_f64(),
        };
        let log_start = reports::append_run_log(
            &mut workbook,
            &template,
            &sources,
            &stamp,
            &result,
            &self.registry,
        )?;
        control.report(
            sources.len(),
            sources.len() + 1,
            "计算完成，正在生成独立结果文件……",
        );
        let parent = output.parent().ok_or_else(|| failure("无效输出路径"))?;
        std::fs::create_dir_all(parent).map_err(failure)?;
        let intermediate = tempfile::Builder::new()
            .prefix(".汇总umya_")
            .suffix(".xlsx")
            .tempfile_in(parent)
            .map_err(failure)?;
        umya_spreadsheet::writer::xlsx::write(&workbook, intermediate.path())
            .map_err(|error| failure(format!("结果保存失败：{error}")))?;
        let generated = read_package(intermediate.path())?;
        let merged = preserve::merge(original, generated, &values, &comment_sheets, log_start)?;
        let temporary = tempfile::Builder::new()
            .prefix(".汇总临时_")
            .suffix(".xlsx")
            .tempfile_in(parent)
            .map_err(failure)?;
        save_package(&merged, temporary.path())?;
        read_package(temporary.path())?;
        control.check()?;
        if fingerprint(&template)? != template_state {
            return Err(failure("汇总期间模板发生变化，请保存好模板后重新运行。"));
        }
        for (source, state) in source_states {
            if fingerprint(&source)? != state {
                return Err(failure(format!(
                    "汇总期间源文件发生变化，请重新运行：{}",
                    source.display()
                )));
            }
        }
        temporary
            .persist_noclobber(&output)
            .map_err(|error| failure(format!("结果发布失败，不覆盖现有文件：{}", error.error)))?;
        result.duration_seconds = started.elapsed().as_secs_f64();
        control.report(
            sources.len() + 1,
            sources.len() + 1,
            &format!("汇总完成：{}", output.display()),
        );
        Ok(result)
    }
}

/// Aggregate with the built-in eight algorithms. Configuration is snapshotted per call.
pub fn aggregate(
    request: &AggregationRequest,
    control: &RunControl<'_>,
) -> Result<AggregationResult> {
    Aggregator::default().aggregate(request, control)
}
