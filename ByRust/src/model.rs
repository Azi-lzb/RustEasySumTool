use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};

#[derive(Clone, Copy, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum Severity {
    #[default]
    Info,
    Warning,
    Error,
}

#[derive(Clone, Debug, Default, Deserialize, PartialEq, Serialize)]
pub struct Issue {
    pub severity: Severity,
    pub code: String,
    pub message: String,
    pub source_file: String,
    pub source_sheet: String,
    pub source_cell: String,
    pub template_sheet: String,
    pub template_cell: String,
}

impl Issue {
    pub fn new(severity: Severity, code: &str, message: impl Into<String>) -> Self {
        Self {
            severity,
            code: code.into(),
            message: message.into(),
            ..Self::default()
        }
    }

    pub fn at_template(mut self, sheet: &str, cell: &str) -> Self {
        self.template_sheet = sheet.into();
        self.template_cell = cell.into();
        self
    }
}

#[derive(Clone, Debug, Deserialize, Serialize)]
pub struct AggregationError {
    pub message: String,
    pub issues: Vec<Issue>,
    pub cancelled: bool,
}

impl AggregationError {
    pub fn new(message: impl Into<String>) -> Self {
        Self {
            message: message.into(),
            issues: Vec::new(),
            cancelled: false,
        }
    }

    pub fn with_issues(message: impl Into<String>, issues: Vec<Issue>) -> Self {
        Self {
            message: message.into(),
            issues,
            cancelled: false,
        }
    }
}

impl std::fmt::Display for AggregationError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(&self.message)
    }
}

impl std::error::Error for AggregationError {}

pub type Result<T> = std::result::Result<T, AggregationError>;

#[derive(Clone, Debug, Deserialize, Serialize)]
pub struct AggregationRequest {
    pub template_path: PathBuf,
    pub source_paths: Vec<PathBuf>,
    pub output_path: PathBuf,
    pub config_path: Option<PathBuf>,
}

#[derive(Clone, Debug, Deserialize, Serialize)]
pub struct AggregationResult {
    pub output_path: PathBuf,
    pub source_count: usize,
    pub source_sheet_count: usize,
    pub template_sheet_count: usize,
    pub rule_count: usize,
    pub sum_rule_count: usize,
    pub avg_rule_count: usize,
    pub wavg_rule_count: usize,
    pub function_counts: BTreeMap<String, usize>,
    pub function_bindings: BTreeMap<String, String>,
    pub config_path: Option<PathBuf>,
    pub issues: Vec<Issue>,
    pub duration_seconds: f64,
}

impl AggregationResult {
    pub fn warning_count(&self) -> usize {
        self.issues
            .iter()
            .filter(|issue| issue.severity == Severity::Warning)
            .count()
    }
}

pub type ProgressCallback<'a> = &'a dyn Fn(usize, usize, &str);

#[derive(Default)]
pub struct RunControl<'a> {
    pub cancel: Option<&'a AtomicBool>,
    pub progress: Option<ProgressCallback<'a>>,
}

impl RunControl<'_> {
    pub(crate) fn check(&self) -> Result<()> {
        if self.cancel.is_some_and(|flag| flag.load(Ordering::Relaxed)) {
            return Err(AggregationError {
                message: "已取消汇总，没有发布结果文件。".into(),
                issues: Vec::new(),
                cancelled: true,
            });
        }
        Ok(())
    }

    pub(crate) fn report(&self, done: usize, total: usize, message: &str) {
        if let Some(progress) = self.progress {
            progress(done, total, message);
        }
    }
}
