use crate::model::Severity;
use crate::numbers::{Number, context, divide};
use crate::rules::Reference;
use bigdecimal::{BigDecimal, Zero};

#[derive(Clone, Debug)]
pub struct Notice {
    pub severity: Severity,
    pub code: String,
    pub message: String,
    pub parameter_index: Option<usize>,
}

impl Notice {
    pub fn new(
        severity: Severity,
        code: &str,
        message: &str,
        parameter_index: Option<usize>,
    ) -> Self {
        Self {
            severity,
            code: code.into(),
            message: message.into(),
            parameter_index,
        }
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ResultSource {
    pub source_index: usize,
    pub parameter_index: usize,
}

#[derive(Clone, Debug, Default)]
pub struct CalculationResult {
    pub value: BigDecimal,
    pub notices: Vec<Notice>,
    pub sources: Vec<ResultSource>,
    pub source_note: String,
}

impl CalculationResult {
    pub fn new(value: BigDecimal) -> Self {
        Self {
            value,
            ..Self::default()
        }
    }
}

pub trait Calculator {
    fn add(&mut self, values: &[Number]) -> Vec<Notice>;
    fn finish(&self, source_count: usize) -> CalculationResult;
}

#[derive(Clone)]
pub struct FunctionSpec {
    pub name: String,
    pub parameter_names: Vec<String>,
    pub factory: fn() -> Box<dyn Calculator>,
    pub variadic: bool,
    pub allow_ranges: bool,
    pub matching_shapes: bool,
}

impl FunctionSpec {
    pub fn new(name: &str, parameter_names: &[&str], factory: fn() -> Box<dyn Calculator>) -> Self {
        Self {
            name: name.to_uppercase(),
            parameter_names: parameter_names.iter().map(|name| (*name).into()).collect(),
            factory,
            variadic: false,
            allow_ranges: false,
            matching_shapes: false,
        }
    }

    pub fn signature(&self, name: &str) -> String {
        format!(
            "{}({}{})",
            name,
            self.parameter_names.join(","),
            if self.variadic { ",..." } else { "" }
        )
    }

    pub fn accepts_count(&self, count: usize) -> bool {
        if self.variadic {
            count >= self.parameter_names.len()
        } else {
            count == self.parameter_names.len()
        }
    }

    pub fn expected_count(&self) -> String {
        if self.variadic {
            format!("至少 {}", self.parameter_names.len())
        } else {
            self.parameter_names.len().to_string()
        }
    }

    pub fn validate_references(&self, references: &[Reference]) -> std::result::Result<(), String> {
        if !self.allow_ranges
            && references
                .iter()
                .any(|reference| reference.shape() != (1, 1))
        {
            return Err(format!("{} 的参数只接受单个单元格引用。", self.name));
        }
        if self.matching_shapes
            && references
                .windows(2)
                .any(|pair| pair[0].shape() != pair[1].shape())
        {
            return Err(format!(
                "{} 的参数区域必须具有相同的行数和列数。",
                self.name
            ));
        }
        Ok(())
    }
}

#[derive(Clone)]
pub struct FunctionRegistry {
    specifications: Vec<FunctionSpec>,
}

impl FunctionRegistry {
    pub fn get(&self, name: &str) -> Option<&FunctionSpec> {
        self.specifications
            .iter()
            .find(|specification| specification.name.eq_ignore_ascii_case(name))
    }

    pub fn specifications(&self) -> &[FunctionSpec] {
        &self.specifications
    }

    pub fn register(&mut self, mut specification: FunctionSpec) {
        specification.name = specification.name.to_uppercase();
        self.specifications
            .retain(|existing| existing.name != specification.name);
        self.specifications.push(specification);
    }
}

impl Default for FunctionRegistry {
    fn default() -> Self {
        let mut sum = FunctionSpec::new("SUM", &["取数单元格或区域"], || {
            Box::new(Sum::default())
        });
        sum.variadic = true;
        sum.allow_ranges = true;
        let mut weighted = FunctionSpec::new("WAVG", &["数值组", "权重组"], || {
            Box::new(Weighted::default())
        });
        weighted.allow_ranges = true;
        weighted.matching_shapes = true;
        Self {
            specifications: vec![
                sum,
                FunctionSpec::new("AVG", &["取数单元格"], || Box::new(Average::default())),
                weighted,
                FunctionSpec::new("MAX", &["取数单元格"], || Box::new(Extreme::new(true))),
                FunctionSpec::new("MIN", &["取数单元格"], || {
                    Box::new(Extreme::new(false))
                }),
                FunctionSpec::new("COUNT", &["取数单元格"], || Box::new(Count::default())),
                FunctionSpec::new(
                    "SUMPRODUCT",
                    &["因子1单元格", "因子2单元格"],
                    || Box::new(SumProduct::default()),
                ),
                FunctionSpec::new("RATIO", &["分子单元格", "分母单元格"], || {
                    Box::new(Ratio::default())
                }),
            ],
        }
    }
}

#[derive(Default)]
struct Sum {
    total: BigDecimal,
}
impl Calculator for Sum {
    fn add(&mut self, values: &[Number]) -> Vec<Notice> {
        let subtotal = values.iter().fold(BigDecimal::zero(), |sum, number| {
            context().add_refs(&sum, &number.value)
        });
        self.total = context().add_refs(&self.total, &subtotal);
        Vec::new()
    }
    fn finish(&self, _: usize) -> CalculationResult {
        CalculationResult::new(self.total.clone())
    }
}

#[derive(Default)]
struct Average {
    sum: Sum,
}
impl Calculator for Average {
    fn add(&mut self, values: &[Number]) -> Vec<Notice> {
        self.sum.add(values)
    }
    fn finish(&self, source_count: usize) -> CalculationResult {
        quotient(
            &self.sum.total,
            &BigDecimal::from(source_count as u64),
            "no_sources",
            "AVG至少需要一份有效源文件。",
        )
    }
}

#[derive(Default)]
struct SumProduct {
    total: BigDecimal,
}
impl Calculator for SumProduct {
    fn add(&mut self, values: &[Number]) -> Vec<Notice> {
        let product = context().multiply(&values[0].value, &values[1].value);
        self.total = context().add_refs(&self.total, &product);
        Vec::new()
    }
    fn finish(&self, _: usize) -> CalculationResult {
        CalculationResult::new(self.total.clone())
    }
}

#[derive(Default)]
struct Ratio {
    numerator: BigDecimal,
    denominator: BigDecimal,
}
impl Calculator for Ratio {
    fn add(&mut self, values: &[Number]) -> Vec<Notice> {
        self.numerator = context().add_refs(&self.numerator, &values[0].value);
        self.denominator = context().add_refs(&self.denominator, &values[1].value);
        Vec::new()
    }
    fn finish(&self, _: usize) -> CalculationResult {
        quotient(
            &self.numerator,
            &self.denominator,
            "zero_denominator",
            "RATIO 分母合计为0，无法计算合计比率，请核对空白、零值或正负抵消。",
        )
    }
}

#[derive(Default)]
struct Weighted {
    numerator: BigDecimal,
    denominator: BigDecimal,
}
impl Calculator for Weighted {
    fn add(&mut self, values: &[Number]) -> Vec<Notice> {
        self.numerator = context().add_refs(
            &self.numerator,
            &context().multiply(&values[0].value, &values[1].value),
        );
        self.denominator = context().add_refs(&self.denominator, &values[1].value);
        let mut notices = Vec::new();
        if values[0].blank && !values[1].value.is_zero() {
            notices.push(Notice::new(
                Severity::Warning,
                "blank_rate_with_weight",
                "利率空白按0%，权重仍进入分母。",
                Some(0),
            ));
        }
        if values[1].value < BigDecimal::zero() {
            notices.push(Notice::new(
                Severity::Warning,
                "negative_weight",
                "负权重按原值参与计算，请核对业务口径。",
                Some(1),
            ));
        }
        notices
    }
    fn finish(&self, _: usize) -> CalculationResult {
        quotient(
            &self.numerator,
            &self.denominator,
            "zero_weight",
            "WAVG 权重合计为0，无法计算加权平均值，请核对空白、零权重或正负抵消。",
        )
    }
}

fn quotient(
    numerator: &BigDecimal,
    denominator: &BigDecimal,
    code: &str,
    message: &str,
) -> CalculationResult {
    if denominator.is_zero() {
        return CalculationResult {
            notices: vec![Notice::new(Severity::Error, code, message, None)],
            ..CalculationResult::default()
        };
    }
    CalculationResult::new(divide(numerator, denominator))
}

struct Extreme {
    maximum: bool,
    ordinal: usize,
    best: Option<BigDecimal>,
    sources: Vec<ResultSource>,
}
impl Extreme {
    fn new(maximum: bool) -> Self {
        Self {
            maximum,
            ordinal: 0,
            best: None,
            sources: Vec::new(),
        }
    }
}
impl Calculator for Extreme {
    fn add(&mut self, values: &[Number]) -> Vec<Notice> {
        let index = self.ordinal;
        self.ordinal += 1;
        if values[0].blank {
            return Vec::new();
        }
        let value = &values[0].value;
        let better = self.best.as_ref().is_none_or(|best| {
            if self.maximum {
                value > best
            } else {
                value < best
            }
        });
        if better {
            self.best = Some(value.clone());
            self.sources = vec![ResultSource {
                source_index: index,
                parameter_index: 0,
            }];
        } else if self.best.as_ref() == Some(value) {
            self.sources.push(ResultSource {
                source_index: index,
                parameter_index: 0,
            });
        }
        Vec::new()
    }
    fn finish(&self, _: usize) -> CalculationResult {
        if let Some(value) = &self.best {
            return CalculationResult {
                value: value.clone(),
                sources: self.sources.clone(),
                ..CalculationResult::default()
            };
        }
        let note = "无有效数值，无极值来源，按约定输出0";
        CalculationResult {
            notices: vec![Notice::new(Severity::Info, "no_valid_values", note, None)],
            source_note: note.into(),
            ..CalculationResult::default()
        }
    }
}

#[derive(Default)]
struct Count {
    count: u64,
}
impl Calculator for Count {
    fn add(&mut self, values: &[Number]) -> Vec<Notice> {
        if !values[0].blank {
            self.count += 1;
        }
        Vec::new()
    }
    fn finish(&self, _: usize) -> CalculationResult {
        CalculationResult::new(BigDecimal::from(self.count))
    }
}
