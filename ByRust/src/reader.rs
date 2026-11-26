use crate::{
    model::{AggregationError, Result},
    numbers::{Number, parse_text},
    package::*,
    rules::Reference,
};
use calamine::{Data, Range, Reader, Xlsx};
use std::{collections::BTreeMap, fs::File, io::BufReader, path::Path};
use unicode_casefold::UnicodeCaseFold;

pub(crate) fn fold(text: &str) -> String {
    text.case_fold().collect()
}
#[derive(Clone, Default)]
pub(crate) struct CellInfo {
    pub formula: bool,
    pub cached: bool,
    pub error: Option<String>,
}
pub(crate) fn metadata(package: &Package, part: &str) -> Result<BTreeMap<String, CellInfo>> {
    let text = xml(package, part)?;
    let doc = document(&text)?;
    Ok(doc
        .descendants()
        .filter(|node| node.has_tag_name("c"))
        .filter_map(|node| {
            let coordinate = node.attribute("r")?.to_uppercase();
            let value = children(node, "v").next();
            Some((
                coordinate,
                CellInfo {
                    formula: children(node, "f").next().is_some(),
                    cached: value.is_some_and(|value| {
                        value.text().is_some() || node.attribute("t") == Some("str")
                    }),
                    error: (node.attribute("t") == Some("e")).then(|| {
                        value
                            .and_then(|value| value.text())
                            .unwrap_or("未知Excel错误")
                            .into()
                    }),
                },
            ))
        })
        .collect())
}
#[derive(Debug)]
pub struct SourceReadError {
    pub code: String,
    pub message: String,
    pub sheet: String,
    pub cell: String,
}
impl SourceReadError {
    fn at(code: &str, message: impl Into<String>, reference: &Reference) -> Self {
        Self {
            code: code.into(),
            message: message.into(),
            sheet: reference.sheet.clone(),
            cell: reference.coordinate.clone(),
        }
    }
}
pub struct SourceWorkbook {
    workbook: Xlsx<BufReader<File>>,
    package: Package,
    sheets: BTreeMap<String, Sheet>,
    values: BTreeMap<String, Range<Data>>,
    metadata: BTreeMap<String, BTreeMap<String, CellInfo>>,
    numbers: BTreeMap<(String, String), Number>,
}
impl SourceWorkbook {
    pub fn open(path: &Path) -> Result<Self> {
        let package = read_package(path)?;
        let sheets = sheets(&package)?
            .into_iter()
            .map(|sheet| (fold(&sheet.name), sheet))
            .collect();
        let workbook = calamine::open_workbook(path).map_err(failure)?;
        Ok(Self {
            workbook,
            package,
            sheets,
            values: BTreeMap::new(),
            metadata: BTreeMap::new(),
            numbers: BTreeMap::new(),
        })
    }
    pub fn read_sheet_count(&self) -> usize {
        self.values.len()
    }
    pub fn get(&mut self, reference: &Reference) -> std::result::Result<Number, SourceReadError> {
        let name = fold(&reference.sheet);
        let key = (name.clone(), reference.coordinate.clone());
        if let Some(number) = self.numbers.get(&key) {
            return Ok(number.clone());
        }
        let Some(sheet) = self.sheets.get(&name) else {
            return Err(SourceReadError::at(
                "missing_sheet",
                "找不到工作表，不能按空白0处理。",
                reference,
            ));
        };
        let fail =
            |error: AggregationError| SourceReadError::at("read_failed", error.message, reference);
        if !self.metadata.contains_key(&name) {
            self.metadata.insert(
                name.clone(),
                metadata(&self.package, &sheet.part).map_err(fail)?,
            );
        }
        if let Some(info) = self.metadata[&name].get(&reference.coordinate) {
            if let Some(error) = &info.error {
                return Err(SourceReadError::at(
                    "excel_error",
                    format!("源单元格包含Excel错误：{error}。"),
                    reference,
                ));
            }
            if info.formula && !info.cached {
                return Err(SourceReadError::at(
                    "formula_cache_missing",
                    "源单元格是公式，但没有有效的计算缓存。请用Excel打开源文件、完成重算并保存后重试。",
                    reference,
                ));
            }
        }
        if !self.values.contains_key(&name) {
            let values = self
                .workbook
                .worksheet_range(&sheet.name)
                .map_err(|error| {
                    SourceReadError::at(
                        "read_failed",
                        format!("读取工作表失败：{error}"),
                        reference,
                    )
                })?;
            self.values.insert(name.clone(), values);
        }
        let raw = self.values[&name]
            .get_value((reference.row - 1, reference.column - 1))
            .unwrap_or(&Data::Empty);
        let number = data_number(raw)
            .map_err(|message| SourceReadError::at("invalid_number", message, reference))?;
        self.numbers.insert(key, number.clone());
        Ok(number)
    }
}
pub fn data_number(data: &Data) -> std::result::Result<Number, String> {
    match data {
        Data::Empty => Ok(Number::blank()),
        Data::String(value) => parse_text(value),
        Data::Int(value) => parse_text(&value.to_string()),
        Data::Float(value) if value.is_finite() => parse_text(&value.to_string()),
        Data::Bool(_) => Err("布尔值不是金额或利率，请填写数字。".into()),
        Data::Error(error) => Err(format!("源单元格包含Excel错误：{error}")),
        _ => Err("不支持的数值类型，请检查是否引用到了日期、表头或非有限数值。".into()),
    }
}
