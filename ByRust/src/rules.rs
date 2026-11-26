use crate::functions::FunctionRegistry;
use regex::Regex;
use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;
use std::sync::OnceLock;

#[derive(Clone, Debug, Deserialize, Eq, Hash, PartialEq, Serialize)]
pub struct Reference {
    pub sheet: String,
    pub coordinate: String,
    pub row: u32,
    pub column: u32,
    pub end_row: u32,
    pub end_column: u32,
}

impl Reference {
    pub fn shape(&self) -> (u32, u32) {
        (
            self.end_row - self.row + 1,
            self.end_column - self.column + 1,
        )
    }
    pub fn label(&self) -> String {
        format!("'{}'!{}", self.sheet.replace('\'', "''"), self.coordinate)
    }
    pub fn cells(&self) -> impl Iterator<Item = Self> + '_ {
        (self.row..=self.end_row).flat_map(move |row| {
            (self.column..=self.end_column).map(move |column| Self {
                sheet: self.sheet.clone(),
                coordinate: address(row, column),
                row,
                column,
                end_row: row,
                end_column: column,
            })
        })
    }
}

#[derive(Clone, Debug, Deserialize, Serialize)]
pub struct TemplateRule {
    pub sheet: String,
    pub coordinate: String,
    pub operation: String,
    pub references: Vec<Reference>,
    pub expression: String,
    pub function_name: String,
}

pub fn address(row: u32, mut column: u32) -> String {
    let mut characters = Vec::new();
    while column > 0 {
        column -= 1;
        characters.push(char::from(b'A' + (column % 26) as u8));
        column /= 26;
    }
    format!("{}{row}", characters.iter().rev().collect::<String>())
}

pub fn is_alias(name: &str) -> bool {
    static PATTERN: OnceLock<Regex> = OnceLock::new();
    PATTERN
        .get_or_init(|| Regex::new(r"(?i)^(?:FUN_[1-9][0-9]*|F[1-9][0-9]*_[1-9][0-9]*)$").unwrap())
        .is_match(name)
}

pub fn legacy_binding(name: &str) -> Option<&'static str> {
    match name {
        "F1_1" => Some("SUM"),
        "F1_2" => Some("AVG"),
        "F2_1" => Some("WAVG"),
        "F1_3" => Some("MAX"),
        "F1_4" => Some("MIN"),
        "F1_5" => Some("COUNT"),
        _ => None,
    }
}

pub fn parse_reference(text: &str, current_sheet: &str) -> std::result::Result<Reference, String> {
    static PATTERN: OnceLock<Regex> = OnceLock::new();
    let pattern = PATTERN.get_or_init(|| Regex::new(
        r"^(?:(?P<sheet>'(?:[^']|'')+'|[^'!]+)\s*!\s*)?\$?(?P<column>[A-Za-z]{1,3})\$?(?P<row>[1-9][0-9]{0,6})(?:\s*:\s*\$?(?P<end_column>[A-Za-z]{1,3})\$?(?P<end_row>[1-9][0-9]{0,6}))?$"
    ).unwrap());
    let matched = pattern.captures(text.trim()).ok_or_else(|| {
        format!("无效引用“{text}”，请使用A5、A1:B3或'工作表名'!A1:B3，不支持运算表达式。")
    })?;
    let sheet = if let Some(sheet) = matched.name("sheet") {
        let sheet = sheet.as_str();
        if sheet.starts_with('\'') {
            sheet[1..sheet.len() - 1].replace("''", "'")
        } else {
            let sheet = sheet.trim();
            if sheet.chars().any(char::is_whitespace) {
                return Err("包含空格的工作表名请用单引号括起来。".into());
            }
            sheet.into()
        }
    } else {
        current_sheet.into()
    };
    if sheet.is_empty()
        || sheet.chars().count() > 31
        || sheet
            .chars()
            .any(|character| "[]:*?/\\".contains(character))
    {
        return Err(format!("无效工作表名“{sheet}”，不支持外部工作簿引用。"));
    }
    let column_index = |column: &str| {
        column.bytes().fold(0u32, |index, letter| {
            index * 26 + u32::from(letter.to_ascii_uppercase() - b'A') + 1
        })
    };
    let column = column_index(&matched["column"]);
    let row = matched["row"].parse::<u32>().unwrap();
    let end_column = matched
        .name("end_column")
        .map_or(column, |value| column_index(value.as_str()));
    let end_row = matched
        .name("end_row")
        .map_or(row, |value| value.as_str().parse::<u32>().unwrap());
    if row.max(end_row) > 1_048_576 || column.max(end_column) > 16_384 {
        return Err("单元格地址超出Excel范围（最大XFD1048576）。".into());
    }
    if end_row < row || end_column < column {
        return Err("区域终点必须位于起点的右下方或同一位置。".into());
    }
    let coordinate = if matched.name("end_row").is_some() {
        format!("{}:{}", address(row, column), address(end_row, end_column))
    } else {
        address(row, column)
    };
    Ok(Reference {
        sheet,
        coordinate,
        row,
        column,
        end_row,
        end_column,
    })
}

fn arguments(text: &str) -> std::result::Result<Vec<String>, String> {
    let mut arguments = Vec::new();
    let mut start = 0;
    let mut quoted = false;
    let mut characters = text.char_indices().peekable();
    while let Some((position, character)) = characters.next() {
        if character == '\'' {
            if quoted && characters.peek().is_some_and(|(_, next)| *next == '\'') {
                characters.next();
            } else {
                quoted = !quoted;
            }
        } else if !quoted && [',', '，'].contains(&character) {
            arguments.push(text[start..position].trim().into());
            start = position + character.len_utf8();
        } else if !quoted && ['(', ')'].contains(&character) {
            return Err("参数只支持单格或矩形区域引用，不支持嵌套函数。".into());
        }
    }
    if quoted {
        return Err("工作表名的单引号没有闭合。".into());
    }
    arguments.push(text[start..].trim().into());
    Ok(arguments)
}

pub fn parse_rule(
    value: &str,
    sheet: &str,
    coordinate: &str,
    bindings: &BTreeMap<String, String>,
    registry: &FunctionRegistry,
) -> std::result::Result<Option<TemplateRule>, String> {
    let text = value.trim();
    static FUNCTION: OnceLock<Regex> = OnceLock::new();
    static CANDIDATE: OnceLock<Regex> = OnceLock::new();
    let function = FUNCTION
        .get_or_init(|| Regex::new(r"(?s)^([A-Za-z_][A-Za-z_0-9]*)\s*[(（](.*)[)）]$").unwrap());
    let candidate =
        CANDIDATE.get_or_init(|| Regex::new(r"(?s)^([A-Za-z_][A-Za-z_0-9]*)(.*)$").unwrap());
    let position = format!("模板 '{sheet}'!{coordinate}，原始指令“{value}”");
    if let Some(formula) = text.strip_prefix('=') {
        if let Some(name) = candidate.captures(formula.trim_start()) {
            let identifier = name[1].to_uppercase();
            if (is_alias(&identifier)
                || (identifier != "SUM" && registry.get(&identifier).is_some()))
                && name[2].trim_start().starts_with(['(', '（'])
            {
                return Err(format!(
                    "{position}，{identifier}是工具指令，请去掉开头的等号。"
                ));
            }
        }
        return Ok(None);
    }
    let matched = function.captures(text);
    if matched.is_none() && !is_alias(text) {
        if let Some(name) = candidate.captures(text)
            && name[2]
                .chars()
                .any(|character| "()（）".contains(character))
            && (registry.get(&name[1]).is_some() || is_alias(&name[1]))
        {
            return Err(format!(
                "{position}，汇总指令格式不完整，请使用成对的()或（），并检查参数。"
            ));
        }
        return Ok(None);
    }
    let name = matched
        .as_ref()
        .map_or_else(|| text.to_uppercase(), |value| value[1].to_uppercase());
    let mut position = format!("{position}，函数 {name}");
    let result = (|| {
        let mut parameters = if let Some(value) = &matched {
            arguments(&value[2])?
        } else {
            vec![String::new()]
        };
        let target = if registry.get(&name).is_some() {
            name.clone()
        } else {
            if !is_alias(&name) {
                return Err(format!("不支持指令{name}。"));
            }
            let mut normalized = BTreeMap::new();
            for (key, value) in bindings {
                if normalized
                    .insert(key.to_uppercase(), value.clone())
                    .is_some()
                {
                    return Err(format!("配置中存在大小写重复的别名{key}。"));
                }
            }
            normalized
                .get(&name)
                .cloned()
                .or_else(|| legacy_binding(&name).map(String::from))
                .ok_or_else(|| {
                    format!(
                        "函数别名{name}未配置，请在config.xlsx的“函数映射”页A/B列绑定实际算法。"
                    )
                })?
        };
        let specification = registry
            .get(target.trim())
            .ok_or_else(|| format!("别名{name}对应的算法{target}尚未实现。"))?;
        position.push_str(&format!(" → {}", specification.name));
        for argument in &mut parameters {
            if argument.is_empty() {
                *argument = coordinate.into();
            }
        }
        if !specification.accepts_count(parameters.len()) {
            return Err(format!(
                "期望 {} 个参数，实际 {} 个参数，请使用{}。",
                specification.expected_count(),
                parameters.len(),
                specification.signature(&name)
            ));
        }
        let references = parameters
            .iter()
            .map(|argument| parse_reference(argument, sheet))
            .collect::<std::result::Result<Vec<_>, _>>()?;
        specification.validate_references(&references)?;
        Ok(TemplateRule {
            sheet: sheet.into(),
            coordinate: coordinate.into(),
            operation: specification.name.clone(),
            references,
            expression: text.into(),
            function_name: name,
        })
    })();
    result
        .map(Some)
        .map_err(|error: String| format!("{position}，{error}"))
}
