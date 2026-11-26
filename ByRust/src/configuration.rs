use crate::{
    functions::FunctionRegistry, model::Result, package::*, reader::metadata, rules::is_alias,
};
use calamine::{Data, Reader, Xlsx};
use std::{
    collections::BTreeMap,
    io::Cursor,
    path::{Path, PathBuf},
};

#[derive(Clone, Debug)]
pub struct ConfigurationSnapshot {
    pub bindings: BTreeMap<String, String>,
    pub path: Option<PathBuf>,
}
pub fn load_configuration(
    path: Option<&Path>,
    registry: &FunctionRegistry,
) -> Result<ConfigurationSnapshot> {
    let explicit = path.is_some();
    let path = path.map(Path::to_path_buf).unwrap_or_else(|| {
        std::env::current_exe()
            .ok()
            .and_then(|path| path.parent().map(Path::to_path_buf))
            .unwrap_or_default()
            .join("config.xlsx")
    });
    let path = std::path::absolute(path).map_err(failure)?;
    let fail = |message: String| {
        failure(format!(
            "配置校验失败，配置文件“{}”：{message}",
            path.display()
        ))
    };
    if !path
        .extension()
        .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx"))
    {
        return Err(fail("扩展名必须为.xlsx。".into()));
    }
    let content = match std::fs::read(&path) {
        Ok(content) => content,
        Err(error) if !explicit && error.kind() == std::io::ErrorKind::NotFound => {
            return Ok(ConfigurationSnapshot {
                bindings: BTreeMap::new(),
                path: None,
            });
        }
        Err(error) => return Err(fail(error.to_string())),
    };
    let package = unzip(&content).map_err(|error| fail(error.message))?;
    let sheet = sheets(&package)?
        .into_iter()
        .find(|sheet| sheet.name == "函数映射")
        .ok_or_else(|| fail("缺少“函数映射”工作表。".into()))?;
    let metadata = metadata(&package, &sheet.part)?;
    let mut workbook = Xlsx::new(Cursor::new(content)).map_err(|error| fail(error.to_string()))?;
    let values = workbook
        .worksheet_range("函数映射")
        .map_err(|error| fail(error.to_string()))?;
    let mut bindings = BTreeMap::new();
    let maximum = values.end().map_or(0, |(row, _)| row + 1).max(
        metadata
            .keys()
            .filter_map(|cell| {
                cell.trim_start_matches(|character: char| character.is_ascii_alphabetic())
                    .parse::<u32>()
                    .ok()
            })
            .max()
            .unwrap_or(0),
    );
    for row in 1..=maximum {
        let location = format!("工作表“函数映射”第{row}行（A{row}/B{row}）");
        if ['A', 'B'].iter().any(|column| {
            metadata
                .get(&format!("{column}{row}"))
                .is_some_and(|info| info.formula || info.error.is_some())
        }) {
            return Err(fail(format!(
                "{location}包含Excel公式或错误值，请直接填写文本。"
            )));
        }
        let alias = values.get_value((row - 1, 0)).unwrap_or(&Data::Empty);
        let target = values.get_value((row - 1, 1)).unwrap_or(&Data::Empty);
        let text = |data: &Data| match data {
            Data::String(text) => Some(text.trim().to_owned()),
            Data::Empty => Some(String::new()),
            _ => None,
        };
        let alias = text(alias);
        let target = text(target);
        if row == 1 && alias.as_deref() == Some("函数别名") && target.as_deref() == Some("映射函数")
        {
            continue;
        }
        if alias.as_deref() == Some("") && target.as_deref() == Some("") {
            continue;
        }
        let alias = alias
            .filter(|text| !text.is_empty())
            .ok_or_else(|| fail(format!("{location}别名必须是非空文本。")))?
            .to_uppercase();
        if registry.get(&alias).is_some() {
            return Err(fail(format!("{location}不能覆盖内置算法名{alias}。")));
        }
        if !is_alias(&alias) {
            return Err(fail(format!(
                "{location}别名{alias}必须为FUN_正整数（不带前导0）。"
            )));
        }
        if bindings.contains_key(&alias) {
            return Err(fail(format!(
                "{location}别名{alias}重复，名称不区分大小写。"
            )));
        }
        let target = target
            .filter(|text| !text.is_empty())
            .ok_or_else(|| fail(format!("{location}映射算法必须为非空文本。")))?;
        let specification = registry
            .get(&target)
            .ok_or_else(|| fail(format!("{location}未知算法{target}，不支持别名链。")))?;
        bindings.insert(alias, specification.name.clone());
    }
    Ok(ConfigurationSnapshot {
        bindings,
        path: Some(path),
    })
}
