//! Small OOXML inspection helpers. Cell values themselves are read with calamine.
use crate::model::{AggregationError, Result};
use roxmltree::{Document, Node};
use std::{
    collections::BTreeMap,
    io::{Cursor, Read, Write},
    path::Path,
};

pub(crate) const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) type Package = BTreeMap<String, Vec<u8>>;
pub(crate) fn failure(error: impl std::fmt::Display) -> AggregationError {
    AggregationError::new(error.to_string())
}
pub(crate) fn unzip(bytes: &[u8]) -> Result<Package> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).map_err(failure)?;
    let mut files = BTreeMap::new();
    for index in 0..archive.len() {
        let mut file = archive.by_index(index).map_err(failure)?;
        if !file.is_dir() {
            let mut bytes = Vec::new();
            file.read_to_end(&mut bytes).map_err(failure)?;
            files.insert(file.name().into(), bytes);
        }
    }
    Ok(files)
}
pub(crate) fn read_package(path: &Path) -> Result<Package> {
    unzip(&std::fs::read(path).map_err(failure)?)
}
pub(crate) fn save_package(package: &Package, path: &Path) -> Result<()> {
    let mut archive = zip::ZipWriter::new(std::fs::File::create(path).map_err(failure)?);
    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    for (name, bytes) in package {
        archive.start_file(name, options).map_err(failure)?;
        archive.write_all(bytes).map_err(failure)?;
    }
    archive.finish().map_err(failure)?;
    Ok(())
}
pub(crate) fn xml(package: &Package, name: &str) -> Result<String> {
    String::from_utf8(
        package
            .get(name)
            .ok_or_else(|| failure(format!("缺少OOXML部件 {name}")))?
            .clone(),
    )
    .map_err(failure)
}
pub(crate) fn document(text: &str) -> Result<Document<'_>> {
    Document::parse(text).map_err(failure)
}
pub(crate) fn children<'a, 'input>(
    node: Node<'a, 'input>,
    name: &'a str,
) -> impl Iterator<Item = Node<'a, 'input>> + 'a {
    node.children()
        .filter(move |child| child.has_tag_name(name))
}
pub(crate) fn escape(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\r', "&#13;")
}
pub(crate) fn resolve(part: &str, target: &str) -> String {
    let decoded = percent_decode(target);
    let full = if decoded.starts_with('/') {
        decoded
    } else {
        format!(
            "{}/{decoded}",
            part.rsplit_once('/').map_or("", |(directory, _)| directory)
        )
    };
    let mut pieces = Vec::new();
    for piece in full.split('/') {
        match piece {
            "" | "." => (),
            ".." => {
                pieces.pop();
            }
            _ => pieces.push(piece),
        }
    }
    pieces.join("/")
}
fn percent_decode(value: &str) -> String {
    let bytes = value.as_bytes();
    let mut result = Vec::new();
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'%'
            && index + 2 < bytes.len()
            && let (Some(high), Some(low)) = (
                (bytes[index + 1] as char).to_digit(16),
                (bytes[index + 2] as char).to_digit(16),
            )
        {
            result.push((high * 16 + low) as u8);
            index += 3;
            continue;
        }
        result.push(bytes[index]);
        index += 1;
    }
    String::from_utf8_lossy(&result).into_owned()
}
pub(crate) fn relationships(part: &str) -> String {
    let (directory, filename) = part.rsplit_once('/').unwrap_or(("", part));
    format!("{directory}/_rels/{filename}.rels")
}
#[derive(Clone, Debug)]
pub(crate) struct Sheet {
    pub name: String,
    pub part: String,
}
pub(crate) fn sheets(package: &Package) -> Result<Vec<Sheet>> {
    let relations = xml(package, "xl/_rels/workbook.xml.rels")?;
    let relations = document(&relations)?;
    let targets: BTreeMap<_, _> = relations
        .root_element()
        .children()
        .filter(|node| node.is_element() && node.attribute("TargetMode") != Some("External"))
        .filter_map(|node| {
            Some((
                node.attribute("Id")?,
                resolve("xl/workbook.xml", node.attribute("Target")?),
            ))
        })
        .collect();
    let workbook = xml(package, "xl/workbook.xml")?;
    let workbook = document(&workbook)?;
    Ok(workbook
        .descendants()
        .filter(|node| node.has_tag_name("sheet"))
        .filter_map(|node| {
            Some(Sheet {
                name: node.attribute("name")?.into(),
                part: targets.get(node.attribute((REL, "id"))?)?.clone(),
            })
        })
        .collect())
}
/// Apply disjoint byte-range edits from right to left, retaining all other XML verbatim.
pub(crate) fn edited(mut text: String, mut edits: Vec<(std::ops::Range<usize>, String)>) -> String {
    edits.sort_by_key(|(range, _)| std::cmp::Reverse(range.start));
    for (range, replacement) in edits {
        text.replace_range(range, &replacement);
    }
    text
}
pub(crate) fn append_children(text: &str, name: &str, addition: &str) -> Result<String> {
    let doc = document(text)?;
    let node = doc
        .descendants()
        .find(|node| node.has_tag_name(name))
        .ok_or_else(|| failure(format!("缺少 {name}")))?;
    let raw = &text[node.range()];
    let replacement = if let Some(opening) = raw.strip_suffix("/>") {
        let qualified = opening[1..]
            .split(|character: char| character.is_whitespace())
            .next()
            .unwrap_or(node.tag_name().name());
        format!("{}>{addition}</{}>", opening, qualified)
    } else {
        let end = raw.rfind("</").ok_or_else(|| failure("无效XML结束标记"))?;
        format!("{}{addition}{}", &raw[..end], &raw[end..])
    };
    Ok(edited(text.into(), vec![(node.range(), replacement)]))
}
