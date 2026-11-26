//! Preserve original business XML verbatim around the cells changed by umya.
//!
//! umya 3.1.1 drops dataValidation formula1/formula2 on a normal round trip.
//! Copying only changed values/new comments/new log rows also protects unsupported
//! style attributes, named ranges, relationships, print settings and extensions.
use crate::{model::Result, package::*};
use std::collections::{BTreeMap, BTreeSet};

fn attribute_edit(
    text: &str,
    node: roxmltree::Node<'_, '_>,
    name: &str,
    value: &str,
) -> (std::ops::Range<usize>, String) {
    let raw = &text[node.range()];
    let head = raw.find('>').unwrap();
    let mut opening = raw[..head].to_string();
    if let Some(attribute) = node.attributes().find(|attribute| attribute.name() == name) {
        let range = attribute.range();
        let local = range.start - node.range().start..range.end - node.range().start;
        opening.replace_range(local, &format!("{name}=\"{}\"", escape(value)));
    } else {
        let position = opening
            .strip_suffix('/')
            .map_or(opening.len(), |value| value.len());
        opening.insert_str(position, &format!(" {name}=\"{}\"", escape(value)));
    }
    (node.range(), format!("{}{}", opening, &raw[head..]))
}
fn add_counted(text: &str, name: &str, addition: &str, count: u32) -> Result<String> {
    let expanded = append_children(text, name, addition)?;
    let doc = document(&expanded)?;
    let node = doc
        .descendants()
        .find(|node| node.has_tag_name(name))
        .unwrap();
    let edit = attribute_edit(&expanded, node, "count", &count.to_string());
    Ok(edited(expanded, vec![edit]))
}
fn log_styles(package: &mut Package) -> Result<[u32; 3]> {
    let mut styles = xml(package, "xl/styles.xml")?;
    let doc = document(&styles)?;
    let fonts = doc
        .descendants()
        .find(|node| node.has_tag_name("fonts"))
        .ok_or_else(|| failure("模板缺少fonts"))?;
    let font = fonts.children().filter(|node| node.is_element()).count() as u32;
    let xfs = doc
        .descendants()
        .find(|node| node.has_tag_name("cellXfs"))
        .ok_or_else(|| failure("模板缺少cellXfs"))?;
    let first = xfs.children().filter(|node| node.is_element()).count() as u32;
    styles = add_counted(
        &styles,
        "fonts",
        "<font><b/><sz val=\"11\"/><color rgb=\"FF17365D\"/><name val=\"微软雅黑\"/></font>",
        font + 1,
    )?;
    let additions = format!(
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyAlignment=\"1\"><alignment vertical=\"top\" wrapText=\"1\"/></xf><xf numFmtId=\"0\" fontId=\"{font}\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyAlignment=\"1\"><alignment vertical=\"center\" wrapText=\"1\"/></xf><xf numFmtId=\"0\" fontId=\"{font}\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyAlignment=\"1\"><alignment vertical=\"center\" wrapText=\"1\"/></xf>"
    );
    styles = add_counted(&styles, "cellXfs", &additions, first + 3)?;
    package.insert("xl/styles.xml".into(), styles.into_bytes());
    Ok([first, first + 1, first + 2])
}
fn content_type(package: &mut Package, path: &str, kind: &str) -> Result<()> {
    let name = "[Content_Types].xml";
    let original = xml(package, name)?;
    let doc = document(&original)?;
    if !doc
        .descendants()
        .any(|node| node.attribute("PartName") == Some(&format!("/{path}")))
    {
        let text = append_children(
            &original,
            "Types",
            &format!(
                "<Override PartName=\"/{}\" ContentType=\"{}\"/>",
                escape(path),
                escape(kind)
            ),
        )?;
        package.insert(name.into(), text.into_bytes());
    }
    Ok(())
}
fn shared_strings(package: &Package) -> Result<Vec<String>> {
    if !package.contains_key("xl/sharedStrings.xml") {
        return Ok(Vec::new());
    }
    let text = xml(package, "xl/sharedStrings.xml")?;
    let doc = document(&text)?;
    Ok(doc
        .descendants()
        .filter(|node| node.has_tag_name("si"))
        .map(|node| {
            node.descendants()
                .filter(|node| node.has_tag_name("t"))
                .filter_map(|node| node.text())
                .collect()
        })
        .collect())
}
fn new_log_xml(
    text: &str,
    strings: &[String],
    start: u32,
    header: u32,
    styles: [u32; 3],
) -> Result<String> {
    let doc = document(text)?;
    let mut edits = Vec::new();
    for node in doc.descendants().filter(|node| node.has_tag_name("c")) {
        let coordinate = node.attribute("r").unwrap_or("");
        let row = coordinate
            .trim_start_matches(|character: char| character.is_ascii_alphabetic())
            .parse::<u32>()
            .unwrap_or(0);
        if row < start {
            continue;
        }
        let style = if row == header {
            styles[1]
        } else if row == start {
            styles[2]
        } else {
            styles[0]
        };
        let value = if node.attribute("t") == Some("s") {
            let index = children(node, "v")
                .next()
                .and_then(|node| node.text())
                .unwrap_or("0")
                .parse::<usize>()
                .map_err(failure)?;
            let value = strings
                .get(index)
                .ok_or_else(|| failure("无效共享字符串索引"))?;
            format!(
                "<c r=\"{coordinate}\" s=\"{style}\" t=\"inlineStr\"><is><t xml:space=\"preserve\">{}</t></is></c>",
                escape(value)
            )
        } else {
            let (_, value) = attribute_edit(text, node, "s", &style.to_string());
            value
        };
        edits.push((node.range(), value));
    }
    Ok(edited(text.into(), edits))
}
fn pixels(style: &str, name: &str) -> f64 {
    let value = style
        .split(';')
        .filter_map(|piece| piece.split_once(':'))
        .find(|(key, _)| key.trim() == name)
        .map_or("0", |(_, value)| value.trim());
    if let Some(points) = value.strip_suffix("pt") {
        points.parse::<f64>().unwrap_or(0.0) / 0.75
    } else {
        value.trim_end_matches("px").parse::<f64>().unwrap_or(0.0)
    }
}
/// umya does not recover every openpyxl VML prefix/shape. Preserve existing
/// shapes and repair new ClientData coordinates from the ordered comment list.
fn repair_vml(old: Option<&str>, new: &str, comments: &str) -> Result<String> {
    let comments_doc = document(comments)?;
    let coordinates = comments_doc
        .descendants()
        .filter(|node| node.has_tag_name("comment"))
        .filter_map(|node| node.attribute("ref"))
        .map(|coordinate| crate::rules::parse_reference(coordinate, "Sheet").map_err(failure))
        .collect::<Result<Vec<_>>>()?;
    let new_doc = document(new)?;
    let new_shapes = new_doc
        .descendants()
        .filter(|node| node.has_tag_name("shape"))
        .collect::<Vec<_>>();
    if new_shapes.len() != coordinates.len() {
        return Err(failure("批注框与批注数量不符，未发布结果。"));
    }
    let base = old.unwrap_or(new);
    let base_doc = document(base)?;
    let mut original_shapes = BTreeMap::new();
    let mut maximum_id = 1024u32;
    for shape in base_doc
        .descendants()
        .filter(|node| node.has_tag_name("shape"))
    {
        if let Some(id) = shape
            .attribute("id")
            .and_then(|id| id.strip_prefix("_x0000_s"))
            .and_then(|id| id.parse::<u32>().ok())
        {
            maximum_id = maximum_id.max(id);
        }
        let row = shape
            .descendants()
            .find(|node| node.has_tag_name("Row"))
            .and_then(|node| node.text())
            .and_then(|value| value.parse::<u32>().ok());
        let column = shape
            .descendants()
            .find(|node| node.has_tag_name("Column"))
            .and_then(|node| node.text())
            .and_then(|value| value.parse::<u32>().ok());
        if let (Some(row), Some(column)) = (row, column) {
            original_shapes.insert((row + 1, column + 1), shape);
        }
    }
    let mut edits = Vec::new();
    let mut addition = String::new();
    for (index, (shape, coordinate)) in new_shapes.iter().zip(&coordinates).enumerate() {
        if old.is_some()
            && let Some(original) = original_shapes.get(&(coordinate.row, coordinate.column))
        {
            let new_style = shape.attribute("style").unwrap_or("");
            if pixels(new_style, "width") > 0.0 {
                let old_style = original.attribute("style").unwrap_or("");
                let width = pixels(new_style, "width").max(pixels(old_style, "width"));
                let height = pixels(new_style, "height").max(pixels(old_style, "height"));
                let style = old_style
                    .split(';')
                    .filter(|piece| {
                        !piece.trim_start().starts_with("width:")
                            && !piece.trim_start().starts_with("height:")
                    })
                    .collect::<Vec<_>>()
                    .join(";");
                edits.push(attribute_edit(
                    base,
                    *original,
                    "style",
                    &format!(
                        "{style};width:{}pt;height:{}pt",
                        width * 0.75,
                        height * 0.75
                    ),
                ));
            }
            continue;
        }
        let row = coordinate.row - 1;
        let column = coordinate.column - 1;
        let client = format!(
            "<x:ClientData xmlns:x=\"urn:schemas-microsoft-com:office:excel\" ObjectType=\"Note\"><x:MoveWithCells/><x:ResizeWithCells/><x:Anchor>{column}, 15, {row}, 2, {}, 31, {}, 10</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>{row}</x:Row><x:Column>{column}</x:Column></x:ClientData>",
            column.saturating_add(8).min(16383),
            row.saturating_add(12).min(1048575)
        );
        let raw = &new[shape.range()];
        let fixed = if let Some(client_node) = shape
            .children()
            .find(|node| node.has_tag_name("ClientData"))
        {
            edited(
                raw.into(),
                vec![(
                    client_node.range().start - shape.range().start
                        ..client_node.range().end - shape.range().start,
                    client,
                )],
            )
        } else {
            let end = raw.rfind("</").ok_or_else(|| failure("无效批注框"))?;
            format!("{}{client}{}", &raw[..end], &raw[end..])
        };
        if old.is_none() {
            edits.push((shape.range(), fixed));
        } else {
            let wrapped = format!(
                "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:o=\"urn:schemas-microsoft-com:office:office\">{fixed}</xml>"
            );
            let doc = document(&wrapped)?;
            let node = doc
                .descendants()
                .find(|node| node.has_tag_name("shape"))
                .unwrap();
            let (range, replacement) = attribute_edit(
                &wrapped,
                node,
                "id",
                &format!("_x0000_s{}", maximum_id + index as u32 + 1),
            );
            let replacement=replacement.replacen("<v:shape","<v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:o=\"urn:schemas-microsoft-com:office:office\"",1);
            let _ = range;
            addition.push_str(&replacement);
        }
    }
    let result = edited(base.into(), edits);
    if addition.is_empty() {
        Ok(result)
    } else {
        append_children(&result, "xml", &addition)
    }
}
fn copy_comments(
    original: &mut Package,
    generated: &Package,
    old_sheet: &Sheet,
    new_sheet: &Sheet,
    ordinal: usize,
) -> Result<()> {
    let new_rels = relationships(&new_sheet.part);
    if !generated.contains_key(&new_rels) {
        return Ok(());
    }
    let text = xml(generated, &new_rels)?;
    let doc = document(&text)?;
    let old_rels = relationships(&old_sheet.part);
    let mut rels=original.get(&old_rels).map(|value|String::from_utf8_lossy(value).into_owned()).unwrap_or_else(||"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>".into());
    let old_doc = document(&rels)?;
    let old_vml = old_doc
        .root_element()
        .children()
        .find(|node| {
            node.attribute("Type")
                .is_some_and(|kind| kind.ends_with("/vmlDrawing"))
        })
        .and_then(|node| node.attribute("Target"))
        .map(|target| resolve(&old_sheet.part, target))
        .and_then(|part| original.get(&part))
        .map(|bytes| String::from_utf8_lossy(bytes).into_owned());
    let new_comment_part = doc
        .root_element()
        .children()
        .find(|node| {
            node.attribute("Type")
                .is_some_and(|kind| kind.ends_with("/comments"))
        })
        .and_then(|node| node.attribute("Target"))
        .map(|target| resolve(&new_sheet.part, target));
    let new_comments = new_comment_part
        .as_ref()
        .map(|part| xml(generated, part))
        .transpose()?;
    let removals = old_doc
        .root_element()
        .children()
        .filter(|node| {
            node.attribute("Type")
                .is_some_and(|kind| kind.ends_with("/comments") || kind.ends_with("/vmlDrawing"))
        })
        .map(|node| (node.range(), String::new()))
        .collect();
    rels = edited(rels, removals);
    let mut vml_id = None;
    for relation in doc
        .root_element()
        .children()
        .filter(|node| node.is_element())
    {
        let kind = relation.attribute("Type").unwrap_or("");
        let extension = if kind.ends_with("/comments") {
            "xml"
        } else if kind.ends_with("/vmlDrawing") {
            "vml"
        } else {
            continue;
        };
        let source = resolve(&new_sheet.part, relation.attribute("Target").unwrap_or(""));
        let destination = format!("xl/rust-summary/comments_{ordinal}.{extension}");
        let bytes = generated
            .get(&source)
            .ok_or_else(|| failure("umya批注部件缺失"))?
            .clone();
        let bytes = if extension == "vml" {
            repair_vml(
                old_vml.as_deref(),
                std::str::from_utf8(&bytes).map_err(failure)?,
                new_comments
                    .as_deref()
                    .ok_or_else(|| failure("缺少批注文字"))?,
            )?
            .into_bytes()
        } else {
            bytes
        };
        original.insert(destination.clone(), bytes);
        // IDs are checked against every retained relationship, including hyperlinks.
        let mut id = format!("rIdRust{extension}{ordinal}");
        while rels.contains(&format!("\"{id}\"")) {
            id.push('_');
        }
        rels = append_children(
            &rels,
            "Relationships",
            &format!("<Relationship Id=\"{id}\" Type=\"{kind}\" Target=\"/{destination}\"/>"),
        )?;
        if extension == "vml" {
            vml_id = Some(id);
            content_type(
                original,
                &destination,
                "application/vnd.openxmlformats-officedocument.vmlDrawing",
            )?;
        } else {
            content_type(
                original,
                &destination,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
            )?;
        }
    }
    original.insert(old_rels, rels.into_bytes());
    if let Some(id) = vml_id {
        let text = xml(original, &old_sheet.part)?;
        let doc = document(&text)?;
        let tag = format!("<legacyDrawing xmlns:r=\"{REL}\" r:id=\"{id}\"/>");
        let replacement = if let Some(node) = doc
            .descendants()
            .find(|node| node.has_tag_name("legacyDrawing"))
        {
            edited(text.clone(), vec![(node.range(), tag)])
        } else if let Some(node) = doc
            .root_element()
            .children()
            .find(|node| node.has_tag_name("extLst"))
        {
            edited(
                text.clone(),
                vec![(node.range().start..node.range().start, tag)],
            )
        } else {
            append_children(&text, "worksheet", &tag)?
        };
        original.insert(old_sheet.part.clone(), replacement.into_bytes());
    }
    Ok(())
}

pub(crate) fn merge(
    mut original: Package,
    generated: Package,
    values: &BTreeMap<(String, String), f64>,
    comment_sheets: &BTreeSet<String>,
    log_start: u32,
) -> Result<Package> {
    let old_sheets = sheets(&original)?;
    let new_sheets = sheets(&generated)?;
    for (ordinal, old) in old_sheets.iter().enumerate() {
        if old.name == "运行日志" {
            continue;
        }
        let text = xml(&original, &old.part)?;
        let doc = document(&text)?;
        let mut edits = Vec::new();
        for node in doc.descendants().filter(|node| node.has_tag_name("c")) {
            let coordinate = node.attribute("r").unwrap_or("");
            if let Some(value) = values.get(&(old.name.clone(), coordinate.into())) {
                // Reuse the original style, metadata and extension attributes.
                let attrs = node
                    .attributes()
                    .filter(|attribute| attribute.name() != "t")
                    .map(|attribute| {
                        format!(" {}=\"{}\"", attribute.name(), escape(attribute.value()))
                    })
                    .collect::<String>();
                edits.push((
                    node.range(),
                    format!("<c{attrs} t=\"n\"><v>{value:.15e}</v></c>"),
                ));
            }
        }
        original.insert(old.part.clone(), edited(text, edits).into_bytes());
        if comment_sheets.contains(&old.name) {
            let new = new_sheets
                .iter()
                .find(|sheet| sheet.name == old.name)
                .ok_or_else(|| failure("umya输出缺少模板页"))?;
            copy_comments(&mut original, &generated, old, new, ordinal)?;
        }
    }
    let styles = log_styles(&mut original)?;
    let strings = shared_strings(&generated)?;
    let new_log = new_sheets
        .iter()
        .find(|sheet| sheet.name == "运行日志")
        .ok_or_else(|| failure("缺少运行日志"))?;
    let new_text = new_log_xml(
        &xml(&generated, &new_log.part)?,
        &strings,
        log_start,
        log_start + 2,
        styles,
    )?;
    if let Some(old) = old_sheets.iter().find(|sheet| sheet.name == "运行日志") {
        let new_doc = document(&new_text)?;
        let rows = new_doc
            .descendants()
            .filter(|node| {
                node.has_tag_name("row")
                    && node
                        .attribute("r")
                        .and_then(|value| value.parse::<u32>().ok())
                        .is_some_and(|row| row >= log_start)
            })
            .map(|node| &new_text[node.range()])
            .collect::<String>();
        let text = append_children(&xml(&original, &old.part)?, "sheetData", &rows)?;
        let doc = document(&text)?;
        let dimension = doc
            .descendants()
            .find(|node| node.has_tag_name("dimension"));
        let end = new_doc
            .descendants()
            .filter(|node| node.has_tag_name("row"))
            .filter_map(|node| node.attribute("r")?.parse::<u32>().ok())
            .max()
            .unwrap_or(1);
        let column = doc
            .descendants()
            .filter(|node| node.has_tag_name("c"))
            .filter_map(|node| node.attribute("r"))
            .map(|coordinate| {
                coordinate.bytes().take_while(u8::is_ascii_alphabetic).fold(
                    0u32,
                    |number, letter| {
                        number * 26 + u32::from(letter.to_ascii_uppercase() - b'A') + 1
                    },
                )
            })
            .max()
            .unwrap_or(6)
            .max(6);
        let text = if let Some(node) = dimension {
            let edit = attribute_edit(
                &text,
                node,
                "ref",
                &format!("A1:{}", crate::rules::address(end, column)),
            );
            edited(text, vec![edit])
        } else {
            text
        };
        original.insert(old.part.clone(), text.into_bytes());
    } else {
        let destination = "xl/rust-summary/run_log.xml";
        let doc = document(&new_text)?;
        let view = format!(
            "<sheetViews><sheetView workbookViewId=\"0\"><pane xSplit=\"3\" ySplit=\"{}\" topLeftCell=\"D{}\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\" activeCell=\"D{}\" sqref=\"D{}\"/></sheetView></sheetViews>",
            log_start + 2,
            log_start + 3,
            log_start + 3,
            log_start + 3
        );
        let new_text = if let Some(node) = doc
            .descendants()
            .find(|node| node.has_tag_name("sheetViews"))
        {
            edited(new_text.clone(), vec![(node.range(), view)])
        } else {
            let node = doc
                .root_element()
                .children()
                .find(|node| node.has_tag_name("sheetFormatPr") || node.has_tag_name("sheetData"))
                .ok_or_else(|| failure("日志缺少sheetData"))?;
            edited(
                new_text.clone(),
                vec![(node.range().start..node.range().start, view)],
            )
        };
        original.insert(destination.into(), new_text.into_bytes());
        let relname = "xl/_rels/workbook.xml.rels";
        let rels = xml(&original, relname)?;
        let mut id = "rIdRustLog".to_string();
        while rels.contains(&format!("\"{id}\"")) {
            id.push('_');
        }
        let rels = append_children(
            &rels,
            "Relationships",
            &format!(
                "<Relationship Id=\"{id}\" Type=\"{REL}/worksheet\" Target=\"/{destination}\"/>"
            ),
        )?;
        original.insert(relname.into(), rels.into_bytes());
        let workbook = xml(&original, "xl/workbook.xml")?;
        let doc = document(&workbook)?;
        let next = doc
            .descendants()
            .filter(|node| node.has_tag_name("sheet"))
            .filter_map(|node| node.attribute("sheetId")?.parse::<u32>().ok())
            .max()
            .unwrap_or(0)
            + 1;
        let workbook = append_children(
            &workbook,
            "sheets",
            &format!(
                "<sheet xmlns:r=\"{REL}\" name=\"运行日志\" sheetId=\"{next}\" r:id=\"{id}\"/>"
            ),
        )?;
        original.insert("xl/workbook.xml".into(), workbook.into_bytes());
        content_type(
            &mut original,
            destination,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        )?;
    }
    let workbook = xml(&original, "xl/workbook.xml")?;
    let doc = document(&workbook)?;
    let sheet_nodes = doc
        .descendants()
        .filter(|node| node.has_tag_name("sheet"))
        .collect::<Vec<_>>();
    let log_index = sheet_nodes
        .iter()
        .position(|node| node.attribute("name") == Some("运行日志"))
        .unwrap();
    let mut reordered = sheet_nodes
        .iter()
        .filter(|node| node.attribute("name") != Some("运行日志"))
        .map(|node| &workbook[node.range()])
        .collect::<String>();
    reordered.push_str(&workbook[sheet_nodes[log_index].range()]);
    let sheets = doc
        .descendants()
        .find(|node| node.has_tag_name("sheets"))
        .unwrap();
    let mut edits = vec![(sheets.range(), format!("<sheets>{reordered}</sheets>"))];
    for node in doc
        .descendants()
        .filter(|node| node.attribute("localSheetId").is_some())
    {
        if let Some(index) = node
            .attribute("localSheetId")
            .and_then(|value| value.parse::<usize>().ok())
        {
            let new = if index == log_index {
                sheet_nodes.len() - 1
            } else if index > log_index {
                index - 1
            } else {
                index
            };
            if new != index {
                edits.push(attribute_edit(
                    &workbook,
                    node,
                    "localSheetId",
                    &new.to_string(),
                ));
            }
        }
    }
    if let Some(node) = doc.descendants().find(|node| node.has_tag_name("calcPr")) {
        edits.push((
            node.range(),
            "<calcPr calcMode=\"auto\" fullCalcOnLoad=\"0\" forceFullCalc=\"0\"/>".into(),
        ));
    }
    original.insert(
        "xl/workbook.xml".into(),
        edited(workbook, edits).into_bytes(),
    );
    Ok(original)
}
