//! 工作簿结构解析：sheet 清单、隐藏状态、1904 日期系统

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::error::{ExcelError, Result};
use crate::read::images::normalize_path;
use crate::read::source::ZipSource;

/// 一个工作表
#[derive(Debug, Clone)]
pub struct SheetInfo {
    /// 0 基下标
    pub index: u32,
    pub name: String,
    /// zip 内的路径（如 `xl/worksheets/sheet1.xml`）
    pub path: String,
    /// 是否隐藏（`state="hidden"` / `"veryHidden"`）
    pub hidden: bool,
}

/// 解析工作簿：返回 (工作表清单, 是否 1904 日期系统)
pub fn load_sheets(source: &ZipSource) -> Result<(Vec<SheetInfo>, bool)> {
    let workbook_xml = source.read_entry("xl/workbook.xml")?;
    let rels = source
        .read_entry("xl/_rels/workbook.xml.rels")
        .map(|x| parse_rels(&x))
        .unwrap_or_default();

    let mut reader = Reader::from_reader(workbook_xml.as_slice());
    reader.config_mut().trim_text(true);
    let mut buf = Vec::with_capacity(8 * 1024);
    let mut raw: Vec<(String, Option<String>, bool)> = Vec::new();
    let mut date_1904 = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Err(e) => return Err(ExcelError::Xml(format!("workbook.xml 解析失败：{e}"))),
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"workbookPr" => {
                    for a in e.attributes().flatten() {
                        if a.key.as_ref() == b"date1904" || a.key.as_ref().ends_with(b":date1904") {
                            let v = String::from_utf8_lossy(&a.value).to_ascii_lowercase();
                            date_1904 = v == "1" || v == "true";
                        }
                    }
                }
                b"sheet" => {
                    let mut name = String::new();
                    let mut rid: Option<String> = None;
                    let mut hidden = false;
                    for a in e.attributes().flatten() {
                        let key = a.key.as_ref();
                        let val = String::from_utf8_lossy(&a.value).to_string();
                        if key == b"name" {
                            name = val;
                        } else if key == b"state" {
                            hidden = val == "hidden" || val == "veryHidden";
                        } else if key.ends_with(b":id") || key == b"id" {
                            rid = Some(val);
                        }
                    }
                    if !name.is_empty() {
                        raw.push((name, rid, hidden));
                    }
                }
                _ => {}
            },
            _ => {}
        }
        buf.clear();
    }

    let mut sheets = Vec::with_capacity(raw.len());
    for (i, (name, rid, hidden)) in raw.into_iter().enumerate() {
        let path = rid
            .as_ref()
            .and_then(|id| rels.get(id))
            .map(|target| normalize_workbook_target(target))
            .unwrap_or_else(|| fallback_sheet_path(i));
        sheets.push(SheetInfo {
            index: i as u32,
            name,
            path,
            hidden,
        });
    }

    Ok((sheets, date_1904))
}

fn fallback_sheet_path(index: usize) -> String {
    format!("xl/worksheets/sheet{}.xml", index + 1)
}

/// rels 的 Target → zip 内路径（去掉前导 `/`，缺 `xl/` 前缀时补上）
fn normalize_workbook_target(target: &str) -> String {
    let t = target.strip_prefix('/').unwrap_or(target);
    if t.starts_with("xl/") {
        t.to_string()
    } else {
        normalize_path(&format!("xl/{t}"))
    }
}

fn parse_rels(xml: &[u8]) -> std::collections::HashMap<String, String> {
    let mut map = std::collections::HashMap::new();
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                if e.local_name().as_ref() != b"Relationship" {
                    buf.clear();
                    continue;
                }
                let mut id = String::new();
                let mut target = String::new();
                for a in e.attributes().flatten() {
                    let k = a.key.as_ref();
                    let v = String::from_utf8_lossy(&a.value).to_string();
                    if k == b"Id" {
                        id = v;
                    } else if k == b"Target" {
                        target = v;
                    }
                }
                if !id.is_empty() {
                    map.insert(id, target);
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    map
}
