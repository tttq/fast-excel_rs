//! 单元格内嵌图片读取（导入时"图片按行对应"的底座）
//!
//! 支持 Excel / WPS 两种存储方式：
//!
//! 1. **标准 drawing 锚点**（Excel）：`xl/drawings/drawingN.xml` 记录图片锚点
//!    `(行, 列)`，`xl/media/*` 存字节；
//! 2. **WPS「单元格嵌入图片」DISPIMG**：单元格值是 `=DISPIMG("ID_xxx",1)` 公式，
//!    图片数据在 `xl/cellimages.xml`（`cNvPr@name` = 公式里的 ID）+ 对应 rels → `xl/media/*`。
//!
//! 流式读取时先在内存里建立"锚点 → media 路径"的小索引，**图片字节按需读取**，
//! 这样百万行文件也不会把图片全量读进内存。

use std::collections::HashMap;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::error::{ExcelError, Result};
use crate::read::source::ZipSource;
use crate::read::workbook::load_sheets;
// 图片类型与单元格值共用一套定义（`fast_excel::CellImage`），避免两处结构体互转
pub use crate::value::CellImage;

/// 两种存储方式的一次性提取结果
#[derive(Debug, Default)]
pub struct ExtractedImages {
    /// 标准 drawing 锚点：`(行, 列) → 图片`
    pub cells: HashMap<(u32, u16), CellImage>,
    /// WPS DISPIMG：`ID → 图片`
    pub wps: HashMap<String, CellImage>,
}

/// 一个锚点图片：media 路径 + Excel 里记录的原名 / alt 文本
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub(crate) struct AnchorImage {
    /// `xl/media/*` 路径
    pub media: String,
    /// `cNvPr@descr`：写侧 `Image::set_alt_text` 会写到这里，用于保住原文件名
    pub name: Option<String>,
}

/// 锚点索引：`(行, 列) → 锚点图片`
pub(crate) type AnchorIndex = HashMap<(u32, u16), AnchorImage>;

/// 读取指定工作表的 drawing 锚点索引
pub(crate) fn load_sheet_anchors(source: &ZipSource, sheet_path: &str) -> Result<AnchorIndex> {
    let rels_path = format!("xl/worksheets/_rels/{}.rels", file_name(sheet_path));
    let Ok(rels_xml) = source.read_entry(&rels_path) else {
        return Ok(AnchorIndex::new());
    };
    let rels = parse_rels(&rels_xml);

    let mut anchors = AnchorIndex::new();
    for target in rels.values() {
        if !target.contains("drawing") {
            continue;
        }
        let drawing_path = normalize_path(&format!("xl/worksheets/{target}"));
        let Ok(drawing_xml) = source.read_entry(&drawing_path) else {
            continue;
        };
        let cells = parse_drawing_anchors(&drawing_xml)?;
        if cells.is_empty() {
            continue;
        }
        let drawing_rels_path = format!("xl/drawings/_rels/{}.rels", file_name(&drawing_path));
        let Ok(drawing_rels_xml) = source.read_entry(&drawing_rels_path) else {
            continue;
        };
        let drawing_rels = parse_rels(&drawing_rels_xml);
        for (row, col, rid, name) in cells {
            if let Some(media) = drawing_rels.get(&rid) {
                let media_path = normalize_path(&format!("xl/drawings/{media}"));
                anchors.entry((row, col)).or_insert(AnchorImage {
                    media: media_path,
                    name,
                });
            }
        }
    }
    Ok(anchors)
}

/// 读取 WPS `cellimages.xml` 索引：`DISPIMG ID → media 路径`
pub(crate) fn load_wps_index(source: &ZipSource) -> Result<HashMap<String, String>> {
    let Ok(xml) = source.read_entry("xl/cellimages.xml") else {
        return Ok(HashMap::new());
    };
    let rels = source
        .read_entry("xl/_rels/cellimages.xml.rels")
        .map(|x| parse_rels(&x))
        .unwrap_or_default();

    let mut out = HashMap::new();
    for (id, rid) in parse_cellimage_ids(&xml)? {
        if let Some(target) = rels.get(&rid) {
            out.insert(id, normalize_path(&format!("xl/{target}")));
        }
    }
    Ok(out)
}

/// 按 media 路径读图片字节（惰性调用，一次读一张）
pub(crate) fn load_media(source: &ZipSource, media_path: &str) -> Result<CellImage> {
    let bytes = source.read_entry(media_path)?;
    // 兜底名：Excel 没写 alt 文本时，至少保留 `image1.png` 这类媒体文件名，
    // 避免图片往返（导出 → 再导入）后完全没有可读标识
    Ok(CellImage::new(bytes, ext_of(media_path)).named(file_name(media_path)))
}

fn ext_of(path: &str) -> String {
    path.rsplit('.')
        .next()
        .map(|s| s.to_ascii_lowercase())
        .filter(|e| !e.is_empty() && !e.contains('/'))
        .unwrap_or_else(|| "jpg".to_string())
}

// ───────────────────── 兼容旧接口（一次性提取，含图片字节） ─────────────────────

/// 提取**首个工作表**的锚点图片 + 全局 WPS 图片（兼容旧版接口）
pub fn extract_all_images(xlsx_bytes: &[u8]) -> Result<ExtractedImages> {
    let source = ZipSource::from_bytes(xlsx_bytes.to_vec());
    let (sheets, _) = load_sheets(&source)?;
    let cells = match sheets.first() {
        Some(sheet) => extract_sheet_images(&source, &sheet.path)?,
        None => HashMap::new(),
    };
    let wps = extract_wps_map(&source)?;
    Ok(ExtractedImages { cells, wps })
}

/// 提取首个工作表中按单元格锚定的图片
pub fn extract_embedded_images(xlsx_bytes: &[u8]) -> Result<HashMap<(u32, u16), CellImage>> {
    let source = ZipSource::from_bytes(xlsx_bytes.to_vec());
    let (sheets, _) = load_sheets(&source)?;
    match sheets.first() {
        Some(sheet) => extract_sheet_images(&source, &sheet.path),
        None => Ok(HashMap::new()),
    }
}

/// 提取 WPS DISPIMG 图片：`ID → 图片`
pub fn extract_wps_cellimages(xlsx_bytes: &[u8]) -> Result<HashMap<String, CellImage>> {
    let source = ZipSource::from_bytes(xlsx_bytes.to_vec());
    extract_wps_map(&source)
}

/// 提取指定工作表（0 基下标）的锚点图片
pub fn extract_images_of_sheet(
    xlsx_bytes: &[u8],
    sheet_index: usize,
) -> Result<HashMap<(u32, u16), CellImage>> {
    let source = ZipSource::from_bytes(xlsx_bytes.to_vec());
    let (sheets, _) = load_sheets(&source)?;
    match sheets.get(sheet_index) {
        Some(sheet) => extract_sheet_images(&source, &sheet.path),
        None => Ok(HashMap::new()),
    }
}

fn extract_sheet_images(
    source: &ZipSource,
    sheet_path: &str,
) -> Result<HashMap<(u32, u16), CellImage>> {
    let anchors = load_sheet_anchors(source, sheet_path)?;
    let mut out = HashMap::with_capacity(anchors.len());
    for (key, anchor) in anchors {
        match load_media(source, &anchor.media) {
            Ok(mut img) => {
                // 写侧把原文件名写到 alt 文本（descr），读到它就用它覆盖媒体文件名
                if let Some(name) = anchor.name {
                    img.name = Some(name);
                }
                out.insert(key, img);
            }
            Err(e) => log::warn!("读取内嵌图片失败（{}）：{e}", anchor.media),
        }
    }
    Ok(out)
}

fn extract_wps_map(source: &ZipSource) -> Result<HashMap<String, CellImage>> {
    let index = load_wps_index(source)?;
    let mut out = HashMap::with_capacity(index.len());
    for (id, media) in index {
        match load_media(source, &media) {
            Ok(img) => {
                out.insert(id, img);
            }
            Err(e) => log::warn!("读取 WPS 内嵌图片失败（{}）：{e}", media),
        }
    }
    Ok(out)
}

/// 从单元格文本中提取 DISPIMG 的图片 ID：`=DISPIMG("ID_ABC",1)` → `ID_ABC`
pub fn parse_dispimg_id(cell_text: &str) -> Option<String> {
    let pos = cell_text.find("DISPIMG(")?;
    let rest = &cell_text[pos + "DISPIMG(".len()..];
    let first = rest.find('"')?;
    let inner = &rest[first + 1..];
    let end = inner.find('"')?;
    let id = &inner[..end];
    if id.is_empty() {
        None
    } else {
        Some(id.to_string())
    }
}

// ───────────────────────────── XML 解析细节 ─────────────────────────────

/// 解析 drawing XML：返回 `(from_row, from_col, rId, alt文本)`
///
/// 只认第一个锚点元素里的 `from/col|row`（`to` 里也有 col/row，必须区分），
/// 文本采集靠状态机而不是嵌套 `read_text`：嵌套读取会让外层状态被后续
/// 兄弟元素的 `End` 事件清掉（这正是旧实现里"标准锚点图片读不出来"的原因）。
#[allow(clippy::type_complexity)]
fn parse_drawing_anchors(xml: &[u8]) -> Result<Vec<(u32, u16, String, Option<String>)>> {
    #[derive(Clone, Copy, PartialEq, Eq)]
    enum Cap {
        Col,
        Row,
    }

    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::with_capacity(8 * 1024);

    let mut in_from = false;
    let mut cap: Option<Cap> = None;
    let mut from_col: Option<u16> = None;
    let mut from_row: Option<u32> = None;
    let mut pending: Option<(u32, u16)> = None;
    let mut pending_name: Option<String> = None;
    let mut results: Vec<(u32, u16, String, Option<String>)> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Err(e) => return Err(ExcelError::Xml(format!("drawing 解析失败：{e}"))),
            Ok(Event::Start(e)) => match e.local_name().as_ref() {
                b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor" => {
                    pending = None;
                    pending_name = None;
                    from_col = None;
                    from_row = None;
                    in_from = false;
                    cap = None;
                }
                b"from" => {
                    from_col = None;
                    from_row = None;
                    in_from = true;
                }
                b"col" if in_from => cap = Some(Cap::Col),
                b"row" if in_from => cap = Some(Cap::Row),
                // `descr` 是图片 alt 文本：写侧用它保存原文件名
                b"cNvPr" => {
                    if let Some(name) = descr_of(&e) {
                        pending_name = Some(name);
                    }
                }
                b"blip" => {
                    if let (Some((r, c)), Some(rid)) = (pending, blip_r_id(&e)) {
                        results.push((r, c, rid, pending_name.clone()));
                    }
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"cNvPr" => {
                    if let Some(name) = descr_of(&e) {
                        pending_name = Some(name);
                    }
                }
                b"blip" => {
                    if let (Some((r, c)), Some(rid)) = (pending, blip_r_id(&e)) {
                        results.push((r, c, rid, pending_name.clone()));
                    }
                }
                _ => {}
            },
            Ok(Event::Text(t)) => match cap {
                Some(Cap::Col) => {
                    from_col = t.unescape().ok().and_then(|s| s.trim().parse::<u16>().ok());
                }
                Some(Cap::Row) => {
                    from_row = t.unescape().ok().and_then(|s| s.trim().parse::<u32>().ok());
                }
                None => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                b"from" => {
                    in_from = false;
                    if let (Some(c), Some(r)) = (from_col, from_row) {
                        pending = Some((r, c));
                    }
                }
                b"col" | b"row" => cap = None,
                _ => {}
            },
            _ => {}
        }
        buf.clear();
    }
    Ok(results)
}
/// `<cNvPr descr="main.png"/>` → `main.png`（图片 alt 文本，写侧用来保住原文件名）
fn descr_of(e: &BytesStart<'_>) -> Option<String> {
    for a in e.attributes().flatten() {
        if !is_local_attr(a.key.as_ref(), b"descr") {
            continue;
        }
        let v = String::from_utf8_lossy(&a.value).to_string();
        if !v.is_empty() {
            return Some(v);
        }
    }
    None
}

/// `<blip r:embed="rId1"/>` → `rId1`
fn blip_r_id(e: &BytesStart<'_>) -> Option<String> {
    for a in e.attributes().flatten() {
        let key = a.key.as_ref();
        if key == b"embed" || key.ends_with(b":embed") {
            return Some(String::from_utf8_lossy(&a.value).to_string());
        }
    }
    None
}

/// `xl/cellimages.xml`：`<etc:cellImage>` → `(cNvPr@name, blip@r:embed)`
fn parse_cellimage_ids(xml: &[u8]) -> Result<Vec<(String, String)>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut cur_id: Option<String> = None;
    let mut cur_rid: Option<String> = None;
    let mut out: Vec<(String, String)> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Err(e) => return Err(ExcelError::Xml(format!("cellimages 解析失败：{e}"))),
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                b"cellImage" => {
                    cur_id = None;
                    cur_rid = None;
                }
                b"cNvPr" => {
                    cur_id = e
                        .attributes()
                        .flatten()
                        .find(|a| is_local_attr(a.key.as_ref(), b"name"))
                        .map(|a| String::from_utf8_lossy(&a.value).to_string());
                }
                b"blip" => {
                    cur_rid = e
                        .attributes()
                        .flatten()
                        .find(|a| is_local_attr(a.key.as_ref(), b"embed"))
                        .map(|a| String::from_utf8_lossy(&a.value).to_string());
                }
                _ => {}
            },
            Ok(Event::End(e)) if e.local_name().as_ref() == b"cellImage" => {
                if let (Some(id), Some(rid)) = (cur_id.take(), cur_rid.take()) {
                    out.push((id, rid));
                }
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

/// `.rels` → `(Id → Target)`
fn parse_rels(xml: &[u8]) -> HashMap<String, String> {
    let mut map = HashMap::new();
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
                    } else if k == b"Target" || k.ends_with(b":Target") {
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

fn is_local_attr(key: &[u8], local: &[u8]) -> bool {
    key == local || key.ends_with(local)
}

/// 折叠 `..` / `.` 的相对路径为 zip 内绝对路径
pub(crate) fn normalize_path(p: &str) -> String {
    let p = p.strip_prefix('/').unwrap_or(p);
    let mut parts: Vec<&str> = Vec::new();
    for seg in p.split('/') {
        match seg {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            s => parts.push(s),
        }
    }
    parts.join("/")
}

/// 取 basename（含扩展名）：`xl/drawings/drawing1.xml` → `drawing1.xml`
///
/// 注意 rels 文件名是「完整文件名 + .rels」（`drawing1.xml.rels`），
/// 少了扩展名就永远找不到关系表（图片/超链接都会静默丢失）。
pub(crate) fn file_name(p: &str) -> String {
    p.rsplit('/').next().unwrap_or(p).to_string()
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    /// 构造最小 WPS DISPIMG 结构的 xlsx
    fn wps_xlsx_bytes() -> Vec<u8> {
        let cellimages = br#"<etc:cellImages xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData"><etc:cellImage><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="ID_TEST0001" descr="x"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic></etc:cellImage></etc:cellImages>"#;
        let rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>"#;
        let media: &[u8] = b"\x89PNG fake-bytes";

        let mut buf = Vec::new();
        let mut w = zip::ZipWriter::new(std::io::Cursor::new(&mut buf));
        let opts = zip::write::SimpleFileOptions::default();
        w.start_file("xl/cellimages.xml", opts).unwrap();
        w.write_all(cellimages).unwrap();
        w.start_file("xl/_rels/cellimages.xml.rels", opts).unwrap();
        w.write_all(rels).unwrap();
        w.start_file("xl/media/image1.png", opts).unwrap();
        w.write_all(media).unwrap();
        w.finish().unwrap();
        buf
    }

    #[test]
    fn test_parse_drawing_with_prefixes() {
        // 与 rust_xlsxwriter / Excel 真实输出同构的 drawing XML
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>76200</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:row>2</xdr:row></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="609600" y="266700"/><a:ext cx="609600" cy="609600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#;
        let cells = parse_drawing_anchors(xml.as_bytes()).unwrap();
        assert_eq!(cells, vec![(1u32, 1u16, "rId1".to_string(), None)]);
    }

    #[test]
    fn test_parse_drawing_keeps_alt_text() {
        // 写侧 set_alt_text 会落到 cNvPr@descr；读侧要把它带回来当图片原名
        let xml = r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1" descr="main.png"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId7"/></xdr:blipFill></xdr:pic></xdr:twoCellAnchor></xdr:wsDr>"#;
        let cells = parse_drawing_anchors(xml.as_bytes()).unwrap();
        assert_eq!(
            cells,
            vec![(2u32, 5u16, "rId7".to_string(), Some("main.png".to_string()))]
        );
    }

    #[test]
    fn test_rels_with_prefixes() {
        let rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#;
        let map = parse_rels(rels);
        assert_eq!(
            map.get("rId1").map(|s| s.as_str()),
            Some("../drawings/drawing1.xml")
        );
    }

    #[test]
    fn test_extract_wps_cellimages() {
        let map = extract_wps_cellimages(&wps_xlsx_bytes()).unwrap();
        let img = map.get("ID_TEST0001").expect("DISPIMG ID 映射缺失");
        assert_eq!(img.ext, "png");
        assert_eq!(img.bytes, b"\x89PNG fake-bytes");
    }

    #[test]
    fn test_parse_dispimg_id() {
        let s = r#"=DISPIMG("ID_FACFC7A1F6FC4324AB81109E79BFF87A",1)"#;
        assert_eq!(
            parse_dispimg_id(s).as_deref(),
            Some("ID_FACFC7A1F6FC4324AB81109E79BFF87A")
        );
        assert_eq!(
            parse_dispimg_id(r#"=DISPIMG("ID_X",2)"#).as_deref(),
            Some("ID_X")
        );
        assert_eq!(parse_dispimg_id("普通文件名.png"), None);
    }

    #[test]
    fn test_normalize_path() {
        assert_eq!(
            normalize_path("xl/drawings/../media/a.png"),
            "xl/media/a.png"
        );
        assert_eq!(
            normalize_path("xl/drawings/.././media/a.png"),
            "xl/media/a.png"
        );
    }
}
