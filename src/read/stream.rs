//! 工作表行流：SAX 式逐行解析（百万行常量内存）
//!
//! 与 `calamine` 的 `worksheet_range`（一次性把整表读成二维数组）不同，这里
//! 直接用 `quick-xml` 事件流扫 `xl/worksheets/sheetN.xml`，每行解析完立刻交给
//! 调用方处理：内存占用与**列数**相关，与行数无关。

use std::collections::HashMap;
use std::io::BufReader;
use std::sync::Arc;

use chrono::NaiveTime;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

use crate::column::ColumnDef;
use crate::error::Result;
use crate::header::{HeaderAnalysis, HeaderMap, resolve_headers};
use crate::model::{RowData, RowImage};
use crate::read::ReadOptions;
use crate::read::images::{
    AnchorImage, AnchorIndex, load_media, load_sheet_anchors, load_wps_index, parse_dispimg_id,
};
use crate::read::shared_strings::{SharedStrings, decode_text};
use crate::read::source::{EntryReader, ZipSource};
use crate::read::styles::Styles;
use crate::value::{Cell, CellValue, parse_datetime_text, serial_to_datetime};

/// 当前正在采集的文本目标
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Capture {
    None,
    Value,
    Formula,
    Inline,
}

/// 工作表行流
pub struct SheetStream {
    xml: Reader<BufReader<EntryReader>>,
    buf: Vec<u8>,
    source: ZipSource,
    sheet_name: String,
    date_1904: bool,
    shared: Arc<SharedStrings>,
    styles: Arc<Styles>,
    defs: Vec<ColumnDef>,
    options: ReadOptions,

    // 图片
    anchors_by_row: HashMap<u32, Vec<(u16, AnchorImage)>>,
    wps_index: Option<HashMap<String, String>>,
    pending_dispimg: Vec<(u16, String)>,

    // 当前行
    current_row: u32,
    last_row: u32,
    cells: Vec<Cell>,
    images: Vec<RowImage>,
    cell_active: bool,
    cell_col: u16,
    cell_last_col: u16,
    cell_type: Vec<u8>,
    cell_style: u32,
    value_buf: String,
    formula_buf: String,
    inline_buf: String,
    capture: Capture,
    in_inline_str: bool,

    // 表头
    header: Option<HeaderMap>,
    header_cells: Vec<Vec<String>>,
    header_done: bool,
    /// 实际需要读图的文件列（None = 全部；Some(空) = 不读）
    image_cols: Option<Vec<u16>>,

    // 统计
    pub rows_scanned: u64,
    pub rows_yielded: u64,
    /// 跳过的空行数
    pub skipped_rows: u64,
    /// 是否因 max_rows 提前停止（预览时判断还有更多行）
    limit_hit: bool,
    finished: bool,
}

impl SheetStream {
    #[allow(clippy::too_many_arguments)]
    pub(crate) fn new(
        source: ZipSource,
        sheet_path: &str,
        sheet_name: &str,
        date_1904: bool,
        shared: Arc<SharedStrings>,
        styles: Arc<Styles>,
        defs: &[ColumnDef],
        options: ReadOptions,
    ) -> Result<Self> {
        let entry = source.entry_reader(sheet_path)?;
        let xml = Reader::from_reader(BufReader::with_capacity(256 * 1024, entry));

        // 图片锚点索引：只在需要读图时建立（索引很小；图片字节仍按需读）
        let mut anchors_by_row: HashMap<u32, Vec<(u16, AnchorImage)>> = HashMap::new();
        if options.read_images {
            let anchors: AnchorIndex = load_sheet_anchors(&source, sheet_path)?;
            for ((row, col), anchor) in anchors {
                anchors_by_row.entry(row).or_default().push((col, anchor));
            }
            for list in anchors_by_row.values_mut() {
                list.sort_by_key(|(c, _)| *c);
            }
        }
        Ok(Self {
            xml,
            buf: Vec::with_capacity(64 * 1024),
            source,
            sheet_name: sheet_name.to_string(),
            date_1904,
            shared,
            styles,
            defs: defs.to_vec(),
            options,
            anchors_by_row,
            wps_index: None,
            pending_dispimg: Vec::new(),
            current_row: 0,
            last_row: 0,
            cells: Vec::new(),
            images: Vec::new(),
            cell_active: false,
            cell_col: 0,
            cell_last_col: 0,
            cell_type: Vec::new(),
            cell_style: 0,
            value_buf: String::new(),
            formula_buf: String::new(),
            inline_buf: String::new(),
            capture: Capture::None,
            in_inline_str: false,
            header: None,
            header_cells: Vec::new(),
            header_done: false,
            image_cols: None,
            rows_scanned: 0,
            rows_yielded: 0,
            skipped_rows: 0,
            limit_hit: false,
            finished: false,
        })
    }

    // ───────────────────────────── 对外接口 ─────────────────────────────

    /// 表头解析结果（进入数据行后才有值）
    pub fn header(&self) -> Option<&HeaderMap> {
        self.header.as_ref()
    }

    /// 表头自适应报告（列匹配情况 / 缺失必填列 / 多出来的列）
    pub fn header_analysis(&self) -> Option<HeaderAnalysis> {
        self.header.as_ref().map(|h| h.analysis())
    }

    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    pub fn is_finished(&self) -> bool {
        self.finished
    }

    /// 是否因 max_rows 提前停止（预览用）
    pub fn limit_hit(&self) -> bool {
        self.limit_hit
    }

    /// 取下一行数据；`None` = 读完
    pub fn next_row(&mut self) -> Result<Option<RowData<'_>>> {
        if self.finished {
            return Ok(None);
        }
        if let Some(max) = self.options.max_rows {
            if self.rows_yielded >= max {
                self.limit_hit = true;
                self.finished = true;
                return Ok(None);
            }
        }

        loop {
            if !self.read_next_row_element()? {
                if !self.header_done {
                    self.build_header();
                }
                self.finished = true;
                return Ok(None);
            }

            self.rows_scanned += 1;
            let row_index = self.current_row;

            if !self.header_done {
                let start = self.options.header.row_index + 1;
                let end = start + self.options.header.row_span.max(1) - 1;
                if row_index < start {
                    continue;
                }
                if row_index <= end {
                    self.collect_header_row(row_index - start);
                    continue;
                }
                self.build_header();
            }

            self.collect_row_images(row_index)?;
            let empty = self.cells.iter().all(|c| c.value.is_empty());
            if empty && self.images.is_empty() && self.options.skip_empty_rows {
                self.skipped_rows += 1;
                continue;
            }
            if !self.header_done {
                self.build_header();
            }
            self.rows_yielded += 1;
            let header = self.header.as_ref().expect("表头已构建");
            return Ok(Some(RowData {
                row_index,
                sheet: &self.sheet_name,
                header,
                cells: &self.cells,
                images: &self.images,
            }));
        }
    }

    // ───────────────────────────── 行解析 ─────────────────────────────

    /// 读到一行结束（`</row>`）为止
    fn read_next_row_element(&mut self) -> Result<bool> {
        self.cells.clear();
        self.images.clear();
        let mut row_seen = false;

        loop {
            self.buf.clear();
            let event = match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Eof) => return Ok(row_seen),
                Ok(ev) => ev,
                Err(e) => return Err(e.into()),
            };

            let mut begin: Option<(Option<u16>, Vec<u8>, u32)> = None;
            let mut cell_end = false;
            let mut row_end = false;

            match event {
                Event::Start(ref e) => match e.local_name().as_ref() {
                    b"row" => {
                        row_seen = true;
                        self.current_row = row_index_of(e).unwrap_or(self.last_row + 1);
                        self.last_row = self.current_row;
                        self.cell_last_col = 0;
                        self.cells.clear();
                    }
                    b"c" => begin = Some(parse_cell_attrs(e)),
                    b"v" => self.capture = Capture::Value,
                    b"f" => self.capture = Capture::Formula,
                    b"is" => self.in_inline_str = true,
                    b"t" if self.in_inline_str => self.capture = Capture::Inline,
                    _ => {}
                },
                Event::Empty(ref e) => match e.local_name().as_ref() {
                    b"row" => {
                        row_seen = true;
                        self.current_row = row_index_of(e).unwrap_or(self.last_row + 1);
                        self.last_row = self.current_row;
                        self.cell_last_col = 0;
                        row_end = true;
                    }
                    b"c" => {
                        begin = Some(parse_cell_attrs(e));
                        cell_end = true;
                    }
                    _ => {}
                },
                Event::End(ref e) => match e.local_name().as_ref() {
                    b"row" => row_end = true,
                    b"c" => cell_end = true,
                    b"v" | b"f" => self.capture = Capture::None,
                    b"is" => {
                        self.in_inline_str = false;
                        self.capture = Capture::None;
                    }
                    b"t" if self.in_inline_str => self.capture = Capture::None,
                    _ => {}
                },
                Event::Text(ref t) => match self.capture {
                    Capture::Value => self.value_buf.push_str(&decode_text(t)),
                    Capture::Formula => self.formula_buf.push_str(&decode_text(t)),
                    Capture::Inline => self.inline_buf.push_str(&decode_text(t)),
                    Capture::None => {}
                },
                Event::CData(ref t) => match self.capture {
                    Capture::Value => self
                        .value_buf
                        .push_str(&String::from_utf8_lossy(t.as_ref())),
                    Capture::Formula => self
                        .formula_buf
                        .push_str(&String::from_utf8_lossy(t.as_ref())),
                    Capture::Inline => self
                        .inline_buf
                        .push_str(&String::from_utf8_lossy(t.as_ref())),
                    Capture::None => {}
                },
                _ => {}
            }

            if let Some((col, cell_type, style)) = begin {
                self.begin_cell(col, cell_type, style);
            }
            if cell_end {
                self.finish_cell();
            }
            if row_end {
                return Ok(true);
            }
        }
    }

    fn begin_cell(&mut self, col: Option<u16>, cell_type: Vec<u8>, style: u32) {
        self.cell_active = true;
        self.cell_type = cell_type;
        self.cell_style = style;
        self.cell_col = col.unwrap_or_else(|| {
            if self.cells.is_empty() {
                0
            } else {
                self.cell_last_col + 1
            }
        });
        self.cell_last_col = self.cell_col;
        self.value_buf.clear();
        self.formula_buf.clear();
        self.inline_buf.clear();
        self.capture = Capture::None;
    }

    fn finish_cell(&mut self) {
        if !self.cell_active {
            return;
        }
        self.cell_active = false;

        let raw_value = std::mem::take(&mut self.value_buf);
        let raw_inline = std::mem::take(&mut self.inline_buf);
        let formula_raw = std::mem::take(&mut self.formula_buf);
        let formula = if formula_raw.is_empty() {
            None
        } else {
            Some(formula_raw)
        };
        let cell_type = self.cell_type.clone();
        let style = self.cell_style;

        let mut value = match cell_type.as_slice() {
            b"s" => match raw_value.trim().parse::<usize>() {
                Ok(idx) => CellValue::Text(self.shared.get(idx).unwrap_or_default().to_string()),
                Err(_) => CellValue::Text(String::new()),
            },
            b"inlineStr" => CellValue::Text(raw_inline),
            b"str" => CellValue::Text(raw_value),
            b"b" => CellValue::Bool(matches!(raw_value.trim(), "1" | "true" | "TRUE")),
            b"e" => CellValue::Error(raw_value),
            b"d" => match parse_datetime_text(&raw_value) {
                Some(dt) => CellValue::DateTime(dt),
                None => CellValue::Text(raw_value),
            },
            _ => {
                let s = raw_value.trim();
                if s.is_empty() {
                    match &formula {
                        Some(f) => CellValue::Formula(f.clone()),
                        None => CellValue::Empty,
                    }
                } else {
                    match s.parse::<f64>() {
                        Ok(n) => self.number_value(n, style),
                        Err(_) => CellValue::Text(s.to_string()),
                    }
                }
            }
        };

        // WPS「单元格嵌入图片」：单元格是 DISPIMG 公式 → 记下 ID，按行 + 列取图
        if let Some(f) = &formula {
            if let Some(id) = parse_dispimg_id(f) {
                self.pending_dispimg.push((self.cell_col, id));
                value = CellValue::Text(f.clone());
            }
        }

        if let CellValue::Text(s) = &mut value {
            if self.options.trim_text {
                let trimmed = s.trim();
                if trimmed.len() != s.len() {
                    *s = trimmed.to_string();
                }
            }
        }

        self.cells.push(Cell {
            row: self.current_row,
            col: self.cell_col,
            value,
            style,
        });
    }

    /// 数字单元格：命中日期样式则转成日期/日期时间
    fn number_value(&self, n: f64, style: u32) -> CellValue {
        if self.styles.is_date(style) {
            if let Some(dt) = serial_to_datetime(n, self.date_1904) {
                let midnight = NaiveTime::from_hms_opt(0, 0, 0).unwrap();
                return if dt.time() == midnight {
                    CellValue::Date(dt.date())
                } else {
                    CellValue::DateTime(dt)
                };
            }
        }
        CellValue::Number(n)
    }

    // ───────────────────────────── 表头 ─────────────────────────────

    fn collect_header_row(&mut self, span_index: u32) {
        let span = self.options.header.row_span.max(1) as usize;
        let span_index = span_index as usize;
        let mut row_cells: Vec<(u16, String)> = self
            .cells
            .iter()
            .map(|c| (c.col, c.text()))
            .filter(|(_, t)| !t.is_empty())
            .collect();
        row_cells.sort_by_key(|(c, _)| *c);
        for (col, text) in row_cells {
            let col = col as usize;
            if self.header_cells.len() <= col {
                self.header_cells.resize(col + 1, vec![String::new(); span]);
            }
            if let Some(slot) = self.header_cells[col].get_mut(span_index) {
                *slot = text;
            }
        }
    }

    fn build_header(&mut self) {
        let span = self.options.header.row_span.max(1) as usize;
        let mut file_headers: Vec<String> = Vec::with_capacity(self.header_cells.len());
        for col in &self.header_cells {
            let parts: Vec<String> = (0..span)
                .filter_map(|i| col.get(i))
                .map(|s| s.trim().to_string())
                .filter(|s| !s.is_empty())
                .collect();
            file_headers.push(parts.join(" "));
        }
        // 去掉尾部空表头（resize 出来的占位列）
        while matches!(file_headers.last(), Some(h) if h.trim().is_empty()) {
            file_headers.pop();
        }

        let map = if self.defs.is_empty() {
            HeaderMap::dynamic(file_headers)
        } else {
            resolve_headers(&file_headers, &self.defs, &self.options.header)
        };
        self.image_cols = match &self.options.image_columns {
            Some(cols) => Some(cols.clone()),
            None => {
                if self.defs.is_empty() {
                    // 动态模式：不知道哪些是图片列，全部按需读
                    None
                } else if self.defs.iter().any(|d| d.image) {
                    // 只读模型声明的图片列，避免把旁列图片也读进内存
                    Some(
                        self.defs
                            .iter()
                            .filter(|d| d.image)
                            .filter_map(|d| map.index_of(&d.header))
                            .collect(),
                    )
                } else {
                    Some(Vec::new())
                }
            }
        };
        self.header = Some(map);
        self.header_done = true;
    }

    // ───────────────────────────── 图片 ─────────────────────────────

    fn ensure_wps_index(&mut self) -> Result<()> {
        if self.wps_index.is_none() {
            let index = load_wps_index(&self.source)?;
            self.wps_index = Some(index);
        }
        Ok(())
    }

    fn collect_row_images(&mut self, row_index: u32) -> Result<()> {
        self.images.clear();
        if !self.options.read_images {
            self.pending_dispimg.clear();
            return Ok(());
        }

        // 1) 标准 drawing 锚点
        //    注意坐标口径：drawing 里的 row 是 0 基，行流的 row_index 是 1 基（与 Excel 界面一致）
        let anchor_row = row_index.saturating_sub(1);
        let mut found: Vec<(u16, String, Option<String>)> = Vec::new();
        if let Some(list) = self.anchors_by_row.get(&anchor_row) {
            for (col, anchor) in list {
                // 优先用写侧写入的 alt 文本（原文件名），其次退回媒体文件名
                let name = anchor
                    .name
                    .clone()
                    .or_else(|| anchor.media.rsplit('/').next().map(|s| s.to_string()));
                found.push((*col, anchor.media.clone(), name));
            }
        }
        // 2) WPS DISPIMG
        if !self.pending_dispimg.is_empty() {
            let pendings: Vec<(u16, String)> = self.pending_dispimg.drain(..).collect();
            if !pendings.is_empty() {
                self.ensure_wps_index()?;
            }
            let index = self.wps_index.as_ref().expect("WPS 索引已加载");
            for (col, id) in pendings {
                if found.iter().any(|(c, _, _)| *c == col) {
                    continue;
                }
                if let Some(media) = index.get(&id) {
                    found.push((col, media.clone(), Some(format!("DISPIMG:{id}"))));
                }
            }
        }

        // 3) 列过滤 + 按需读字节
        found.sort_by_key(|(c, _, _)| *c);
        let filter = self.image_cols.clone();
        for (col, media, name) in found {
            if let Some(cols) = &filter {
                if !cols.contains(&col) {
                    continue;
                }
            }
            match load_media(&self.source, &media) {
                Ok(img) => {
                    let header = self.header_name_of(col);
                    self.images.push(RowImage {
                        column: col,
                        header,
                        bytes: img.bytes,
                        ext: img.ext,
                        name,
                    });
                }
                Err(e) => log::warn!("读取内嵌图片失败（{media}）：{e}"),
            }
        }
        Ok(())
    }

    /// 文件列 → 表头名（优先用模型声明列，其次用文件表头）
    fn header_name_of(&self, col: u16) -> String {
        if let Some(header) = &self.header {
            if let Some(c) = header.columns.iter().find(|c| c.file_index == Some(col)) {
                return c.header.clone();
            }
            if let Some(h) = header.file_headers.get(col as usize) {
                if !h.is_empty() {
                    return h.clone();
                }
            }
        }
        format!("第{}列", crate::header::excel_column_name(col))
    }
}

/// 解析 `<c r="B12" t="s" s="3">` 的属性 → (列索引, 类型, 样式)
fn parse_cell_attrs(e: &BytesStart<'_>) -> (Option<u16>, Vec<u8>, u32) {
    let mut col: Option<u16> = None;
    let mut cell_type: Vec<u8> = Vec::new();
    let mut style: u32 = 0;
    for a in e.attributes().flatten() {
        match a.key.as_ref() {
            b"r" => col = col_from_ref(&String::from_utf8_lossy(&a.value)),
            b"t" => cell_type = a.value.to_vec(),
            b"s" => {
                style = String::from_utf8_lossy(&a.value)
                    .trim()
                    .parse()
                    .unwrap_or(0)
            }
            _ => {}
        }
    }
    (col, cell_type, style)
}

/// `<row r="12">` → 12
fn row_index_of(e: &BytesStart<'_>) -> Option<u32> {
    for a in e.attributes().flatten() {
        if a.key.as_ref() == b"r" {
            return String::from_utf8_lossy(&a.value).trim().parse().ok();
        }
    }
    None
}
/// `B12` → 列索引（B → 1）
pub(crate) fn col_from_ref(r: &str) -> Option<u16> {
    let mut n: u32 = 0;
    let mut has = false;
    for c in r.chars() {
        if c.is_ascii_alphabetic() {
            has = true;
            n = n * 26 + (c.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
        } else {
            break;
        }
    }
    if has && n > 0 {
        Some((n - 1) as u16)
    } else {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_col_from_ref() {
        assert_eq!(col_from_ref("A1"), Some(0));
        assert_eq!(col_from_ref("B12"), Some(1));
        assert_eq!(col_from_ref("Z1"), Some(25));
        assert_eq!(col_from_ref("AA1"), Some(26));
        assert_eq!(col_from_ref("AB100"), Some(27));
        assert_eq!(col_from_ref("1"), None);
    }
}
