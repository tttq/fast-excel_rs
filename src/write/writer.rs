//! 流式 Excel 写出（`rust_xlsxwriter`，常量内存模式）
//!
//! - 打开 `constant_memory` 后，工作表数据边写边落临时文件，**内存与行数无关**，
//!   百万行导出的内存占用是一条行缓冲的量级；
//! - 代价：只能按行号递增写入（`rust_xlsxwriter` 会忽略对已 flush 行的写入），
//!   因此本模块用 `row_cursor` 强制顺序推进，避免"静默丢行"。

use std::collections::HashMap;
use std::path::Path;

use rust_xlsxwriter::{
    Color, ExcelDateTime, Format, FormatAlign, FormatBorder, Image, Workbook, Worksheet,
};

use crate::column::{ColumnDef, ColumnKind};
use crate::error::Result;
use crate::model::ExcelRow;
use crate::value::{CellImage, CellValue};
use crate::write::{SheetOptions, WriteOptions, default_column_width};

/// Excel 写出器（可包含多个工作表）
pub struct ExcelWriter {
    workbook: Workbook,
    sheet_count: u32,
}

impl Default for ExcelWriter {
    fn default() -> Self {
        Self::new()
    }
}

impl ExcelWriter {
    pub fn new() -> Self {
        Self {
            workbook: Workbook::new(),
            sheet_count: 0,
        }
    }

    pub fn sheet_count(&self) -> u32 {
        self.sheet_count
    }

    /// 新增一个工作表（列宽 / 表头 / 冻结窗格在写入前配置好）
    pub fn add_sheet(&mut self, options: &SheetOptions) -> Result<SheetWriter<'_>> {
        let name = if options.name.is_empty() {
            format!("Sheet{}", self.sheet_count + 1)
        } else {
            options.name.clone()
        };
        let ws = if options.write.constant_memory {
            self.workbook.add_worksheet_with_constant_memory()
        } else {
            self.workbook.add_worksheet()
        };
        self.sheet_count += 1;
        ws.set_name(&name)?;
        let mut writer = SheetWriter::new(ws, options.clone());
        writer.prepare()?;
        Ok(writer)
    }

    /// 保存到文件；返回文件字节数（大文件不会进内存）
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> Result<u64> {
        let path = path.as_ref();
        self.workbook.save(path)?;
        Ok(std::fs::metadata(path).map(|m| m.len()).unwrap_or(0))
    }

    /// 保存到内存缓冲（HTTP 响应直接下载；小数据量场景）
    pub fn save_to_buffer(&mut self) -> Result<Vec<u8>> {
        Ok(self.workbook.save_to_buffer()?)
    }

    /// 逃生口：需要合并单元格、批注、条件格式等高级能力时直接用底层工作簿
    pub fn workbook(&mut self) -> &mut Workbook {
        &mut self.workbook
    }
}

/// 单个工作表的写出句柄
pub struct SheetWriter<'a> {
    ws: &'a mut Worksheet,
    options: SheetOptions,
    /// 下一个可写行（0 基）
    row_cursor: u32,
    /// 已写入的最大列 / 行（autofilter 用）
    last_col: u16,
    /// 写入的列数（用于表头与筛选范围）
    col_count: u16,
    formats: HashMap<String, Format>,
}

impl<'a> SheetWriter<'a> {
    fn new(ws: &'a mut Worksheet, options: SheetOptions) -> Self {
        let col_count = options.columns.len() as u16;
        Self {
            ws,
            options,
            row_cursor: 0,
            last_col: col_count.saturating_sub(1),
            col_count,
            formats: HashMap::new(),
        }
    }

    /// 列宽 + 冻结窗格
    fn prepare(&mut self) -> Result<()> {
        if !self.options.columns.is_empty() {
            for (i, col) in self.options.columns.iter().enumerate() {
                let width = col.width.unwrap_or_else(|| default_column_width(col));
                self.ws.set_column_width(i as u16, width)?;
            }
            if self.options.write.freeze_header {
                self.ws.set_freeze_panes(1, 0)?;
            }
        }
        Ok(())
    }

    pub fn options(&self) -> &SheetOptions {
        &self.options
    }

    /// 当前待写行号（0 基）
    pub fn current_row(&self) -> u32 {
        self.row_cursor
    }

    /// 写表头（按列定义的 `header`）
    pub fn write_headers(&mut self) -> Result<()> {
        self.write_headers_with(&self.options.columns.clone())
    }

    /// 写表头（自定义列定义）
    pub fn write_headers_with(&mut self, columns: &[ColumnDef]) -> Result<()> {
        let row = self.row_cursor;
        let header_format = self.header_format();
        for (i, col) in columns.iter().enumerate() {
            if self.options.write.header_style {
                let fmt = header_format.clone();
                self.ws
                    .write_string_with_format(row, i as u16, &col.header, &fmt)?;
            } else {
                self.ws.write_string(row, i as u16, &col.header)?;
            }
        }
        if columns.len() as u16 > self.col_count {
            self.col_count = columns.len() as u16;
        }
        self.last_col = self.last_col.max(columns.len().saturating_sub(1) as u16);
        self.row_cursor += 1;
        Ok(())
    }

    /// 写表头（必填列红字标出，导入模板场景）
    pub fn write_headers_marked(&mut self, columns: &[ColumnDef]) -> Result<()> {
        let row = self.row_cursor;
        let header_format = self.header_format();
        let required_format = self.required_header_format();
        for (i, col) in columns.iter().enumerate() {
            let fmt = if col.required {
                &required_format
            } else {
                &header_format
            };
            self.ws
                .write_string_with_format(row, i as u16, &col.header, fmt)?;
        }
        self.last_col = self.last_col.max(columns.len().saturating_sub(1) as u16);
        self.col_count = self.col_count.max(columns.len() as u16);
        self.row_cursor += 1;
        Ok(())
    }

    fn required_header_format(&mut self) -> Format {
        let key = "header:required".to_string();
        if let Some(f) = self.formats.get(&key) {
            return f.clone();
        }
        let fmt = Format::new()
            .set_bold()
            .set_align(FormatAlign::Center)
            .set_border(FormatBorder::Thin)
            .set_background_color(Color::RGB(0xFD_E3_E3))
            .set_font_color(Color::Red);
        self.formats.insert(key, fmt.clone());
        fmt
    }

    /// 给某一列挂下拉数据验证（模板场景：`dropdown` 列）
    pub fn add_dropdown(
        &mut self,
        col: u16,
        options: &[String],
        first_row: u32,
        last_row: u32,
        prompt: Option<&str>,
    ) -> Result<()> {
        if options.is_empty() {
            return Ok(());
        }
        let list: Vec<&str> = options.iter().map(|s| s.as_str()).collect();
        let mut dv = rust_xlsxwriter::DataValidation::new().allow_list_strings(&list)?;
        if let Some(p) = prompt {
            dv = dv.set_input_message(p)?;
        }
        self.ws
            .add_data_validation(first_row, col, last_row, col, &dv)?;
        Ok(())
    }
    /// 写一行原始值（顺序与列定义对齐）
    ///
    /// 行高在写单元格**之前**设置：带图行按 `image_row_height` 撑高，
    /// `insert_image_fit_to_cell_centered` 依赖行高计算缩放比例。
    pub fn write_row(&mut self, values: &[CellValue]) -> Result<()> {
        let row = self.row_cursor;
        let has_image = values
            .iter()
            .any(|v| matches!(v, CellValue::Image(img) if !img.bytes.is_empty()));
        if has_image {
            self.ws
                .set_row_height(row, self.options.write.image_row_height)?;
        } else if let Some(h) = self.options.write.default_row_height {
            self.ws.set_row_height(row, h)?;
        }
        let columns = self.options.columns.clone();
        for (i, value) in values.iter().enumerate() {
            let def = columns.get(i);
            self.write_value(row, i as u16, value, def)?;
        }
        self.row_cursor += 1;
        self.last_col = self.last_col.max(values.len().saturating_sub(1) as u16);
        Ok(())
    }

    /// 写业务模型一行
    pub fn write_model<T: ExcelRow>(&mut self, model: &T) -> Result<()> {
        self.write_row(&model.to_row())
    }

    /// 批量写业务模型（内部仍是逐行 => 常量内存）
    pub fn write_models<T: ExcelRow>(&mut self, models: &[T]) -> Result<()> {
        for m in models {
            self.write_model(m)?;
        }
        Ok(())
    }

    /// 写任意文本行（如"填报说明"）
    pub fn write_text_row<I, S>(&mut self, texts: I) -> Result<()>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        let row = self.row_cursor;
        for (i, t) in texts.into_iter().enumerate() {
            self.ws.write_string(row, i as u16, t.as_ref())?;
        }
        self.row_cursor += 1;
        Ok(())
    }

    /// 空一行
    pub fn skip_rows(&mut self, n: u32) {
        self.row_cursor += n;
    }

    /// 结束本表：应用自动筛选（必须写完数据后调用）
    pub fn finish(&mut self) -> Result<()> {
        if self.options.write.auto_filter
            && self.options.write.header_style
            && self.row_cursor > 1
            && self.col_count > 0
        {
            self.ws
                .autofilter(0, 0, self.row_cursor - 1, self.last_col)?;
        }
        Ok(())
    }

    /// 逃生口：直接操作底层工作表（批注 / 合并 / 条件格式）
    pub fn worksheet(&mut self) -> &mut Worksheet {
        self.ws
    }

    // ─────────────────────────── 单元格写入 ───────────────────────────

    fn write_value(
        &mut self,
        row: u32,
        col: u16,
        value: &CellValue,
        def: Option<&ColumnDef>,
    ) -> Result<()> {
        match value {
            CellValue::Empty => {}
            CellValue::Text(s) => {
                if s.is_empty() {
                    return Ok(());
                }
                match self.cell_format_key(def, true) {
                    Some(key) => {
                        let fmt = self.format_of(&key, def, true);
                        self.ws.write_string_with_format(row, col, s, &fmt)?;
                    }
                    None => {
                        self.ws.write_string(row, col, s)?;
                    }
                }
            }
            CellValue::Number(n) => match self.cell_format_key(def, false) {
                Some(key) => {
                    let fmt = self.format_of(&key, def, false);
                    self.ws.write_number_with_format(row, col, *n, &fmt)?;
                }
                None => {
                    self.ws.write_number(row, col, *n)?;
                }
            },
            CellValue::Bool(b) => {
                self.ws.write_boolean(row, col, *b)?;
            }
            CellValue::Date(d) => {
                let fmt = self.date_format(def, false);
                let dt = ExcelDateTime::from_ymd(
                    chrono::Datelike::year(d) as u16,
                    chrono::Datelike::month(d) as u8,
                    chrono::Datelike::day(d) as u8,
                )?;
                self.ws.write_datetime_with_format(row, col, &dt, &fmt)?;
            }
            CellValue::DateTime(raw) => {
                let fmt = self.date_format(def, true);
                let dt = ExcelDateTime::from_ymd(
                    chrono::Datelike::year(raw) as u16,
                    chrono::Datelike::month(raw) as u8,
                    chrono::Datelike::day(raw) as u8,
                )?
                .and_hms(
                    chrono::Timelike::hour(raw) as u16,
                    chrono::Timelike::minute(raw) as u8,
                    chrono::Timelike::second(raw) as f64,
                )?;
                self.ws.write_datetime_with_format(row, col, &dt, &fmt)?;
            }
            CellValue::Formula(f) => {
                self.ws.write_formula(row, col, f.as_str())?;
            }
            CellValue::Error(e) => {
                self.ws.write_string(row, col, e)?;
            }
            CellValue::Image(img) => {
                self.write_image(row, col, img)?;
            }
        }
        Ok(())
    }

    fn write_image(&mut self, row: u32, col: u16, img: &CellImage) -> Result<()> {
        if img.bytes.is_empty() {
            return Ok(());
        }
        let mut image = Image::new_from_buffer(&img.bytes)?;
        if let Some(name) = &img.name {
            image = image.set_alt_text(name.clone());
        }
        if self.options.write.cell_image {
            // Excel「置于单元格内」图片：随单元格排序/筛选，但常量内存模式下
            // rust_xlsxwriter 0.99 会 panic（内部全局图片索引此时还没准备好）
            self.ws.embed_image(row, col, &image)?;
        } else {
            // 标准锚点图片：按行高自适应缩放 + 居中，常量内存模式下可用
            self.ws
                .insert_image_fit_to_cell_centered(row, col, &image)?;
        }
        Ok(())
    }

    // ─────────────────────────── 格式缓存 ───────────────────────────

    fn header_format(&mut self) -> Format {
        let key = "header".to_string();
        if let Some(f) = self.formats.get(&key) {
            return f.clone();
        }
        let fmt = Format::new()
            .set_bold()
            .set_align(FormatAlign::Center)
            .set_border(FormatBorder::Thin)
            .set_background_color(Color::RGB(0xD9_E1_F2))
            .set_font_color(Color::RGB(0x1F_23_29));
        self.formats.insert(key, fmt.clone());
        fmt
    }

    /// 该列是否需要显式格式；返回缓存 key
    fn cell_format_key(&self, def: Option<&ColumnDef>, text: bool) -> Option<String> {
        let def = def?;
        if text && def.text_format {
            return Some("text:@".to_string());
        }
        if !text {
            if let Some(f) = &def.format {
                return Some(format!("num:{f}"));
            }
            if def.text_format {
                return Some("text:@".to_string());
            }
        }
        None
    }

    fn format_of(&mut self, key: &str, def: Option<&ColumnDef>, text: bool) -> Format {
        if let Some(f) = self.formats.get(key) {
            return f.clone();
        }
        let code = match def.and_then(|d| d.format.clone()) {
            Some(code) => code,
            None => {
                if matches!(def.map(|d| d.kind), Some(ColumnKind::Integer)) {
                    "0".to_string()
                } else if matches!(
                    def.map(|d| d.kind),
                    Some(ColumnKind::Decimal) | Some(ColumnKind::Number)
                ) {
                    "0.00".to_string()
                } else if text {
                    "@".to_string()
                } else {
                    "General".to_string()
                }
            }
        };
        let fmt = Format::new().set_num_format(code);
        self.formats.insert(key.to_string(), fmt.clone());
        fmt
    }

    fn date_format(&mut self, def: Option<&ColumnDef>, with_time: bool) -> Format {
        let default = if with_time {
            self.options.write.datetime_format.clone()
        } else {
            self.options.write.date_format.clone()
        };
        let code = def.and_then(|d| d.format.clone()).unwrap_or(default);
        let key = format!("date:{code}");
        if let Some(f) = self.formats.get(&key) {
            return f.clone();
        }
        let fmt = Format::new().set_num_format(code);
        self.formats.insert(key, fmt.clone());
        fmt
    }
}

/// 便捷：一步导出单表（表头 + 行 + 收尾）
pub fn export_single_sheet<T: ExcelRow>(
    path: impl AsRef<Path>,
    sheet: &SheetOptions,
    rows: &[T],
) -> Result<u64> {
    let mut writer = ExcelWriter::new();
    {
        let mut ws = writer.add_sheet(sheet)?;
        ws.write_headers()?;
        ws.write_models(rows)?;
        ws.finish()?;
    }
    writer.save(path)
}

/// 便捷：把写出选项转成 rust_xlsxwriter 的通用行高（模板用）
pub fn default_write_options() -> WriteOptions {
    WriteOptions::default()
}
