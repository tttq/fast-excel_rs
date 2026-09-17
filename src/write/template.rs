//! 导入模板生成（表头 + 必填标红 + 示例行 + 下拉 + 说明 / 引用 sheet）
//!
//! 与导出共用同一套列定义，所以"模板 / 导出 / 导入校验"三处的表头永远一致——
//! 这是旧实现里最容易走偏的地方（模板改了、导入没改，或者反过来）。

use std::path::Path;

use crate::column::ColumnDef;
use crate::error::Result;
use crate::value::CellValue;
use crate::write::{ExcelWriter, SheetOptions, WriteOptions};

/// 引用 sheet（如"可用工厂列表"，供填写时对照）
#[derive(Debug, Clone, Default)]
pub struct ReferenceSheet {
    pub name: String,
    /// 表头
    pub headers: Vec<String>,
    /// 数据行
    pub rows: Vec<Vec<String>>,
    /// 表格上方的说明
    pub notes: Vec<String>,
}

impl ReferenceSheet {
    pub fn new(name: impl Into<String>, headers: Vec<String>) -> Self {
        Self {
            name: name.into(),
            headers,
            rows: Vec::new(),
            notes: Vec::new(),
        }
    }

    pub fn rows(mut self, rows: Vec<Vec<String>>) -> Self {
        self.rows = rows;
        self
    }

    pub fn note(mut self, note: impl Into<String>) -> Self {
        self.notes.push(note.into());
        self
    }
}

/// 模板定义
#[derive(Debug, Clone)]
pub struct TemplateSpec {
    /// 数据 sheet 名（默认 `导入模板`）
    pub sheet_name: String,
    /// 可选标题行（写在表头之前；写了标题行则表头在第 2 行）
    pub title: Option<String>,
    pub columns: Vec<ColumnDef>,
    /// 填写说明（写到"填写说明" sheet）
    pub notes: Vec<String>,
    /// 示例行（写到表头下方，提交前删除或替换）
    pub sample_rows: Vec<Vec<CellValue>>,
    /// 引用 sheet
    pub reference_sheets: Vec<ReferenceSheet>,
    pub write: WriteOptions,
    /// 下拉验证生效的最大行号（默认 1000）
    pub dropdown_rows: u32,
}

impl TemplateSpec {
    pub fn new(columns: Vec<ColumnDef>) -> Self {
        Self {
            sheet_name: "导入模板".to_string(),
            title: None,
            columns,
            notes: Vec::new(),
            sample_rows: Vec::new(),
            reference_sheets: Vec::new(),
            write: WriteOptions::default().in_memory(),
            dropdown_rows: 1000,
        }
    }

    pub fn sheet_name(mut self, name: impl Into<String>) -> Self {
        self.sheet_name = name.into();
        self
    }

    pub fn title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(title.into());
        self
    }

    pub fn note(mut self, note: impl Into<String>) -> Self {
        self.notes.push(note.into());
        self
    }

    pub fn notes<I, S>(mut self, notes: I) -> Self
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        self.notes.extend(notes.into_iter().map(Into::into));
        self
    }

    pub fn sample_row(mut self, row: Vec<CellValue>) -> Self {
        self.sample_rows.push(row);
        self
    }

    pub fn reference(mut self, sheet: ReferenceSheet) -> Self {
        self.reference_sheets.push(sheet);
        self
    }

    pub fn write_options(mut self, write: WriteOptions) -> Self {
        self.write = write;
        self
    }

    pub fn dropdown_rows(mut self, rows: u32) -> Self {
        self.dropdown_rows = rows;
        self
    }

    /// 表头行号（0 基）：有标题行时为 1
    pub fn header_row(&self) -> u32 {
        if self.title.is_some() { 1 } else { 0 }
    }
}

/// 生成模板字节
pub fn build_template(spec: &TemplateSpec) -> Result<Vec<u8>> {
    let mut writer = ExcelWriter::new();
    let header_row = spec.header_row();
    let columns = spec.columns.clone();

    {
        let sheet_options = SheetOptions::new(&spec.sheet_name)
            .columns(columns.clone())
            .write_options(spec.write.clone());
        let mut ws = writer.add_sheet(&sheet_options)?;
        if let Some(title) = &spec.title {
            ws.write_text_row([title.clone()])?;
        }
        ws.write_headers_marked(&columns)?;
        for row in &spec.sample_rows {
            ws.write_row(row)?;
        }
        // 下拉验证（模板场景数据量小，直接给到 dropdown_rows 行）
        let first = header_row + 1;
        let last = header_row + spec.dropdown_rows;
        for (i, col) in columns.iter().enumerate() {
            if !col.dropdown.is_empty() {
                let prompt = col.note.clone();
                ws.add_dropdown(i as u16, &col.dropdown, first, last, prompt.as_deref())?;
            }
        }
        ws.finish()?;
    }

    // 填写说明 sheet
    {
        let mut ws = writer.add_sheet(&SheetOptions::new("填写说明"))?;
        let notes = build_notes(spec);
        for line in &notes {
            ws.write_text_row([line.clone()])?;
        }
    }

    // 引用 sheet
    for reference in &spec.reference_sheets {
        let mut ws = writer.add_sheet(&SheetOptions::new(&reference.name))?;
        for line in &reference.notes {
            ws.write_text_row([line.clone()])?;
        }
        ws.write_text_row(reference.headers.clone())?;
        for row in &reference.rows {
            ws.write_text_row(row.clone())?;
        }
    }

    writer.save_to_buffer()
}

/// 生成模板文件
pub fn write_template_file<P: AsRef<Path>>(path: P, spec: &TemplateSpec) -> Result<u64> {
    let bytes = build_template(spec)?;
    std::fs::write(path.as_ref(), &bytes)?;
    Ok(bytes.len() as u64)
}

fn build_notes(spec: &TemplateSpec) -> Vec<String> {
    let mut notes = Vec::new();
    notes.push(format!("填写说明（数据 sheet：{}）", spec.sheet_name));
    notes.push(String::new());

    let required: Vec<String> = spec
        .columns
        .iter()
        .filter(|c| c.required)
        .map(|c| c.header.clone())
        .collect();
    if required.is_empty() {
        notes.push("1. 表头列均为选填。".to_string());
    } else {
        notes.push(format!("1. 红色表头为必填列：{}。", required.join("、")));
    }

    let image_cols: Vec<String> = spec
        .columns
        .iter()
        .filter(|c| c.image)
        .map(|c| c.header.clone())
        .collect();
    if !image_cols.is_empty() {
        notes.push(format!(
            "2. 图片列（{}）：把图片直接插入/粘贴到对应单元格，导入时按行自动对应，不需要打包 zip。",
            image_cols.join("、")
        ));
    }

    let mut idx = 3;
    if !spec.sample_rows.is_empty() {
        notes.push(format!(
            "{idx}. 第 {} 行是示例数据，正式导入前请删除或整行替换。",
            spec.header_row() + 2
        ));
        idx += 1;
    }
    for line in &spec.notes {
        notes.push(format!("{idx}. {line}"));
        idx += 1;
    }

    notes.push(String::new());
    notes.push("列说明：".to_string());
    for col in &spec.columns {
        let mut desc = vec![col.header.clone()];
        if col.required {
            desc.push("必填".to_string());
        }
        desc.push(col.kind.as_str().to_string());
        if let Some(note) = &col.note {
            desc.push(note.clone());
        }
        if !col.dropdown.is_empty() {
            desc.push(format!("可选值：{}", col.dropdown.join("/")));
        }
        notes.push(format!("- {}", desc.join(" | ")));
    }
    notes
}
