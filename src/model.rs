//! 行模型与声明式映射
//!
//! 业务侧只需要：`#[derive(ExcelRow)] struct XxxRow { ... }`（或手写 [`ExcelRow`]），
//! 工具就能按表头把 Excel 行映射成结构体，并反向写出。

use serde::Serialize;

use crate::column::ColumnDef;
use crate::header::{HeaderMap, excel_column_name};
use crate::value::{Cell, CellImage, CellValue, FromCell, IntoCell};

/// 行内图片：导入时按"行 + 列"对应到具体单元格
#[derive(Debug, Clone, PartialEq)]
pub struct RowImage {
    /// 所在列（0 基）
    pub column: u16,
    /// 该列对应的表头（已按声明列归一，取不到时用文件表头）
    pub header: String,
    pub bytes: Vec<u8>,
    /// 扩展名（小写，不带点）
    pub ext: String,
    /// 原始名 / 占位名
    pub name: Option<String>,
}

impl RowImage {
    pub fn mime(&self) -> &'static str {
        crate::value::mime_from_ext(&self.ext)
    }

    pub fn image(&self) -> CellImage {
        CellImage {
            bytes: self.bytes.clone(),
            ext: self.ext.clone(),
            name: self.name.clone(),
        }
    }

    /// 识别到的图片数量
    pub fn len(&self) -> usize {
        self.bytes.len()
    }

    pub fn is_empty(&self) -> bool {
        self.bytes.is_empty()
    }
}

/// 单元格错误（列级）
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ColumnError {
    /// 模型列名（标准表头）
    pub column: String,
    /// 文件里的实际表头（缺失时为空）
    pub file_header: String,
    pub message: String,
}

impl ColumnError {
    pub fn new(column: impl Into<String>, message: impl Into<String>) -> Self {
        Self {
            column: column.into(),
            file_header: String::new(),
            message: message.into(),
        }
    }
}

/// 一行原始数据（解析后、映射前）
pub struct RowData<'a> {
    /// Excel 行号（1 基）
    pub row_index: u32,
    /// 工作表名
    pub sheet: &'a str,
    /// 表头解析结果（含列序对应关系）
    pub header: &'a HeaderMap,
    /// 本行单元格（可能稀疏：空单元格不出现）
    pub cells: &'a [Cell],
    /// 本行图片（仅当开启读图）
    pub images: &'a [RowImage],
}

impl<'a> RowData<'a> {
    /// 按列索引取单元格
    pub fn cell_at(&self, col: u16) -> Option<&Cell> {
        self.cells.iter().find(|c| c.col == col)
    }

    /// 按模型列名（或别名）取单元格
    pub fn cell_of(&self, header: &str) -> Option<&Cell> {
        let col = self.header.index_of(header)?;
        self.cell_at(col)
    }

    pub fn value_of(&self, header: &str) -> Option<&CellValue> {
        self.cell_of(header).map(|c| &c.value)
    }

    pub fn text_of(&self, header: &str) -> String {
        self.cell_of(header).map(|c| c.text()).unwrap_or_default()
    }

    /// 取强类型值（列缺失或空值都会报错，用于必填字段）
    pub fn get<T: FromCell>(&self, header: &str) -> Result<T, String> {
        match self.cell_of(header) {
            Some(c) => T::from_cell(c),
            None => Err(format!("缺少列「{}」", header)),
        }
    }

    /// 取可选值（列缺失或空值都返回 `None`）
    pub fn opt<T: FromCell>(&self, header: &str) -> Result<Option<T>, String> {
        match self.cell_of(header) {
            Some(c) if !c.value.is_empty() => T::from_cell(c).map(Some),
            _ => Ok(None),
        }
    }

    /// 取字符串（保持"列缺失"与"空值"都为空串）
    pub fn string(&self, header: &str) -> Option<String> {
        let cell = self.cell_of(header)?;
        if cell.value.is_empty() {
            None
        } else {
            Some(cell.text())
        }
    }

    /// 该列对应的行内图片（按列筛）
    pub fn images_of(&self, header: &str) -> Vec<&RowImage> {
        let Some(col) = self.header.index_of(header) else {
            return Vec::new();
        };
        self.images.iter().filter(|i| i.column == col).collect()
    }

    /// 第 n 张图片（图片列：图片1/图片2… 按列升序）
    pub fn image_at(&self, n: usize) -> Option<&RowImage> {
        self.images.get(n)
    }

    /// 整行是否为空（含图片）
    pub fn is_empty(&self) -> bool {
        if !self.images.is_empty() {
            return false;
        }
        self.cells.iter().all(|c| c.value.is_empty())
    }

    /// 把本行的单元格按文件列序转换为文本（错误报告 / 预览用）
    pub fn to_text_row(&self) -> Vec<String> {
        (0..self.header.file_headers.len() as u16)
            .map(|col| self.cell_at(col).map(|c| c.text()).unwrap_or_default())
            .collect()
    }

    /// 未在表头出现但模型声明过的列（用于区分"缺列"与"空值"）
    pub fn missing_columns(&self) -> Vec<String> {
        self.header
            .columns
            .iter()
            .filter(|c| !c.is_matched())
            .map(|c| c.header.clone())
            .collect()
    }
}

/// 一行数据的业务模型
///
/// 通常用 `#[derive(ExcelRow)]` 生成实现，也可以手写以处理非常规逻辑。
pub trait ExcelRow: Sized + Send + 'static {
    /// 列定义（表头 / 别名 / 必填 / 类型 / 图片列）
    fn columns() -> Vec<ColumnDef>;

    /// 解析一行；返回逐列错误（非空表示该行失败）
    fn from_row(row: &RowData<'_>) -> Result<Self, Vec<ColumnError>>;

    /// 写出一行（与 [`ExcelRow::columns`] 顺序对齐）
    fn to_row(&self) -> Vec<CellValue>;

    /// 该模型默认落到哪个 sheet（多 sheet 导出时使用）
    fn sheet_name() -> Option<&'static str> {
        None
    }

    /// 行级业务校验（在解析成功后调用，错误会记入该行）
    fn validate(&self) -> Result<(), Vec<String>> {
        Ok(())
    }
}

/// 完全动态的行：不预声明列，文件有什么列收什么列（"动态导入表格头适配"）
#[derive(Debug, Clone, Default)]
pub struct DynamicRow {
    /// (表头, 值)，按文件列序
    pub entries: Vec<(String, CellValue)>,
}

impl DynamicRow {
    pub fn get(&self, header: &str) -> Option<&CellValue> {
        self.entries
            .iter()
            .find(|(h, _)| h == header)
            .map(|(_, v)| v)
    }

    pub fn text(&self, header: &str) -> String {
        self.get(header).map(|v| v.to_text()).unwrap_or_default()
    }

    pub fn len(&self) -> usize {
        self.entries.len()
    }

    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    pub fn headers(&self) -> Vec<String> {
        self.entries.iter().map(|(h, _)| h.clone()).collect()
    }
}

impl ExcelRow for DynamicRow {
    fn columns() -> Vec<ColumnDef> {
        // 空声明 = 动态模式（不校验必填列，所有文件列都收下）
        Vec::new()
    }

    fn from_row(row: &RowData<'_>) -> Result<Self, Vec<ColumnError>> {
        let entries = row
            .header
            .file_headers
            .iter()
            .enumerate()
            .map(|(i, h)| {
                let v = row
                    .cell_at(i as u16)
                    .map(|c| c.value.clone())
                    .unwrap_or(CellValue::Empty);
                (h.clone(), v)
            })
            .collect();
        Ok(DynamicRow { entries })
    }

    fn to_row(&self) -> Vec<CellValue> {
        self.entries.iter().map(|(_, v)| v.clone()).collect()
    }
}

/// 便捷：把一组单元格值转成模型（写库前手工转换用）
pub fn values_of<T: IntoCell>(items: &[T]) -> Vec<CellValue> {
    items.iter().map(|v| v.into_cell()).collect()
}

/// 列名 → Excel 列号（错误提示里定位到具体列）
pub fn excel_column_of(index: u16) -> String {
    excel_column_name(index)
}
