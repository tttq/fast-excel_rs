//! 写侧：流式导出（多 sheet / 单元格图片 / 模板生成）

pub mod csv;
pub mod template;
pub mod writer;

pub use csv::write_csv;
pub use template::{ReferenceSheet, TemplateSpec, build_template};
pub use writer::{ExcelWriter, SheetWriter};

use crate::column::ColumnDef;

/// 写出选项
#[derive(Debug, Clone)]
pub struct WriteOptions {
    /// 表头样式（加粗 + 底色 + 居中）
    pub header_style: bool,
    /// 冻结表头行
    pub freeze_header: bool,
    /// 自动筛选
    pub auto_filter: bool,
    /// 默认行高
    pub default_row_height: Option<f64>,
    /// 含图片的行高（图片嵌入后按行高缩放）
    pub image_row_height: f64,
    /// 日期默认格式
    pub date_format: String,
    /// 日期时间默认格式
    pub datetime_format: String,
    /// 常量内存模式（**百万行必须开启**；开启后只能按行号递增写入）
    pub constant_memory: bool,
    /// 是否写 Excel「置于单元格内」图片（`embed_image`）
    ///
    /// 默认写标准锚点图片（`insert_image`）：常量内存模式下可用、且能被
    /// 本工具的读取端（drawing 锚点）与 Excel / WPS 正常识别。
    /// 注意：`embed_image` 在 `rust_xlsxwriter` 常量内存模式下会报错，
    /// 且生成的是 `xl/richData/*`（Excel 内嵌图片）；请用 [`WriteOptions::cell_image`]
    /// 开启——它会自动关闭常量内存；直接改字段的话请配合 `constant_memory = false`。
    pub cell_image: bool,
}

impl Default for WriteOptions {
    fn default() -> Self {
        Self {
            header_style: true,
            freeze_header: true,
            auto_filter: true,
            default_row_height: None,
            image_row_height: 60.0,
            date_format: "yyyy-mm-dd".to_string(),
            datetime_format: "yyyy-mm-dd hh:mm:ss".to_string(),
            constant_memory: true,
            cell_image: false,
        }
    }
}

impl WriteOptions {
    pub fn new() -> Self {
        Self::default()
    }

    /// 小数据量 / 需要随机写（合并单元格、绝对定位图片）时关闭常量内存
    pub fn in_memory(mut self) -> Self {
        self.constant_memory = false;
        self
    }

    /// 写 Excel 内嵌单元格图片（`embed_image`）
    ///
    /// 内嵌图片在 `rust_xlsxwriter` 常量内存模式下支持不完整（会直接报错），
    /// 因此这里**自动关闭常量内存**（等价于 `in_memory()`），只需调用本方法即可。
    pub fn cell_image(mut self) -> Self {
        self.cell_image = true;
        self.constant_memory = false;
        self
    }

    pub fn without_header_style(mut self) -> Self {
        self.header_style = false;
        self
    }

    pub fn auto_filter(mut self, enable: bool) -> Self {
        self.auto_filter = enable;
        self
    }

    pub fn freeze_header(mut self, enable: bool) -> Self {
        self.freeze_header = enable;
        self
    }

    pub fn image_row_height(mut self, height: f64) -> Self {
        self.image_row_height = height;
        self
    }

    pub fn date_format(mut self, format: impl Into<String>) -> Self {
        self.date_format = format.into();
        self
    }

    pub fn datetime_format(mut self, format: impl Into<String>) -> Self {
        self.datetime_format = format.into();
        self
    }
}

/// 单个工作表的配置
#[derive(Debug, Clone)]
pub struct SheetOptions {
    pub name: String,
    /// 列定义（表头、别名、宽度、格式、图片列）
    pub columns: Vec<ColumnDef>,
    pub write: WriteOptions,
}

impl SheetOptions {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            columns: Vec::new(),
            write: WriteOptions::default(),
        }
    }

    pub fn columns<I>(mut self, columns: I) -> Self
    where
        I: IntoIterator<Item = ColumnDef>,
    {
        self.columns = columns.into_iter().collect();
        self
    }

    pub fn write_options(mut self, write: WriteOptions) -> Self {
        self.write = write;
        self
    }
}

impl Default for SheetOptions {
    fn default() -> Self {
        Self::new("Sheet1")
    }
}

/// 列默认宽度（按语义类型 + 表头长度估算）
pub(crate) fn default_column_width(col: &ColumnDef) -> f64 {
    let base: f64 = match col.kind {
        crate::column::ColumnKind::DateTime => 20.0,
        crate::column::ColumnKind::Date => 13.0,
        crate::column::ColumnKind::Image => 16.0,
        crate::column::ColumnKind::Text => 16.0,
        _ => 12.0,
    };
    let header_width = col.header.chars().count() as f64 + 3.0;
    base.max(header_width)
}
