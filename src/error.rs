//! 工具内部统一错误类型
//!
//! 设计原则：库内部不依赖 Web 层，所有错误先收敛成 [`ExcelError`]。
//! 业务侧需要时再自行实现 `From<ExcelError> for AppError` 等映射
//! （[`ExcelError::code`] 给出 `excel_xxx` key，可拼成 `@excel_xxx` 消息便于前端 i18n）。

use std::fmt;

/// Excel 工具错误
#[derive(Debug, Clone)]
pub enum ExcelError {
    /// 文件 / IO 错误
    Io(String),
    /// xlsx 容器（zip）错误
    Zip(String),
    /// XML 解析错误
    Xml(String),
    /// 不是合法的 xlsx 文件
    NotXlsx(String),
    /// 找不到指定工作表
    SheetNotFound(String),
    /// 表头缺少必填列（动态表头适配失败）
    MissingColumns(Vec<String>),
    /// 单元格级错误
    InvalidCell {
        row: u32,
        column: String,
        message: String,
    },
    /// 行级错误（同一行多列错误）
    InvalidRow { row: u32, messages: Vec<String> },
    /// 不支持的能力（例如加密 xlsx、非 deflate 压缩）
    Unsupported(String),
    /// 批量落库失败
    Sink(String),
    /// 其它
    Other(String),
}

impl ExcelError {
    /// 错误码（前端 i18n key 后缀，与项目 `@xxx` 消息约定一致）
    pub fn code(&self) -> &'static str {
        match self {
            ExcelError::Io(_) => "excel_io_failed",
            ExcelError::Zip(_) => "excel_zip_failed",
            ExcelError::Xml(_) => "excel_xml_failed",
            ExcelError::NotXlsx(_) => "excel_not_xlsx",
            ExcelError::SheetNotFound(_) => "excel_sheet_not_found",
            ExcelError::MissingColumns(_) => "excel_missing_column",
            ExcelError::InvalidCell { .. } => "excel_invalid_cell",
            ExcelError::InvalidRow { .. } => "excel_invalid_row",
            ExcelError::Unsupported(_) => "excel_unsupported",
            ExcelError::Sink(_) => "excel_sink_failed",
            ExcelError::Other(_) => "excel_error",
        }
    }

    pub fn other(msg: impl Into<String>) -> Self {
        ExcelError::Other(msg.into())
    }

    pub fn unsupported(msg: impl Into<String>) -> Self {
        ExcelError::Unsupported(msg.into())
    }
}

impl fmt::Display for ExcelError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            ExcelError::Io(m) => write!(f, "文件读写失败：{m}"),
            ExcelError::Zip(m) => write!(f, "Excel 文件损坏或格式不支持：{m}"),
            ExcelError::Xml(m) => write!(f, "Excel 内部结构解析失败：{m}"),
            ExcelError::NotXlsx(m) => write!(f, "不是有效的 xlsx 文件：{m}"),
            ExcelError::SheetNotFound(m) => write!(f, "工作表不存在：{m}"),
            ExcelError::MissingColumns(cols) => write!(f, "缺少必填列：{}", cols.join("、")),
            ExcelError::InvalidCell {
                row,
                column,
                message,
            } => {
                write!(f, "第 {row} 行「{column}」列：{message}")
            }
            ExcelError::InvalidRow { row, messages } => {
                write!(f, "第 {row} 行：{}", messages.join("；"))
            }
            ExcelError::Unsupported(m) => write!(f, "暂不支持的 Excel 特性：{m}"),
            ExcelError::Sink(m) => write!(f, "数据写入失败：{m}"),
            ExcelError::Other(m) => write!(f, "{m}"),
        }
    }
}

impl std::error::Error for ExcelError {}

impl From<std::io::Error> for ExcelError {
    fn from(e: std::io::Error) -> Self {
        ExcelError::Io(e.to_string())
    }
}

impl From<zip::result::ZipError> for ExcelError {
    fn from(e: zip::result::ZipError) -> Self {
        ExcelError::Zip(e.to_string())
    }
}

impl From<quick_xml::Error> for ExcelError {
    fn from(e: quick_xml::Error) -> Self {
        ExcelError::Xml(e.to_string())
    }
}

impl From<rust_xlsxwriter::XlsxError> for ExcelError {
    fn from(e: rust_xlsxwriter::XlsxError) -> Self {
        ExcelError::Other(format!("Excel 生成失败：{e}"))
    }
}

impl From<serde_json::Error> for ExcelError {
    fn from(e: serde_json::Error) -> Self {
        ExcelError::Other(format!("JSON 序列化失败：{e}"))
    }
}

impl From<chrono::ParseError> for ExcelError {
    fn from(e: chrono::ParseError) -> Self {
        ExcelError::Other(format!("日期解析失败：{e}"))
    }
}

/// 便捷别名
pub type Result<T> = std::result::Result<T, ExcelError>;
