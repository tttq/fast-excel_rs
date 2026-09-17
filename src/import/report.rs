//! 导入报告 / 预览结果

use serde::Serialize;

use crate::header::HeaderAnalysis;

/// 行级错误
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct RowError {
    /// Excel 行号（1 基）
    pub row: u32,
    /// 工作表名（多 sheet 导入时区分来源）
    pub sheet: String,
    pub messages: Vec<String>,
}

/// 单个工作表的表头匹配情况
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetHeader {
    pub sheet: String,
    #[serde(flatten)]
    pub analysis: HeaderAnalysis,
}

/// 导入进度（大文件时用于前端进度条）
#[derive(Debug, Clone, Copy, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ImportProgress {
    pub rows_read: u64,
    pub rows_ok: u64,
    pub rows_error: u64,
    pub batches_done: u64,
}

/// 导入结果
#[derive(Debug, Clone, Default, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ImportReport {
    /// 实际读取的工作表
    pub sheets: Vec<String>,
    /// 表头自适应结果（每个 sheet 一份）
    pub headers: Vec<SheetHeader>,
    /// 数据行总数（不含表头 / 空行）
    pub total_rows: u64,
    /// 成功（解析 + 校验通过）
    pub success_rows: u64,
    /// 失败行
    pub error_rows: u64,
    /// 跳过的空行
    pub skipped_rows: u64,
    /// 落库批次数
    pub batches: u64,
    pub elapsed_ms: u64,
    /// 行级错误（受 `max_errors` 限制）
    pub errors: Vec<RowError>,
    /// 错误是否被截断（超过上限只保留前 N 条）
    pub errors_truncated: bool,
    /// 落库失败的批次数
    pub failed_batches: u64,
}

impl ImportReport {
    pub fn is_success(&self) -> bool {
        self.error_rows == 0 && self.failed_batches == 0
    }
}

/// 预览行
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PreviewRow<T> {
    pub row: u32,
    /// 解析成功时的模型（失败为 `None`）
    pub data: Option<T>,
    /// 原始文本（按文件列序，供前端展示"用户填了什么"）
    pub values: Vec<String>,
    /// 该行错误
    pub errors: Vec<String>,
}

/// 导入预览结果（不落库）
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ImportPreview<T> {
    pub sheet: String,
    #[serde(flatten)]
    pub header: HeaderAnalysis,
    /// 读取开关：`true` 表示还有更多行（预览只取前 N 行）
    pub has_more: bool,
    pub total_rows: u64,
    pub error_rows: u64,
    pub rows: Vec<PreviewRow<T>>,
}
