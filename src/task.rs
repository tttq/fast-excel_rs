//! 异步导出任务：注册即函数，不再需要「执行器」概念。
//!
//! 业务侧只用 [`crate::export_task!`] 注册一个异步函数即可：
//! - 行数据导出：业务返回 `Vec<Vec<String>>`，引擎负责写 xlsx；
//! - 文件导出：业务直接把文件写到 `ctx.out_dir`，引擎负责收尾（行数 / 大小 / MIME）。
//!
//! 注册通过 `inventory` 链接期收集，零初始化、无需在 `main` 里聚合。

use std::fmt::Display;
use std::future::Future;
use std::path::PathBuf;
use std::pin::Pin;

use crate::column::ColumnDef;
use crate::value::CellValue;
use crate::write::{ExcelWriter, SheetOptions};

/// xlsx MIME。
pub const XLSX_MIME: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/// 异步导出处理函数返回的 future。
pub type ExportTaskFuture =
    Pin<Box<dyn Future<Output = Result<ExportTaskOutput, String>> + Send + 'static>>;

/// 异步导出任务上下文（由导出中心构建）。
#[derive(Debug, Clone)]
pub struct ExportTaskContext {
    /// 任务编号（业务可用于生成文件名）。
    pub task_no: String,
    /// 创建任务时保存的筛选参数（JSON 对象，业务自行解析）。
    pub query: serde_json::Value,
    /// 输出目录；文件必须落在此目录下。
    pub out_dir: PathBuf,
}

/// 异步导出产物。
#[derive(Debug, Clone)]
pub struct ExportTaskOutput {
    /// 数据行数（不含表头）。
    pub total_rows: i32,
    /// 文件名（仅文件名，位于 `ctx.out_dir` 下）。
    pub file_name: String,
    pub file_size: i64,
    /// 文件 MIME；`None` 时由导出中心按扩展名推断。
    pub mime: Option<String>,
}

/// 任务形态：决定引擎如何收尾。
#[derive(Debug, Clone, Copy)]
pub enum ExportTaskKind {
    /// 业务只提供表头 + 行数据，引擎写 xlsx。
    Rows {
        sheet_name: &'static str,
        headers: &'static [&'static str],
    },
    /// 业务自行生成文件，引擎只读取文件信息。
    File { mime: &'static str },
}

/// 导出处理函数（宏展开生成，业务无感）。
pub type ExportTaskFn = fn(ExportTaskContext) -> ExportTaskFuture;

/// 一条异步导出任务定义（`inventory` 链接期收集项）。
pub struct ExportTaskEntry {
    pub task_type: &'static str,
    pub kind: ExportTaskKind,
    pub run: ExportTaskFn,
}

inventory::collect!(ExportTaskEntry);

/// 已注册的全部任务。
pub fn registered_export_tasks() -> Vec<&'static ExportTaskEntry> {
    inventory::iter::<ExportTaskEntry>().into_iter().collect()
}

/// 按 `task_type` 查找任务。
pub fn export_task_entry(task_type: &str) -> Option<&'static ExportTaskEntry> {
    inventory::iter::<ExportTaskEntry>()
        .into_iter()
        .find(|e| e.task_type == task_type)
}

/// 是否已注册该 `task_type`。
pub fn export_task_registered(task_type: &str) -> bool {
    export_task_entry(task_type).is_some()
}

/// 直接执行已注册的异步导出任务；未注册返回错误。
pub async fn run_export_task(
    task_type: &str,
    ctx: ExportTaskContext,
) -> Result<ExportTaskOutput, String> {
    let entry = export_task_entry(task_type)
        .ok_or_else(|| format!("导出类型「{task_type}」未注册"))?;
    (entry.run)(ctx).await
}

/// 行数据任务的通用实现：业务回调 -> 引擎写 xlsx。
///
/// 业务错误类型只需实现 [`Display`]；宏会自动完成错误转换。
pub async fn run_rows_task<F, Fut, E>(
    ctx: ExportTaskContext,
    task_type: &'static str,
    sheet_name: &'static str,
    headers: &'static [&'static str],
    provider: F,
) -> Result<ExportTaskOutput, String>
where
    F: FnOnce(ExportTaskContext) -> Fut + Send + 'static,
    Fut: Future<Output = Result<Vec<Vec<String>>, E>> + Send + 'static,
    E: Display,
{
    let rows = provider(ctx.clone()).await.map_err(|e| e.to_string())?;
    let total_rows = rows.len() as i32;
    let file_name = format!("{}-{}.xlsx", ctx.task_no, task_type);
    let path = ctx.out_dir.join(&file_name);
    std::fs::create_dir_all(&ctx.out_dir)
        .map_err(|e| format!("创建导出目录失败: {e}"))?;

    // 写盘是 CPU/IO 阻塞操作，放到阻塞线程；常量内存模式下行数再多也不涨内存。
    let sheet = SheetOptions::new(sheet_name).columns(
        headers
            .iter()
            .map(|h| ColumnDef::new(*h))
            .collect::<Vec<_>>(),
    );
    let write_path = path.clone();
    let file_size = tokio::task::spawn_blocking(move || -> Result<u64, String> {
        let mut writer = ExcelWriter::new();
        {
            let mut ws = writer
                .add_sheet(&sheet)
                .map_err(|e| format!("创建工作表失败: {e}"))?;
            ws.write_headers()
                .map_err(|e| format!("写表头失败: {e}"))?;
            for row in &rows {
                let cells: Vec<CellValue> = row.iter().map(|v| CellValue::Text(v.clone())).collect();
                ws.write_row(&cells)
                    .map_err(|e| format!("写数据行失败: {e}"))?;
            }
            ws.finish().map_err(|e| format!("收尾工作表失败: {e}"))?;
        }
        writer
            .save(&write_path)
            .map_err(|e| format!("保存导出文件失败: {e}"))
    })
    .await
    .map_err(|e| format!("导出写入线程异常: {e}"))??;

    Ok(ExportTaskOutput {
        total_rows,
        file_name,
        file_size: file_size as i64,
        mime: Some(XLSX_MIME.to_string()),
    })
}

/// 文件任务的通用实现：业务保证文件已写入 `ctx.out_dir`。
pub async fn run_file_task<F, Fut, E>(
    ctx: ExportTaskContext,
    mime: &'static str,
    provider: F,
) -> Result<ExportTaskOutput, String>
where
    F: FnOnce(ExportTaskContext) -> Fut + Send + 'static,
    Fut: Future<Output = Result<(i64, String), E>> + Send + 'static,
    E: Display,
{
    let (total_rows, file_name) = provider(ctx.clone()).await.map_err(|e| e.to_string())?;
    let file_size = std::fs::metadata(ctx.out_dir.join(&file_name))
        .map_err(|e| format!("读取导出文件大小失败: {e}"))?
        .len();
    Ok(ExportTaskOutput {
        total_rows: total_rows as i32,
        file_name,
        file_size: file_size as i64,
        mime: Some(mime.to_string()),
    })
}
