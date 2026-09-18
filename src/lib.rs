//! # excel — 通用 Excel 导入导出工具
//!
//! 对标 Java 生态的 EasyExcel，面向"百万行 + 批量落库 + 单元格图片 + 多 sheet +
//! 动态表头"的真实业务场景。整个工具只依赖纯 Rust 库（无 JVM / 无 Office 组件）。
//!
//! ## 能力总览
//!
//! | 能力 | 实现方式 |
//! |---|---|
//! | 百万行读取 | `quick-xml` SAX 事件流逐行解析，内存与行数无关（[`SheetStream`]） |
//! | 百万行写出 | `rust_xlsxwriter` 常量内存工作表，边写边落临时文件 |
//! | 批量落库 | [`ImportRunner::commit`]：`spawn_blocking` 解析 + 有界通道 + 分批事务（[`BatchSink`]） |
//! | 单元格图片 | 标准 drawing 锚点 + WPS `DISPIMG`；按行对应，图片字节按需读 |
//! | 多 sheet | 读：`SheetSelector::{Name,Names,All}`；写：[`ExcelWriter::add_sheet`] / [`ExportRunner::export_multi_sheet`] |
//! | 动态表头 | 别名 / 模糊匹配 / 归一化（全角、空格、`*` 必填标记）/ 列序回退（[`HeaderOptions`]） |
//! | 注解式映射 | `#[derive(ExcelRow)]`（`excel-macros`） |
//! | 执行器 | `#[derive(ExcelExecutor)]` 自动注册执行器工厂（[`inventory`](https://docs.rs/inventory) 静态收集，零初始化）：[`executor_for::<T>`] / [`executor_by_name`] |
//! | 导入模板 | 必填红字表头 + 示例行 + 下拉 + 说明 sheet（[`TemplateSpec`]） |
//! | 导出 | 游标分批拉取 + 流式写文件（[`ExportRunner`]）；CSV 兜底（[`write::write_csv`]） |
//!
//! 完整用法与百万行调优见 crate 根目录 `README.md`。
//!
//! ## 最小示例（声明式）
//!
//! ```no_run
//! use std::sync::Arc;
//! use excel::{BatchSink, ExcelError, ExcelReader, ExcelRow, ImportRunner, SheetSelector, ZipSource};
//!
//! #[derive(ExcelRow, Debug)]
//! #[excel(sheet = "产品导入")]
//! struct ProductRow {
//!     #[excel(header = "工厂名称", alias = ["工厂", "厂名"], required)]
//!     factory_name: String,
//!     #[excel(header = "品名", required)]
//!     name: String,
//!     #[excel(header = "MOQ")]
//!     moq: Option<i32>,
//! }
//!
//! # async fn demo(db_sink: Arc<dyn BatchSink<ProductRow>>) -> Result<(), ExcelError> {
//! let report = ImportRunner::new()
//!     .sheets(SheetSelector::All)      // 多 sheet 合并导入
//!     .batch_size(500)                 // 每 500 行一批
//!     .commit(ZipSource::open("products.xlsx")?, db_sink)
//!     .await?;
//! println!("成功 {} 行，失败 {} 行", report.success_rows, report.error_rows);
//! # Ok(()) }
//! ```
//!
//! ## 落库实现约定
//!
//! [`BatchSink::save`] 拿到的是一批已解析、已校验的行，事务边界由业务侧决定
//! （推荐"一批一个事务 + `insert_many`"）。项目内 sea-orm 实体用
//! `Entity::insert_many_with_fill(models, &tx)` 即可同时拿到主键生成、审计字段
//! 自动填充与租户注入。

// 让派生宏在 crate 内部（单元测试 / 文档测试）也能用 `::excel::` 路径
extern crate self as excel;

pub mod column;
pub mod error;
pub mod executor;
pub mod export;
pub mod header;
pub mod import;
pub mod model;
pub mod read;
pub mod task;
pub mod value;
pub mod write;

pub use column::{ColumnDef, ColumnKind};
pub use error::{ExcelError, Result};
pub use executor::{
    ErasedExecutor, ExcelExecutor, Executor, ExecutorEntry, executor_by_name, executor_entry,
    executor_for, preview_by_name, registered_executors, registered_names,
};
pub use export::{ExportRunner, ExportStats, Page};
pub use header::{
    HeaderAnalysis, HeaderMap, HeaderOptions, MatchKind, ResolvedColumn, excel_column_name,
    resolve_headers,
};
pub use import::{
    BatchConfig, BatchSink, FnSink, ImportPreview, ImportProgress, ImportReport, ImportRunner,
    NoopSink, PreviewRow, ProgressFn, RowError, SheetHeader,
};
pub use model::{ColumnError, DynamicRow, ExcelRow, RowData, RowImage};
pub use read::{
    CellImage, EntryReader, ExcelReader, ReadOptions, SheetInfo, SheetSelector, SheetStream,
    ZipSource, extract_all_images, extract_embedded_images, extract_images_of_sheet,
    extract_wps_cellimages, parse_dispimg_id,
};
pub use task::{
    ExportTaskContext, ExportTaskEntry, ExportTaskFn, ExportTaskFuture, ExportTaskKind,
    ExportTaskOutput, XLSX_MIME, export_task_entry, export_task_registered, registered_export_tasks,
    run_export_task, run_file_task, run_rows_task,
};
pub use value::{Cell, CellValue, FromCell, IntoCell};
pub use write::{
    ExcelWriter, ReferenceSheet, SheetOptions, SheetWriter, TemplateSpec, WriteOptions,
    build_template, write_csv,
};

/// `#[derive(ExcelRow)]` 派生宏（与 [`ExcelRow`] trait 同名，不同命名空间）
pub use excel_macros::ExcelRow;

/// `#[derive(ExcelExecutor)]` 派生宏：注册执行器工厂（与 [`ExcelExecutor`] trait 同名）
pub use excel_macros::ExcelExecutor;

/// 派生宏生成的 `inventory::submit!` 通过此路径引用；通常无需直接使用
#[doc(hidden)]
pub use ::inventory;
/// 注册一个异步导出任务（唯一宏：行导出 / 文件导出都用它）。
///
/// 展开后直接向 `inventory` 提交任务定义，链接期即完成注册：
/// **没有执行器结构体、没有 `install()`、不需要在 `main.rs` 聚合**。
/// 导出中心按 `task_type` 查到函数后直接调用。
///
/// # 形态一：行数据导出（业务只查数据，引擎写 xlsx）
/// ```ignore
/// excel::export_task! {
///     task_type = "factory",
///     sheet_name = "工厂数据",
///     headers = FACTORY_HEADERS,
///     rows = factory_export_rows, // async fn(ExportTaskContext) -> Result<Vec<Vec<String>>, E>
/// }
/// ```
///
/// # 形态二：文件导出（业务自带流式生成，如带图 xlsx）
/// ```ignore
/// excel::export_task! {
///     task_type = "product",
///     mime = excel::XLSX_MIME,
///     file = product_export_file, // async fn(ExportTaskContext) -> Result<(i64, String), E>
/// }
/// ```
///
/// `rows` / `file` 也接受非捕获闭包（如 `|ctx| rows_for(kind, ctx)`），
/// 方便一个模块用同一份逻辑注册多张表。
#[macro_export]
macro_rules! export_task {
    (
        task_type = $task:expr,
        sheet_name = $sheet:expr,
        headers = $headers:expr,
        rows = $rows:expr
        $(,)?
    ) => {
        const _: () = {
            fn __excel_export_task(
                ctx: $crate::ExportTaskContext,
            ) -> $crate::ExportTaskFuture {
                ::std::boxed::Box::pin($crate::run_rows_task(
                    ctx,
                    $task,
                    $sheet,
                    $headers,
                    $rows,
                ))
            }

            $crate::inventory::submit! {
                $crate::ExportTaskEntry {
                    task_type: $task,
                    kind: $crate::ExportTaskKind::Rows {
                        sheet_name: $sheet,
                        headers: $headers,
                    },
                    run: __excel_export_task,
                }
            }
        };
    };
    (
        task_type = $task:expr,
        mime = $mime:expr,
        file = $file:expr
        $(,)?
    ) => {
        const _: () = {
            fn __excel_export_task(
                ctx: $crate::ExportTaskContext,
            ) -> $crate::ExportTaskFuture {
                ::std::boxed::Box::pin($crate::run_file_task(ctx, $mime, $file))
            }

            $crate::inventory::submit! {
                $crate::ExportTaskEntry {
                    task_type: $task,
                    kind: $crate::ExportTaskKind::File { mime: $mime },
                    run: __excel_export_task,
                }
            }
        };
    };
}
