//! 执行器：注册一次，全链路通用（预览 / 导入 / 导出 / 模板）
//!
//! 业务端只需在行模型上挂 `#[derive(ExcelExecutor)]`（配合 `#[excel(register = "...")]`），
//! 本模块会自动生成该模型的执行器工厂并**静态注册进全局表**（`inventory` 链接期收集，
//! 无需任何启动初始化代码），随后按类型或按注册名统一调用：
//!
//! ```no_run
//! use serde::Serialize;
//! use excel::{ExcelExecutor, ExcelRow, ZipSource};
//!
//! #[derive(ExcelRow, ExcelExecutor, Serialize)]
//! #[excel(sheet = "产品导入", register = "product")]
//! struct ProductRow {
//!     #[excel(header = "品名", required)]
//!     name: String,
//! }
//!
//! # async fn demo(sink: std::sync::Arc<dyn excel::BatchSink<ProductRow>>) -> Result<(), excel::ExcelError> {
//! let source = ZipSource::open("products.xlsx")?;
//!
//! // ① 按类型调用（编译期保证类型安全）
//! let exec = excel::executor_for::<ProductRow>();
//! let preview = exec.preview(source.clone(), 200)?;
//! let report = exec.commit(source.clone(), sink).await?;
//! let bytes = exec.export_bytes(&products())?;
//!
//! // ② 按注册名字符串调度（Web 层按请求参数路由）
//! let erased = excel::executor_by_name("product").expect("已注册");
//! let json = erased.preview(source, 200)?;
//! # Ok(()) }
//! # fn products() -> Vec<ProductRow> { vec![] }
//! ```

use std::any::type_name;
use std::fmt;
use std::future::Future;
use std::marker::PhantomData;
use std::path::Path;
use std::sync::Arc;

use serde::Serialize;

use crate::column::ColumnDef;
use crate::error::{ExcelError, Result};
use crate::export::{ExportRunner, ExportStats, Page};
use crate::import::{BatchSink, ImportPreview, ImportReport, ImportRunner};
use crate::model::{ExcelRow, RowData};
use crate::read::{SheetSelector, ZipSource};
use crate::write::{SheetOptions, TemplateSpec};

/// 行模型注册契约：`#[derive(ExcelExecutor)]` 生成实现
///
/// 通常不手写；实现该 trait 后即可用 [`executor_for::<T>`] 拿到通用执行器。
pub trait ExcelExecutor: ExcelRow {
    /// 全局注册名（`#[excel(register = "...")]` 指定）
    const EXECUTOR_NAME: &'static str;
}

/// 按类型获取注册模型的通用执行器（自动创建，业务端无需管理工厂）
pub fn executor_for<T: ExcelExecutor>() -> Executor<T> {
    Executor::new(T::EXECUTOR_NAME, T::sheet_name())
}

/// 一个行模型的通用执行器：预览 / 导入 / 导出 / 模板的统一入口
///
/// 所有方法都基于模型自带的列定义与表名自动构造 runner，
/// 需要微调时再通过 [`Executor::import_runner`] / [`Executor::export_runner`] 拿到
/// 可配置的 runner（默认值已经可用，一般不需要）。
#[derive(Clone)]
pub struct Executor<T> {
    /// 注册名（仅用于标识 / 按名调度）
    name: &'static str,
    /// 默认工作表名（来自 `#[excel(sheet = "...")]` 或 `ExcelRow::sheet_name`）
    sheet: Option<&'static str>,
    columns: Vec<ColumnDef>,
    _marker: PhantomData<fn() -> T>,
}

impl<T: ExcelRow> Executor<T> {
    /// 手动构造（未注册模型也能用；注册模型请用 [`executor_for`]）
    pub fn new(name: &'static str, sheet: Option<&'static str>) -> Self {
        Self {
            name,
            sheet,
            columns: T::columns(),
            _marker: PhantomData,
        }
    }

    pub fn name(&self) -> &'static str {
        self.name
    }

    pub fn type_name(&self) -> &'static str {
        type_name::<T>()
    }

    pub fn sheet_name(&self) -> Option<&'static str> {
        self.sheet
    }

    pub fn columns(&self) -> &[ColumnDef] {
        &self.columns
    }

    /// 模型名不足时的兜底工作表名
    fn default_sheet(&self) -> String {
        self.sheet.unwrap_or("Sheet1").to_string()
    }

    /// 导入 runner（默认 sheets：单模型 = 按 `sheet` 名，多模型 = `All`）
    pub fn import_runner(&self) -> ImportRunner {
        let sheets = match self.sheet {
            Some(name) => SheetSelector::Name(name.to_string()),
            None => SheetSelector::All,
        };
        ImportRunner::new().sheets(sheets)
    }

    /// 导出 runner（列宽 / 格式 / 图片行高来自列定义）
    pub fn export_runner(&self) -> ExportRunner {
        ExportRunner::new(
            SheetOptions::new(self.default_sheet()).columns(self.columns.clone()),
        )
    }

    /// 导出 runner（显式指定工作表名）
    pub fn export_runner_for(&self, sheet_name: impl Into<String>) -> ExportRunner {
        ExportRunner::new(
            SheetOptions::new(sheet_name).columns(self.columns.clone()),
        )
    }

    /// 导入模板（填入示例行 / 引用列表后交给 [`build_template`](crate::write::build_template)）
    pub fn template(&self) -> TemplateSpec {
        TemplateSpec::new(self.columns.clone()).sheet_name(self.default_sheet())
    }
}

impl<T: ExcelRow> Executor<T> {
    /// 校验预览（不落库，只读前 `limit` 行）
    pub fn preview(
        &self,
        source: ZipSource,
        limit: usize,
    ) -> Result<ImportPreview<T>>
    where
        T: Serialize,
    {
        self.import_runner().preview::<T>(source, limit)
    }

    /// 预览转 JSON（按名调度 / HTTP 响应直接返回用）
    pub fn preview_json(
        &self,
        source: ZipSource,
        limit: usize,
    ) -> Result<serde_json::Value>
    where
        T: Serialize,
    {
        let preview = self.preview(source, limit)?;
        serde_json::to_value(&preview).map_err(ExcelError::from)
    }

    /// 正式导入：流式解析 + 分批落库（解析在阻塞线程，落库走异步任务）
    pub async fn commit<S>(&self, source: ZipSource, sink: Arc<S>) -> Result<ImportReport>
    where
        S: BatchSink<T> + ?Sized,
    {
        self.import_runner().commit::<T, S>(source, sink).await
    }

    /// 只解析 + 自定义逐行处理（同步处理，不落库、不攒批）
    pub fn for_each_row<F>(&self, source: ZipSource, on_row: F) -> Result<ImportReport>
    where
        F: FnMut(&RowData<'_>, T) -> std::result::Result<(), Vec<String>>,
    {
        self.import_runner()
            .for_each_row::<T, F>(source, on_row)
    }

    /// 小数据量：写回内存（HTTP 直接下载）
    pub fn export_bytes(&self, rows: &[T]) -> Result<Vec<u8>> {
        self.export_runner().export_bytes(rows)
    }

    /// 小数据量：写文件
    pub fn export_rows<P: AsRef<Path>>(&self, path: P, rows: &[T]) -> Result<u64> {
        self.export_runner().export_rows(path, rows)
    }

    /// 小数据量：一个文件多张工作表（每张独立数据）
    pub fn export_sheets<P: AsRef<Path>>(
        &self,
        path: P,
        named_rows: Vec<(String, Vec<T>)>,
    ) -> Result<ExportStats> {
        let sheets = named_rows
            .into_iter()
            .map(|(name, rows)| (SheetOptions::new(name).columns(self.columns.clone()), rows))
            .collect();
        self.export_runner().export_sheets(path, sheets)
    }

    /// 游标分批拉取 + 流式写文件（百万行导出推荐）
    pub async fn export_xlsx<P, F, Fut>(&self, path: P, fetch: F) -> Result<ExportStats>
    where
        P: AsRef<Path>,
        F: FnMut(Option<String>) -> Fut,
        Fut: Future<Output = Result<Page<T>>> + Send,
    {
        self.export_runner().export_xlsx(path, fetch).await
    }

    /// 多工作表游标导出：每张表独立分页拉取，顺序写出
    pub async fn export_multi_sheet<P, F, Fut>(
        &self,
        path: P,
        sheets: Vec<String>,
        fetch: F,
    ) -> Result<ExportStats>
    where
        P: AsRef<Path>,
        F: FnMut(usize, Option<String>) -> Fut,
        Fut: Future<Output = Result<Page<T>>> + Send,
    {
        let options: Vec<SheetOptions> = sheets
            .into_iter()
            .map(|name| SheetOptions::new(name).columns(self.columns.clone()))
            .collect();
        self.export_runner().export_multi_sheet(path, options, fetch).await
    }
}

impl<T: ExcelRow + Serialize> crate::executor::ErasedExecutor for Executor<T> {
    fn name(&self) -> &'static str {
        self.name
    }

    fn type_name(&self) -> &'static str {
        self.type_name()
    }

    fn preview(&self, source: ZipSource, limit: usize) -> Result<serde_json::Value> {
        self.preview_json(source, limit)
    }
}

impl<T> fmt::Debug for Executor<T> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Executor")
            .field("name", &self.name)
            .field("sheet", &self.sheet)
            .field("model", &type_name::<T>())
            .field("columns", &self.columns.len())
            .finish()
    }
}

// ───────────────────────── 全局注册表（inventory 静态收集） ─────────────────────────

/// 注册表条目：`#[derive(ExcelExecutor)]` 自动提交
///
/// 通过 [`registered_executors`] / [`executor_entry`] 查询，
/// 想要运行时按名拿一个可用的执行器用 [`executor_by_name`]。
#[derive(Debug)]
pub struct ExecutorEntry {
    /// 注册名（`#[excel(register = "...")]`）
    pub name: &'static str,
    /// 模型类型名（调试 / 前端展示用）
    pub type_name: &'static str,
    /// 工厂：创建一个擦除了具体类型的执行器（预览返回 JSON）
    pub make: fn() -> Box<dyn ErasedExecutor>,
}

inventory::collect!(ExecutorEntry);

/// 已注册的全部执行器条目
pub fn registered_executors() -> Vec<&'static ExecutorEntry> {
    inventory::iter::<ExecutorEntry>
        .into_iter()
        .collect()
}

/// 已注册的注册名清单
pub fn registered_names() -> Vec<&'static str> {
    registered_executors()
        .into_iter()
        .map(|e| e.name)
        .collect()
}

/// 按注册名查条目
pub fn executor_entry(name: &str) -> Option<&'static ExecutorEntry> {
    inventory::iter::<ExecutorEntry>
        .into_iter()
        .find(|e| e.name == name)
}

/// 按注册名创建擦除类型的执行器（preview 返回 JSON，适合 Web 层按请求参数路由）
pub fn executor_by_name(name: &str) -> Option<Box<dyn ErasedExecutor>> {
    executor_entry(name).map(|e| (e.make)())
}

/// 按注册名直接预览，返回 JSON；未注册返回 `Ok(None)`
pub fn preview_by_name(
    name: &str,
    source: ZipSource,
    limit: usize,
) -> Result<Option<serde_json::Value>> {
    match executor_by_name(name) {
        Some(exec) => Ok(Some(exec.preview(source, limit)?)),
        None => Ok(None),
    }
}

/// 擦除具体类型后的执行器（按名调度用）
pub trait ErasedExecutor: Send + Sync {
    fn name(&self) -> &'static str;
    fn type_name(&self) -> &'static str;
    /// 预览前 `limit` 行并转成 JSON
    fn preview(&self, source: ZipSource, limit: usize) -> Result<serde_json::Value>;
}