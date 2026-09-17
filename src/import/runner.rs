//! 通用导入执行器
//!
//! 三条链路：
//!
//! | 方法 | 用途 | 是否落库 | 内存 |
//! |---|---|---|---|
//! | [`ImportRunner::preview`] | 上传后先给用户看校验结果 | 否 | 只装前 N 行 |
//! | [`ImportRunner::commit`] | 正式导入（解析 → 分批 → 落库） | 是 | 有界（通道容量 × 批大小） |
//! | [`ImportRunner::for_each_row`] | 只解析 + 自定义逐行处理（同步） | 否 | 单行 |
//!
//! `commit` 的内部结构（百万行的关键）：
//!
//! ```text
//!   spawn_blocking(解析线程)                 tokio 任务（落库）
//!   ┌──────────────────────┐   mpsc(有界)  ┌──────────────────────┐
//!   │ 逐行 SAX 解析 + 校验 │ ───────────▶ │ BatchSink::save      │
//!   │ 攒满 batch_size 就发 │  背压 = 内存  │ 事务 / insert_many   │
//!   └──────────────────────┘               └──────────────────────┘
//! ```
//!
//! 解析线程是阻塞式 CPU 活，放到 `spawn_blocking`；落库是 IO 活，走异步任务，
//! 通道有界 ⇒ 解析再快也不会把行堆在内存里。

use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::time::Instant;

use serde::Serialize;
use tokio::sync::mpsc;
use tokio::task::JoinSet;

use crate::column::ColumnDef;
use crate::error::{ExcelError, Result};
use crate::header::{HeaderAnalysis, HeaderMap, HeaderOptions};
use crate::model::{ExcelRow, RowData};
use crate::read::{ExcelReader, ReadOptions, SheetInfo, SheetSelector, ZipSource};

use super::report::{
    ImportPreview, ImportProgress, ImportReport, PreviewRow, RowError, SheetHeader,
};
use super::sink::BatchSink;

/// 进度回调（可跨线程调用，内部请勿做重活）
pub type ProgressFn = Arc<dyn Fn(&ImportProgress) + Send + Sync>;

/// 分批落库配置
#[derive(Debug, Clone)]
pub struct BatchConfig {
    /// 每批行数（一条 `insert_many` / 一个事务）
    pub batch_size: usize,
    /// 并发落库批次数；`1` = 严格串行（默认，事务友好、内存最小）
    pub max_in_flight: usize,
    /// 解析线程与落库之间的缓冲批次数（背压水位）
    pub queue_batches: usize,
    /// 行级错误是否跳过（`false` = 首行错误即中断）
    pub continue_on_row_error: bool,
    /// 落库出错是否立即中断（`false` = 记录后继续下一批）
    pub stop_on_error: bool,
    /// 每读取多少行回调一次进度
    pub progress_every: u64,
}

impl Default for BatchConfig {
    fn default() -> Self {
        Self {
            batch_size: 500,
            max_in_flight: 1,
            queue_batches: 2,
            continue_on_row_error: true,
            stop_on_error: false,
            progress_every: 1000,
        }
    }
}

/// 导入执行器
#[derive(Clone)]
pub struct ImportRunner {
    /// 读哪些工作表（`All` = 多 sheet 合并导入）
    pub sheets: SheetSelector,
    pub read: ReadOptions,
    pub batch: BatchConfig,
    /// 报告里最多保留多少条行级错误
    pub max_errors: usize,
    pub progress: Option<ProgressFn>,
}

impl Default for ImportRunner {
    fn default() -> Self {
        Self::new()
    }
}

impl ImportRunner {
    pub fn new() -> Self {
        Self {
            sheets: SheetSelector::First,
            read: ReadOptions::new(),
            batch: BatchConfig::default(),
            max_errors: 1000,
            progress: None,
        }
    }

    pub fn sheets(mut self, sheets: SheetSelector) -> Self {
        self.sheets = sheets;
        self
    }

    pub fn header_row(mut self, row_index: u32) -> Self {
        self.read.header.row_index = row_index;
        self
    }

    pub fn header_span(mut self, span: u32) -> Self {
        self.read.header.row_span = span.max(1);
        self
    }

    /// 严格表头匹配（默认允许别名 / 模糊 / 归一化）
    pub fn strict_header(mut self, strict: bool) -> Self {
        self.read.header = if strict {
            HeaderOptions {
                fuzzy: false,
                allow_index_fallback: false,
                ..self.read.header.clone()
            }
        } else {
            self.read.header.clone()
        };
        self
    }

    pub fn read_options(mut self, read: ReadOptions) -> Self {
        self.read = read;
        self
    }

    pub fn batch_size(mut self, size: usize) -> Self {
        self.batch.batch_size = size.max(1);
        self
    }

    /// 并发落库批次数（>1 需要业务侧自行保证批次间顺序无关）
    pub fn concurrency(mut self, n: usize) -> Self {
        self.batch.max_in_flight = n.max(1);
        self
    }

    pub fn queue_batches(mut self, n: usize) -> Self {
        self.batch.queue_batches = n.max(1);
        self
    }

    pub fn stop_on_error(mut self, stop: bool) -> Self {
        self.batch.stop_on_error = stop;
        self
    }

    pub fn continue_on_row_error(mut self, yes: bool) -> Self {
        self.batch.continue_on_row_error = yes;
        self
    }

    pub fn max_errors(mut self, n: usize) -> Self {
        self.max_errors = n;
        self
    }

    pub fn progress<F>(mut self, f: F) -> Self
    where
        F: Fn(&ImportProgress) + Send + Sync + 'static,
    {
        self.progress = Some(Arc::new(f));
        self
    }

    /// 模型声明了图片列时自动开启读图
    fn options_for<T: ExcelRow>(&self) -> ReadOptions {
        let mut read = self.read.clone();
        if read.image_columns.is_none() && T::columns().iter().any(|c| c.image) {
            read.read_images = true;
        }
        read
    }

    fn columns_of<T: ExcelRow>() -> Vec<ColumnDef> {
        T::columns()
    }
}

/// 进度计数（读取线程与落库任务共享）
#[derive(Default)]
pub(crate) struct ProgressState {
    rows_read: AtomicU64,
    rows_ok: AtomicU64,
    rows_error: AtomicU64,
    batches: AtomicU64,
    cb: Option<ProgressFn>,
}

impl ProgressState {
    fn new(cb: Option<ProgressFn>) -> Arc<Self> {
        Arc::new(Self {
            cb,
            ..Default::default()
        })
    }

    fn notify(&self) {
        if let Some(cb) = &self.cb {
            cb(&ImportProgress {
                rows_read: self.rows_read.load(Ordering::Relaxed),
                rows_ok: self.rows_ok.load(Ordering::Relaxed),
                rows_error: self.rows_error.load(Ordering::Relaxed),
                batches_done: self.batches.load(Ordering::Relaxed),
            });
        }
    }
}

/// 解析线程的产出
#[derive(Default)]
pub(crate) struct ReadOutcome {
    pub sheets: Vec<String>,
    pub headers: Vec<SheetHeader>,
    pub total_rows: u64,
    pub rows_sent: u64,
    pub error_rows: u64,
    pub skipped_rows: u64,
    pub batches: u64,
    pub errors: Vec<RowError>,
    pub errors_truncated: bool,
}

impl ReadOutcome {
    pub(crate) fn push_error(&mut self, sheet: &str, row: u32, messages: Vec<String>, max: usize) {
        if self.errors.len() < max {
            self.errors.push(RowError {
                row,
                sheet: sheet.to_string(),
                messages,
            });
        } else {
            self.errors_truncated = true;
        }
    }

    pub(crate) fn push_header(&mut self, sheet: &str, analysis: HeaderAnalysis) {
        self.headers.push(SheetHeader {
            sheet: sheet.to_string(),
            analysis,
        });
    }
}

/// 行错误 → 面向用户的中文消息
pub(crate) fn format_column_errors(errors: &[crate::model::ColumnError]) -> Vec<String> {
    errors
        .iter()
        .map(|e| {
            if e.file_header.is_empty() {
                format!("{}：{}", e.column, e.message)
            } else {
                format!("{}（{}）：{}", e.column, e.file_header, e.message)
            }
        })
        .collect()
}

impl ImportRunner {
    /// 校验预览（不落库，只读前 `limit` 行）
    pub fn preview<T>(&self, source: ZipSource, limit: usize) -> Result<ImportPreview<T>>
    where
        T: ExcelRow + Serialize,
    {
        let started = Instant::now();
        let defs = Self::columns_of::<T>();
        let reader = ExcelReader::from_source(source)?;
        let sheet = reader
            .select(&self.sheets)?
            .into_iter()
            .next()
            .cloned()
            .ok_or_else(|| ExcelError::SheetNotFound("工作表".to_string()))?;

        let mut options = self.options_for::<T>();
        options.max_rows = Some(limit.max(1) as u64);
        let mut stream = reader.stream_sheet(&sheet, &defs, options)?;

        let mut rows: Vec<PreviewRow<T>> = Vec::new();
        let mut total_rows = 0u64;
        let mut error_rows = 0u64;

        while let Some(row) = stream.next_row()? {
            total_rows += 1;
            let values = row.to_text_row();
            match T::from_row(&row) {
                Ok(data) => {
                    let errors = match data.validate() {
                        Ok(()) => Vec::new(),
                        Err(messages) => messages,
                    };
                    if !errors.is_empty() {
                        error_rows += 1;
                    }
                    rows.push(PreviewRow {
                        row: row.row_index,
                        data: Some(data),
                        values,
                        errors,
                    });
                }
                Err(column_errors) => {
                    error_rows += 1;
                    rows.push(PreviewRow {
                        row: row.row_index,
                        data: None,
                        values,
                        errors: format_column_errors(&column_errors),
                    });
                }
            }
        }

        let header = stream
            .header_analysis()
            .unwrap_or_else(|| HeaderMap::dynamic(Vec::new()).analysis());
        log::debug!(
            "导入预览完成：{} 行 / {} 错误 / {} ms",
            total_rows,
            error_rows,
            started.elapsed().as_millis()
        );
        Ok(ImportPreview {
            sheet: sheet.name.clone(),
            header,
            has_more: stream.limit_hit(),
            total_rows,
            error_rows,
            rows,
        })
    }
}
impl ImportRunner {
    /// 只解析 + 自定义逐行处理（同步；不落库、不攒批）
    ///
    /// 适合"解析后落一条非结构化数据"或自定义并发策略的场景；
    /// 需要事务 / 批量 insert 请用 [`ImportRunner::commit`]。
    pub fn for_each_row<T, F>(&self, source: ZipSource, mut on_row: F) -> Result<ImportReport>
    where
        T: ExcelRow,
        F: FnMut(&RowData<'_>, T) -> std::result::Result<(), Vec<String>>,
    {
        let started = Instant::now();
        let defs = Self::columns_of::<T>();
        let reader = ExcelReader::from_source(source)?;
        let sheets: Vec<SheetInfo> = reader.select(&self.sheets)?.into_iter().cloned().collect();
        let options = self.options_for::<T>();
        let mut outcome = ReadOutcome::default();

        for sheet in &sheets {
            outcome.sheets.push(sheet.name.clone());
            let mut stream = reader.stream_sheet(sheet, &defs, options.clone())?;
            let mut header_checked = false;
            while let Some(row) = stream.next_row()? {
                if !header_checked {
                    header_checked = true;
                    let missing = row.header.missing_required();
                    if !missing.is_empty() {
                        return Err(ExcelError::MissingColumns(missing));
                    }
                    outcome.push_header(&sheet.name, row.header.analysis());
                }
                outcome.total_rows += 1;
                match T::from_row(&row) {
                    Ok(data) => {
                        let mut messages = data.validate().err().unwrap_or_default();
                        if messages.is_empty() {
                            if let Err(errs) = on_row(&row, data) {
                                messages = errs;
                            }
                        }
                        if messages.is_empty() {
                            outcome.rows_sent += 1;
                        } else {
                            outcome.error_rows += 1;
                            outcome.push_error(
                                &sheet.name,
                                row.row_index,
                                messages,
                                self.max_errors,
                            );
                            if !self.batch.continue_on_row_error {
                                break;
                            }
                        }
                    }
                    Err(column_errors) => {
                        outcome.error_rows += 1;
                        outcome.push_error(
                            &sheet.name,
                            row.row_index,
                            format_column_errors(&column_errors),
                            self.max_errors,
                        );
                        if !self.batch.continue_on_row_error {
                            break;
                        }
                    }
                }
            }
            if !header_checked {
                if let Some(header) = stream.header() {
                    let missing = header.missing_required();
                    if !missing.is_empty() {
                        return Err(ExcelError::MissingColumns(missing));
                    }
                    outcome.push_header(&sheet.name, header.analysis());
                }
            }
            outcome.skipped_rows += stream.skipped_rows;
        }

        Ok(build_report(outcome, 0, 0, 0, started))
    }

    /// 正式导入：流式解析 + 分批落库
    ///
    /// - 解析在 `spawn_blocking` 线程，落库在异步任务；
    /// - 通道有界 ⇒ 内存占用 ≈ `queue_batches × batch_size` 行；
    /// - `stop_on_error = true` 时首批落库失败即中断解析（已提交的批次不回滚）。
    pub async fn commit<T, S>(&self, source: ZipSource, sink: Arc<S>) -> Result<ImportReport>
    where
        T: ExcelRow,
        // 允许 `Arc<dyn BatchSink<T>>`：业务侧可以用 trait object 把不同 sink 装进同一个调度器
        S: BatchSink<T> + ?Sized,
    {
        let started = Instant::now();
        let defs = Self::columns_of::<T>();
        let reader = ExcelReader::from_source(source)?;
        let sheets: Vec<SheetInfo> = reader.select(&self.sheets)?.into_iter().cloned().collect();
        let options = self.options_for::<T>();
        let config = self.batch.clone();
        let max_errors = self.max_errors;

        let progress = ProgressState::new(self.progress.clone());
        let cancel = Arc::new(AtomicBool::new(false));
        let (tx, mut rx) = mpsc::channel::<Vec<T>>(config.queue_batches.max(1));

        // ── 解析线程 ──
        let reader_progress = progress.clone();
        let reader_cancel = cancel.clone();
        let reader_config = config.clone();
        let reader_task = tokio::task::spawn_blocking(move || -> Result<ReadOutcome> {
            let mut outcome = ReadOutcome::default();
            let batch_size = reader_config.batch_size.max(1);
            let mut batch: Vec<T> = Vec::with_capacity(batch_size);
            let mut seq: u64 = 0;

            for sheet in &sheets {
                if reader_cancel.load(Ordering::Relaxed) {
                    break;
                }
                outcome.sheets.push(sheet.name.clone());
                let mut stream = reader.stream_sheet(sheet, &defs, options.clone())?;
                let mut header_checked = false;

                while let Some(row) = stream.next_row()? {
                    if !header_checked {
                        header_checked = true;
                        let missing = row.header.missing_required();
                        if !missing.is_empty() {
                            return Err(ExcelError::MissingColumns(missing));
                        }
                        outcome.push_header(&sheet.name, row.header.analysis());
                    }

                    outcome.total_rows += 1;
                    seq += 1;
                    reader_progress.rows_read.fetch_add(1, Ordering::Relaxed);

                    match T::from_row(&row) {
                        Ok(data) => {
                            let messages = data.validate().err().unwrap_or_default();
                            if !messages.is_empty() {
                                outcome.error_rows += 1;
                                reader_progress.rows_error.fetch_add(1, Ordering::Relaxed);
                                outcome.push_error(
                                    &sheet.name,
                                    row.row_index,
                                    messages,
                                    max_errors,
                                );
                                if !reader_config.continue_on_row_error {
                                    return Err(ExcelError::InvalidRow {
                                        row: row.row_index,
                                        messages: vec!["行校验失败，已按配置中断导入".to_string()],
                                    });
                                }
                                continue;
                            }
                            batch.push(data);
                            reader_progress.rows_ok.fetch_add(1, Ordering::Relaxed);
                            if batch.len() >= batch_size {
                                let chunk = std::mem::take(&mut batch);
                                if tx.blocking_send(chunk).is_err() {
                                    // 落库端已停止（例如 stop_on_error）
                                    return Ok(outcome);
                                }
                                outcome.batches += 1;
                            }
                        }
                        Err(column_errors) => {
                            outcome.error_rows += 1;
                            reader_progress.rows_error.fetch_add(1, Ordering::Relaxed);
                            let messages = format_column_errors(&column_errors);
                            outcome.push_error(
                                &sheet.name,
                                row.row_index,
                                messages.clone(),
                                max_errors,
                            );
                            if !reader_config.continue_on_row_error {
                                return Err(ExcelError::InvalidRow {
                                    row: row.row_index,
                                    messages,
                                });
                            }
                        }
                    }

                    if reader_config.progress_every > 0 && seq % reader_config.progress_every == 0 {
                        reader_progress.notify();
                    }
                }

                // 数据行一个都没有时，表头也已在 EOF 处构建好
                if !header_checked {
                    if let Some(header) = stream.header() {
                        let missing = header.missing_required();
                        if !missing.is_empty() {
                            return Err(ExcelError::MissingColumns(missing));
                        }
                        outcome.push_header(&sheet.name, header.analysis());
                    }
                }
                outcome.skipped_rows += stream.skipped_rows;
            }

            if !batch.is_empty() && tx.blocking_send(batch).is_ok() {
                outcome.batches += 1;
            }
            reader_progress.notify();
            Ok(outcome)
        });

        // ── 落库任务（有界并发） ──
        let mut set: JoinSet<Result<usize>> = JoinSet::new();
        let mut success_rows: u64 = 0;
        let mut sink_batches: u64 = 0;
        let mut failed_batches: u64 = 0;
        let mut sink_error: Option<ExcelError> = None;
        let mut stopped = false;
        let max_in_flight = config.max_in_flight.max(1);

        while let Some(chunk) = rx.recv().await {
            if !stopped {
                while set.len() >= max_in_flight {
                    if let Some(result) = set.join_next().await {
                        handle_batch_result(
                            result,
                            &mut success_rows,
                            &mut sink_batches,
                            &mut failed_batches,
                            &mut sink_error,
                            &mut stopped,
                            &config,
                            &cancel,
                        );
                    }
                }
                let s = sink.clone();
                set.spawn(async move {
                    let n = chunk.len();
                    match s.save(chunk).await {
                        Ok(written) => Ok(written),
                        Err(e) => {
                            s.on_batch_error(n, &e).await;
                            Err(e)
                        }
                    }
                });
            }
        }
        while let Some(result) = set.join_next().await {
            handle_batch_result(
                result,
                &mut success_rows,
                &mut sink_batches,
                &mut failed_batches,
                &mut sink_error,
                &mut stopped,
                &config,
                &cancel,
            );
        }

        progress.batches.store(sink_batches, Ordering::Relaxed);
        progress.notify();

        let outcome = reader_task
            .await
            .map_err(|e| ExcelError::Other(format!("解析线程异常：{e}")))??;

        if config.stop_on_error {
            if let Some(err) = sink_error {
                return Err(err);
            }
        }

        Ok(build_report(
            outcome,
            success_rows,
            sink_batches,
            failed_batches,
            started,
        ))
    }
}

#[allow(clippy::too_many_arguments)]
fn handle_batch_result(
    result: std::result::Result<Result<usize>, tokio::task::JoinError>,
    success_rows: &mut u64,
    batches: &mut u64,
    failed_batches: &mut u64,
    sink_error: &mut Option<ExcelError>,
    stopped: &mut bool,
    config: &BatchConfig,
    cancel: &Arc<AtomicBool>,
) {
    match result {
        Ok(Ok(written)) => {
            *success_rows += written as u64;
            *batches += 1;
        }
        Ok(Err(e)) => {
            *failed_batches += 1;
            log::error!("批量落库失败：{e}");
            if config.stop_on_error && !*stopped {
                *stopped = true;
                cancel.store(true, Ordering::Relaxed);
                *sink_error = Some(e);
            }
        }
        Err(join_error) => {
            *failed_batches += 1;
            let e = ExcelError::Sink(format!("落库任务异常：{join_error}"));
            log::error!("{e}");
            if config.stop_on_error && !*stopped {
                *stopped = true;
                cancel.store(true, Ordering::Relaxed);
                *sink_error = Some(e);
            }
        }
    }
}

fn build_report(
    outcome: ReadOutcome,
    success_rows: u64,
    batches: u64,
    failed_batches: u64,
    started: Instant,
) -> ImportReport {
    ImportReport {
        sheets: outcome.sheets,
        headers: outcome.headers,
        total_rows: outcome.total_rows,
        success_rows,
        error_rows: outcome.error_rows,
        skipped_rows: outcome.skipped_rows,
        batches,
        failed_batches,
        elapsed_ms: started.elapsed().as_millis() as u64,
        errors: outcome.errors,
        errors_truncated: outcome.errors_truncated,
    }
}
