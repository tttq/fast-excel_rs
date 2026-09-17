//! 导出侧：游标分批拉取 + 流式写出（百万行不进内存）
//!
//! 用法（sea-orm 游标分页为例）：
//!
//! ```ignore
//! let runner = ExportRunner::new(SheetOptions::new("选品数据").columns(ProductRow::columns()));
//! let stats = runner
//!     .export_xlsx("uploads/export/products.xlsx", |cursor| async move {
//!         let mut q = product::Entity::find().order_by_asc(product::Column::Id).limit(1000);
//!         if let Some(last_id) = cursor {
//!             q = q.filter(product::Column::Id.gt(last_id));
//!         }
//!         let rows = q.all(&db).await?;
//!         let done = rows.len() < 1000;
//!         let last_id = rows.last().map(|r| r.id.clone());
//!         Ok(Page { rows: rows.into_iter().map(Into::into).collect(), cursor: last_id, done })
//!     })
//!     .await?;
//! ```
//!
//! 多工作表导出用 [`ExportRunner::export_multi_sheet`]：每个 sheet 各自游标分页，
//! 依次写出，单表百万行依旧是常量内存。

use std::future::Future;
use std::path::{Path, PathBuf};
use std::time::Instant;

use serde::Serialize;
use tokio::sync::mpsc;

use crate::error::{ExcelError, Result};
use crate::model::ExcelRow;
use crate::write::{ExcelWriter, SheetOptions};

/// 一页导出数据（游标由业务自定义：上一页最后一条 id / 时间戳 / 主键组合都可以）
#[derive(Debug, Clone)]
pub struct Page<T> {
    pub rows: Vec<T>,
    /// 下一页游标；`None` 表示没有更多（配合 `done` 使用）
    pub cursor: Option<String>,
    /// 是否已是最后一页
    pub done: bool,
}

impl<T> Page<T> {
    /// 恰好一页（拉完就结束）
    pub fn last(rows: Vec<T>) -> Self {
        Self {
            rows,
            cursor: None,
            done: true,
        }
    }

    /// 还有下一页
    pub fn more(rows: Vec<T>, cursor: impl Into<String>) -> Self {
        Self {
            rows,
            cursor: Some(cursor.into()),
            done: false,
        }
    }
}

/// 导出统计
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ExportStats {
    /// 写出行数（不含表头）
    pub rows: u64,
    /// 拉取页数
    pub pages: u64,
    pub file_size: u64,
    pub elapsed_ms: u64,
    pub path: String,
}

/// 导出执行器
#[derive(Debug, Clone)]
pub struct ExportRunner {
    pub sheet: SheetOptions,
    /// 解析与写出之间的缓冲页数（越多越吃内存，默认 2）
    pub queue_pages: usize,
}

/// 多 sheet 写出线程的消息：行批次 / 本表结束
enum WriteMsg<T> {
    Rows(Vec<T>),
    EndSheet,
}

impl ExportRunner {
    pub fn new(sheet: SheetOptions) -> Self {
        Self {
            sheet,
            queue_pages: 2,
        }
    }

    pub fn queue_pages(mut self, pages: usize) -> Self {
        self.queue_pages = pages.max(1);
        self
    }

    /// 游标分批拉取 + 流式写文件（推荐：百万行导出走它）
    pub async fn export_xlsx<T, F, Fut>(
        &self,
        path: impl AsRef<Path>,
        mut fetch: F,
    ) -> Result<ExportStats>
    where
        T: ExcelRow,
        F: FnMut(Option<String>) -> Fut,
        Fut: Future<Output = Result<Page<T>>> + Send,
    {
        let started = Instant::now();
        let out_path: PathBuf = path.as_ref().to_path_buf();
        let (tx, mut rx) = mpsc::channel::<Vec<T>>(self.queue_pages);

        let sheet = self.sheet.clone();
        let write_path = out_path.clone();
        let writer_task = tokio::task::spawn_blocking(move || -> Result<u64> {
            let mut writer = ExcelWriter::new();
            let mut ws = writer.add_sheet(&sheet)?;
            ws.write_headers()?;
            while let Some(rows) = rx.blocking_recv() {
                for row in &rows {
                    ws.write_model(row)?;
                }
            }
            ws.finish()?;
            drop(ws);
            writer.save(&write_path)
        });

        let mut cursor: Option<String> = None;
        let mut pages = 0u64;
        let mut rows_total = 0u64;
        let fetch_result: Result<()> = async {
            loop {
                let page = fetch(cursor.clone()).await?;
                pages += 1;
                let done = page.done;
                let next = page.cursor.clone();
                rows_total += page.rows.len() as u64;
                if !page.rows.is_empty() {
                    tx.send(page.rows)
                        .await
                        .map_err(|_| ExcelError::Other("导出写出任务已结束".to_string()))?;
                }
                if done {
                    break;
                }
                if next.is_none() {
                    // 业务没给游标也没标记结束：避免死循环
                    log::warn!("导出分页既没有 done 也没有 cursor，提前结束");
                    break;
                }
                cursor = next;
            }
            Ok(())
        }
        .await;

        drop(tx);
        let file_size = writer_task
            .await
            .map_err(|e| ExcelError::Other(format!("导出写出线程异常：{e}")))??;
        if let Err(e) = fetch_result {
            // 拉取失败时不要留下半截文件
            let _ = std::fs::remove_file(&out_path);
            return Err(e);
        }

        Ok(ExportStats {
            rows: rows_total,
            pages,
            file_size,
            elapsed_ms: started.elapsed().as_millis() as u64,
            path: out_path.to_string_lossy().into_owned(),
        })
    }

    /// 多工作表游标导出：每个 sheet 独立分页拉取，顺序写出
    ///
    /// 与 [`ExportRunner::export_xlsx`] 的差别只有一个：`fetch` 额外收到
    /// `sheet_index`，可以给每个 sheet 拉不同筛选条件的数据；
    /// 写出仍是常量内存，单表百万行不会因为拆表而进内存。
    pub async fn export_multi_sheet<T, F, Fut>(
        &self,
        path: impl AsRef<Path>,
        sheets: Vec<SheetOptions>,
        mut fetch: F,
    ) -> Result<ExportStats>
    where
        T: ExcelRow,
        F: FnMut(usize, Option<String>) -> Fut,
        Fut: Future<Output = Result<Page<T>>> + Send,
    {
        if sheets.is_empty() {
            return Err(ExcelError::Other("导出至少需要一个工作表".to_string()));
        }
        let started = Instant::now();
        let out_path: PathBuf = path.as_ref().to_path_buf();
        let sheet_count = sheets.len();
        let write_path = out_path.clone();
        let (tx, mut rx) = mpsc::channel::<WriteMsg<T>>(self.queue_pages.max(1));

        // 写线程：每个 sheet 写完才取下一个（只持有一个 SheetWriter 借用）
        let writer_task = tokio::task::spawn_blocking(move || -> Result<u64> {
            let mut writer = ExcelWriter::new();
            for options in &sheets {
                let mut ws = writer.add_sheet(options)?;
                ws.write_headers()?;
                while let Some(WriteMsg::Rows(rows)) = rx.blocking_recv() {
                    for row in &rows {
                        ws.write_model(row)?;
                    }
                }
                ws.finish()?;
            }
            writer.save(&write_path)
        });

        let mut pages = 0u64;
        let mut rows_total = 0u64;
        let fetch_result: Result<()> = async {
            for index in 0..sheet_count {
                let mut cursor: Option<String> = None;
                loop {
                    let page = fetch(index, cursor.clone()).await?;
                    pages += 1;
                    let done = page.done;
                    let next = page.cursor.clone();
                    rows_total += page.rows.len() as u64;
                    if !page.rows.is_empty() {
                        tx.send(WriteMsg::Rows(page.rows))
                            .await
                            .map_err(|_| ExcelError::Other("导出写出任务已结束".to_string()))?;
                    }
                    if done {
                        break;
                    }
                    if next.is_none() {
                        log::warn!("导出分页既没有 done 也没有 cursor，提前结束（sheet {index}）");
                        break;
                    }
                    cursor = next;
                }
                tx.send(WriteMsg::EndSheet)
                    .await
                    .map_err(|_| ExcelError::Other("导出写出任务已结束".to_string()))?;
            }
            Ok(())
        }
        .await;

        drop(tx);
        let file_size = writer_task
            .await
            .map_err(|e| ExcelError::Other(format!("导出写出线程异常：{e}")))??;
        if let Err(e) = fetch_result {
            let _ = std::fs::remove_file(&out_path);
            return Err(e);
        }

        Ok(ExportStats {
            rows: rows_total,
            pages,
            file_size,
            elapsed_ms: started.elapsed().as_millis() as u64,
            path: out_path.to_string_lossy().into_owned(),
        })
    }

    /// 小数据量：一次性写多个工作表（每个 sheet 独立列定义／数据）
    pub fn export_sheets<T: ExcelRow>(
        &self,
        path: impl AsRef<Path>,
        sheets: Vec<(SheetOptions, Vec<T>)>,
    ) -> Result<ExportStats> {
        let started = Instant::now();
        let out_path = path.as_ref().to_path_buf();
        let mut rows_total = 0u64;
        let mut writer = ExcelWriter::new();
        for (options, rows) in &sheets {
            let mut ws = writer.add_sheet(options)?;
            ws.write_headers()?;
            ws.write_models(rows)?;
            ws.finish()?;
            rows_total += rows.len() as u64;
        }
        let file_size = writer.save(&out_path)?;
        Ok(ExportStats {
            rows: rows_total,
            pages: sheets.len() as u64,
            file_size,
            elapsed_ms: started.elapsed().as_millis() as u64,
            path: out_path.to_string_lossy().into_owned(),
        })
    }

    /// 小数据量：一次性写出文件
    pub fn export_rows<T: ExcelRow>(&self, path: impl AsRef<Path>, rows: &[T]) -> Result<u64> {
        let mut writer = ExcelWriter::new();
        {
            let mut ws = writer.add_sheet(&self.sheet)?;
            ws.write_headers()?;
            ws.write_models(rows)?;
            ws.finish()?;
        }
        writer.save(path)
    }

    /// 小数据量：写出到内存（HTTP 直接下载）
    pub fn export_bytes<T: ExcelRow>(&self, rows: &[T]) -> Result<Vec<u8>> {
        let mut writer = ExcelWriter::new();
        {
            let mut ws = writer.add_sheet(&self.sheet)?;
            ws.write_headers()?;
            ws.write_models(rows)?;
            ws.finish()?;
        }
        writer.save_to_buffer()
    }
}
