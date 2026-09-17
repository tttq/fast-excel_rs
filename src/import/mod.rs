//! 导入侧：解析 → 校验 → 分批落库（可选并发）

pub mod report;
pub mod runner;
pub mod sink;

pub use report::{ImportPreview, ImportProgress, ImportReport, PreviewRow, RowError, SheetHeader};
pub use runner::{BatchConfig, ImportRunner, ProgressFn};
pub use sink::{BatchSink, FnSink, NoopSink, SinkFuture};
