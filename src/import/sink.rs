//! 批量落库抽象：读取线程只负责"解析 + 攒批"，落库由 [`BatchSink`] 实现决定

use std::future::Future;
use std::pin::Pin;

use async_trait::async_trait;

use crate::error::ExcelError;

/// 一批数据的落库入口
///
/// 业务侧通常这样实现（以 sea-orm 为例）：
///
/// ```ignore
/// #[async_trait]
/// impl BatchSink<ProductRow> for ProductSink {
///     async fn save(&self, batch: Vec<ProductRow>) -> Result<usize, ExcelError> {
///         let models: Vec<product::ActiveModel> = batch.into_iter().map(Into::into).collect();
///         let tx = self.db.inner().begin().await.map_err(|e| ExcelError::Sink(e.to_string()))?;
///         let n = models.len();
///         // sea-orm-ext 生成的批量插入：自动填 id / 审计字段 / 租户
///         product::Entity::insert_many_with_fill(models, &tx)
///             .await
///             .map_err(|e| ExcelError::Sink(e.to_string()))?;
///         tx.commit().await.map_err(|e| ExcelError::Sink(e.to_string()))?;
///         Ok(n)
///     }
/// }
/// ```
#[async_trait]
pub trait BatchSink<T>: Send + Sync + 'static {
    /// 落一批数据，返回实际写入行数
    async fn save(&self, batch: Vec<T>) -> Result<usize, ExcelError>;

    /// 一批失败后的回调（默认只记日志；`stop_on_error = false` 时用来做补偿 / 告警）
    async fn on_batch_error(&self, count: usize, error: &ExcelError) {
        log::warn!("批量落库失败（{count} 行）：{error}");
    }
}

/// 类型别名：boxed future（自定义 sink 需要时用）
pub type SinkFuture<'a, T> = Pin<Box<dyn Future<Output = T> + Send + 'a>>;

/// 用闭包当 sink（不想定义结构体时用）
pub struct FnSink<F> {
    f: F,
}

impl<F> FnSink<F> {
    pub fn new(f: F) -> Self {
        Self { f }
    }
}

#[async_trait]
impl<T, F, Fut> BatchSink<T> for FnSink<F>
where
    T: Send + 'static,
    F: Fn(Vec<T>) -> Fut + Send + Sync + 'static,
    Fut: Future<Output = Result<usize, ExcelError>> + Send,
{
    async fn save(&self, batch: Vec<T>) -> Result<usize, ExcelError> {
        (self.f)(batch).await
    }
}

/// 什么都不做的 sink（只做校验 / 只数行数时用）
pub struct NoopSink;

#[async_trait]
impl<T: Send + Sync + 'static> BatchSink<T> for NoopSink {
    async fn save(&self, batch: Vec<T>) -> Result<usize, ExcelError> {
        Ok(batch.len())
    }
}
