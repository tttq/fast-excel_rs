//! 执行器：注册宏 + 全局注册表 + 通用入口（消费者视角）
//!
//! 本文件模拟"业务端"使用 `#[derive(ExcelExecutor)]`：
//! 不管理任何工厂，注册 + 调度全靠宏与全局表。

use std::path::PathBuf;
use std::sync::Arc;

use fast_excel::{
    ExcelExecutor, ExcelReader, ExcelRow, Executor, NoopSink, SheetSelector, ZipSource,
    build_template, executor_by_name, executor_entry, executor_for, preview_by_name,
    registered_executors, registered_names,
};
use serde::Serialize;

#[derive(ExcelRow, ExcelExecutor, Serialize, Debug, Clone, PartialEq)]
#[excel(sheet = "产品", register = "product")]
struct Product {
    #[excel(header = "品名", required)]
    name: String,
    #[excel(header = "数量")]
    qty: Option<i32>,
}

#[derive(ExcelRow, ExcelExecutor, Serialize, Debug, Clone, PartialEq)]
#[excel(sheet = "订单", register = "order")]
struct Order {
    #[excel(header = "订单号", required)]
    no: String,
}

fn temp_path(name: &str) -> PathBuf {
    let mut p = std::env::temp_dir();
    p.push(format!("excel-executor-{}-{name}", std::process::id()));
    p
}

fn products() -> Vec<Product> {
    vec![
        Product { name: "螺丝".into(), qty: Some(100) },
        Product { name: "螺母".into(), qty: None },
    ]
}

#[test]
fn test_register_and_discover() {
    // 注册表自动收集（inventory 链接期），无需任何初始化
    let names = registered_names();
    assert!(names.contains(&"product"), "缺 product：{names:?}");
    assert!(names.contains(&"order"), "缺 order：{names:?}");

    let entry = executor_entry("product").expect("product 应已注册");
    assert_eq!(entry.name, "product");
    assert_eq!(entry.type_name, "Product");

    let all = registered_executors();
    assert_eq!(all.len(), 2, "恰好两个模型注册：{:?}", all);
}

#[test]
fn test_executor_for_uses_model_defaults() {
    let exec = executor_for::<Product>();
    assert_eq!(exec.name(), "product");
    assert_eq!(exec.sheet_name(), Some("产品"));
    assert_eq!(exec.columns().len(), 2);

    // 派生态生成的内置关联函数
    assert_eq!(Product::executor().name(), "product");

    // 手动构造（未注册模型也能用）
    let manual = Executor::<Product>::new("custom", None);
    assert_eq!(manual.name(), "custom");
    assert_eq!(manual.sheet_name(), None);
}

#[test]
fn test_preview_by_type_and_by_name() {
    let path = temp_path("preview.xlsx");
    let exec = executor_for::<Product>();
    exec.export_rows(&path, &products()).unwrap();

    let source = ZipSource::open(&path).unwrap();
    // 类型化预览
    let preview = exec.preview(source.clone(), 10).unwrap();
    assert_eq!(preview.total_rows, 2);
    assert_eq!(preview.error_rows, 0);
    assert_eq!(preview.rows[0].data.as_ref().unwrap().name, "螺丝");

    // 按名预览 → JSON（Web 层用）
    let json = preview_by_name("product", source, 10)
        .unwrap()
        .expect("product 已注册");
    let total = json.get("totalRows").and_then(|v| v.as_u64()).unwrap();
    assert_eq!(total, 2, "JSON 预览行数：{json}");
}

#[tokio::test]
async fn test_commit_via_executor() {
    let path = temp_path("commit.xlsx");
    executor_for::<Product>().export_rows(&path, &products()).unwrap();

    let report = executor_for::<Product>()
        .commit(ZipSource::open(&path).unwrap(), Arc::new(NoopSink))
        .await
        .unwrap();
    assert_eq!(report.success_rows, 2, "{report:?}");
    assert_eq!(report.error_rows, 0);
}

#[test]
fn test_export_and_template_via_executor() {
    // 导出到内存 → 读回
    let bytes = executor_for::<Product>().export_bytes(&products()).unwrap();
    let reader = ExcelReader::from_bytes(bytes).unwrap();
    let mut stream = reader
        .stream(
            SheetSelector::First,
            &Product::columns(),
            fast_excel::ReadOptions::new(),
        )
        .unwrap();
    let mut count = 0;
    while let Some(row) = stream.next_row().unwrap() {
        Product::from_row(&row).unwrap();
        count += 1;
    }
    assert_eq!(count, 2);

    // 模板
    let spec = executor_for::<Product>().template();
    assert_eq!(spec.sheet_name, "产品");
    let template_bytes = build_template(&spec).unwrap();
    assert!(!template_bytes.is_empty());
}

#[test]
fn test_erased_executor_by_name() {
    let path = temp_path("erased.xlsx");
    executor_for::<Product>().export_rows(&path, &products()).unwrap();

    let erased = executor_by_name("product").expect("product 已注册");
    assert_eq!(erased.name(), "product");
    assert!(erased.type_name().contains("Product"));

    // 手动构造路径（未注册模型也能用）
    let manual = Executor::<Product>::new("x", None);
    assert_eq!(manual.name(), "x");
    assert_eq!(manual.columns().len(), 2);
}

#[test]
fn test_missing_name_returns_none() {
    assert!(executor_entry("not-exists").is_none());
    assert!(executor_by_name("not-exists").is_none());
}