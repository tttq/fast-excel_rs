//! 验证「一个宏注册异步导出任务」的两种形态。

use fast_excel::{ExportTaskContext, export_task_registered, run_export_task};

const HEADERS: &[&str] = &["名称", "数量"];

async fn unit_rows(_ctx: ExportTaskContext) -> Result<Vec<Vec<String>>, String> {
    Ok(vec![
        vec!["A".to_string(), "1".to_string()],
        vec!["B".to_string(), "2".to_string()],
    ])
}

fast_excel::export_task! {
    task_type = "unit-rows",
    sheet_name = "测试数据",
    headers = HEADERS,
    rows = unit_rows,
}

async fn unit_file(ctx: ExportTaskContext) -> Result<(i64, String), String> {
    let name = "unit.csv".to_string();
    std::fs::write(ctx.out_dir.join(&name), b"a,b\n1,2\n")
        .map_err(|e| e.to_string())?;
    Ok((1, name))
}

fast_excel::export_task! {
    task_type = "unit-file",
    mime = "text/csv",
    file = unit_file,
}

#[test]
fn task_is_registered_at_link_time() {
    assert!(export_task_registered("unit-rows"));
    assert!(export_task_registered("unit-file"));
    assert!(!export_task_registered("unit-missing"));
}

#[tokio::test]
async fn rows_task_writes_xlsx() {
    let dir = std::env::temp_dir().join(format!("excel-rows-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    let out = run_export_task(
        "unit-rows",
        ExportTaskContext {
            task_no: "T1".to_string(),
            query: serde_json::json!({}),
            out_dir: dir.clone(),
        },
    )
    .await
    .unwrap();
    assert_eq!(out.total_rows, 2);
    assert!(out.file_size > 0);
    assert!(dir.join(&out.file_name).is_file());
    let _ = std::fs::remove_dir_all(&dir);
}

#[tokio::test]
async fn file_task_reports_size() {
    let dir = std::env::temp_dir().join(format!("excel-file-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    let out = run_export_task(
        "unit-file",
        ExportTaskContext {
            task_no: "T2".to_string(),
            query: serde_json::json!({}),
            out_dir: dir.clone(),
        },
    )
    .await
    .unwrap();
    assert_eq!(out.total_rows, 1);
    assert!(out.file_size > 0);
    assert_eq!(out.mime.as_deref(), Some("text/csv"));
    let _ = std::fs::remove_dir_all(&dir);
}
