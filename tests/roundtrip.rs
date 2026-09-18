//! 端到端：写出 → 读回（动态表头 / 多 sheet / 单元格图片 / 模板）

use std::path::PathBuf;
use std::sync::Arc;

use base64::Engine;
use chrono::NaiveDate;
use fast_excel::{
    CellImage, CellValue, ExcelReader, ExcelRow, ExcelWriter, ExportRunner, Page, ReadOptions,
    ReferenceSheet, SheetOptions, SheetSelector, TemplateSpec, ZipSource, build_template,
};
use rust_decimal::Decimal;

fn temp_path(name: &str) -> PathBuf {
    let mut p = std::env::temp_dir();
    p.push(format!("excel-rs-{}-{name}", std::process::id()));
    p
}

/// 1x1 PNG（真图，给 embed_image / drawing 解析用）
fn png_1x1() -> Vec<u8> {
    base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==")
        .expect("base64 解码失败")
}

#[derive(ExcelRow, Debug, Clone, PartialEq)]
#[excel(sheet = "产品")]
struct Product {
    #[excel(header = "工厂名称", alias = ["工厂", "厂名"], required)]
    factory: String,
    #[excel(header = "品名", required)]
    name: String,
    #[excel(header = "报价(USD)", kind = "decimal", format = "#,##0.00")]
    price: Option<Decimal>,
    #[excel(header = "MOQ")]
    moq: Option<i32>,
    #[excel(header = "上架日期")]
    listed: Option<NaiveDate>,
    #[excel(header = "图片1", image)]
    image1: Option<CellImage>,
    #[excel(header = "备注")]
    note: Option<String>,
}

fn date(y: i32, m: u32, d: u32) -> NaiveDate {
    NaiveDate::from_ymd_opt(y, m, d).unwrap()
}

fn products(count: usize) -> Vec<Product> {
    (0..count)
        .map(|i| Product {
            factory: format!("工厂{}", i % 3 + 1),
            name: format!("品名-{i}"),
            price: Some(Decimal::new(12_900 + i as i64, 2)),
            moq: Some(100 + i as i32),
            listed: if i % 5 == 0 {
                None
            } else {
                Some(date(2025, 1, 1))
            },
            image1: if i % 10 == 0 {
                Some(CellImage::new(png_1x1(), "png").named("main.png"))
            } else {
                None
            },
            note: if i % 7 == 0 {
                Some(format!("备注 {i}"))
            } else {
                None
            },
        })
        .collect()
}

#[test]
fn test_derive_columns() {
    let cols = Product::columns();
    assert_eq!(cols.len(), 7);
    assert_eq!(cols[0].header, "工厂名称");
    assert!(cols[0].required);
    assert_eq!(cols[0].aliases, vec!["工厂", "厂名"]);
    assert_eq!(cols[2].kind, fast_excel::ColumnKind::Decimal);
    assert_eq!(cols[5].kind, fast_excel::ColumnKind::Image);
    assert!(cols[5].image);
    assert_eq!(Product::sheet_name(), Some("产品"));
}

#[test]
fn test_write_then_read_roundtrip_multi_sheet() {
    let path = temp_path("roundtrip.xlsx");
    let rows = products(500);

    // 写：两个 sheet（第二个只有文本）
    let mut writer = ExcelWriter::new();
    {
        let mut ws = writer
            .add_sheet(&SheetOptions::new("产品").columns(Product::columns()))
            .unwrap();
        ws.write_headers().unwrap();
        ws.write_models(&rows).unwrap();
        ws.finish().unwrap();
    }
    {
        let mut ws = writer.add_sheet(&SheetOptions::new("汇总")).unwrap();
        ws.write_text_row(["合计", "500"]).unwrap();
        ws.finish().unwrap();
    }
    writer.save(&path).unwrap();

    // 读：多 sheet 名称
    let reader = ExcelReader::open(&path).unwrap();
    assert_eq!(reader.sheet_names(), vec!["产品", "汇总"]);

    let mut stream = reader
        .stream(
            SheetSelector::name("产品"),
            &Product::columns(),
            ReadOptions::new().with_images(),
        )
        .unwrap();

    let mut read_back = Vec::new();
    while let Some(row) = stream.next_row().unwrap() {
        read_back.push(Product::from_row(&row).unwrap());
    }
    assert_eq!(read_back.len(), rows.len(), "行数应一致");
    for (i, (actual, expected)) in read_back.iter().zip(rows.iter()).enumerate() {
        assert_eq!(
            actual, expected,
            "第 {i} 行不一致：\n  读回 = {actual:?}\n  原值 = {expected:?}"
        );
    }
}

#[test]
fn test_dynamic_header_adaptation() {
    let path = temp_path("legacy.xlsx");

    // 旧模板：列顺序打乱 + 全角括号 + 多余列 + 别名列
    let legacy_headers = [
        "厂名", // 别名
        "MOQ",
        "品名",
        "上架日期",
        "报价（USD）(含税)", // 全角 + 后缀 → 模糊匹配
        "图片1",
        "未知列",
        "备注",
    ];

    let mut writer = ExcelWriter::new();
    {
        let mut ws = writer.add_sheet(&SheetOptions::new("Sheet1")).unwrap();
        ws.write_text_row(legacy_headers).unwrap();
        ws.write_row(&[
            CellValue::Text("工厂A".into()),
            CellValue::Number(200.0),
            CellValue::Text("商品X".into()),
            CellValue::Date(date(2025, 3, 9)),
            CellValue::Number(19.99),
            CellValue::Image(CellImage::new(png_1x1(), "png")),
            CellValue::Text("忽略我".into()),
            CellValue::Text("急单".into()),
        ])
        .unwrap();
        ws.finish().unwrap();
    }
    writer.save(&path).unwrap();

    let reader = ExcelReader::open(&path).unwrap();
    let mut stream = reader
        .stream(
            SheetSelector::First,
            &Product::columns(),
            ReadOptions::new().with_images(),
        )
        .unwrap();

    let row = stream.next_row().unwrap().expect("应有一行数据");
    let analysis = row.header.analysis();
    assert_eq!(analysis.missing, Vec::<String>::new(), "不应缺必填列");
    assert!(
        analysis.matched.iter().any(|m| m.matched_by == "alias"),
        "厂名 应命中别名"
    );
    assert!(
        analysis.matched.iter().any(|m| m.matched_by == "fuzzy"),
        "报价（USD）(含税) 应模糊命中"
    );
    assert_eq!(analysis.unknown, vec!["未知列".to_string()]);

    let parsed = Product::from_row(&row).unwrap();
    assert_eq!(parsed.factory, "工厂A");
    assert_eq!(parsed.name, "商品X");
    assert_eq!(parsed.moq, Some(200));
    assert_eq!(parsed.price, Some(Decimal::new(1999, 2)));
    assert_eq!(parsed.listed, Some(date(2025, 3, 9)));
    assert_eq!(parsed.note.as_deref(), Some("急单"));

    // 单元格内嵌图片按行读回
    eprintln!("row.images = {:?}", row.images);
    let image = parsed.image1.expect("图片应被解析出来");
    assert_eq!(image.ext, "png");
    assert_eq!(image.bytes, png_1x1());
}

#[test]
fn test_missing_required_column_reports() {
    let path = temp_path("missing.xlsx");
    let mut writer = ExcelWriter::new();
    {
        let mut ws = writer.add_sheet(&SheetOptions::new("Sheet1")).unwrap();
        ws.write_text_row(["品名"]).unwrap();
        ws.write_row(&[CellValue::Text("只有品名".into())]).unwrap();
        ws.finish().unwrap();
    }
    writer.save(&path).unwrap();

    let reader = ExcelReader::open(&path).unwrap();
    let analysis = reader
        .analyze_header(
            &SheetSelector::First,
            &Product::columns(),
            ReadOptions::new(),
        )
        .unwrap();
    assert_eq!(analysis.missing, vec!["工厂名称".to_string()]);
}

#[test]
fn test_dynamic_row_collects_everything() {
    let path = temp_path("dynamic.xlsx");
    let mut writer = ExcelWriter::new();
    {
        let mut ws = writer.add_sheet(&SheetOptions::new("Sheet1")).unwrap();
        ws.write_text_row(["甲", "乙", "丙"]).unwrap();
        ws.write_row(&[
            CellValue::Text("1".into()),
            CellValue::Number(2.0),
            CellValue::Bool(true),
        ])
        .unwrap();
        ws.finish().unwrap();
    }
    writer.save(&path).unwrap();

    let reader = ExcelReader::open(&path).unwrap();
    let mut stream = reader
        .stream(SheetSelector::First, &[], ReadOptions::new())
        .unwrap();
    let row = stream.next_row().unwrap().unwrap();
    let dynamic = fast_excel::DynamicRow::from_row(&row).unwrap();
    assert_eq!(dynamic.text("甲"), "1");
    assert_eq!(dynamic.text("乙"), "2");
    assert_eq!(dynamic.text("丙"), "TRUE");
}

#[test]
fn test_template_generation() {
    let spec = TemplateSpec::new(Product::columns())
        .sheet_name("导入模板")
        .title("产品导入模板")
        .note("图片直接粘贴到单元格")
        .sample_row(vec![
            CellValue::Text("示例工厂".into()),
            CellValue::Text("示例品名".into()),
            CellValue::Empty,
            CellValue::Empty,
            CellValue::Empty,
            CellValue::Empty,
            CellValue::Empty,
        ])
        .reference(
            ReferenceSheet::new("工厂列表", vec!["工厂名称".to_string(), "短码".to_string()])
                .rows(vec![vec!["工厂A".to_string(), "A".to_string()]]),
        );

    let bytes = build_template(&spec).unwrap();
    let reader = ExcelReader::from_bytes(bytes).unwrap();
    assert_eq!(
        reader.sheet_names(),
        vec!["导入模板", "填写说明", "工厂列表"]
    );

    // 标题行占了第 1 行，所以表头在第 2 行（0 基 1）
    let mut stream = reader
        .stream(
            SheetSelector::First,
            &Product::columns(),
            ReadOptions::new().header_row(1),
        )
        .unwrap();
    let row = stream.next_row().unwrap().expect("模板里应有示例行");
    assert_eq!(row.text_of("工厂名称"), "示例工厂");
    assert_eq!(row.text_of("品名"), "示例品名");
}

#[test]
fn test_zip_source_from_bytes_matches_path() {
    let path = temp_path("bytes.xlsx");
    let rows = products(20);
    let mut writer = ExcelWriter::new();
    {
        let mut ws = writer
            .add_sheet(&SheetOptions::new("产品").columns(Product::columns()))
            .unwrap();
        ws.write_headers().unwrap();
        ws.write_models(&rows).unwrap();
        ws.finish().unwrap();
    }
    writer.save(&path).unwrap();

    let bytes = std::fs::read(&path).unwrap();
    let reader = ExcelReader::from_bytes(bytes).unwrap();
    let mut stream = reader
        .stream(
            SheetSelector::First,
            &Product::columns(),
            ReadOptions::new(),
        )
        .unwrap();
    let mut count = 0;
    while let Some(row) = stream.next_row().unwrap() {
        Product::from_row(&row).unwrap();
        count += 1;
    }
    assert_eq!(count, 20);

    // ZipSource 直接走一遍（导入管道就是这么用的）
    let source = ZipSource::from_bytes(std::fs::read(&path).unwrap());
    let reader = ExcelReader::from_source(source).unwrap();
    assert_eq!(reader.sheet_names().len(), 1);
}
#[tokio::test]
async fn test_export_multi_sheet_paged() {
    let path = temp_path("multi-export.xlsx");
    let runner = ExportRunner::new(SheetOptions::new("ignored"));
    let data = Arc::new(products(30));

    let stats = runner
        .export_multi_sheet(
            &path,
            vec![
                SheetOptions::new("甲").columns(Product::columns()),
                SheetOptions::new("乙").columns(Product::columns()),
            ],
            {
                let data = data.clone();
                move |_sheet, cursor| {
                    let data = data.clone();
                    async move {
                        let offset: usize =
                            cursor.as_deref().and_then(|c| c.parse().ok()).unwrap_or(0);
                        if offset >= data.len() {
                            return Ok(Page::last(Vec::new()));
                        }
                        let end = (offset + 10).min(data.len());
                        let rows = data[offset..end].to_vec();
                        if end >= data.len() {
                            Ok(Page::last(rows))
                        } else {
                            Ok(Page::more(rows, end.to_string()))
                        }
                    }
                }
            },
        )
        .await
        .unwrap();

    assert_eq!(stats.rows, 60);
    assert_eq!(stats.pages, 6);

    let reader = ExcelReader::open(&path).unwrap();
    assert_eq!(reader.sheet_names(), vec!["甲", "乙"]);
    for name in ["甲", "乙"] {
        let mut stream = reader
            .stream(
                SheetSelector::name(name),
                &Product::columns(),
                ReadOptions::new().with_images(),
            )
            .unwrap();
        let mut count = 0;
        while let Some(row) = stream.next_row().unwrap() {
            Product::from_row(&row).unwrap();
            count += 1;
        }
        assert_eq!(count, 30, "sheet {name} 行数不一致");
    }
}

#[test]
fn test_export_sheets_in_memory() {
    let path = temp_path("multi-rows.xlsx");
    let runner = ExportRunner::new(SheetOptions::new("ignored"));
    let stats = runner
        .export_sheets(
            &path,
            vec![
                (
                    SheetOptions::new("A").columns(Product::columns()),
                    products(3),
                ),
                (
                    SheetOptions::new("B").columns(Product::columns()),
                    products(5),
                ),
            ],
        )
        .unwrap();
    assert_eq!(stats.rows, 8);

    let reader = ExcelReader::open(&path).unwrap();
    assert_eq!(reader.sheet_names(), vec!["A", "B"]);
}

/// 百万行压力测试：验证“边写边落盘 + SAX 逐行读”不依赖行数。
/// 默认 `#[ignore]`，手动跑：`cargo test -p fast-excel --test roundtrip -- --ignored`
#[test]
#[ignore = "million-row stress test; run manually with --ignored"]
fn stress_one_million_rows_streaming() {
    let path = temp_path("million.xlsx");
    let mut writer = ExcelWriter::new();
    {
        let mut ws = writer.add_sheet(&SheetOptions::new("big")).unwrap();
        ws.write_text_row(["编号", "名称", "数量"]).unwrap();
        for i in 0..1_000_000u32 {
            ws.write_text_row([i.to_string(), format!("行-{i}"), "1".to_string()])
                .unwrap();
        }
        ws.finish().unwrap();
    }
    let size = writer.save(&path).unwrap();

    let reader = ExcelReader::open(&path).unwrap();
    let mut stream = reader
        .stream(SheetSelector::First, &[], ReadOptions::new())
        .unwrap();
    let mut count = 0u64;
    while let Some(_row) = stream.next_row().unwrap() {
        count += 1;
    }
    assert_eq!(count, 1_000_000);
    eprintln!("million-row xlsx = {} MB", size / 1024 / 1024);
    std::fs::remove_file(&path).ok();
}
