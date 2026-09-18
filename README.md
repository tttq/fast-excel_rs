# fast-excel\_rs —— 通用 Excel 导入导出工具（Rust 版 EasyExcel）

面向真实业务场景：**百万行 + 批量落库 + 单元格图片 + 多 sheet + 动态表头**。
纯 Rust（`quick-xml` + `rust_xlsxwriter` + `zip`），不依赖 JVM、Office 组件或外部服务。

本 crate 是独立工具库；采购模块的 Excel 导入 / 导出中心共用它的一套列定义，
保证「模板 / 导出 / 导入校验」三处表头永远一致。

## 能力对照

| 能力      | EasyExcel 对应         | 本工具实现                                                                                             |
| ------- | -------------------- | ------------------------------------------------------------------------------------------------- |
| 注解式映射   | `@ExcelProperty`     | `#[derive(ExcelRow)]` + `#[excel(header = "...")]`                                                |
| 动态表头适配  | 表头匹配器                | 精确 / 别名 / 模糊 / 列序回退（`HeaderOptions`）                                                              |
| 百万行读取   | SAX 逐行读              | `SheetStream`：`quick-xml` 事件流，内存与行数无关                                                             |
| 百万行写出   | `constant_memory`    | `ExcelWriter` / `SheetWriter`：边写边落临时文件                                                            |
| 分批落库    | `ReadListener` + 手动批 | `ImportRunner::commit` + `BatchSink`（`batch_size` 一批一事务）                                          |
| 分页导出    | 分页查询写                | `ExportRunner::export_xlsx`：游标分页 `Page { rows, cursor, done }`                                    |
| 单元格图片   | 图片列 / 自定义转换          | 读：标准 drawing 锚点 + WPS `DISPIMG`；写：`CellValue::Image`                                              |
| 多 sheet | 多 `sheet()` 写        | 读：`SheetSelector::{Name,Names,All}`；写：`ExcelWriter::add_sheet`、`ExportRunner::export_multi_sheet` |
| 导入模板    | 模板导出                 | `TemplateSpec` / `build_template`：必填红字 + 示例行 + 下拉 + 说明 sheet                                      |
| 导入预览    | 校验后再入库               | `ImportRunner::preview`（只读前 N 行，返回列匹配与行级错误）                                                       |
| 执行器注册   | 注册后全链路复用             | `#[derive(ExcelExecutor)]` + 全局注册表：预览 / 导入 / 导出 / 模板一条龙（按类型或按名字符串）                           |
| 纯文本兜底   | —                    | `write_csv`（UTF-8 BOM，Excel 双击不乱码）                                                                |

## 快速开始

### 1）声明行模型

```rust
use excel::{CellImage, ExcelRow};

#[derive(ExcelRow, Debug, Clone)]
#[excel(sheet = "产品导入")]
struct ProductRow {
    #[excel(header = "工厂名称", alias = ["工厂", "厂名"], required)]
    factory: String,
    #[excel(header = "品名", required)]
    name: String,
    #[excel(header = "报价(USD)", kind = "decimal", format = "#,##0.00")]
    price: Option<rust_decimal::Decimal>,
    #[excel(header = "手机号", text)]
    phone: Option<String>,
    #[excel(header = "图片1", image)]
    image1: Option<CellImage>,
}
```

字段属性：`header` / `alias` / `required` / `kind` / `width` / `format` / `text` /
`image` / `default` / `note` / `dropdown` / `ignore`。

### 2）导入：预览 → 分批落库

```rust
use std::sync::Arc;
use excel::{BatchSink, ExcelError, ImportRunner, SheetSelector, ZipSource};

// 每批一个事务，业务侧自己决定事务边界
struct ProductSink { /* db: DbConn, ... */ }

#[async_trait::async_trait]
impl BatchSink<ProductRow> for ProductSink {
    async fn save(&self, batch: Vec<ProductRow>) -> Result<usize, ExcelError> {
        let n = batch.len();
        // product::Entity::insert_many_with_fill(models, &tx).await?; tx.commit().await?;
        Ok(n)
    }
}

# async fn demo() -> Result<(), ExcelError> {
let source = ZipSource::open("products.xlsx")?;

// 预览：不落库，只读前 200 行，返回表头匹配 + 行级错误
let preview = ImportRunner::new()
    .sheets(SheetSelector::All)
    .preview::<ProductRow>(source.clone(), 200)?;

// 正式导入：解析在阻塞线程，落库在异步任务，中间有界队列做背压
let report = ImportRunner::new()
    .sheets(SheetSelector::All)
    .batch_size(500)          // 每 500 行一批
    .concurrency(1)           // 批次并发（>1 需业务侧保证批次无序安全）
    .stop_on_error(true)
    .commit(source, Arc::new(ProductSink { /* ... */ }))
    .await?;
# Ok(()) }
```

### 3）导出：游标分页 + 流式写

```rust
use excel::{ExportRunner, Page, SheetOptions};

# async fn demo(db: sea_orm::DatabaseConnection) -> Result<(), excel::ExcelError> {
let runner = ExportRunner::new(SheetOptions::new("选品数据").columns(ProductRow::columns()));
let stats = runner
    .export_xlsx("uploads/export/products.xlsx", |cursor| {
        let db = db.clone();
        async move {
            let mut q = product::Entity::find()
                .order_by_asc(product::Column::Id)
                .limit(1000);
            if let Some(last_id) = cursor {
                q = q.filter(product::Column::Id.gt(last_id));
            }
            let rows = q.all(&db).await.map_err(|e| excel::ExcelError::Sink(e.to_string()))?;
            let done = rows.len() < 1000;
            let last_id = rows.last().map(|r| r.id.clone());
            let rows = rows.into_iter().map(Into::into).collect();
            Ok(match last_id {
                Some(id) if !done => Page::more(rows, id),
                _ => Page::last(rows),
            })
        }
    })
    .await?;
# Ok(()) }
```

## 图片支持

**导入**：图片直接插在单元格里即可（不需要「文件名 + zip 资料包」那套约定）。

- Excel 标准 drawing 锚点（`xl/drawings/drawingN.xml` → `xl/media/*`）；

- WPS 单元格嵌入图片 `=DISPIMG("ID_xxx",1)`（`xl/cellimages.xml`）；

- 图片字节按需读取，只有模型声明了图片列时才读对应列，避免百万行文件把图片全量读进内存；

- 导出再导入能保住原文件名：写侧把 `CellImage::name` 写进图片 alt 文本（`cNvPr@descr`），读侧优先取它，取不到才回退 `image1.png`。

**导出**：

```rust
use excel::{CellImage, CellValue, SheetOptions, WriteOptions};

let sheet = SheetOptions::new("带图数据")
    .columns(ProductRow::columns())
    .write_options(WriteOptions::default().image_row_height(72.0));
// 行值里放 CellValue::Image(CellImage::new(bytes, "png").named("main.png"))
```

- 默认写标准锚点图片（`insert_image_fit_to_cell_centered`），**常量内存模式下可用**；

- `WriteOptions::cell_image()` 会写 Excel「置于单元格内」图片（随筛选/排序移动）；
  该方法**自动关闭常量内存**（`rust_xlsxwriter` 的内嵌图片不支持常量内存模式）；

- 行高在写单元格之前设置，图片按行高自适应缩放。

## 百万级要点

| 环节   | 做法                                                 | 内存                               |
| ---- | -------------------------------------------------- | -------------------------------- |
| 读取   | `quick-xml` SAX 逐行，`next_row()` 一次一行               | ≈ 单行                             |
| 解析并发 | `spawn_blocking` 解析 + 有界 `mpsc` 通道背压               | ≈ `queue_batches × batch_size` 行 |
| 落库   | `BatchSink::save(batch)`，一批一事务；`concurrency` 控制批并发 | 由业务侧决定                           |
| 写出   | `constant_memory` 工作表，只能按行号递增写                     | ≈ 单行缓冲                           |
| 导出拉取 | 游标分页 `Page`，`queue_pages` 控制预取页数                   | ≈ `queue_pages × page_size` 行    |

注意事项：

- 常量内存模式**只能顺序写行**（`SheetWriter` 用 `row_cursor` 强制推进，避免静默丢行）；

- 拉取失败时不会留下半截导出文件（自动删除），避免「能下载但内容不全」；

- 行级错误默认跳过并记录（`max_errors` 截断），`continue_on_row_error(false)` 可改成首错中断。

## 多工作表

```rust
// 小数据量：一次写多个 sheet
runner.export_sheets("out.xlsx", vec![
    (SheetOptions::new("汇总").columns(ProductRow::columns()), summary_rows),
    (SheetOptions::new("明细").columns(ProductRow::columns()), detail_rows),
])?;

// 大数据量：每个 sheet 独立游标分页，顺序写出（单表仍是常量内存）
runner.export_multi_sheet("out.xlsx", sheet_options, |sheet_index, cursor| async move {
    // 按 sheet_index 拉不同的筛选条件
    # unimplemented!()
}).await?;
```

## 导入模板

```rust
use excel::{build_template, ColumnDef, ReferenceSheet, TemplateSpec};

let bytes = build_template(&TemplateSpec::new(ProductRow::columns())
    .sheet_name("产品导入")
    .title("产品导入模板")
    .note("图片直接粘贴到单元格")
    .sample_row(vec![/* 示例行 */])
    .reference(ReferenceSheet::new("工厂列表", vec!["工厂名称".into(), "短码".into()])
        .rows(vec![vec!["工厂A".into(), "A".into()]])))?;
```

生成的工作簿包含：数据 sheet（必填列红字表头 + 示例行 + 下拉验证）、
「填写说明」sheet、若干引用 sheet。

## 动态表头适配

`HeaderOptions` 的匹配顺序（可在 `ImportRunner` / `ReadOptions` 上调整）：

1. **精确**：归一化后一致（去空白、全角转半角、去 `*` 必填标记）；
2. **别名**：命中 `ColumnDef::aliases`；
3. **模糊**：互相包含，取差异最小（如 `报价(USD)(含税)` 命中 `报价(USD)`）；
4. **列序回退**：显式开启后，表头缺失时按声明顺序落到第 N 列。

`analyze_header` / `analyze_all_headers` 可只做预检，把「匹配到哪些列、缺哪些必填列、
多出哪些列」返回给前端确认。

## 执行器（推荐）：注册一次，全链路通用

不想每次手拼 `ImportRunner` / `ExportRunner` 的，给行模型挂一个派生宏即可。
宏自动创建执行器工厂并**静态注册进全局表**（`inventory` 链接期收集，零初始化），
业务端既不用管理工厂，也不用写任何注册代码：

```rust
use excel::{ExcelExecutor, ExcelRow, ZipSource};
use serde::Serialize;

#[derive(ExcelRow, ExcelExecutor, Serialize)]
#[excel(sheet = "产品导入", register = "product")]
struct ProductRow {
    #[excel(header = "品名", required)]
    name: String,
    // ...
}

# async fn demo(sink: std::sync::Arc<dyn excel::BatchSink<ProductRow>>) -> Result<(), excel::ExcelError> {
let source = ZipSource::open("products.xlsx")?;

// ① 按类型：编译期类型安全，一个入口覆盖全部
let exec = excel::executor_for::<ProductRow>();
let preview = exec.preview(source.clone(), 200)?;        // 校验预览
let report = exec.commit(source.clone(), sink).await?;   // 分批落库
let bytes = exec.export_bytes(&[])?;                     // 导出到内存（HTTP 下载）
exec.export_xlsx("out.xlsx", /* 游标分页闭包 */ |cursor| async move { todo!() }).await?;

// ② 按注册名：字符串调度（Web 层按请求参数路由到不同模型）
let json = excel::preview_by_name("product", source, 200)?;
println!("已注册执行器：{:?}", excel::registered_names());
# Ok(()) }
```

- `register = "name"` 是全局注册名，**必填**；模型需实现 `Serialize`（按名预览返回 JSON）；
- 需要微调默认行为时，用 `exec.import_runner()` / `exec.export_runner()` 拿到切好的 runner 再改；
- 全部能力：`preview` / `preview_json` / `commit` / `for_each_row` / `export_bytes` /
  `export_rows` / `export_sheets` / `export_xlsx` / `export_multi_sheet` / `template`。

## 错误与国际化

库内错误统一收敛为 `ExcelError`，`ExcelError::code()` 给出 `excel_*` key；
业务侧实现 `From<ExcelError> for AppError` 时可按 `@excel_xxx:原始信息` 拼接，
由全局 i18n 中间件按 `Accept-Language` / `x-locale` 翻译，和业务错误走同一条链路。

## 测试

```bash
cargo test -p excel                                                   # 单测 + 往返 + 文档测试
cargo test -p excel --test roundtrip -- --ignored --nocapture         # 100 万行流式压力测试
```

`stress_one_million_rows_streaming` 会真实写出 100 万行 xlsx 再逐行读回，
验证「边写边落盘 + SAX 逐行读」不随行数增长内存，默认不参与 CI。
