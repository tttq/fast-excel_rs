//! 读侧：流式 xlsx 读取（动态表头 / 多 sheet / 单元格图片）
//!
//! ```no_run
//! use fast_excel::read::{ExcelReader, ReadOptions, SheetSelector};
//! # fn main() -> Result<(), fast_excel::ExcelError> {
//! let reader = ExcelReader::open("products.xlsx")?;
//! println!("工作表：{:?}", reader.sheet_names());
//! let options = ReadOptions::new().header_row(0).with_images();
//! let mut stream = reader.stream(SheetSelector::First, &[], options)?;
//! while let Some(row) = stream.next_row()? {
//!     println!("第 {} 行：{}", row.row_index, row.text_of("品名"));
//! }
//! # Ok(()) }
//! ```

pub mod images;
pub mod shared_strings;
pub mod source;
pub mod stream;
pub mod styles;
pub mod workbook;

use std::path::Path;
use std::sync::Arc;

use crate::column::ColumnDef;
use crate::error::{ExcelError, Result};
use crate::header::{HeaderAnalysis, HeaderMap, HeaderOptions};

pub use crate::value::CellImage;
pub use images::{
    ExtractedImages, extract_all_images, extract_embedded_images, extract_images_of_sheet,
    extract_wps_cellimages, parse_dispimg_id,
};
pub use shared_strings::SharedStrings;
pub use source::{EntryReader, ZipSource};
pub use stream::SheetStream;
pub use styles::Styles;
pub use workbook::SheetInfo;

/// 工作表选择器
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SheetSelector {
    /// 第一个工作表
    First,
    /// 按下标（0 基）
    Index(u32),
    /// 按名称（精确）
    Name(String),
    /// 按名称集合（多 sheet 导入：逐个读，顺序与传入一致）
    Names(Vec<String>),
    /// 全部工作表（多 sheet 导入）
    All,
}

impl SheetSelector {
    pub fn name(name: impl Into<String>) -> Self {
        SheetSelector::Name(name.into())
    }

    pub fn names<I, S>(names: I) -> Self
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        SheetSelector::Names(names.into_iter().map(Into::into).collect())
    }
}

/// 读取选项
#[derive(Debug, Clone)]
pub struct ReadOptions {
    /// 表头定位与匹配策略
    pub header: HeaderOptions,
    /// 最多产出多少行（预览 / 抽样用）
    pub max_rows: Option<u64>,
    /// 跳过整行空白
    pub skip_empty_rows: bool,
    /// 是否读取单元格内嵌图片
    pub read_images: bool,
    /// 只读这些文件列（0 基）的图片；`None` = 按模型声明自动推导
    pub image_columns: Option<Vec<u16>>,
    /// 文本是否 trim
    pub trim_text: bool,
}

impl Default for ReadOptions {
    fn default() -> Self {
        Self {
            header: HeaderOptions::default(),
            max_rows: None,
            skip_empty_rows: true,
            read_images: false,
            image_columns: None,
            trim_text: true,
        }
    }
}

impl ReadOptions {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn header_options(mut self, header: HeaderOptions) -> Self {
        self.header = header;
        self
    }

    pub fn header_row(mut self, row_index: u32) -> Self {
        self.header.row_index = row_index;
        self
    }

    pub fn header_span(mut self, span: u32) -> Self {
        self.header.row_span = span.max(1);
        self
    }

    pub fn max_rows(mut self, rows: u64) -> Self {
        self.max_rows = Some(rows);
        self
    }

    /// 打开图片读取（模型里声明了图片列时，只读这些列）
    pub fn with_images(mut self) -> Self {
        self.read_images = true;
        self
    }

    /// 显式指定要读图的文件列
    pub fn image_columns(mut self, cols: Vec<u16>) -> Self {
        self.image_columns = Some(cols);
        self
    }

    pub fn skip_empty_rows(mut self, skip: bool) -> Self {
        self.skip_empty_rows = skip;
        self
    }

    pub fn trim_text(mut self, trim: bool) -> Self {
        self.trim_text = trim;
        self
    }
}

/// xlsx 读取器（可同时解析多个工作表；每次 `stream` 都是独立的流式读取）
pub struct ExcelReader {
    source: ZipSource,
    sheets: Vec<SheetInfo>,
    date_1904: bool,
    shared: Arc<SharedStrings>,
    styles: Arc<Styles>,
}

impl ExcelReader {
    /// 打开磁盘文件（推荐：不把整个文件读进内存）
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::from_source(ZipSource::open(path)?)
    }

    /// 以内存字节打开（HTTP 上传的 multipart 场景）
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_source(ZipSource::from_bytes(bytes))
    }

    /// 复用已有数据源
    pub fn from_source(source: ZipSource) -> Result<Self> {
        let (sheets, date_1904) = workbook::load_sheets(&source)?;
        // 共享字符串与样式表体积小、且需要随机访问，策略是整体加载一次
        let shared = match source.entry_reader("xl/sharedStrings.xml") {
            Ok(reader) => SharedStrings::parse(std::io::BufReader::new(reader))?,
            Err(_) => SharedStrings::empty(),
        };
        let styles = match source.entry_reader("xl/styles.xml") {
            Ok(reader) => Styles::parse(std::io::BufReader::new(reader))?,
            Err(_) => Styles::empty(),
        };
        Ok(Self {
            source,
            sheets,
            date_1904,
            shared: Arc::new(shared),
            styles: Arc::new(styles),
        })
    }

    pub fn source(&self) -> &ZipSource {
        &self.source
    }

    pub fn sheets(&self) -> &[SheetInfo] {
        &self.sheets
    }

    pub fn sheet_names(&self) -> Vec<String> {
        self.sheets.iter().map(|s| s.name.clone()).collect()
    }

    pub fn date_1904(&self) -> bool {
        self.date_1904
    }

    pub fn shared_strings(&self) -> &SharedStrings {
        &self.shared
    }

    /// 解析选择器 → 工作表列表
    pub fn select(&self, selector: &SheetSelector) -> Result<Vec<&SheetInfo>> {
        let found: Vec<&SheetInfo> = match selector {
            SheetSelector::First => self.sheets.first().into_iter().collect(),
            SheetSelector::Index(i) => self.sheets.iter().filter(|s| s.index == *i).collect(),
            SheetSelector::Name(name) => self.sheets.iter().filter(|s| &s.name == name).collect(),
            SheetSelector::Names(names) => names
                .iter()
                .map(|n| {
                    self.sheets
                        .iter()
                        .find(|s| &s.name == n)
                        .ok_or_else(|| ExcelError::SheetNotFound(n.clone()))
                })
                .collect::<Result<Vec<_>>>()?,
            SheetSelector::All => self.sheets.iter().collect(),
        };
        if found.is_empty() {
            return Err(ExcelError::SheetNotFound(match selector {
                SheetSelector::First => "第一个工作表".to_string(),
                SheetSelector::Index(i) => format!("下标 {i}"),
                SheetSelector::Name(n) => n.clone(),
                SheetSelector::Names(n) => n.join("、"),
                SheetSelector::All => "任意工作表".to_string(),
            }));
        }
        Ok(found)
    }

    /// 打开一个工作表的行流（单个 sheet）
    pub fn stream(
        &self,
        selector: SheetSelector,
        defs: &[ColumnDef],
        options: ReadOptions,
    ) -> Result<SheetStream> {
        let sheets = self.select(&selector)?;
        let sheet = sheets[0];
        self.stream_sheet(sheet, defs, options)
    }

    /// 打开指定工作表的行流
    pub fn stream_sheet(
        &self,
        sheet: &SheetInfo,
        defs: &[ColumnDef],
        options: ReadOptions,
    ) -> Result<SheetStream> {
        SheetStream::new(
            self.source.clone(),
            &sheet.path,
            &sheet.name,
            self.date_1904,
            self.shared.clone(),
            self.styles.clone(),
            defs,
            options,
        )
    }

    /// 只做表头自适应分析（不消费数据，供"导入预检"接口使用）
    pub fn analyze_header(
        &self,
        selector: &SheetSelector,
        defs: &[ColumnDef],
        options: ReadOptions,
    ) -> Result<HeaderAnalysis> {
        let sheets = self.select(selector)?;
        let mut stream = self.stream_sheet(sheets[0], defs, options)?;
        // 触发一次解析即可确保表头已构建（空文件也会在 EOF 处构建）
        let _ = stream.next_row()?;
        Ok(stream
            .header_analysis()
            .unwrap_or_else(|| HeaderMap::dynamic(Vec::new()).analysis()))
    }

    /// 读取所有工作表的表头分析（多 sheet 预检）
    pub fn analyze_all_headers(
        &self,
        defs: &[ColumnDef],
        options: ReadOptions,
    ) -> Result<Vec<(String, HeaderAnalysis)>> {
        let mut out = Vec::with_capacity(self.sheets.len());
        for sheet in &self.sheets {
            let mut stream = self.stream_sheet(sheet, defs, options.clone())?;
            let _ = stream.next_row()?;
            let analysis = stream
                .header_analysis()
                .unwrap_or_else(|| HeaderMap::dynamic(Vec::new()).analysis());
            out.push((sheet.name.clone(), analysis));
        }
        Ok(out)
    }
}
