//! 列定义：一个业务字段 = 一列（对标 EasyExcel 的 `@ExcelProperty`）

/// 列的语义类型：读时决定解析方式，写时决定单元格格式
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ColumnKind {
    /// 自动（按单元格实际类型解析）
    #[default]
    Any,
    /// 文本（导出时按文本格式写，长数字/手机号不丢前导 0、不变科学计数）
    Text,
    Integer,
    Number,
    /// 金额 / 高精度小数（对应 `rust_decimal::Decimal`）
    Decimal,
    Bool,
    Date,
    DateTime,
    /// 图片列：导入按行读内嵌图片，导出按行嵌入图片
    Image,
}

impl ColumnKind {
    pub fn is_image(self) -> bool {
        matches!(self, ColumnKind::Image)
    }

    pub fn is_numeric(self) -> bool {
        matches!(
            self,
            ColumnKind::Integer | ColumnKind::Number | ColumnKind::Decimal
        )
    }

    pub fn is_temporal(self) -> bool {
        matches!(self, ColumnKind::Date | ColumnKind::DateTime)
    }

    pub fn as_str(self) -> &'static str {
        match self {
            ColumnKind::Any => "any",
            ColumnKind::Text => "text",
            ColumnKind::Integer => "integer",
            ColumnKind::Number => "number",
            ColumnKind::Decimal => "decimal",
            ColumnKind::Bool => "bool",
            ColumnKind::Date => "date",
            ColumnKind::DateTime => "datetime",
            ColumnKind::Image => "image",
        }
    }
}

/// 列定义
#[derive(Debug, Clone)]
pub struct ColumnDef {
    /// 标准表头（模板/导出使用）
    pub header: String,
    /// 别名（动态表头适配：历史模板改过列名的场景）
    pub aliases: Vec<String>,
    /// 必填
    pub required: bool,
    /// 语义类型
    pub kind: ColumnKind,
    /// 导出列宽
    pub width: Option<f64>,
    /// 导出数值格式（如 `#,##0.00`、`yyyy-mm-dd`）
    pub format: Option<String>,
    /// 导出是否按文本格式写（`@`）
    pub text_format: bool,
    /// 是否图片列
    pub image: bool,
    /// 导入时忽略（仅参与模板/导出展示）
    pub ignore: bool,
    /// 空值默认值（模板示例行与导入兜底）
    pub default: Option<String>,
    /// 模板备注（写入说明 sheet）
    pub note: Option<String>,
    /// 模板下拉选项（数据验证）
    pub dropdown: Vec<String>,
}

impl ColumnDef {
    pub fn new(header: impl Into<String>) -> Self {
        Self {
            header: header.into(),
            aliases: Vec::new(),
            required: false,
            kind: ColumnKind::Any,
            width: None,
            format: None,
            text_format: false,
            image: false,
            ignore: false,
            default: None,
            note: None,
            dropdown: Vec::new(),
        }
    }

    pub fn alias(mut self, alias: impl Into<String>) -> Self {
        let a = alias.into();
        if !a.is_empty() && a != self.header {
            self.aliases.push(a);
        }
        self
    }

    pub fn aliases<I, S>(mut self, list: I) -> Self
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        for a in list {
            self = self.alias(a);
        }
        self
    }

    pub fn required(mut self) -> Self {
        self.required = true;
        self
    }

    pub fn optional(mut self) -> Self {
        self.required = false;
        self
    }

    pub fn kind(mut self, kind: ColumnKind) -> Self {
        self.kind = kind;
        if kind.is_image() {
            self.image = true;
        }
        self
    }

    pub fn width(mut self, width: f64) -> Self {
        self.width = Some(width);
        self
    }

    pub fn format(mut self, format: impl Into<String>) -> Self {
        self.format = Some(format.into());
        self
    }

    /// 文本格式（长数字、手机号、编码类字段）
    pub fn as_text(mut self) -> Self {
        self.text_format = true;
        self
    }

    pub fn image(mut self) -> Self {
        self.image = true;
        self.kind = ColumnKind::Image;
        self
    }

    pub fn ignore(mut self) -> Self {
        self.ignore = true;
        self
    }

    pub fn default_value(mut self, value: impl Into<String>) -> Self {
        self.default = Some(value.into());
        self
    }

    pub fn note(mut self, note: impl Into<String>) -> Self {
        self.note = Some(note.into());
        self
    }

    pub fn dropdown<I, S>(mut self, list: I) -> Self
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        self.dropdown = list.into_iter().map(Into::into).collect();
        self
    }

    /// 表头 + 别名（动态匹配用）
    pub fn names(&self) -> impl Iterator<Item = &str> {
        std::iter::once(self.header.as_str()).chain(self.aliases.iter().map(|s| s.as_str()))
    }

    /// 是否参与导入解析
    pub fn is_importable(&self) -> bool {
        !self.ignore
    }

    /// 该列是否需要在模板里标红（必填）
    pub fn mark_required(&self) -> bool {
        self.required
    }
}

impl From<&str> for ColumnDef {
    fn from(header: &str) -> Self {
        ColumnDef::new(header)
    }
}

impl From<String> for ColumnDef {
    fn from(header: String) -> Self {
        ColumnDef::new(header)
    }
}

impl From<(&str, ColumnKind)> for ColumnDef {
    fn from((header, kind): (&str, ColumnKind)) -> Self {
        ColumnDef::new(header).kind(kind)
    }
}
