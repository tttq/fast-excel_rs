//! 单元格值模型 + 类型转换（读时 `from_cell` 解析、写时 `into_cell` 生成）
//!
//! 解析侧的宽容度是刻意设计的：Excel 里同一个"数字列"可能是文本、带千分位、
//! 带货币符号或全角字符，[`CellValue::to_f64`] / [`CellValue::to_i64`] 统一兜底，
//! 避免业务侧各写一套 `parse().unwrap_or_default()`。

use chrono::{DateTime, NaiveDate, NaiveDateTime, TimeZone, Utc};
use rust_decimal::Decimal;

/// 单元格内嵌图片
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CellImage {
    pub bytes: Vec<u8>,
    /// 扩展名（小写，不带点）：png / jpg / gif / webp / bmp
    pub ext: String,
    /// 原始文件名或占位名（仅用于展示 / 落库时取名字）
    pub name: Option<String>,
}

impl CellImage {
    pub fn new(bytes: Vec<u8>, ext: impl Into<String>) -> Self {
        Self {
            bytes,
            ext: ext.into(),
            name: None,
        }
    }

    pub fn named(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    pub fn is_empty(&self) -> bool {
        self.bytes.is_empty()
    }

    pub fn mime(&self) -> &'static str {
        mime_from_ext(&self.ext)
    }
}

/// 扩展名 → mime
pub fn mime_from_ext(ext: &str) -> &'static str {
    match ext.to_ascii_lowercase().as_str() {
        "png" => "image/png",
        "gif" => "image/gif",
        "webp" => "image/webp",
        "bmp" => "image/bmp",
        "emf" => "image/emf",
        "wmf" => "image/wmf",
        _ => "image/jpeg",
    }
}

/// 单元格值
#[derive(Debug, Clone, PartialEq, Default)]
pub enum CellValue {
    /// 空单元格
    #[default]
    Empty,
    /// 文本
    Text(String),
    /// 数值
    Number(f64),
    /// 布尔
    Bool(bool),
    /// 日期
    Date(NaiveDate),
    /// 日期时间（无时区，Excel 语义）
    DateTime(NaiveDateTime),
    /// 公式（写出时按公式写；读入时保留原始公式文本）
    Formula(String),
    /// Excel 错误值（#DIV/0! 等）
    Error(String),
    /// 图片（写出时嵌入单元格；读入来自内嵌图片）
    Image(CellImage),
}

impl CellValue {
    /// 是否为空（空白文本视为空）
    pub fn is_empty(&self) -> bool {
        match self {
            CellValue::Empty => true,
            CellValue::Text(s) => s.trim().is_empty(),
            _ => false,
        }
    }

    pub fn type_name(&self) -> &'static str {
        match self {
            CellValue::Empty => "空",
            CellValue::Text(_) => "文本",
            CellValue::Number(_) => "数字",
            CellValue::Bool(_) => "布尔",
            CellValue::Date(_) => "日期",
            CellValue::DateTime(_) => "日期时间",
            CellValue::Formula(_) => "公式",
            CellValue::Error(_) => "错误值",
            CellValue::Image(_) => "图片",
        }
    }

    /// 显示文本（导入预览 / 错误提示 / 导出兜底都用它）
    pub fn to_text(&self) -> String {
        match self {
            CellValue::Empty => String::new(),
            CellValue::Text(s) => s.clone(),
            CellValue::Number(n) => fmt_number(*n),
            CellValue::Bool(b) => {
                if *b {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            }
            CellValue::Date(d) => d.format("%Y-%m-%d").to_string(),
            CellValue::DateTime(dt) => dt.format("%Y-%m-%d %H:%M:%S").to_string(),
            CellValue::Formula(f) => f.clone(),
            CellValue::Error(e) => e.clone(),
            // 图片没有文本表示：用占位名，便于预览展示"已识别到图片"
            CellValue::Image(img) => img.name.clone().unwrap_or_default(),
        }
    }

    /// 去空白后的显示文本
    pub fn trimmed(&self) -> String {
        self.to_text().trim().to_string()
    }

    pub fn to_f64(&self) -> Option<f64> {
        match self {
            CellValue::Number(n) => Some(*n),
            CellValue::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
            CellValue::Date(d) => Some(date_to_serial(*d, false)),
            CellValue::DateTime(dt) => Some(datetime_to_serial(*dt, false)),
            CellValue::Text(s) => parse_number(s),
            CellValue::Formula(f) => parse_number(f),
            _ => None,
        }
    }

    pub fn to_i64(&self) -> Option<i64> {
        match self {
            CellValue::Text(s) => parse_number(s).map(round_to_i64),
            CellValue::Formula(f) => parse_number(f).map(round_to_i64),
            other => other.to_f64().map(round_to_i64),
        }
    }

    pub fn to_bool(&self) -> Option<bool> {
        match self {
            CellValue::Bool(b) => Some(*b),
            CellValue::Number(n) => Some(*n != 0.0),
            CellValue::Text(s) => match normalize_text(s).to_ascii_lowercase().as_str() {
                "true" | "yes" | "y" | "1" | "是" | "√" | "对" => Some(true),
                "false" | "no" | "n" | "0" | "否" | "×" | "错" => Some(false),
                _ => None,
            },
            _ => None,
        }
    }

    pub fn to_decimal(&self) -> Option<Decimal> {
        match self {
            // 优先按"最短十进制表示"解析：Excel 里的 19.99 不该变成 19.98999999999999843680598132
            CellValue::Number(n) => fmt_number(*n)
                .parse::<Decimal>()
                .ok()
                .or_else(|| Decimal::from_f64_retain(*n)),
            other => {
                let t = normalize_number(other.to_text().as_str())?;
                t.parse::<Decimal>().ok()
            }
        }
    }

    pub fn to_date(&self) -> Option<NaiveDate> {
        match self {
            CellValue::Date(d) => Some(*d),
            CellValue::DateTime(dt) => Some(dt.date()),
            CellValue::Number(n) => Some(serial_to_datetime(*n, false)?.date()),
            other => parse_date_text(&other.to_text()),
        }
    }

    pub fn to_datetime(&self) -> Option<NaiveDateTime> {
        match self {
            CellValue::DateTime(dt) => Some(*dt),
            CellValue::Date(d) => d.and_hms_opt(0, 0, 0),
            CellValue::Number(n) => serial_to_datetime(*n, false),
            other => parse_datetime_text(&other.to_text()),
        }
    }
}

impl From<&str> for CellValue {
    fn from(v: &str) -> Self {
        CellValue::Text(v.to_string())
    }
}
impl From<String> for CellValue {
    fn from(v: String) -> Self {
        CellValue::Text(v)
    }
}
impl From<f64> for CellValue {
    fn from(v: f64) -> Self {
        CellValue::Number(v)
    }
}
impl From<i64> for CellValue {
    fn from(v: i64) -> Self {
        CellValue::Number(v as f64)
    }
}
impl From<i32> for CellValue {
    fn from(v: i32) -> Self {
        CellValue::Number(v as f64)
    }
}
impl From<bool> for CellValue {
    fn from(v: bool) -> Self {
        CellValue::Bool(v)
    }
}

impl From<CellValue> for String {
    fn from(v: CellValue) -> Self {
        v.to_text()
    }
}

/// 单元格（含位置）
#[derive(Debug, Clone, PartialEq)]
pub struct Cell {
    /// Excel 行号（1 基，与界面一致）
    pub row: u32,
    /// 列索引（0 基）
    pub col: u16,
    pub value: CellValue,
    /// 原始样式索引（判断日期格式 / 自定义格式用）
    pub style: u32,
}

impl Cell {
    pub fn new(row: u32, col: u16, value: CellValue) -> Self {
        Self {
            row,
            col,
            value,
            style: 0,
        }
    }

    pub fn text(&self) -> String {
        self.value.to_text()
    }

    pub fn is_empty(&self) -> bool {
        self.value.is_empty()
    }
}
// ─────────────────────────── 文本归一化 / 解析工具 ───────────────────────────

/// 全角字符（含全角空格）转半角，去掉 BOM 与制表符
pub fn normalize_text(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.trim().trim_start_matches('\u{feff}').chars() {
        match c {
            '\u{3000}' => out.push(' '),
            '\u{ff01}'..='\u{ff5e}' => {
                // 全角 ASCII → 半角
                out.push(char::from_u32(c as u32 - 0xfee0).unwrap_or(c))
            }
            '\u{feff}' => {}
            c => out.push(c),
        }
    }
    out
}

/// 数字文本归一化：去货币符号 / 千分位 / 空白；百分号换算成小数
pub fn normalize_number(s: &str) -> Option<String> {
    let t = normalize_text(s);
    if t.is_empty() {
        return None;
    }
    let (body, percent) = match t.strip_suffix('%') {
        Some(rest) => (rest, true),
        None => (t.as_str(), false),
    };
    let cleaned: String = body
        .chars()
        .filter(|c| !matches!(c, ',' | ' ' | '¥' | '￥' | '$' | '€' | '£' | '\'' | '_'))
        .collect();
    if cleaned.is_empty() {
        return None;
    }
    if percent {
        let v: f64 = cleaned.parse().ok()?;
        Some(((v / 100.0).to_string()).parse().unwrap_or(cleaned))
    } else {
        Some(cleaned)
    }
}

/// 宽松数字解析（文本 / 千分位 / 货币符号 / 百分号）
pub fn parse_number(s: &str) -> Option<f64> {
    let t = normalize_text(s);
    if t.is_empty() {
        return None;
    }
    let (body, percent) = match t.strip_suffix('%') {
        Some(rest) => (rest, true),
        None => (t.as_str(), false),
    };
    let cleaned: String = body
        .chars()
        .filter(|c| !matches!(c, ',' | ' ' | '¥' | '￥' | '$' | '€' | '£' | '\'' | '_'))
        .collect();
    let v: f64 = cleaned.parse().ok()?;
    Some(if percent { v / 100.0 } else { v })
}

fn round_to_i64(v: f64) -> i64 {
    if v.fract() == 0.0 {
        v as i64
    } else {
        v.round() as i64
    }
}

/// 日期文本解析（支持 `2024-01-02` / `2024/1/2` / `2024.1.2` / `2024年1月2日` / `20240102`）
pub fn parse_date_text(s: &str) -> Option<NaiveDate> {
    let t = normalize_text(s)
        .replace(['年', '月'], "-")
        .replace('日', "")
        .replace(['.', '/'], "-");
    let t = t.trim();
    if t.is_empty() {
        return None;
    }
    for f in [
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y%m%d",
        "%Y-%m",
    ] {
        if let Ok(d) = NaiveDate::parse_from_str(t, f) {
            return Some(d);
        }
        if let Ok(dt) = NaiveDateTime::parse_from_str(t, f) {
            return Some(dt.date());
        }
    }
    None
}

/// 日期时间文本解析
pub fn parse_datetime_text(s: &str) -> Option<NaiveDateTime> {
    let t = normalize_text(s)
        .replace(['年', '月'], "-")
        .replace('日', "")
        .replace(['.', '/', 'T'], "-");
    let t = t.trim();
    if t.is_empty() {
        return None;
    }
    for f in [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%Y%m%d%H%M%S",
        "%Y%m%d",
    ] {
        if let Ok(dt) = NaiveDateTime::parse_from_str(t, f) {
            return Some(dt);
        }
    }
    parse_date_text(t).and_then(|d| d.and_hms_opt(0, 0, 0))
}

/// 数字 → 显示文本（整数不带 `.0`，避免导出 1 显示成 1.0）
pub fn fmt_number(v: f64) -> String {
    if v.is_nan() {
        return "NaN".to_string();
    }
    if v.fract() == 0.0 && v.abs() < 1e15 {
        format!("{}", v as i64)
    } else {
        let s = format!("{v}");
        s
    }
}

// ─────────────────────────── Excel 日期序列号 ───────────────────────────

/// Excel 序列号 → 日期时间
///
/// 1900 日期系统以 `1899-12-30` 为原点（自动跳过 Excel 的 1900 闰年 bug）；
/// 1904 日期系统（Mac 老文件）以 `1904-01-01` 为原点。
pub fn serial_to_datetime(serial: f64, date_1904: bool) -> Option<NaiveDateTime> {
    if !serial.is_finite() || serial < 0.0 {
        return None;
    }
    // 1900 日期系统：序列号 < 61 时基点为 1899-12-31（Excel 把 1900 当闰年，
    // 序列号 60 = 不存在的 1900-02-29），>= 61 时基点为 1899-12-30
    let base = if date_1904 {
        NaiveDate::from_ymd_opt(1904, 1, 1)?
    } else if serial < 61.0 {
        NaiveDate::from_ymd_opt(1899, 12, 31)?
    } else {
        NaiveDate::from_ymd_opt(1899, 12, 30)?
    };
    let ms = (serial * 86_400_000.0).round() as i64;
    Some(base.and_hms_opt(0, 0, 0)? + chrono::Duration::milliseconds(ms))
}

/// 日期时间 → Excel 序列号
pub fn datetime_to_serial(dt: NaiveDateTime, date_1904: bool) -> f64 {
    let early = dt.date() < NaiveDate::from_ymd_opt(1900, 3, 1).unwrap();
    let base = if date_1904 {
        NaiveDate::from_ymd_opt(1904, 1, 1).unwrap()
    } else if early {
        NaiveDate::from_ymd_opt(1899, 12, 31).unwrap()
    } else {
        NaiveDate::from_ymd_opt(1899, 12, 30).unwrap()
    };
    let base = base.and_hms_opt(0, 0, 0).unwrap();
    let dur = dt - base;
    dur.num_milliseconds() as f64 / 86_400_000.0
}

/// 日期 → Excel 序列号
pub fn date_to_serial(d: NaiveDate, date_1904: bool) -> f64 {
    datetime_to_serial(d.and_hms_opt(0, 0, 0).unwrap(), date_1904)
}

// ─────────────────────────── 类型转换 trait ───────────────────────────

/// 从单元格解析为业务类型（错误信息直接面向用户，中文）
pub trait FromCell: Sized {
    fn from_cell(cell: &Cell) -> std::result::Result<Self, String>;

    /// 类型名（错误提示用）
    fn type_name() -> &'static str {
        "值"
    }

    /// 该列在文件里完全缺失时的取值：普通类型报错，`Option<T>` 返回 None
    fn from_missing_column() -> std::result::Result<Self, String> {
        Err("文件里缺少该列".to_string())
    }
}

/// 业务类型 → 单元格值
///
/// 与 [`FromCell`] 构成 `from_cell` / `into_cell` 对称命名，因此 `&self` 接收者被 clippy 点名，
/// 但改名会破坏公开 API，这里保持现状并压制该 lint。
#[allow(clippy::wrong_self_convention)]
pub trait IntoCell {
    fn into_cell(&self) -> CellValue;
}

impl FromCell for CellValue {
    fn from_cell(cell: &Cell) -> std::result::Result<Self, String> {
        Ok(cell.value.clone())
    }

    fn type_name() -> &'static str {
        "单元格值"
    }
}

impl FromCell for Cell {
    fn from_cell(cell: &Cell) -> std::result::Result<Self, String> {
        Ok(cell.clone())
    }

    fn type_name() -> &'static str {
        "单元格"
    }
}

impl FromCell for CellImage {
    fn from_cell(cell: &Cell) -> std::result::Result<Self, String> {
        match &cell.value {
            CellValue::Image(img) => Ok(img.clone()),
            other => Err(format!("期望图片，实际是{}", other.type_name())),
        }
    }

    fn type_name() -> &'static str {
        "图片"
    }
}

impl FromCell for String {
    fn from_cell(cell: &Cell) -> std::result::Result<Self, String> {
        Ok(cell.value.trimmed())
    }

    fn type_name() -> &'static str {
        "文本"
    }
}

impl<T: FromCell> FromCell for Option<T> {
    fn from_cell(cell: &Cell) -> std::result::Result<Self, String> {
        if cell.value.is_empty() {
            Ok(None)
        } else {
            T::from_cell(cell).map(Some)
        }
    }

    fn type_name() -> &'static str {
        "可选值"
    }

    fn from_missing_column() -> std::result::Result<Self, String> {
        Ok(None)
    }
}

impl FromCell for bool {
    fn from_cell(cell: &Cell) -> std::result::Result<Self, String> {
        cell.value
            .to_bool()
            .ok_or_else(|| format!("期望布尔值，实际「{}」", cell.value.to_text()))
    }

    fn type_name() -> &'static str {
        "布尔"
    }
}

macro_rules! impl_from_cell_int {
    ($($t:ty),* $(,)?) => {
        $(
            impl FromCell for $t {
                fn from_cell(cell: &Cell) -> std::result::Result<Self, String> {
                    let v = cell
                        .value
                        .to_i64()
                        .ok_or_else(|| format!("期望整数，实际「{}」", cell.value.to_text()))?;
                    <$t>::try_from(v)
                        .map_err(|_| format!("整数「{v}」超出 {} 取值范围", stringify!($t)))
                }

                fn type_name() -> &'static str {
                    "整数"
                }
            }
        )*
    };
}

impl_from_cell_int!(i8, i16, i32, i64, isize, u8, u16, u32, u64, usize);

impl FromCell for f32 {
    fn from_cell(cell: &Cell) -> std::result::Result<Self, String> {
        cell.value
            .to_f64()
            .map(|v| v as f32)
            .ok_or_else(|| format!("期望数字，实际「{}」", cell.value.to_text()))
    }

    fn type_name() -> &'static str {
        "数字"
    }
}

impl FromCell for f64 {
    fn from_cell(cell: &Cell) -> std::result::Result<Self, String> {
        cell.value
            .to_f64()
            .ok_or_else(|| format!("期望数字，实际「{}」", cell.value.to_text()))
    }

    fn type_name() -> &'static str {
        "数字"
    }
}

impl FromCell for Decimal {
    fn from_cell(cell: &Cell) -> std::result::Result<Self, String> {
        cell.value
            .to_decimal()
            .ok_or_else(|| format!("期望数字，实际「{}」", cell.value.to_text()))
    }

    fn type_name() -> &'static str {
        "金额 / 小数"
    }
}

impl FromCell for NaiveDate {
    fn from_cell(cell: &Cell) -> std::result::Result<Self, String> {
        cell.value
            .to_date()
            .ok_or_else(|| format!("期望日期，实际「{}」", cell.value.to_text()))
    }

    fn type_name() -> &'static str {
        "日期"
    }
}

impl FromCell for NaiveDateTime {
    fn from_cell(cell: &Cell) -> std::result::Result<Self, String> {
        cell.value
            .to_datetime()
            .ok_or_else(|| format!("期望日期时间，实际「{}」", cell.value.to_text()))
    }

    fn type_name() -> &'static str {
        "日期时间"
    }
}

impl FromCell for DateTime<Utc> {
    fn from_cell(cell: &Cell) -> std::result::Result<Self, String> {
        let dt = cell
            .value
            .to_datetime()
            .ok_or_else(|| format!("期望日期时间，实际「{}」", cell.value.to_text()))?;
        Ok(Utc.from_utc_datetime(&dt))
    }

    fn type_name() -> &'static str {
        "日期时间(UTC)"
    }
}

impl FromCell for serde_json::Value {
    fn from_cell(cell: &Cell) -> std::result::Result<Self, String> {
        Ok(match &cell.value {
            CellValue::Empty => serde_json::Value::Null,
            CellValue::Number(n) => serde_json::json!(n),
            CellValue::Bool(b) => serde_json::json!(b),
            other => serde_json::Value::String(other.to_text()),
        })
    }

    fn type_name() -> &'static str {
        "JSON 值"
    }
}

impl IntoCell for CellValue {
    fn into_cell(&self) -> CellValue {
        self.clone()
    }
}

impl IntoCell for Cell {
    fn into_cell(&self) -> CellValue {
        self.value.clone()
    }
}

impl IntoCell for str {
    fn into_cell(&self) -> CellValue {
        CellValue::Text(self.to_string())
    }
}

impl IntoCell for String {
    fn into_cell(&self) -> CellValue {
        CellValue::Text(self.clone())
    }
}

impl IntoCell for &String {
    fn into_cell(&self) -> CellValue {
        CellValue::Text((*self).clone())
    }
}

impl IntoCell for bool {
    fn into_cell(&self) -> CellValue {
        CellValue::Bool(*self)
    }
}

impl IntoCell for f64 {
    fn into_cell(&self) -> CellValue {
        CellValue::Number(*self)
    }
}

impl IntoCell for f32 {
    fn into_cell(&self) -> CellValue {
        CellValue::Number(*self as f64)
    }
}

impl IntoCell for Decimal {
    fn into_cell(&self) -> CellValue {
        CellValue::Number(self.to_string().parse().unwrap_or_default())
    }
}

impl IntoCell for NaiveDate {
    fn into_cell(&self) -> CellValue {
        CellValue::Date(*self)
    }
}

impl IntoCell for NaiveDateTime {
    fn into_cell(&self) -> CellValue {
        CellValue::DateTime(*self)
    }
}

impl IntoCell for DateTime<Utc> {
    fn into_cell(&self) -> CellValue {
        CellValue::DateTime(self.naive_utc())
    }
}

impl IntoCell for CellImage {
    fn into_cell(&self) -> CellValue {
        CellValue::Image(self.clone())
    }
}

macro_rules! impl_into_cell_num {
    ($($t:ty),* $(,)?) => {
        $(
            impl IntoCell for $t {
                fn into_cell(&self) -> CellValue {
                    CellValue::Number(*self as f64)
                }
            }
        )*
    };
}

impl_into_cell_num!(i8, i16, i32, i64, isize, u8, u16, u32, u64, usize);

impl<T: IntoCell> IntoCell for Option<T> {
    fn into_cell(&self) -> CellValue {
        match self {
            Some(v) => v.into_cell(),
            None => CellValue::Empty,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_number_variants() {
        assert_eq!(parse_number("1,299.50"), Some(1299.5));
        assert_eq!(parse_number("￥1 299"), Some(1299.0));
        assert_eq!(parse_number("１２"), Some(12.0));
        assert_eq!(parse_number("50%"), Some(0.5));
        assert_eq!(parse_number("abc"), None);
    }

    #[test]
    fn test_parse_date_variants() {
        let d = NaiveDate::from_ymd_opt(2024, 1, 2).unwrap();
        assert_eq!(parse_date_text("2024-01-02"), Some(d));
        assert_eq!(parse_date_text("2024/1/2"), Some(d));
        assert_eq!(parse_date_text("2024.1.2"), Some(d));
        assert_eq!(parse_date_text("2024年1月2日"), Some(d));
        assert_eq!(parse_date_text("20240102"), Some(d));
        assert_eq!(parse_date_text("not a date"), None);
    }

    #[test]
    fn test_serial_roundtrip() {
        let dt = NaiveDate::from_ymd_opt(2024, 1, 2)
            .unwrap()
            .and_hms_opt(12, 30, 0)
            .unwrap();
        let s = datetime_to_serial(dt, false);
        assert_eq!(serial_to_datetime(s, false), Some(dt));
        // Excel 著名的 1900 闰年 bug：序列号 1 = 1900-01-01
        assert_eq!(
            serial_to_datetime(1.0, false).unwrap().date(),
            NaiveDate::from_ymd_opt(1900, 1, 1).unwrap()
        );
    }
}
