//! 动态表头适配
//!
//! 目标：同一份业务模型能接住"各种版本 / 各种人改过列名"的 Excel。匹配顺序：
//!
//! 1. **精确**：归一化后与标准表头一致；
//! 2. **别名**：命中 [`ColumnDef::aliases`]；
//! 3. **模糊**：互相包含（`报价(USD)(单价)` 命中 `报价(USD)`），按差异最小的列优先；
//! 4. **列序回退**：表头缺失时按声明顺序落到第 N 列（需显式开启）。
//!
//! 归一化会去掉所有空白、`*` 必填标记，统一全角/半角与中英文括号、冒号。

use std::collections::HashMap;

use serde::Serialize;

use crate::column::{ColumnDef, ColumnKind};
use crate::value::normalize_text;

/// 表头定位与匹配选项
#[derive(Debug, Clone)]
pub struct HeaderOptions {
    /// 表头所在行（0 基，允许前面有标题行 / 说明行）
    pub row_index: u32,
    /// 表头占几行（多行表头按列纵向拼接，如 `报价\n(USD)`）
    pub row_span: u32,
    /// 是否启用模糊匹配（互相包含）
    pub fuzzy: bool,
    /// 是否归一化（去空白/全角半角/必填标记）
    pub normalize: bool,
    /// 未匹配到的文件列是否忽略（false = 收集到 `unknown` 由业务侧决定）
    pub ignore_unknown: bool,
    /// 表头完全缺失时按声明顺序回退到同序列
    pub allow_index_fallback: bool,
    /// 单元格文本是否 trim
    pub trim: bool,
}

impl Default for HeaderOptions {
    fn default() -> Self {
        Self {
            row_index: 0,
            row_span: 1,
            fuzzy: true,
            normalize: true,
            ignore_unknown: true,
            allow_index_fallback: false,
            trim: true,
        }
    }
}

impl HeaderOptions {
    /// 严格模式：只认精确表头（不做模糊、不许回退）
    pub fn strict() -> Self {
        Self {
            fuzzy: false,
            allow_index_fallback: false,
            ..Default::default()
        }
    }

    pub fn row_index(mut self, row_index: u32) -> Self {
        self.row_index = row_index;
        self
    }

    pub fn row_span(mut self, span: u32) -> Self {
        self.row_span = span.max(1);
        self
    }

    pub fn strict_match(mut self) -> Self {
        self.fuzzy = false;
        self
    }

    pub fn index_fallback(mut self, enable: bool) -> Self {
        self.allow_index_fallback = enable;
        self
    }

    pub fn normalize_enabled(mut self, enable: bool) -> Self {
        self.normalize = enable;
        self
    }
}

/// 匹配方式
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MatchKind {
    /// 与标准表头精确一致
    Exact,
    /// 命中别名
    Alias,
    /// 模糊包含
    Fuzzy,
    /// 按列序回退
    IndexFallback,
    /// 缺失（该列在本文件里没有对应列）
    Missing,
}

impl MatchKind {
    pub fn as_str(self) -> &'static str {
        match self {
            MatchKind::Exact => "exact",
            MatchKind::Alias => "alias",
            MatchKind::Fuzzy => "fuzzy",
            MatchKind::IndexFallback => "index",
            MatchKind::Missing => "missing",
        }
    }
}

/// 一个声明列的解析结果
#[derive(Debug, Clone)]
pub struct ResolvedColumn {
    /// 在声明列表中的下标
    pub order: usize,
    pub header: String,
    pub required: bool,
    pub kind: ColumnKind,
    /// 命中的文件列（0 基）；缺失为 `None`
    pub file_index: Option<u16>,
    /// 命中的文件表头原文
    pub file_header: Option<String>,
    pub matched_by: MatchKind,
}

impl ResolvedColumn {
    pub fn is_matched(&self) -> bool {
        self.file_index.is_some()
    }
}

/// 表头解析结果
#[derive(Debug, Clone)]
pub struct HeaderMap {
    /// 按声明顺序对齐的列（含缺失列）
    pub columns: Vec<ResolvedColumn>,
    /// 文件表头原文（按列序，多行表头已拼接）
    pub file_headers: Vec<String>,
    /// 文件名里没被任何声明列用到的列
    pub unmapped: Vec<(u16, String)>,
    /// 实际使用的表头行（0 基）
    pub header_row: u32,
    /// 表头占几行
    pub header_span: u32,
}

impl HeaderMap {
    /// 声明列 → 文件列索引
    pub fn index_of(&self, header: &str) -> Option<u16> {
        self.columns
            .iter()
            .find(|c| c.header == header && c.file_index.is_some())
            .and_then(|c| c.file_index)
    }

    pub fn resolved(&self, header: &str) -> Option<&ResolvedColumn> {
        self.columns.iter().find(|c| c.header == header)
    }

    /// 文件列 → 文件表头
    pub fn file_header(&self, col: u16) -> Option<&str> {
        self.file_headers.get(col as usize).map(|s| s.as_str())
    }

    /// 缺哪些必填列
    pub fn missing_required(&self) -> Vec<String> {
        self.columns
            .iter()
            .filter(|c| c.required && !c.is_matched())
            .map(|c| c.header.clone())
            .collect()
    }

    /// 未匹配的文件列
    pub fn unknown_columns(&self) -> Vec<String> {
        self.unmapped.iter().map(|(_, h)| h.clone()).collect()
    }

    /// 表头自适应报告（预览接口直接返回给前端）
    pub fn analysis(&self) -> HeaderAnalysis {
        HeaderAnalysis {
            header_row: self.header_row + 1,
            header_span: self.header_span,
            total_columns: self.file_headers.len(),
            matched: self
                .columns
                .iter()
                .filter(|c| c.is_matched())
                .map(|c| MatchedColumn {
                    header: c.header.clone(),
                    file_header: c.file_header.clone().unwrap_or_default(),
                    excel_column: excel_column_name(c.file_index.unwrap_or(0)),
                    matched_by: c.matched_by.as_str().to_string(),
                })
                .collect(),
            missing: self.missing_required(),
            unknown: self.unknown_columns(),
        }
    }

    /// 完全动态：不声明列，文件的每一列都收下来
    pub fn dynamic(file_headers: Vec<String>) -> Self {
        let columns = file_headers
            .iter()
            .enumerate()
            .map(|(i, h)| ResolvedColumn {
                order: i,
                header: h.clone(),
                required: false,
                kind: ColumnKind::Any,
                file_index: Some(i as u16),
                file_header: Some(h.clone()),
                matched_by: MatchKind::Exact,
            })
            .collect();
        Self {
            columns,
            file_headers,
            unmapped: Vec::new(),
            header_row: 0,
            header_span: 1,
        }
    }
}

/// 表头自适应报告（返回给前端的"列匹配结果"）
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HeaderAnalysis {
    /// 表头行号（1 基，前端展示用）
    pub header_row: u32,
    pub header_span: u32,
    pub total_columns: usize,
    pub matched: Vec<MatchedColumn>,
    /// 缺失的必填列
    pub missing: Vec<String>,
    /// 文件中未识别的列（可作为"多出来的列"提示）
    pub unknown: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MatchedColumn {
    /// 模型字段对应的标准表头
    pub header: String,
    /// 文件里的实际表头
    pub file_header: String,
    /// Excel 列号（A/B/AA...）
    pub excel_column: String,
    /// exact / alias / fuzzy / index
    pub matched_by: String,
}

/// 表头归一化
pub fn normalize_header(raw: &str) -> String {
    let t = normalize_text(raw);
    let mut out = String::with_capacity(t.len());
    for c in t.chars() {
        match c {
            ' ' | '\t' | '\r' | '\n' | '\u{3000}' | '\u{00a0}' => {}
            // 必填标记 / 常见装饰
            '*' | '★' | '☆' | '●' | '▲' => {}
            '（' => out.push('('),
            '）' => out.push(')'),
            '：' => out.push(':'),
            '，' | '、' => out.push(','),
            '／' => out.push('/'),
            '－' | '—' | '–' => out.push('-'),
            c => out.push(c.to_ascii_lowercase()),
        }
    }
    // 去掉尾部冒号
    while out.ends_with(':') || out.ends_with('(') {
        out.pop();
    }
    out
}

/// 列索引 → Excel 列号（0 → A，25 → Z，26 → AA）
pub fn excel_column_name(index: u16) -> String {
    let mut n = index as u32;
    let mut out = String::new();
    loop {
        let r = (n % 26) as u8;
        out.insert(0, (b'A' + r) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }
    out
}

/// 按声明列动态解析文件表头
pub fn resolve_headers(
    file_headers: &[String],
    defs: &[ColumnDef],
    opts: &HeaderOptions,
) -> HeaderMap {
    let norm: Vec<String> = file_headers
        .iter()
        .map(|h| normalize_key(h, opts))
        .collect();
    let mut claimed = vec![false; file_headers.len()];
    let mut columns = Vec::with_capacity(defs.len());

    for (order, def) in defs.iter().enumerate() {
        if !def.is_importable() {
            continue;
        }
        let target = normalize_key(&def.header, opts);
        let alias_keys: Vec<String> = def.aliases.iter().map(|a| normalize_key(a, opts)).collect();

        let mut matched: Option<(usize, MatchKind)> = None;

        // 1) 标准表头精确
        if let Some(i) = find_exact(&norm, &claimed, std::slice::from_ref(&target)) {
            matched = Some((i, MatchKind::Exact));
        }
        // 2) 别名
        if matched.is_none() {
            if let Some(i) = find_exact(&norm, &claimed, &alias_keys) {
                matched = Some((i, MatchKind::Alias));
            }
        }
        // 3) 模糊包含（差异最小的优先）
        if matched.is_none() && opts.fuzzy {
            let mut keys = vec![target.clone()];
            keys.extend(alias_keys.iter().cloned());
            if let Some(i) = find_fuzzy(&norm, &claimed, &keys) {
                matched = Some((i, MatchKind::Fuzzy));
            }
        }
        // 4) 列序回退
        if matched.is_none() && opts.allow_index_fallback && order < norm.len() && !claimed[order] {
            matched = Some((order, MatchKind::IndexFallback));
        }

        match matched {
            Some((i, kind)) => {
                claimed[i] = true;
                columns.push(ResolvedColumn {
                    order,
                    header: def.header.clone(),
                    required: def.required,
                    kind: def.kind,
                    file_index: Some(i as u16),
                    file_header: Some(file_headers[i].clone()),
                    matched_by: kind,
                });
            }
            None => columns.push(ResolvedColumn {
                order,
                header: def.header.clone(),
                required: def.required,
                kind: def.kind,
                file_index: None,
                file_header: None,
                matched_by: MatchKind::Missing,
            }),
        }
    }

    let unmapped: Vec<(u16, String)> = file_headers
        .iter()
        .enumerate()
        .filter(|(i, _)| !claimed[*i])
        .map(|(i, h)| (i as u16, h.clone()))
        .collect();

    HeaderMap {
        columns,
        file_headers: file_headers.to_vec(),
        unmapped,
        header_row: opts.row_index,
        header_span: opts.row_span.max(1),
    }
}

fn normalize_key(raw: &str, opts: &HeaderOptions) -> String {
    if opts.normalize {
        normalize_header(raw)
    } else {
        raw.trim().to_ascii_lowercase()
    }
}

fn find_exact(norm: &[String], claimed: &[bool], keys: &[String]) -> Option<usize> {
    keys.iter().find_map(|k| {
        if k.is_empty() {
            return None;
        }
        norm.iter()
            .enumerate()
            .find(|(i, h)| !claimed[*i] && h.as_str() == k.as_str())
            .map(|(i, _)| i)
    })
}

/// 模糊匹配：互相包含，取差异最小（最贴近）的未占用列
fn find_fuzzy(norm: &[String], claimed: &[bool], keys: &[String]) -> Option<usize> {
    let mut best: Option<(usize, usize)> = None;
    for key in keys {
        if key.chars().count() < 2 {
            continue;
        }
        for (i, h) in norm.iter().enumerate() {
            if claimed[i] || h.is_empty() {
                continue;
            }
            if h.contains(key.as_str()) || key.contains(h.as_str()) {
                let diff = h.chars().count().abs_diff(key.chars().count());
                if best.map(|(_, d)| diff < d).unwrap_or(true) {
                    best = Some((i, diff));
                }
            }
        }
    }
    best.map(|(i, _)| i)
}

/// 便捷：声明列名 → 声明的 [`ColumnDef`] 列表（默认 `Any` 类型、非必填）
pub fn simple_columns<I, S>(headers: I) -> Vec<ColumnDef>
where
    I: IntoIterator<Item = S>,
    S: Into<String>,
{
    headers.into_iter().map(ColumnDef::new).collect()
}

/// 便捷：把文件表头映射成 `声明表头 → 文件列索引`
pub fn index_map(map: &HeaderMap) -> HashMap<String, u16> {
    map.columns
        .iter()
        .filter_map(|c| c.file_index.map(|i| (c.header.clone(), i)))
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn headers(v: &[&str]) -> Vec<String> {
        v.iter().map(|s| s.to_string()).collect()
    }

    #[test]
    fn test_exact_and_alias_and_fuzzy() {
        let defs = vec![
            ColumnDef::new("工厂名称").required(),
            ColumnDef::new("工厂品号").alias("品号").required(),
            ColumnDef::new("报价(USD)").kind(ColumnKind::Decimal),
            ColumnDef::new("图片1").image(),
        ];
        let file = headers(&[" 工厂名称 ", "品 号", "报价（USD）(含税)", "备注", "图片1"]);
        let map = resolve_headers(&file, &defs, &HeaderOptions::default());
        assert_eq!(map.index_of("工厂名称"), Some(0));
        assert_eq!(map.index_of("工厂品号"), Some(1));
        assert_eq!(map.index_of("报价(USD)"), Some(2));
        assert_eq!(map.index_of("图片1"), Some(4));
        assert_eq!(map.missing_required(), Vec::<String>::new());
        assert_eq!(map.unknown_columns(), vec!["备注".to_string()]);
    }

    #[test]
    fn test_missing_required() {
        let defs = vec![
            ColumnDef::new("工厂名称").required(),
            ColumnDef::new("品名").required(),
        ];
        let file = headers(&["工厂名称"]);
        let map = resolve_headers(&file, &defs, &HeaderOptions::default());
        assert_eq!(map.missing_required(), vec!["品名".to_string()]);
    }

    #[test]
    fn test_strict_mode_rejects_fuzzy() {
        let defs = vec![ColumnDef::new("报价(USD)")];
        let file = headers(&["报价(USD)(含税)"]);
        let loose = resolve_headers(&file, &defs, &HeaderOptions::default());
        assert_eq!(loose.index_of("报价(USD)"), Some(0));
        let strict = resolve_headers(&file, &defs, &HeaderOptions::strict());
        assert_eq!(strict.index_of("报价(USD)"), None);
    }

    #[test]
    fn test_excel_column_name() {
        assert_eq!(excel_column_name(0), "A");
        assert_eq!(excel_column_name(25), "Z");
        assert_eq!(excel_column_name(26), "AA");
        assert_eq!(excel_column_name(701), "ZZ");
        assert_eq!(excel_column_name(702), "AAA");
    }
}
