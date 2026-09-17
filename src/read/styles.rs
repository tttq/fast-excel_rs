//! `xl/styles.xml` 解析：判断哪些样式是日期/时间格式
//!
//! Excel 把日期存成数字（序列号）+ 单元格样式，所以"这一列是不是日期"必须靠
//! 样式里的 `numFmtId` / `formatCode` 判断，否则日期会变成 `45293` 这种数字。

use std::collections::{HashMap, HashSet};
use std::io::BufRead;

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::error::Result;

/// 单元格样式表
#[derive(Debug, Default)]
pub struct Styles {
    /// 日期样式的 `cellXfs` 下标
    date_styles: HashSet<u32>,
}

impl Styles {
    pub fn empty() -> Self {
        Self::default()
    }

    /// 该样式索引是否按日期解析
    pub fn is_date(&self, style: u32) -> bool {
        self.date_styles.contains(&style)
    }

    pub fn parse<R: BufRead>(reader: R) -> Result<Self> {
        let mut xml = Reader::from_reader(reader);
        xml.config_mut().trim_text(true);
        let mut buf = Vec::with_capacity(8 * 1024);

        let mut custom: HashMap<u32, String> = HashMap::new();
        let mut xf_num_fmts: Vec<u32> = Vec::new();
        let mut in_cell_xfs = false;

        loop {
            match xml.read_event_into(&mut buf) {
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                    b"numFmt" => {
                        let mut id = None;
                        let mut code = None;
                        for attr in e.attributes().flatten() {
                            match attr.key.as_ref() {
                                b"numFmtId" => {
                                    id = String::from_utf8_lossy(&attr.value).parse::<u32>().ok()
                                }
                                b"formatCode" => {
                                    code = Some(String::from_utf8_lossy(&attr.value).to_string())
                                }
                                _ => {}
                            }
                        }
                        if let (Some(id), Some(code)) = (id, code) {
                            custom.insert(id, code);
                        }
                    }
                    b"cellXfs" => in_cell_xfs = true,
                    b"xf" if in_cell_xfs => {
                        let mut num_fmt = 0u32;
                        for attr in e.attributes().flatten() {
                            if attr.key.as_ref() == b"numFmtId" {
                                num_fmt = String::from_utf8_lossy(&attr.value).parse().unwrap_or(0);
                            }
                        }
                        xf_num_fmts.push(num_fmt);
                    }
                    _ => {}
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"cellXfs" {
                        in_cell_xfs = false;
                    }
                }
                _ => {}
            }
            buf.clear();
        }

        let date_styles = xf_num_fmts
            .iter()
            .enumerate()
            .filter(|(_, id)| is_date_format(**id, custom.get(id).map(|s| s.as_str())))
            .map(|(i, _)| i as u32)
            .collect();

        Ok(Self { date_styles })
    }
}

/// 内置日期格式 ID（ECMA-376 预定义）
const BUILTIN_DATE_FORMATS: &[u32] = &[
    14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51,
    52, 53, 54, 55, 56, 57, 58,
];

/// 判断 `numFmtId` + 自定义格式串是否为日期/时间
pub fn is_date_format(num_fmt_id: u32, code: Option<&str>) -> bool {
    match code {
        Some(code) => is_date_code(code),
        None => BUILTIN_DATE_FORMATS.contains(&num_fmt_id),
    }
}

/// 格式串是否表示日期/时间（去引号字面量、取第一个分号段后按字符判断）
pub fn is_date_code(code: &str) -> bool {
    let mut cleaned = String::with_capacity(code.len());
    let mut in_quote = false;
    let mut escaped = false;
    for c in code.chars() {
        if escaped {
            escaped = false;
            continue;
        }
        match c {
            '\\' => escaped = true,
            '"' => in_quote = !in_quote,
            '[' => in_quote = true,
            ']' => in_quote = false,
            _ if in_quote => {}
            _ => cleaned.push(c.to_ascii_lowercase()),
        }
    }
    let first = cleaned.split(';').next().unwrap_or("").trim().to_string();
    if first.is_empty() {
        return false;
    }
    if first.contains('y') || first.contains('d') || first.contains('h') || first.contains('s') {
        return true;
    }
    // 只有 m 时（可能是分钟/月份，也可能是 "0.00" 里的字母）需要配合分隔符
    first.contains('m') && (first.contains('/') || first.contains('-') || first.contains(':'))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_date_codes() {
        assert!(is_date_code("yyyy-mm-dd"));
        assert!(is_date_code("yyyy\"年\"m\"月\"d\"日\""));
        assert!(is_date_code("h:mm:ss"));
        assert!(is_date_code("mm:ss"));
        assert!(!is_date_code("#,##0.00"));
        assert!(!is_date_code("@"));
        assert!(!is_date_code("General"));
        assert!(!is_date_code("0.00%"));
        assert!(is_date_format(14, None));
        assert!(!is_date_format(2, None));
        assert!(is_date_format(2, Some("yyyy/mm/dd")));
    }
}
