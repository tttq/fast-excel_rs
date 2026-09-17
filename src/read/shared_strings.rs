//! `xl/sharedStrings.xml` 解析（字符串池）
//!
//! 池子必须整体常驻（一个字符串的索引可以出现在任意位置），但**只装字符串本身**，
//! 一个 100 万行 x 10 列的文件池子通常是几十 MB，属于可控量级。
//! 需要更省内存时，写出侧请用 inline string 或数字列。

use std::io::BufRead;

use quick_xml::Reader;
use quick_xml::events::{BytesText, Event};

use crate::error::Result;

/// 共享字符串池
#[derive(Debug, Default)]
pub struct SharedStrings {
    items: Vec<String>,
}

impl SharedStrings {
    pub fn empty() -> Self {
        Self { items: Vec::new() }
    }

    pub fn len(&self) -> usize {
        self.items.len()
    }

    pub fn is_empty(&self) -> bool {
        self.items.is_empty()
    }

    /// 按索引取字符串
    pub fn get(&self, index: usize) -> Option<&str> {
        self.items.get(index).map(|s| s.as_str())
    }

    pub fn push(&mut self, s: String) {
        self.items.push(s);
    }

    /// 流式解析：只保留文本，富文本（`<r><t>`）拼接成一段，拼音（`<rPh>`）丢弃
    pub fn parse<R: BufRead>(reader: R) -> Result<Self> {
        let mut xml = Reader::from_reader(reader);
        xml.config_mut().trim_text(false);
        let mut buf = Vec::with_capacity(8 * 1024);
        let mut items: Vec<String> = Vec::new();
        let mut cur = String::new();
        let mut in_si = false;
        let mut in_rph = false;
        let mut in_t = false;

        loop {
            match xml.read_event_into(&mut buf) {
                Ok(Event::Eof) => break,
                Err(e) => return Err(e.into()),
                Ok(Event::Start(e)) => match e.local_name().as_ref() {
                    b"si" => {
                        in_si = true;
                        cur.clear();
                    }
                    b"rPh" => in_rph = true,
                    b"t" => in_t = true,
                    _ => {}
                },
                Ok(Event::End(e)) => match e.local_name().as_ref() {
                    b"si" => {
                        in_si = false;
                        items.push(std::mem::take(&mut cur));
                    }
                    b"rPh" => in_rph = false,
                    b"t" => in_t = false,
                    _ => {}
                },
                Ok(Event::Text(t)) => {
                    if in_t && !in_rph && in_si {
                        cur.push_str(&decode_text(&t));
                    }
                }
                Ok(Event::CData(t)) => {
                    if in_t && !in_rph && in_si {
                        cur.push_str(&String::from_utf8_lossy(t.as_ref()));
                    }
                }
                _ => {}
            }
            buf.clear();
        }
        Ok(Self { items })
    }
}

/// XML 文本解码（实体转义 → 原文）
pub(crate) fn decode_text(t: &BytesText<'_>) -> String {
    match t.unescape() {
        Ok(c) => c.into_owned(),
        Err(e) => {
            log::warn!("Excel 文本实体解码失败，按原样保留：{e}");
            String::from_utf8_lossy(t.as_ref()).into_owned()
        }
    }
}
