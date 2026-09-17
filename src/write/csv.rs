//! CSV 写出（导出兜底格式：不依赖 Excel 生态，纯文本流式写）
//!
//! 默认带 UTF-8 BOM，Excel 双击打开中文才不会乱码。

use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::Path;

use crate::error::Result;

/// 流式 CSV 写出器
pub struct CsvWriter<W: Write> {
    inner: W,
    /// 是否已经写过记录（用于插入换行）
    started: bool,
}

impl CsvWriter<BufWriter<File>> {
    /// 创建文件（自动带 UTF-8 BOM）
    pub fn create<P: AsRef<Path>>(path: P) -> Result<Self> {
        let file = File::create(path)?;
        let mut inner = BufWriter::with_capacity(128 * 1024, file);
        inner.write_all(&[0xEF, 0xBB, 0xBF])?;
        Ok(Self {
            inner,
            started: false,
        })
    }
}

impl<W: Write> CsvWriter<W> {
    /// 用已有 writer 构造
    pub fn new(inner: W) -> Self {
        Self {
            inner,
            started: false,
        }
    }

    /// 写一条记录（自动转义逗号 / 引号 / 换行）
    pub fn write_record<I, S>(&mut self, fields: I) -> Result<()>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        if self.started {
            self.inner.write_all(b"\r\n")?;
        }
        let mut first = true;
        for f in fields {
            if !first {
                self.inner.write_all(b",")?;
            }
            first = false;
            write_field(&mut self.inner, f.as_ref())?;
        }
        self.started = true;
        Ok(())
    }

    pub fn flush(&mut self) -> Result<()> {
        self.inner.flush()?;
        Ok(())
    }

    pub fn into_inner(self) -> W {
        self.inner
    }
}

/// 写一个字段：含有分隔符或引号时用引号包裹，内部引号双写
fn write_field<W: Write>(w: &mut W, field: &str) -> Result<()> {
    let need_quote =
        field.contains(',') || field.contains('"') || field.contains('\n') || field.contains('\r');
    if !need_quote {
        w.write_all(field.as_bytes())?;
        return Ok(());
    }
    w.write_all(b"\x22")?;
    for (i, chunk) in field.split('"').enumerate() {
        if i > 0 {
            w.write_all(b"\x22\x22")?;
        }
        w.write_all(chunk.as_bytes())?;
    }
    w.write_all(b"\x22")?;
    Ok(())
}

/// 一步导出 CSV 文件（表头 + 行）
pub fn write_csv<P, I, S>(path: P, headers: &[String], rows: I) -> Result<u64>
where
    P: AsRef<Path>,
    I: IntoIterator<Item = Vec<S>>,
    S: AsRef<str>,
{
    let mut writer = CsvWriter::create(path.as_ref())?;
    writer.write_record(headers.iter().map(|s| s.as_str()))?;
    for row in rows {
        writer.write_record(row.iter().map(|s| s.as_ref()))?;
    }
    writer.flush()?;
    Ok(std::fs::metadata(path.as_ref())
        .map(|m| m.len())
        .unwrap_or(0))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_csv_escaping() {
        let mut buf: Vec<u8> = Vec::new();
        {
            let mut w = CsvWriter::new(&mut buf);
            w.write_record(["a,b", "c?d", "e\nf"]).unwrap();
            w.flush().unwrap();
        }
        let s = String::from_utf8(buf).unwrap();
        assert_eq!(s, "\x22a,b\x22,c?d,\x22e\nf\x22");
    }
}
