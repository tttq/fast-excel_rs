//! xlsx 容器（zip）访问层
//!
//! 关键点：**同一份文件可以同时打开多个 zip 句柄**。
//! 大文件流式读取时，工作表 XML 需要长时间持有自己的读取器，而图片 /
//! 共享字符串等条目又要能"随用随读"。为了避免自引用结构体（`ZipFile` 借用
//! `ZipArchive`），这里按条目元信息（压缩方式 + 数据起始偏移 + 压缩长度）
//! 自己拼一个**独占所有权**的读取器：`data_start` 之后接 `Take` + `DeflateDecoder`。

use std::fs::File;
use std::io::{BufReader, Read, Seek, SeekFrom};
use std::path::{Path, PathBuf};
use std::sync::Arc;

use flate2::read::DeflateDecoder;
use zip::CompressionMethod;
use zip::read::ZipArchive;

use crate::error::{ExcelError, Result};

/// `Read + Seek` 组合 trait（用于装箱成 trait object）
pub trait ReadSeek: Read + Seek {}
impl<T: Read + Seek> ReadSeek for T {}

/// 装箱后的 zip 读取器
pub type BoxedReader = Box<dyn ReadSeek + Send>;

/// xlsx 数据来源：磁盘文件或内存字节（可克隆，克隆体共享同一份数据）
#[derive(Clone)]
pub enum ZipSource {
    Path(PathBuf),
    Bytes(Arc<Vec<u8>>),
}

impl std::fmt::Debug for ZipSource {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ZipSource::Path(p) => write!(f, "ZipSource::Path({})", p.display()),
            ZipSource::Bytes(b) => write!(f, "ZipSource::Bytes({} bytes)", b.len()),
        }
    }
}

impl ZipSource {
    /// 以磁盘文件为数据源
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let path = path.as_ref().to_path_buf();
        if !path.exists() {
            return Err(ExcelError::Io(format!("文件不存在：{}", path.display())));
        }
        Ok(ZipSource::Path(path))
    }

    /// 以内存字节为数据源
    pub fn from_bytes(bytes: Vec<u8>) -> Self {
        ZipSource::Bytes(Arc::new(bytes))
    }

    pub fn from_arc(bytes: Arc<Vec<u8>>) -> Self {
        ZipSource::Bytes(bytes)
    }

    pub fn path(&self) -> Option<&Path> {
        match self {
            ZipSource::Path(p) => Some(p.as_path()),
            ZipSource::Bytes(_) => None,
        }
    }

    pub fn len(&self) -> u64 {
        match self {
            ZipSource::Path(p) => std::fs::metadata(p).map(|m| m.len()).unwrap_or(0),
            ZipSource::Bytes(b) => b.len() as u64,
        }
    }

    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }

    fn open_raw(&self) -> Result<BoxedReader> {
        match self {
            // 带缓冲，避免 XML 逐字节读时大量系统调用
            ZipSource::Path(p) => Ok(Box::new(BufReader::with_capacity(
                128 * 1024,
                File::open(p)?,
            ))),
            ZipSource::Bytes(b) => Ok(Box::new(SharedBytes::new(b.clone()))),
        }
    }

    /// 打开一个 zip 归档句柄（互不影响，可同时持有多份用于并发读取）
    pub fn open_archive(&self) -> Result<ZipArchive<BoxedReader>> {
        ZipArchive::new(self.open_raw()?).map_err(|e| ExcelError::NotXlsx(e.to_string()))
    }

    /// 完整读取一个条目到内存
    pub fn read_entry(&self, name: &str) -> Result<Vec<u8>> {
        let meta = self.entry_meta(name)?;
        let mut reader = self.open_entry_reader(&meta)?;
        let mut buf = Vec::with_capacity(meta.uncompressed_size.min(64 * 1024 * 1024) as usize);
        reader.read_to_end(&mut buf)?;
        Ok(buf)
    }

    /// 条目是否存在
    pub fn has_entry(&self, name: &str) -> bool {
        match self.open_archive() {
            Ok(mut a) => a.by_name(name).is_ok(),
            Err(_) => false,
        }
    }

    /// 条目名列表（兼容性探测 / 调试用）
    pub fn entry_names(&self) -> Result<Vec<String>> {
        let archive = self.open_archive()?;
        Ok(archive.file_names().map(|s| s.to_string()).collect())
    }

    /// 取条目元信息
    pub fn entry_meta(&self, name: &str) -> Result<EntryMeta> {
        let mut archive = self.open_archive()?;
        let file = archive
            .by_name(name)
            .map_err(|_| ExcelError::NotXlsx(format!("缺少内部条目：{name}")))?;
        if file.encrypted() {
            return Err(ExcelError::Unsupported(
                "加密的 xlsx 文件暂不支持".to_string(),
            ));
        }
        Ok(EntryMeta {
            name: name.to_string(),
            data_start: file.data_start(),
            compression: file.compression(),
            compressed_size: file.compressed_size(),
            uncompressed_size: file.size(),
        })
    }

    /// 按元信息打开一个**独占所有权**的条目读取器（无借用，可长期持有）
    pub fn open_entry_reader(&self, meta: &EntryMeta) -> Result<EntryReader> {
        let mut raw = self.open_raw()?;
        raw.seek(SeekFrom::Start(meta.data_start))?;
        let limited = raw.take(meta.compressed_size);
        let inner: Box<dyn Read + Send> = match meta.compression {
            CompressionMethod::Stored => Box::new(limited),
            CompressionMethod::Deflated => Box::new(DeflateDecoder::new(limited)),
            other => {
                return Err(ExcelError::Unsupported(format!(
                    "条目不支持的压缩方式 {other:?}（{}）",
                    meta.name
                )));
            }
        };
        Ok(EntryReader {
            inner,
            remaining: meta.uncompressed_size,
        })
    }

    /// 一步到位：按条目名打开流式读取器
    pub fn entry_reader(&self, name: &str) -> Result<EntryReader> {
        let meta = self.entry_meta(name)?;
        self.open_entry_reader(&meta)
    }
}

/// 条目元信息
#[derive(Debug, Clone)]
pub struct EntryMeta {
    pub name: String,
    pub data_start: u64,
    pub compression: CompressionMethod,
    pub compressed_size: u64,
    pub uncompressed_size: u64,
}

/// 独占所有权的条目读取器
pub struct EntryReader {
    inner: Box<dyn Read + Send>,
    /// 未压缩长度（预估容量用）
    remaining: u64,
}

impl EntryReader {
    pub fn uncompressed_size(&self) -> u64 {
        self.remaining
    }
}

impl Read for EntryReader {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        self.inner.read(buf)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_missing_file() {
        let err = ZipSource::open("definitely-not-exists.xlsx").unwrap_err();
        assert!(matches!(err, ExcelError::Io(_)));
    }
}
/// `Arc<Vec<u8>>` 上的零拷贝游标（`std::io::Cursor<Arc<Vec<u8>>>` 不满足 `Read`）
pub(crate) struct SharedBytes {
    data: Arc<Vec<u8>>,
    pos: u64,
}

impl SharedBytes {
    pub(crate) fn new(data: Arc<Vec<u8>>) -> Self {
        Self { data, pos: 0 }
    }
}

impl Read for SharedBytes {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        let data: &[u8] = &self.data;
        let start = self.pos as usize;
        if start >= data.len() {
            return Ok(0);
        }
        let n = (data.len() - start).min(buf.len());
        buf[..n].copy_from_slice(&data[start..start + n]);
        self.pos += n as u64;
        Ok(n)
    }
}

impl Seek for SharedBytes {
    fn seek(&mut self, pos: SeekFrom) -> std::io::Result<u64> {
        let len = self.data.len() as i64;
        let next = match pos {
            SeekFrom::Start(n) => n as i64,
            SeekFrom::End(n) => len + n,
            SeekFrom::Current(n) => self.pos as i64 + n,
        };
        if next < 0 {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                "seek 到负数位置",
            ));
        }
        self.pos = next as u64;
        Ok(self.pos)
    }
}

#[cfg(test)]
mod shared_bytes_tests {
    use super::*;

    #[test]
    fn test_shared_bytes_seek_read() {
        let data = Arc::new(vec![1u8, 2, 3, 4, 5]);
        let mut r = SharedBytes::new(data);
        assert_eq!(r.seek(SeekFrom::Start(2)).unwrap(), 2);
        let mut buf = [0u8; 2];
        r.read_exact(&mut buf).unwrap();
        assert_eq!(buf, [3, 4]);
        assert_eq!(r.seek(SeekFrom::End(-1)).unwrap(), 4);
        assert_eq!(r.read(&mut buf).unwrap(), 1);
        assert_eq!(buf[0], 5);
    }
}
