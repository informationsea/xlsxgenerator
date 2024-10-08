use super::super::FormatManager;
use anyhow::Result;
use std::collections::{HashMap, HashSet};
use std::io::Write;
use std::str;
use xlsxwriter::worksheet::{WorksheetCol, WorksheetRow};

const XLSX_MAX_ROW: WorksheetRow = 1048576;

pub trait TableWriter {
    fn set_header(&mut self, items: &[String]);
    fn header(&self) -> &[String];
    fn write_row(&mut self, items: &[&str]) -> Result<()>;

    fn write_header(&mut self) -> Result<()> {
        let header: Vec<_> = self.header().iter().map(|x| x.to_string()).collect();
        self.write_row(&header.iter().map(|x| x.as_str()).collect::<Vec<&str>>())
    }

    fn write_row_bytes(&mut self, items: &[&[u8]]) -> Result<()> {
        self.write_row(
            &items
                .iter()
                .try_fold::<_, _, Result<_>>(Vec::new(), |mut v, x| {
                    let s = str::from_utf8(x)?;
                    v.push(s);
                    Ok(v)
                })?,
        )
    }

    fn write_dict(&mut self, items: &HashMap<&str, &str>) -> Result<()> {
        let mut row: Vec<&str> = Vec::new();
        for one in self.header() {
            let s: &str = one;
            row.push(items.get(s).copied().unwrap_or(""));
        }
        self.write_row(&row)
    }

    fn write_dict_bytes(&mut self, items: &HashMap<&[u8], &[u8]>) -> Result<()> {
        let mut row: Vec<&[u8]> = Vec::new();
        for one in self.header() {
            row.push(items.get(one.as_bytes()).copied().unwrap_or(b""));
        }
        self.write_row_bytes(&row)
    }

    fn column_widths(&mut self, widths: &[f64]) -> Result<()>;

    fn is_next_row_allowed(&self) -> bool {
        true
    }
}

impl<T: TableWriter + ?Sized> TableWriter for Box<T> {
    fn set_header(&mut self, items: &[String]) {
        (**self).set_header(items)
    }
    fn header(&self) -> &[String] {
        (**self).header()
    }
    fn write_row(&mut self, items: &[&str]) -> Result<()> {
        (**self).write_row(items)
    }
    fn write_header(&mut self) -> Result<()> {
        (**self).write_header()
    }
    fn write_row_bytes(&mut self, items: &[&[u8]]) -> Result<()> {
        (**self).write_row_bytes(items)
    }
    fn write_dict(&mut self, items: &HashMap<&str, &str>) -> Result<()> {
        (**self).write_dict(items)
    }
    fn write_dict_bytes(&mut self, items: &HashMap<&[u8], &[u8]>) -> Result<()> {
        (**self).write_dict_bytes(items)
    }
    fn column_widths(&mut self, widths: &[f64]) -> Result<()> {
        (**self).column_widths(widths)
    }
}

impl<T: TableWriter + ?Sized> TableWriter for &mut T {
    fn set_header(&mut self, items: &[String]) {
        (**self).set_header(items)
    }
    fn header(&self) -> &[String] {
        (**self).header()
    }
    fn write_row(&mut self, items: &[&str]) -> Result<()> {
        (**self).write_row(items)
    }
    fn write_header(&mut self) -> Result<()> {
        (**self).write_header()
    }
    fn write_row_bytes(&mut self, items: &[&[u8]]) -> Result<()> {
        (**self).write_row_bytes(items)
    }
    fn write_dict(&mut self, items: &HashMap<&str, &str>) -> Result<()> {
        (**self).write_dict(items)
    }
    fn write_dict_bytes(&mut self, items: &HashMap<&[u8], &[u8]>) -> Result<()> {
        (**self).write_dict_bytes(items)
    }
    fn column_widths(&mut self, widths: &[f64]) -> Result<()> {
        (**self).column_widths(widths)
    }
}

#[derive(Debug)]
pub struct TSVWriter<W: Write> {
    writer: W,
    header: Vec<String>,
}

impl<W: Write> TSVWriter<W> {
    pub fn new(writer: W) -> Self {
        TSVWriter {
            writer,
            header: Vec::new(),
        }
    }
}

impl<W: Write> TableWriter for TSVWriter<W> {
    fn set_header(&mut self, items: &[String]) {
        self.header.clear();
        self.header.extend_from_slice(items);
    }

    fn header(&self) -> &[String] {
        &self.header
    }

    fn write_row(&mut self, items: &[&str]) -> Result<()> {
        for (i, data) in items.iter().enumerate() {
            if i != 0 {
                self.writer.write_all(b"\t")?;
            }
            write!(self.writer, "{}", data)?;
        }
        writeln!(self.writer)?;
        Ok(())
    }

    fn column_widths(&mut self, _widths: &[f64]) -> Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
pub struct CSVWriter<W: Write> {
    writer: csv::Writer<W>,
    header: Vec<String>,
}

impl<W: Write> CSVWriter<W> {
    pub fn new(writer: W) -> Self {
        CSVWriter {
            writer: csv::Writer::from_writer(writer),
            header: Vec::new(),
        }
    }
}

impl<W: Write> TableWriter for CSVWriter<W> {
    fn set_header(&mut self, items: &[String]) {
        self.header.clear();
        self.header.extend_from_slice(items);
    }

    fn header(&self) -> &[String] {
        &self.header
    }

    fn write_row(&mut self, items: &[&str]) -> Result<()> {
        self.writer.write_record(items)?;
        Ok(())
    }

    fn column_widths(&mut self, _widths: &[f64]) -> Result<()> {
        Ok(())
    }
}

pub struct XlsxSheetWriter<'a, 'b> {
    writer: &'a mut xlsxwriter::Worksheet<'b>,
    pub header: Vec<String>,
    pub header_comment: Vec<String>,
    data_type: Vec<XlsxDataType>,
    current_row: WorksheetRow,
    offset_col: WorksheetCol,
    column_filter_index: Option<usize>,
    column_filter_list: &'a HashSet<String>,
    format_manager: &'a FormatManager,
}

impl<'a, 'b> XlsxSheetWriter<'a, 'b> {
    pub fn new(
        worksheet: &'a mut xlsxwriter::Worksheet<'b>,
        format_manager: &'a FormatManager,
        offset_row: WorksheetRow,
        offset_col: WorksheetCol,
        column_filter_index: Option<usize>,
        column_filter_list: &'a HashSet<String>,
    ) -> Self {
        XlsxSheetWriter {
            writer: worksheet,
            header: Vec::new(),
            header_comment: Vec::new(),
            data_type: Vec::new(),
            current_row: offset_row,
            offset_col,
            column_filter_index,
            column_filter_list,
            format_manager,
        }
    }

    pub fn set_data_type(&mut self, data_type: &[XlsxDataType]) {
        self.data_type.clear();
        self.data_type.extend_from_slice(data_type);
    }

    pub fn set_header_comment(&mut self, items: &[String]) {
        self.header_comment.clear();
        self.header_comment.extend_from_slice(items);
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum XlsxDataType {
    String,
    Boolean,
    Number,
    Integer,
}

impl<'a, 'b> TableWriter for XlsxSheetWriter<'a, 'b> {
    fn set_header(&mut self, items: &[String]) {
        self.header.clear();
        self.header.extend_from_slice(items);
    }
    fn header(&self) -> &[String] {
        &self.header
    }

    fn write_header(&mut self) -> Result<()> {
        let header: Vec<_> = self.header().iter().map(|x| x.to_string()).collect();
        for (i, column) in header.iter().enumerate() {
            self.writer
                .write_string(self.current_row, i as u16 + self.offset_col, column, None)?;
            if let Some(comment) = self.header_comment.get(i) {
                if comment != "" {
                    self.writer.write_comment(
                        self.current_row,
                        i as u16 + self.offset_col,
                        &comment,
                    )?;
                }
            }
        }
        self.current_row += 1;
        Ok(())
    }

    fn write_row(&mut self, items: &[&str]) -> Result<()> {
        for (i, column) in items.iter().enumerate() {
            let write_col = i as WorksheetCol + self.offset_col;
            if column.is_empty() {
                self.writer.write_blank(self.current_row, i as u16, None)?;
            } else {
                let data_type = self
                    .data_type
                    .get(i)
                    .copied()
                    .unwrap_or(XlsxDataType::String);
                match data_type {
                    XlsxDataType::String => {
                        if column.len() > 32766 {
                            self.writer.write_string(
                                self.current_row,
                                write_col,
                                &format!("{}...", &column[0..32763]),
                                self.format_manager
                                    .get_format(None, super::super::CellType::String),
                            )?;
                        } else {
                            self.writer.write_string(
                                self.current_row,
                                write_col,
                                column,
                                self.format_manager
                                    .get_format(None, super::super::CellType::String),
                            )?;
                        }
                    }
                    XlsxDataType::Number | XlsxDataType::Integer => {
                        if let Ok(f) = column.parse() {
                            self.writer.write_number(
                                self.current_row,
                                write_col,
                                f,
                                self.format_manager.get_format(
                                    None,
                                    match data_type {
                                        XlsxDataType::Integer => super::super::CellType::Integer,
                                        _ => super::super::CellType::Number,
                                    },
                                ),
                            )?;
                        } else {
                            self.writer.write_string(
                                self.current_row,
                                write_col,
                                column,
                                self.format_manager
                                    .get_format(None, super::super::CellType::String),
                            )?;
                        }
                    }
                    XlsxDataType::Boolean => {
                        self.writer.write_boolean(
                            self.current_row,
                            write_col,
                            *column == "TRUE" || *column == "True" || *column == "true",
                            self.format_manager
                                .get_format(None, super::super::CellType::Boolean),
                        )?;
                    }
                }
            }
        }

        if self
            .column_filter_index
            .map(|i| items.get(i))
            .flatten()
            .map(|x| !self.column_filter_list.contains(*x))
            .unwrap_or(false)
        {
            self.writer.set_row_opt(
                self.current_row,
                xlsxwriter::worksheet::LXW_DEF_ROW_HEIGHT,
                None,
                &xlsxwriter::worksheet::RowColOptions::new(true, 0, false),
            )?;
        }

        self.current_row += 1;
        Ok(())
    }

    fn column_widths(&mut self, widths: &[f64]) -> Result<()> {
        for (i, one) in widths.iter().enumerate() {
            self.writer.set_column(
                self.offset_col + i as WorksheetCol,
                self.offset_col + i as WorksheetCol,
                *one * xlsxwriter::worksheet::LXW_DEF_COL_WIDTH,
                None,
            )?;
        }
        Ok(())
    }

    fn is_next_row_allowed(&self) -> bool {
        self.current_row < XLSX_MAX_ROW
    }
}

#[cfg(test)]
mod test {
    use super::*;

    #[test]
    fn test_table_writer() -> Result<()> {
        let mut write_buf: Vec<u8> = Vec::new();
        let mut tsv_writer = TSVWriter::new(&mut write_buf);
        tsv_writer.set_header(&["Col1".to_string(), "Col2".to_string(), "Col3".to_string()]);
        tsv_writer.write_header()?;
        tsv_writer.write_row(&["val1", "val2", "val3"])?;
        tsv_writer.write_row_bytes(&[b"bin1", b"bin2", b"bin3"])?;

        let str_map: HashMap<&str, &str> =
            vec![("Col1", "hash1"), ("Col2", "hash2"), ("Col3", "hash3")]
                .into_iter()
                .collect();
        tsv_writer.write_dict(&str_map)?;

        let bytes_map: HashMap<&[u8], &[u8]> = vec![
            (&b"Col1"[..], &b"bin_hash1"[..]),
            (&b"Col2"[..], &b"bin_hash2"[..]),
            (&b"Col3"[..], &b"bin_hash3"[..]),
        ]
        .into_iter()
        .collect();
        tsv_writer.write_dict_bytes(&bytes_map)?;

        let str_map2: HashMap<&str, &str> =
            vec![("Col1", "hash1"), ("Col3", "hash3"), ("Col4", "hash4")]
                .into_iter()
                .collect();
        tsv_writer.write_dict(&str_map2)?;

        let expected_bytes = br#"Col1	Col2	Col3
val1	val2	val3
bin1	bin2	bin3
hash1	hash2	hash3
bin_hash1	bin_hash2	bin_hash3
hash1		hash3
"#;
        assert_eq!(&expected_bytes[..], &write_buf[..]);

        Ok(())
    }

    #[test]
    fn test_csv_writer() -> Result<()> {
        let mut write_buf: Vec<u8> = Vec::new();
        {
            let mut csv_writer = CSVWriter::new(&mut write_buf);
            csv_writer.set_header(&["Col1".to_string(), "Col2".to_string(), "Col3".to_string()]);
            csv_writer.write_header()?;
            csv_writer.write_row(&["val1", "val2,val", "val3"])?;
            csv_writer.write_row_bytes(&[b"bin1", b"bin2", b"bin3,bin4"])?;
        }

        let expected_bytes = br#"Col1,Col2,Col3
val1,"val2,val",val3
bin1,bin2,"bin3,bin4"
"#;
        assert_eq!(&expected_bytes[..], &write_buf[..]);

        Ok(())
    }
}
