use super::cell::{actual_cell_type, parse_cell, parse_cell_value, write_cell};
use super::FormatManager;
use crate::model::*;
use anyhow::Context;
use std::collections::HashSet;
use std::io::BufRead;
use std::path::Path;
use xlsxwriter::worksheet::{Worksheet, WorksheetCol, WorksheetRow};

pub fn insert_csv<P: AsRef<Path>>(
    worksheet: &mut Worksheet,
    source_def: &SheetSourceDef,
    formats: &FormatManager,
    base_path: P,
) -> anyhow::Result<()> {
    let reader: Box<dyn BufRead> = if let Some(file) = source_def.file.as_deref() {
        Box::new(std::io::BufReader::new(
            autocompress::autodetect_open(base_path.as_ref().join(file))
                .with_context(|| format!("Cannot open \"{}\"", file))?,
        ))
    } else if let Some(data) = source_def.data.as_deref() {
        Box::new(&data.as_bytes()[..])
    } else {
        return Err(anyhow::anyhow!("No data found for CSV/TSV"));
    };

    let mut csv_reader_builder = csv::ReaderBuilder::new();
    csv_reader_builder.has_headers(false);
    csv_reader_builder.flexible(true);
    if let Some(comment_line_prefix) = source_def.comment_line_prefix.as_deref() {
        if comment_line_prefix.as_bytes().len() == 1 {
            csv_reader_builder.comment(Some(comment_line_prefix.as_bytes()[0]));
        } else {
            return Err(anyhow::anyhow!("Comment line prefix must be length 1"));
        }
    }
    match source_def.suggest_format() {
        SheetSourceType::CSV => {}
        SheetSourceType::TSV => {
            csv_reader_builder
                .delimiter(b'\t')
                .quoting(false)
                .double_quote(false);
        }
        _ => unreachable!(),
    }
    let mut csv_reader = csv_reader_builder.from_reader(reader);
    let mut header_line = Vec::new();
    let mut filter_column_index: Option<usize> = None;
    let filter_list: HashSet<String> = source_def
        .filter_list
        .as_ref()
        .map(|x| x.items.iter().map(|y| y.to_string()).collect())
        .unwrap_or_default();

    if source_def.table && source_def.filter_list.is_some() {
        eprintln!("filter_list cannot be enabled when table mode is enabled");
    }

    let mut maximum_col = 0;
    let mut maximum_row = 0;
    let mut offset_row = source_def.start_row;

    if let Some(columns) = source_def.columns.as_ref() {
        if !source_def.has_header && columns.iter().any(|y| y.header_value.is_some()) {
            for (i, one) in columns.iter().enumerate() {
                if let Some(value) = one.header_value.as_ref() {
                    let value = parse_cell_value(value, one.header_type)?;
                    write_cell(worksheet, offset_row, i as WorksheetCol, &value, None)?;
                }
                if let Some(comment) = one.header_comment.as_ref() {
                    worksheet.write_comment(offset_row, i as WorksheetCol, &comment)?;
                }
            }
            offset_row += 1;
        }
    }

    for (i, row) in csv_reader.records().enumerate() {
        maximum_row = i;
        let row = row?;
        for (j, cell) in row.iter().enumerate() {
            maximum_col = maximum_col.max(j);

            let link_prefix: Option<String> = source_def
                .columns
                .as_ref()
                .map(|x| x.get(j))
                .flatten()
                .map(|x| x.link_prefix.clone())
                .flatten();

            if i == 0 && source_def.has_header {
                if let Some(comment) = source_def
                    .columns
                    .as_ref()
                    .map(|x| x.get(j))
                    .flatten()
                    .map(|x| x.header_comment.as_deref())
                    .flatten()
                {
                    worksheet.write_comment(
                        (i as WorksheetRow) + offset_row,
                        (j as WorksheetCol) + source_def.start_column,
                        comment,
                    )?;
                }

                if source_def
                    .filter_list
                    .as_ref()
                    .map(|x| x.column_header == cell)
                    .unwrap_or(false)
                    && !source_def.table
                {
                    filter_column_index = Some(j);
                }

                header_line.push(cell.to_string());
            } else if filter_column_index.map(|x| x == j).unwrap_or(false) {
                if !filter_list.contains(cell) {
                    worksheet.set_row_opt(
                        (i as WorksheetRow) + offset_row,
                        xlsxwriter::worksheet::LXW_DEF_ROW_HEIGHT,
                        None,
                        &xlsxwriter::worksheet::RowColOptions::new(true, 0, false),
                    )?;
                }
            }

            let cell_type: CellType = if i == 0 && source_def.has_header {
                CellType::String
            } else {
                source_def
                    .columns
                    .as_ref()
                    .map(|x| x.get(j).map(|y| y.cell_type))
                    .flatten()
                    .unwrap_or(CellType::Auto)
            };

            match parse_cell(cell, cell_type) {
                Ok(value) => {
                    if let Some(link_prefix) = link_prefix.as_deref() {
                        worksheet.write_url(
                            (i as WorksheetRow) + offset_row,
                            (j as WorksheetCol) + source_def.start_column,
                            &format!("{}{}", link_prefix, cell),
                            None,
                        )?;
                    }
                    write_cell(
                        worksheet,
                        (i as WorksheetRow) + offset_row,
                        (j as WorksheetCol) + source_def.start_column,
                        &value,
                        source_def
                            .columns
                            .as_ref()
                            .map(|x| {
                                x.get(j).map(|y| {
                                    formats.get_format(
                                        y.format.as_ref(),
                                        if link_prefix.is_some() {
                                            CellType::Url
                                        } else {
                                            actual_cell_type(&value, y.cell_type)
                                        },
                                    )
                                })
                            })
                            .flatten()
                            .flatten(),
                    )?;
                }
                Err(e) => {
                    eprintln!(
                        "warning: {}: row {}, column {}: {}",
                        source_def.file.as_deref().unwrap_or("embedded data"),
                        i,
                        j,
                        e
                    );
                }
            }
        }
    }

    super::table::setup_table(
        worksheet,
        source_def,
        formats,
        &header_line,
        filter_column_index.map(|x| x as WorksheetCol),
        maximum_row as WorksheetRow,
        maximum_col as WorksheetCol,
    )?;

    Ok(())
}
