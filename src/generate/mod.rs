mod cell;
mod format;
mod insert_csv;
mod insert_vcf;
pub mod table;
pub mod vcf;

use anyhow::Context;
use std::collections::HashSet;
use std::io::BufRead;
use std::path::Path;
use std::str;
use xlsxwriter::worksheet::{ImageOptions, Worksheet, WorksheetCol, WorksheetRow};

use crate::model::*;
use cell::*;
use format::*;
use insert_csv::*;
use insert_vcf::*;

pub fn generate_worksheet<P: AsRef<Path>>(
    worksheet: &mut Worksheet,
    worksheet_def: &WorksheetDef,
    formats: &FormatManager,
    base_path: P,
    canonical_transcripts: Option<HashSet<Vec<u8>>>,
) -> anyhow::Result<()> {
    if let Some(source) = worksheet_def.source.as_ref() {
        let source_array: Vec<SheetSourceDef> = source.clone().into();
        for source in source_array.iter() {
            match source.suggest_format() {
                SheetSourceType::CSV | SheetSourceType::TSV => {
                    insert_csv(worksheet, &source, formats, base_path.as_ref())?;
                }
                SheetSourceType::VCF => {
                    insert_vcf(
                        worksheet,
                        &source,
                        formats,
                        base_path.as_ref(),
                        canonical_transcripts.clone(),
                    )?;
                }
                _ => unreachable!(),
            }
        }
    }

    let mut first_cell = true;
    let mut last_row: WorksheetRow = 0;
    let mut last_col: WorksheetCol = 0;
    let mut last_explicit_col: WorksheetCol = 0;

    for one_cell in worksheet_def.cells.iter() {
        //eprintln!("cell: {:?}", one_cell);
        let row: WorksheetRow = if let Some(row_relative) = one_cell.row_relative {
            //eprint!("row relative : {:?} ", row_relative);
            if row_relative < 0 {
                last_row - (-row_relative).try_into().unwrap_or(0)
            } else {
                last_row + row_relative.try_into().unwrap_or(0)
            }
        } else if let Some(row) = one_cell.row {
            row
        } else {
            last_row
        };
        let column: WorksheetCol = if let Some(col_relative) = one_cell.column_relative {
            let col = if col_relative < 0 {
                last_col - (-col_relative).try_into().unwrap_or(0)
            } else {
                last_col + col_relative.try_into().unwrap_or(0)
            };
            last_explicit_col = col;
            col
        } else if let Some(col) = one_cell.column {
            last_explicit_col = col;
            col
        } else {
            if last_row != row {
                last_explicit_col
            } else if first_cell {
                0
            } else {
                last_col + 1
            }
        };
        last_row = row;
        last_col = column;
        first_cell = false;

        if one_cell.merge_column.is_some() || one_cell.merge_row.is_some() {
            let merge_col = one_cell.merge_column.unwrap_or(1);
            let merge_row = one_cell.merge_row.unwrap_or(1);
            worksheet.merge_range(
                row,
                column,
                row + merge_row - 1,
                column + merge_col - 1,
                "",
                None,
            )?;

            for row in row..(row + merge_row) {
                for col in column..(column + merge_col) {
                    worksheet.write_blank(
                        row,
                        col,
                        formats.get_format(one_cell.format.as_ref(), CellType::Null),
                    )?;
                }
            }
        }
        if let Some(url) = one_cell.url.as_deref() {
            worksheet.write_url(row, column, url, None)?;
        }
        if let Some(value) = one_cell.value.as_ref() {
            let parsed_value = parse_cell_value(&value, one_cell.cell_type)?;
            write_cell(
                worksheet,
                row,
                column,
                &parsed_value,
                formats.get_format(
                    one_cell.format.as_ref(),
                    if one_cell.url.is_some() {
                        CellType::Url
                    } else {
                        actual_cell_type(&parsed_value, one_cell.cell_type)
                    },
                ),
            )?;
            if let Some(comment) = one_cell.comment.as_deref() {
                worksheet.write_comment(row, column, comment)?;
            }
        }
    }

    for (i, one) in worksheet_def.column_widths.iter().enumerate() {
        worksheet.set_column(
            i as xlsxwriter::worksheet::WorksheetCol,
            i as xlsxwriter::worksheet::WorksheetCol,
            *one,
            None,
        )?;
    }

    for (i, one) in worksheet_def.row_heights.iter().enumerate() {
        worksheet.set_row(i as xlsxwriter::worksheet::WorksheetRow, *one, None)?;
    }

    for one_image in worksheet_def.images.iter() {
        let opt = ImageOptions {
            x_scale: one_image.width_scale.unwrap_or(1.),
            y_scale: one_image.height_scale.unwrap_or(1.),
            x_offset: 0,
            y_offset: 0,
        };

        worksheet.insert_image_opt(
            one_image.row,
            one_image.column,
            base_path
                .as_ref()
                .join(&one_image.file)
                .to_string_lossy()
                .as_ref(),
            &opt,
        )?;
    }

    if let Some(freeze) = worksheet_def.freeze.as_ref() {
        worksheet.freeze_panes(freeze.row, freeze.column);
    }
    Ok(())
}

pub fn generate<P: AsRef<Path>>(
    workbook_def: &WorkbookDef,
    filename: &str,
    base_path: P,
    canonical_transcripts: Option<HashSet<Vec<u8>>>,
) -> anyhow::Result<()> {
    let workbook = xlsxwriter::Workbook::new(filename)?;

    let format_defs = collect_format(workbook_def);
    let mut format_manager = FormatManager::new();
    for one in format_defs.iter() {
        format_manager.add_format(one)?;
    }

    for (sheet_index, one_sheet) in workbook_def.sheets.iter().enumerate() {
        let name = if let Some(name) = one_sheet.name.as_ref() {
            name.to_string()
        } else {
            format!("Sheet {}", sheet_index + 1)
        };
        let mut worksheet = workbook.add_worksheet(Some(&name))?;
        generate_worksheet(
            &mut worksheet,
            one_sheet,
            &format_manager,
            base_path.as_ref(),
            canonical_transcripts.clone(),
        )
        .with_context(|| format!("Error on generating \"{}\"", name))?;
    }

    workbook.close()?;
    Ok(())
}

pub fn load_list<P: AsRef<Path>>(path: P) -> anyhow::Result<HashSet<Vec<u8>>> {
    let mut reader = std::io::BufReader::new(autocompress::autodetect_open(path)?);
    let mut line = String::new();
    let mut transcripts = HashSet::new();
    loop {
        line.clear();
        reader.read_line(&mut line)?;
        transcripts.insert(line.trim().as_bytes().to_vec());
        if line.is_empty() {
            break;
        }
    }
    Ok(transcripts)
}

#[cfg(test)]
mod test;
