use xlsxwriter::worksheet::table::{TableColumn, TableOptions, TableStyleType, TableTotalFunction};
use xlsxwriter::worksheet::{Worksheet, WorksheetCol, WorksheetRow};

use crate::model::SheetSourceDef;

use super::format::FormatManager;

pub fn setup_table(
    worksheet: &mut Worksheet,
    source_def: &SheetSourceDef,
    _formats: &FormatManager,
    column_header: &[String],
    filter_col: Option<WorksheetCol>,
    maximum_row: WorksheetRow,
    maximum_col: WorksheetCol,
) -> anyhow::Result<()> {
    if maximum_row <= 1 {
        return Ok(());
    }
    if source_def.table {
        let mut table_options = TableOptions::default();
        table_options.no_autofilter = !source_def.autofilter;
        table_options.no_header_row = !source_def.has_header;
        table_options.columns = Some(
            column_header
                .iter()
                .map(|x| TableColumn {
                    header: Some(x.to_string()),
                    formula: None,
                    total_string: None,
                    total_function: TableTotalFunction::None,
                    header_format: None,
                    format: None,
                    total_value: 0.,
                })
                .collect(),
        );
        table_options.style_type = source_def
            .table_style_type
            .map(|x| x.into())
            .unwrap_or(TableStyleType::Default);
        table_options.style_type_number = source_def.table_style_type_num.unwrap_or(0);
        worksheet.add_table(
            source_def.start_row,
            source_def.start_column,
            source_def.start_row + maximum_row,
            source_def.start_column + maximum_col,
            Some(table_options),
        )?;
    } else if source_def.autofilter {
        worksheet.autofilter(
            source_def.start_row,
            source_def.start_column,
            source_def.start_row + maximum_row,
            source_def.start_column + maximum_col,
        )?;
        if let Some(filter_list) = source_def.filter_list.as_ref() {
            if let Some(column_index) = filter_col {
                let list: Vec<_> = filter_list.items.iter().map(|x| x.as_str()).collect();
                worksheet.filter_list(column_index + source_def.start_column, &list)?;
            }
        }
    }

    Ok(())
}
