use xlsxwriter::worksheet::{Worksheet, WorksheetCol, WorksheetRow};
use xlsxwriter::Format;

use crate::model::*;

pub fn parse_cell(data: &str, cell_type: CellType) -> anyhow::Result<CellValue> {
    if data.is_empty() {
        return Ok(CellValue::Null);
    }
    match cell_type {
        CellType::Boolean => match data.to_uppercase().as_str() {
            "TRUE" | "YES" => Ok(CellValue::Boolean(true)),
            "FALSE" | "NO" => Ok(CellValue::Boolean(false)),
            _ => Err(anyhow::anyhow!("\"{}\" is not boolean value", data)),
        },
        CellType::Integer | CellType::Number => {
            let num = data.trim().parse::<f64>();
            if let Ok(num) = num {
                Ok(CellValue::Number(num))
            } else {
                Err(anyhow::anyhow!("\"{}\" is not number", data))
            }
        }
        CellType::Percent => {
            let num = data.trim().parse::<f64>();
            if let Ok(num) = num {
                Ok(CellValue::Percent(num / 100.0))
            } else {
                Err(anyhow::anyhow!("\"{}\" is not number", data))
            }
        }
        CellType::String => Ok(CellValue::String(data.to_string())),
        CellType::Formula => Ok(CellValue::Formula(data.to_string())),
        CellType::Url => Ok(CellValue::Url(data.to_string())),
        CellType::Auto => {
            if let Ok(num) = data.parse::<f64>() {
                Ok(CellValue::Number(num))
            } else if data.starts_with("=") {
                Ok(CellValue::Formula(data.to_string()))
            } else if data.starts_with("https://")
                || data.starts_with("http://")
                || data.starts_with("mailto:")
                || data.starts_with("internal:")
                || data.starts_with("external:")
            {
                Ok(CellValue::Url(data.to_string()))
            } else {
                Ok(CellValue::String(data.to_string()))
            }
        }
        _ => {
            eprintln!("Not implemented: {:?}", cell_type);
            unimplemented!()
        }
    }
}

pub fn parse_cell_value(data: &CellValue, cell_type: CellType) -> anyhow::Result<CellValue> {
    match data {
        CellValue::String(value) => parse_cell(&value, cell_type),
        CellValue::Number(value) => match cell_type {
            CellType::Percent => Ok(CellValue::Percent(value / 100.)),
            _ => Ok(CellValue::Number(*value)),
        },
        _ => Ok(data.clone()),
    }
}

pub fn actual_cell_type(value: &CellValue, cell_type: CellType) -> CellType {
    match value {
        CellValue::Boolean(_) => CellType::Boolean,
        CellValue::Formula(_) => CellType::Formula,
        CellValue::Null => CellType::Null,
        CellValue::String(_) => CellType::String,
        CellValue::Url(_) => CellType::Url,
        CellValue::Percent(_) => CellType::Percent,
        CellValue::Number(_) => match cell_type {
            CellType::Integer => CellType::Integer,
            CellType::Percent => CellType::Percent,
            _ => CellType::Number,
        },
    }
}

pub fn write_cell(
    worksheet: &mut Worksheet,
    row: WorksheetRow,
    column: WorksheetCol,
    value: &CellValue,
    format: Option<&Format>,
) -> anyhow::Result<()> {
    match value {
        CellValue::Boolean(val) => {
            worksheet.write_boolean(row, column, *val, format)?;
        }
        CellValue::Null => {
            worksheet.write_blank(row, column, format)?;
        }
        CellValue::Number(val) | CellValue::Percent(val) => {
            worksheet.write_number(row, column, *val, format)?;
        }
        CellValue::String(val) => {
            worksheet.write_string(row, column, val, format)?;
        }
        CellValue::Url(val) => {
            worksheet.write_url(row, column, val, format)?;
        }
        CellValue::Formula(val) => {
            worksheet.write_formula(row, column, val, format)?;
        }
    }
    Ok(())
}
