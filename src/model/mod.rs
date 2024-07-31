use serde::{Deserialize, Serialize};
use xlsxwriter::worksheet::{WorksheetCol, WorksheetRow};

#[derive(Debug, Clone, PartialEq, PartialOrd, Deserialize, Serialize)]
#[serde(untagged)]
pub enum CellValue {
    String(String),
    Null,
    Number(f64),
    Percent(f64),
    Boolean(bool),
    Url(String),
    Formula(String),
}

impl From<String> for CellValue {
    fn from(v: String) -> Self {
        CellValue::String(v)
    }
}

impl From<f64> for CellValue {
    fn from(v: f64) -> Self {
        CellValue::Number(v)
    }
}

impl From<bool> for CellValue {
    fn from(v: bool) -> Self {
        CellValue::Boolean(v)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Deserialize, Serialize, Hash)]
#[serde(rename_all = "kebab-case")]
pub enum CellType {
    String,
    Null,
    Boolean,
    Integer,
    Number,
    Percent,
    Datetime,
    Formula,
    Url,
    Auto,
}

impl Default for CellType {
    fn default() -> Self {
        CellType::Auto
    }
}

#[derive(Debug, Clone, PartialEq, PartialOrd, Deserialize, Serialize)]
#[serde(rename_all = "kebab-case")]
pub struct CellDef {
    #[serde(rename = "type", default)]
    pub cell_type: CellType,
    pub value: Option<CellValue>,
    pub row: Option<WorksheetRow>,
    pub row_relative: Option<i32>,
    pub column: Option<WorksheetCol>,
    pub column_relative: Option<i32>,
    pub format: Option<FormatDef>,
    pub comment: Option<String>,
    pub url: Option<String>,
    pub merge_row: Option<WorksheetRow>,
    pub merge_column: Option<WorksheetCol>,
}

#[derive(Debug, Clone, Copy, PartialEq, PartialOrd, Deserialize, Serialize, Eq, Hash, Ord)]
#[serde(rename_all = "kebab-case")]
pub enum BorderType {
    None,
    Thin,
    Medium,
    Dashed,
    Dotted,
    Thick,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}

#[derive(Debug, Clone, PartialEq, PartialOrd, Deserialize, Serialize, Eq, Hash, Ord)]
#[serde(rename_all = "kebab-case")]
pub struct BorderFormatDef {
    #[serde(rename = "type")]
    pub border_type: BorderType,
    pub color: Option<String>,
}

#[derive(Debug, Clone, PartialEq, PartialOrd, Deserialize, Serialize, Eq, Hash, Ord)]
#[serde(untagged)]
pub enum BorderFormatDefChoice {
    TypeOnly(BorderType),
    One(BorderFormatDef),
    Multi(Vec<BorderFormatDef>),
}

#[derive(Debug, Clone, PartialEq, PartialOrd, Deserialize, Serialize, Eq, Hash, Ord)]
pub struct BorderFormatAll {
    pub top: BorderFormatDef,
    pub bottom: BorderFormatDef,
    pub left: BorderFormatDef,
    pub right: BorderFormatDef,
}

impl BorderFormatDefChoice {
    pub fn parse(&self) -> BorderFormatAll {
        match self {
            BorderFormatDefChoice::TypeOnly(border_type) => {
                let border = BorderFormatDef {
                    border_type: *border_type,
                    color: None,
                };
                BorderFormatAll {
                    top: border.clone(),
                    bottom: border.clone(),
                    left: border.clone(),
                    right: border.clone(),
                }
            }
            BorderFormatDefChoice::One(border) => BorderFormatAll {
                top: border.clone(),
                bottom: border.clone(),
                left: border.clone(),
                right: border.clone(),
            },
            BorderFormatDefChoice::Multi(array) => match array.len() {
                1 => BorderFormatAll {
                    top: array[0].clone(),
                    bottom: array[0].clone(),
                    left: array[0].clone(),
                    right: array[0].clone(),
                },
                2 => BorderFormatAll {
                    top: array[0].clone(),
                    bottom: array[0].clone(),
                    left: array[1].clone(),
                    right: array[1].clone(),
                },
                3 => BorderFormatAll {
                    top: array[0].clone(),
                    bottom: array[2].clone(),
                    left: array[1].clone(),
                    right: array[1].clone(),
                },
                4 => BorderFormatAll {
                    top: array[0].clone(),
                    bottom: array[2].clone(),
                    left: array[3].clone(),
                    right: array[1].clone(),
                },
                _ => panic!("Invalid number of border content"),
            },
        }
    }
}

#[derive(Debug, Clone, PartialEq, PartialOrd, Deserialize, Serialize, Eq, Hash, Ord)]
#[serde(rename_all = "kebab-case")]
pub struct FormatDef {
    pub font_name: Option<String>,
    pub font_size: Option<u16>,
    pub font_color: Option<String>,
    #[serde(default)]
    pub underline: bool,
    pub background_color: Option<String>,
    pub num_format: Option<String>,
    pub border: Option<BorderFormatDefChoice>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Deserialize, Serialize, Hash)]
pub enum SheetSourceType {
    Auto,
    CSV,
    TSV,
    VCF,
}

impl Default for SheetSourceType {
    fn default() -> Self {
        SheetSourceType::Auto
    }
}

#[derive(Debug, Clone, PartialEq, PartialOrd, Deserialize, Serialize)]
#[serde(rename_all = "kebab-case")]
pub struct SheetSourceColumnDef {
    pub format: Option<FormatDef>,
    #[serde(rename = "type", default)]
    pub cell_type: CellType,
    #[serde(default)]
    pub header_type: CellType,
    pub header_value: Option<CellValue>,
    pub header_comment: Option<String>,
    pub link_prefix: Option<String>,
}

#[derive(Debug, Clone, PartialEq, PartialOrd, Deserialize, Serialize, Hash, Default)]
#[serde(rename_all = "kebab-case")]
pub struct VCFConfigDef {
    #[serde(default)]
    pub split_multi_allelic: bool,
    #[serde(default)]
    pub decode_genotype: bool,
    #[serde(default)]
    pub info: Option<Vec<String>>,
    #[serde(default)]
    pub format: Option<Vec<String>>,
    #[serde(default)]
    pub priority_info: Option<Vec<String>>,
    #[serde(default)]
    pub priority_format: Option<Vec<String>>,
}

#[derive(Debug, Clone, PartialEq, PartialOrd, Deserialize, Serialize)]
#[serde(untagged)]
pub enum SheetSource {
    Path(String),
    Def(Vec<SheetSourceDef>),
}

#[derive(Debug, Clone, Copy, PartialEq, PartialOrd, Deserialize, Serialize, Hash)]
#[serde(rename_all = "kebab-case")]
pub enum TableStyleType {
    Default,
    Light,
    Medium,
    Dark,
}

impl From<TableStyleType> for xlsxwriter::worksheet::table::TableStyleType {
    fn from(t: TableStyleType) -> Self {
        match t {
            TableStyleType::Default => xlsxwriter::worksheet::table::TableStyleType::Default,
            TableStyleType::Light => xlsxwriter::worksheet::table::TableStyleType::Light,
            TableStyleType::Medium => xlsxwriter::worksheet::table::TableStyleType::Medium,
            TableStyleType::Dark => xlsxwriter::worksheet::table::TableStyleType::Dark,
        }
    }
}

#[derive(Debug, Clone, PartialEq, PartialOrd, Deserialize, Serialize, Hash)]
#[serde(rename_all = "kebab-case")]
pub struct TableFilterList {
    pub column_header: String,
    pub items: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, PartialOrd, Deserialize, Serialize)]
#[serde(rename_all = "kebab-case")]
pub struct SheetSourceDef {
    pub file: Option<String>,
    pub data: Option<String>,
    #[serde(default)]
    pub format: SheetSourceType,
    pub columns: Option<Vec<SheetSourceColumnDef>>,
    #[serde(default = "true_value")]
    pub autofilter: bool,
    #[serde(default = "true_value")]
    pub table: bool,
    pub table_style_type: Option<TableStyleType>,
    pub table_style_type_num: Option<u8>,
    #[serde(default = "true_value")]
    pub has_header: bool,
    #[serde(default)]
    pub start_row: WorksheetRow,
    #[serde(default)]
    pub start_column: WorksheetCol,
    pub vcf_config: Option<VCFConfigDef>,
    pub comment_line_prefix: Option<String>,
    pub filter_list: Option<TableFilterList>,
}

impl SheetSourceDef {
    pub fn suggest_format(&self) -> SheetSourceType {
        if self.format == SheetSourceType::Auto {
            if let Some(file) = self.file.as_deref() {
                if file.ends_with(".vcf") || file.ends_with(".vcf.gz") {
                    SheetSourceType::VCF
                } else if file.ends_with(".csv") || file.ends_with(".csv.gz") {
                    SheetSourceType::CSV
                } else {
                    SheetSourceType::TSV
                }
            } else {
                SheetSourceType::TSV
            }
        } else {
            self.format
        }
    }
}

impl From<SheetSource> for Vec<SheetSourceDef> {
    fn from(x: SheetSource) -> Self {
        match x {
            SheetSource::Def(def) => def,
            SheetSource::Path(path) => vec![SheetSourceDef {
                file: Some(path),
                data: None,
                format: SheetSourceType::Auto,
                columns: None,
                autofilter: true,
                table: true,
                table_style_type: None,
                table_style_type_num: None,
                has_header: true,
                start_row: 0,
                start_column: 0,
                vcf_config: None,
                comment_line_prefix: None,
                filter_list: None,
            }],
        }
    }
}

#[derive(Debug, Clone, PartialEq, PartialOrd, Deserialize, Serialize, Hash)]
pub struct SheetFreeze {
    pub row: WorksheetRow,
    pub column: WorksheetCol,
}

#[derive(Debug, Clone, PartialEq, PartialOrd, Deserialize, Serialize)]
#[serde(rename_all = "kebab-case")]
pub struct SheetImage {
    pub file: String,
    pub row: WorksheetRow,
    pub column: WorksheetCol,
    #[serde(default)]
    pub width_scale: Option<f64>,
    #[serde(default)]
    pub height_scale: Option<f64>,
}

#[derive(Debug, Clone, PartialEq, PartialOrd, Deserialize, Serialize)]
#[serde(rename_all = "kebab-case")]
pub struct WorksheetDef {
    pub name: Option<String>,
    pub source: Option<SheetSource>,
    pub freeze: Option<SheetFreeze>,
    #[serde(default)]
    pub cells: Vec<CellDef>,
    #[serde(default)]
    pub column_widths: Vec<f64>,
    #[serde(default)]
    pub row_heights: Vec<f64>,
    #[serde(default)]
    pub images: Vec<SheetImage>,
}

#[derive(Debug, Clone, PartialEq, PartialOrd, Deserialize, Serialize)]
pub struct WorkbookDef {
    pub sheets: Vec<WorksheetDef>,
}

fn true_value() -> bool {
    true
}

#[cfg(test)]
mod test;
