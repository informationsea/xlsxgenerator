use crate::model::*;
use once_cell::sync::Lazy;
use std::collections::{HashMap, HashSet};
use std::rc::Rc;
use xlsxwriter::{format::FormatUnderline, Format};

pub static EMPTY_FORMAT: FormatDef = FormatDef {
    font_name: None,
    font_size: None,
    font_color: None,
    background_color: None,
    num_format: None,
    border: None,
    underline: false,
};

pub static PERCENT_FORMAT: Lazy<FormatDef> = Lazy::new(|| FormatDef {
    font_name: None,
    font_size: None,
    font_color: None,
    background_color: None,
    num_format: Some("0.0%".to_string()),
    border: None,
    underline: false,
});

pub static URL_FORMAT: Lazy<FormatDef> = Lazy::new(|| FormatDef {
    font_name: None,
    font_size: None,
    font_color: Some("blue".to_string()),
    background_color: None,
    num_format: None,
    border: None,
    underline: true,
});

#[derive(Default)]
pub struct FormatManager {
    string_format: HashMap<FormatDef, Rc<xlsxwriter::Format>>,
    integer_format: HashMap<FormatDef, Rc<xlsxwriter::Format>>,
    float_format: HashMap<FormatDef, Rc<xlsxwriter::Format>>,
    percent_format: HashMap<FormatDef, Rc<xlsxwriter::Format>>,
    date_format: HashMap<FormatDef, Rc<xlsxwriter::Format>>,
    url_format: HashMap<FormatDef, Rc<xlsxwriter::Format>>,
    general_format: HashMap<FormatDef, Rc<xlsxwriter::Format>>,
}

fn convert_border_type(border_type: BorderType) -> xlsxwriter::format::FormatBorder {
    match border_type {
        BorderType::None => xlsxwriter::format::FormatBorder::None,
        BorderType::Thin => xlsxwriter::format::FormatBorder::Thin,
        BorderType::Medium => xlsxwriter::format::FormatBorder::Medium,
        BorderType::Dashed => xlsxwriter::format::FormatBorder::Dashed,
        BorderType::Dotted => xlsxwriter::format::FormatBorder::Dotted,
        BorderType::Thick => xlsxwriter::format::FormatBorder::Thick,
        BorderType::Double => xlsxwriter::format::FormatBorder::Double,
        BorderType::Hair => xlsxwriter::format::FormatBorder::Hair,
        BorderType::MediumDashed => xlsxwriter::format::FormatBorder::MediumDashed,
        BorderType::MediumDashDot => xlsxwriter::format::FormatBorder::MediumDashDot,
        BorderType::DashDotDot => xlsxwriter::format::FormatBorder::DashDotDot,
        BorderType::MediumDashDotDot => xlsxwriter::format::FormatBorder::MediumDashDotDot,
        BorderType::SlantDashDot => xlsxwriter::format::FormatBorder::SlantDashDot,
        BorderType::DashDot => xlsxwriter::format::FormatBorder::DashDot,
    }
}

fn create_format_base<'a>(new_format: &mut Format, format_def: &FormatDef) -> anyhow::Result<()> {
    if let Some(font_name) = format_def.font_name.as_deref() {
        new_format.set_font_name(font_name);
    }
    if let Some(font_size) = format_def.font_size {
        new_format.set_font_size(font_size.into());
    }
    if let Some(color) = format_def.font_color.as_deref() {
        new_format.set_font_color(color_parse(color)?);
    }
    if let Some(color) = format_def.background_color.as_deref() {
        new_format.set_bg_color(color_parse(color)?);
    }
    if let Some(border) = format_def.border.as_ref() {
        let parsed = border.parse();
        new_format
            .set_border_top(convert_border_type(parsed.top.border_type))
            .set_border_bottom(convert_border_type(parsed.bottom.border_type))
            .set_border_left(convert_border_type(parsed.left.border_type))
            .set_border_right(convert_border_type(parsed.right.border_type));
        if let Some(color) = parsed.top.color.as_deref() {
            new_format.set_border_top_color(color_parse(color)?);
        }
        if let Some(color) = parsed.bottom.color.as_deref() {
            new_format.set_border_bottom_color(color_parse(color)?);
        }
        if let Some(color) = parsed.left.color.as_deref() {
            new_format.set_border_left_color(color_parse(color)?);
        }
        if let Some(color) = parsed.right.color.as_deref() {
            new_format.set_border_right_color(color_parse(color)?);
        }
    }
    if format_def.underline {
        new_format.set_underline(FormatUnderline::Single);
    }
    Ok(())
}

impl FormatManager {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn add_format(&mut self, format_def: &FormatDef) -> anyhow::Result<()> {
        if let Some(num_format) = format_def.num_format.as_deref() {
            let mut new_format = Format::new();
            create_format_base(&mut new_format, format_def)?;
            new_format.set_num_format(num_format);
            let new_format = Rc::new(new_format);
            self.string_format
                .insert(format_def.clone(), new_format.clone());
            self.integer_format
                .insert(format_def.clone(), new_format.clone());
            self.float_format
                .insert(format_def.clone(), new_format.clone());
            self.percent_format
                .insert(format_def.clone(), new_format.clone());
            self.date_format
                .insert(format_def.clone(), new_format.clone());
            self.general_format
                .insert(format_def.clone(), new_format.clone());
            self.url_format
                .insert(format_def.clone(), new_format.clone());
        } else {
            let mut string_format = Format::new();
            create_format_base(&mut string_format, format_def)?;
            string_format.set_num_format("@");
            self.string_format
                .insert(format_def.clone(), Rc::new(string_format));

            let mut integer_format = Format::new();
            create_format_base(&mut integer_format, format_def)?;
            integer_format.set_num_format("@");
            self.integer_format
                .insert(format_def.clone(), Rc::new(integer_format));

            let mut percent_format = Format::new();
            create_format_base(&mut percent_format, format_def)?;
            percent_format.set_num_format("0.0%");
            self.percent_format
                .insert(format_def.clone(), Rc::new(percent_format));

            let mut float_format = Format::new();
            create_format_base(&mut float_format, format_def)?;
            self.float_format
                .insert(format_def.clone(), Rc::new(float_format));

            let mut date_format = Format::new();
            create_format_base(&mut date_format, format_def)?;
            date_format.set_num_format("@");
            self.date_format
                .insert(format_def.clone(), Rc::new(date_format));

            let mut general_format = Format::new();
            create_format_base(&mut general_format, format_def)?;
            self.general_format
                .insert(format_def.clone(), Rc::new(general_format));

            if format_def == &EMPTY_FORMAT {
                let mut f = Format::new();
                create_format_base(&mut f, format_def)?;
                f.set_underline(xlsxwriter::format::FormatUnderline::Single)
                    .set_font_color(xlsxwriter::format::FormatColor::Blue);
                self.url_format.insert(format_def.clone(), Rc::new(f));
            } else {
                let mut f = Format::new();
                create_format_base(&mut f, format_def)?;
                self.url_format.insert(format_def.clone(), Rc::new(f));
            }
        }
        Ok(())
    }

    pub fn get_format(
        &self,
        format_def: Option<&FormatDef>,
        cell_type: CellType,
    ) -> Option<&Format> {
        let format_def = format_def.unwrap_or(&EMPTY_FORMAT);

        match cell_type {
            CellType::Integer => self
                .integer_format
                .get(format_def)
                .or_else(|| self.integer_format.get(&EMPTY_FORMAT)),
            CellType::Number => self
                .float_format
                .get(format_def)
                .or_else(|| self.float_format.get(&EMPTY_FORMAT)),
            CellType::Percent => self
                .percent_format
                .get(format_def)
                .or_else(|| self.float_format.get(&EMPTY_FORMAT)),
            CellType::Datetime => self
                .date_format
                .get(format_def)
                .or_else(|| self.date_format.get(&EMPTY_FORMAT)),
            CellType::String => self
                .string_format
                .get(format_def)
                .or_else(|| self.string_format.get(&EMPTY_FORMAT)),
            CellType::Url => self
                .url_format
                .get(format_def)
                .or_else(|| self.string_format.get(&EMPTY_FORMAT)),
            _ => self.general_format.get(format_def),
        }
        .map(|x| x.as_ref())
    }
}

pub fn color_parse(color: &str) -> anyhow::Result<xlsxwriter::format::FormatColor> {
    match color {
        "black" | "#000000" => Ok(xlsxwriter::format::FormatColor::Black),
        "blue" => Ok(xlsxwriter::format::FormatColor::Blue),
        "brown" => Ok(xlsxwriter::format::FormatColor::Brown),
        "cyan" => Ok(xlsxwriter::format::FormatColor::Cyan),
        "gray" => Ok(xlsxwriter::format::FormatColor::Gray),
        "green" => Ok(xlsxwriter::format::FormatColor::Green),
        "lime" => Ok(xlsxwriter::format::FormatColor::Lime),
        "magenta" => Ok(xlsxwriter::format::FormatColor::Magenta),
        "navy" => Ok(xlsxwriter::format::FormatColor::Navy),
        "orange" => Ok(xlsxwriter::format::FormatColor::Orange),
        "purple" => Ok(xlsxwriter::format::FormatColor::Purple),
        "red" => Ok(xlsxwriter::format::FormatColor::Red),
        "sliver" => Ok(xlsxwriter::format::FormatColor::Silver),
        "pink" => Ok(xlsxwriter::format::FormatColor::Pink),
        "white" => Ok(xlsxwriter::format::FormatColor::White),
        "yellow" => Ok(xlsxwriter::format::FormatColor::Yellow),
        _ => {
            if color.starts_with("#") && color.len() == 7 {
                let val = u32::from_str_radix(&color[1..], 16)?;
                Ok(xlsxwriter::format::FormatColor::Custom(val))
            } else {
                Err(anyhow::anyhow!("{} is not valid color", color))
            }
        }
    }
}

pub fn collect_format(workbook_def: &WorkbookDef) -> HashSet<FormatDef> {
    let mut set = HashSet::new();

    for one_sheet in &workbook_def.sheets {
        if let Some(source) = one_sheet.source.as_ref() {
            if let SheetSource::Def(def_array) = source {
                for def in def_array.iter() {
                    if let Some(columns) = &def.columns {
                        for one_column in columns {
                            if let Some(format) = one_column.format.as_ref() {
                                set.insert(format.clone());
                            }
                        }
                    }
                }
            }
        }

        for one_cell in &one_sheet.cells {
            if let Some(format) = one_cell.format.as_ref() {
                set.insert(format.clone());
            }
        }
    }

    set.insert(EMPTY_FORMAT.clone());
    set.insert(PERCENT_FORMAT.clone());
    set.insert(URL_FORMAT.clone());

    set
}
