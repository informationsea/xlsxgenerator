use super::vcf::{self, VCF2CSVConfig};
use super::FormatManager;
use crate::model::*;
use anyhow::Context;
use std::collections::HashSet;
use std::io::BufRead;
use std::path::Path;
use xlsxwriter::worksheet::{Worksheet, WorksheetCol, WorksheetRow};

pub fn insert_vcf<P: AsRef<Path>>(
    worksheet: &mut Worksheet,
    source_def: &SheetSourceDef,
    formats: &FormatManager,
    base_path: P,
    canonical_transcripts: Option<HashSet<Vec<u8>>>,
) -> anyhow::Result<()> {
    let mut vcf_data_reader: Box<dyn BufRead> = if let Some(file) = source_def.file.as_deref() {
        Box::new(std::io::BufReader::new(
            autocompress::autodetect_open(base_path.as_ref().join(file))
                .with_context(|| format!("Cannot open \"{}\"", file))?,
        ))
    } else if let Some(data) = source_def.data.as_deref() {
        Box::new(&data.as_bytes()[..])
    } else {
        return Err(anyhow::anyhow!("No data found for VCF"));
    };

    if source_def.table && source_def.filter_list.is_some() {
        eprintln!("filter_list cannot be enabled when table mode is enabled");
    }

    let mut vcf_reader = ::vcf::VCFReader::new(&mut vcf_data_reader)?;
    let priority_info_list = source_def
        .vcf_config
        .as_ref()
        .map(|x| {
            x.priority_info
                .as_ref()
                .map(|y| y.iter().map(|z| z.as_bytes().to_vec()).collect::<Vec<_>>())
        })
        .flatten()
        .unwrap_or_else(|| vec![]);
    let priority_info_hash = priority_info_list.iter().collect::<HashSet<_>>();
    let priority_format_list = source_def
        .vcf_config
        .as_ref()
        .map(|x| {
            x.priority_format
                .as_ref()
                .map(|y| y.iter().map(|z| z.as_bytes().to_vec()).collect::<Vec<_>>())
        })
        .flatten()
        .unwrap_or_else(|| vec![]);
    let priority_format_hash = priority_format_list.iter().collect::<HashSet<_>>();
    let info_list: Vec<Vec<u8>> = source_def
        .vcf_config
        .as_ref()
        .map(|x| x.info.as_ref())
        .flatten()
        .map(|x| x.iter().map(|y| y.as_bytes().to_vec()).collect())
        .unwrap_or_else(|| {
            let mut l = vcf_reader
                .header()
                .info_list()
                .cloned()
                .filter(|x| !priority_info_hash.contains(&x))
                .collect::<Vec<_>>();
            l.sort();
            l
        });
    let format_list: Vec<Vec<u8>> = source_def
        .vcf_config
        .as_ref()
        .map(|x| x.format.as_ref())
        .flatten()
        .map(|x| x.iter().map(|y| y.as_bytes().to_vec()).collect())
        .unwrap_or_else(|| {
            let mut l = vcf_reader
                .header()
                .format_list()
                .cloned()
                .filter(|x| !priority_format_hash.contains(&x))
                .collect::<Vec<_>>();
            l.sort();
            l
        });

    let config = VCF2CSVConfig {
        split_multi_allelic: source_def
            .vcf_config
            .as_ref()
            .map(|x| x.split_multi_allelic)
            .unwrap_or(false),
        decoded_genotype: source_def
            .vcf_config
            .as_ref()
            .map(|x| x.decode_genotype)
            .unwrap_or(false),
        canonical_list: canonical_transcripts,
        priority_info_list,
        priority_format_list,
        info_list,
        format_list,
        replace_sample_name: None,
        group_names: None,
    };

    let header_contents = vcf::create_header_line(&vcf_reader.header(), &config);
    let filter_column_index = source_def
        .filter_list
        .as_ref()
        .map(|x| {
            header_contents.iter().enumerate().find_map(|(i, y)| {
                if y.to_string() == x.column_header {
                    Some(i)
                } else {
                    None
                }
            })
        })
        .flatten();
    let filter_list: HashSet<String> = source_def
        .filter_list
        .as_ref()
        .map(|x| x.items.iter().map(|y| y.to_string()).collect())
        .unwrap_or_default();

    let mut writer = vcf::tablewriter::XlsxSheetWriter::new(
        worksheet,
        formats,
        source_def.start_row,
        source_def.start_column,
        filter_column_index,
        &filter_list,
    );

    vcf::vcf2table_set_data_type(&header_contents, &mut writer)?;
    let row_num = vcf::vcf2table(
        &mut vcf_reader,
        &header_contents,
        &config,
        None,
        true,
        &mut writer,
    )?;

    let column_widths = vcf::column_widths(&header_contents);
    for (i, one) in column_widths.iter().enumerate() {
        worksheet.set_column(
            source_def.start_column + i as WorksheetCol,
            source_def.start_column + i as WorksheetCol,
            8.0 * one,
            None,
        )?;
    }

    let column_header: Vec<_> = header_contents.iter().map(|x| x.to_string()).collect();

    super::table::setup_table(
        worksheet,
        source_def,
        formats,
        &column_header,
        filter_column_index.map(|x| x as WorksheetCol),
        row_num as WorksheetRow,
        column_header.len() as WorksheetCol - 1,
    )?;

    Ok(())
}
