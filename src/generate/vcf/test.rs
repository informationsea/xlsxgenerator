use super::super::{FormatManager, EMPTY_FORMAT};
use super::tablewriter;
use super::*;
use std::io::BufReader;

#[test]
fn test_write_comma_separated_values() {
    let mut result = Vec::new();
    write_comma_separated_values(&mut result, &[b"AB".to_vec(), b"CD".to_vec()][..]);
    assert_eq!(result, b"AB,CD");
    result.clear();
    write_comma_separated_values(&mut result, &[b"AB".to_vec()][..]);
    assert_eq!(result, b"AB");
    result.clear();
    write_comma_separated_values(
        &mut result,
        &[b"AB".to_vec(), b"CD".to_vec(), b"EF".to_vec()][..],
    );
    assert_eq!(result, b"AB,CD,EF");
}

#[test]
fn test_write_value_for_alt_index() {
    let mut buffer = Vec::new();
    let values = &[b"A".to_vec(), b"B".to_vec(), b"C".to_vec()][..];
    write_value_for_alt_index(&mut buffer, values, &vcf::Number::Unknown, 0, Some(0));
    assert_eq!(buffer, b"A,B,C");
    buffer.clear();

    write_value_for_alt_index(&mut buffer, values, &vcf::Number::Number(3), 0, Some(0));
    assert_eq!(buffer, b"A");
    buffer.clear();

    write_value_for_alt_index(&mut buffer, values, &vcf::Number::Number(3), 1, Some(0));
    assert_eq!(buffer, b"B");
    buffer.clear();

    write_value_for_alt_index(&mut buffer, values, &vcf::Number::Number(3), 2, Some(0));
    assert_eq!(buffer, b"C");
    buffer.clear();

    write_value_for_alt_index(&mut buffer, values, &vcf::Number::Allele, 0, Some(0));
    assert_eq!(buffer, b"A");
    buffer.clear();

    write_value_for_alt_index(&mut buffer, values, &vcf::Number::Allele, 0, Some(1));
    assert_eq!(buffer, b"B");
    buffer.clear();

    write_value_for_alt_index(&mut buffer, values, &vcf::Number::Reference, 0, Some(1));
    assert_eq!(buffer, b"A");
    buffer.clear();

    write_value_for_alt_index(&mut buffer, values, &vcf::Number::Reference, 1, Some(1));
    assert_eq!(buffer, b"C");
    buffer.clear();
}

#[test]
fn test_vcf2table_csv_split_multi() -> Result<()> {
    let vcf_data = include_bytes!("../../../examples/vcf/simple1.vcf");
    let config = VCF2CSVConfig {
        split_multi_allelic: true,
        decoded_genotype: false,
        canonical_list: None,
        priority_info_list: vec![],
        priority_format_list: vec![],
        info_list: vec![
            b"AC".to_vec(),
            b"AF".to_vec(),
            b"AN".to_vec(),
            b"DP".to_vec(),
        ],
        format_list: vec![b"AD".to_vec(), b"DP".to_vec(), b"GT".to_vec()],
        replace_sample_name: None,
        group_names: None,
    };
    let mut vcf_data_reader = BufReader::new(&vcf_data[..]);
    let mut vcf_reader = vcf::VCFReader::new(&mut vcf_data_reader)?;
    let mut write_bytes = Vec::<u8>::new();
    let header_contents = create_header_line(&vcf_reader.header(), &config);
    vcf2table(
        &mut vcf_reader,
        &header_contents,
        &config,
        None,
        true,
        &mut tablewriter::CSVWriter::new(&mut write_bytes),
    )?;
    std::fs::File::create("./target/simple1-expected-multiallelic-split.csv")?
        .write_all(&write_bytes)?;
    assert_eq!(
        &write_bytes[..],
        &include_bytes!("../../../examples/vcf/simple1-expected-multiallelic-split.csv")[..]
    );
    Ok(())
}

#[test]
fn test_vcf2table_csv_split_multi_snpeff() -> Result<()> {
    let vcf_data = include_bytes!("../../../examples/vcf/simple1-snpeff.vcf");
    let config = VCF2CSVConfig {
        split_multi_allelic: true,
        decoded_genotype: true,
        canonical_list: None,
        priority_info_list: vec![],
        priority_format_list: vec![],
        info_list: vec![
            b"AC".to_vec(),
            b"AF".to_vec(),
            b"AN".to_vec(),
            b"DP".to_vec(),
            b"ANN".to_vec(),
            b"FLAG".to_vec(),
        ],
        format_list: vec![b"AD".to_vec(), b"DP".to_vec(), b"GT".to_vec()],
        replace_sample_name: None,
        group_names: None,
    };
    let mut vcf_data_reader = BufReader::new(&vcf_data[..]);
    let mut vcf_reader = vcf::VCFReader::new(&mut vcf_data_reader)?;
    let mut write_bytes = Vec::<u8>::new();
    let header_contents = create_header_line(&vcf_reader.header(), &config);
    vcf2table(
        &mut vcf_reader,
        &header_contents,
        &config,
        None,
        true,
        &mut tablewriter::CSVWriter::new(&mut write_bytes),
    )?;
    std::fs::File::create("./target/split-snpeff.csv")?.write_all(&write_bytes)?;
    assert_eq!(
        &write_bytes[..],
        &include_bytes!("../../../examples/vcf/simple1-expected-multiallelic-split-snpeff.csv")[..]
    );
    Ok(())
}

#[test]
fn test_vcf2table_csv_split_multi_snpeff_with_canonical() -> Result<()> {
    let vcf_data = include_bytes!("../../../examples/vcf/1kGP-subset-snpeff.vcf");
    let config = VCF2CSVConfig {
        split_multi_allelic: true,
        decoded_genotype: true,
        canonical_list: Some(
            [b"ENST00000380152.7_1"]
                .iter()
                .map(|x| x.to_vec())
                .collect(),
        ),
        priority_info_list: vec![],
        priority_format_list: vec![],
        info_list: vec![
            b"AC".to_vec(),
            b"AF".to_vec(),
            b"AN".to_vec(),
            b"DP".to_vec(),
            b"ANN".to_vec(),
        ],
        format_list: vec![b"AD".to_vec(), b"DP".to_vec(), b"GT".to_vec()],
        replace_sample_name: None,
        group_names: None,
    };
    let mut vcf_data_reader = BufReader::new(&vcf_data[..]);
    let mut vcf_reader = vcf::VCFReader::new(&mut vcf_data_reader)?;
    let mut write_bytes = Vec::<u8>::new();
    let header_contents = create_header_line(&vcf_reader.header(), &config);
    vcf2table(
        &mut vcf_reader,
        &header_contents,
        &config,
        None,
        true,
        &mut tablewriter::CSVWriter::new(&mut write_bytes),
    )?;
    std::fs::File::create("./target/split-snpeff-canonical.csv")?.write_all(&write_bytes)?;
    // assert_eq!(
    //     &write_bytes[..],
    //     &include_bytes!("../../testfiles/simple1-expected-multiallelic-split-snpeff.csv")[..]
    // );
    Ok(())
}

#[test]
fn test_vcf2table_csv_no_split() -> Result<()> {
    let vcf_data = include_bytes!("../../../examples/vcf/simple1.vcf");
    let config = VCF2CSVConfig {
        split_multi_allelic: false,
        decoded_genotype: false,
        canonical_list: None,
        priority_info_list: vec![],
        priority_format_list: vec![],
        info_list: vec![
            b"AC".to_vec(),
            b"AF".to_vec(),
            b"AN".to_vec(),
            b"DP".to_vec(),
        ],
        format_list: vec![b"AD".to_vec(), b"DP".to_vec(), b"GT".to_vec()],
        replace_sample_name: None,
        group_names: None,
    };
    let mut vcf_data_reader = BufReader::new(&vcf_data[..]);
    let mut vcf_reader = vcf::VCFReader::new(&mut vcf_data_reader)?;
    let mut write_bytes = Vec::<u8>::new();
    let header_contents = create_header_line(&vcf_reader.header(), &config);
    vcf2table(
        &mut vcf_reader,
        &header_contents,
        &config,
        None,
        true,
        &mut tablewriter::CSVWriter::new(&mut write_bytes),
    )?;
    assert_eq!(
        &write_bytes[..],
        &include_bytes!("../../../examples/vcf/simple1-expected-no-split.csv")[..]
    );
    Ok(())
}

#[test]
fn test_vcf2table_xlsx_split_multi() -> Result<()> {
    let vcf_data = include_bytes!("../../../examples/vcf/simple1.vcf");
    let config = VCF2CSVConfig {
        split_multi_allelic: true,
        decoded_genotype: false,
        canonical_list: None,
        priority_info_list: vec![],
        priority_format_list: vec![],
        info_list: vec![
            b"AC".to_vec(),
            b"AN".to_vec(),
            b"AF".to_vec(),
            b"DP".to_vec(),
        ],
        format_list: vec![b"GT".to_vec(), b"AD".to_vec(), b"DP".to_vec()],
        replace_sample_name: None,
        group_names: None,
    };
    let mut vcf_data_reader = BufReader::new(&vcf_data[..]);
    let mut vcf_reader = vcf::VCFReader::new(&mut vcf_data_reader)?;
    let workbook = xlsxwriter::Workbook::new("./target/table-split-multi.xlsx")?;
    let mut format_manager = FormatManager::new();
    format_manager.add_format(&EMPTY_FORMAT)?;
    let mut sheet = workbook.add_worksheet(None)?;
    let empty_hash = HashSet::new();
    let mut writer =
        tablewriter::XlsxSheetWriter::new(&mut sheet, &format_manager, 2, 3, None, &empty_hash);
    let header_contents = create_header_line(&vcf_reader.header(), &config);
    vcf2table_set_data_type(&header_contents, &mut writer)?;
    let row = vcf2table(
        &mut vcf_reader,
        &header_contents,
        &config,
        None,
        true,
        &mut writer,
    )?;
    sheet.add_table(2, 3, row + 2, (header_contents.len() - 1) as u16 + 3, None)?;
    workbook.close()?;
    Ok(())
}

#[test]
fn test_vcf2table_xlsx_no_split_multi() -> Result<()> {
    let vcf_data = include_bytes!("../../../examples/vcf/simple1.vcf");
    let config = VCF2CSVConfig {
        split_multi_allelic: false,
        decoded_genotype: false,
        canonical_list: None,
        priority_info_list: vec![],
        priority_format_list: vec![],
        info_list: vec![
            b"AC".to_vec(),
            b"AN".to_vec(),
            b"AF".to_vec(),
            b"DP".to_vec(),
        ],
        format_list: vec![b"GT".to_vec(), b"AD".to_vec(), b"DP".to_vec()],
        replace_sample_name: None,
        group_names: None,
    };
    let mut vcf_data_reader = BufReader::new(&vcf_data[..]);
    let mut vcf_reader = vcf::VCFReader::new(&mut vcf_data_reader)?;
    let workbook = xlsxwriter::Workbook::new("./target/table-no-split.xlsx")?;
    let mut format_manager = FormatManager::new();
    format_manager.add_format(&EMPTY_FORMAT)?;
    let mut sheet = workbook.add_worksheet(None)?;
    let empty_hash = HashSet::new();
    let mut writer =
        tablewriter::XlsxSheetWriter::new(&mut sheet, &format_manager, 0, 0, None, &empty_hash);
    let header_contents = create_header_line(&vcf_reader.header(), &config);
    vcf2table_set_data_type(&header_contents, &mut writer)?;
    let row = vcf2table(
        &mut vcf_reader,
        &header_contents,
        &config,
        None,
        true,
        &mut writer,
    )?;
    sheet.autofilter(0, 0, row, (header_contents.len() - 1) as u16)?;
    workbook.close()?;
    Ok(())
}

#[test]
fn test_vcf2table_xlsx_split_multi_with_group_name() -> Result<()> {
    let vcf_data = include_bytes!("../../../examples/vcf/simple1.vcf");
    let config = VCF2CSVConfig {
        split_multi_allelic: true,
        decoded_genotype: false,
        canonical_list: None,
        priority_info_list: vec![],
        priority_format_list: vec![],
        info_list: vec![
            b"AC".to_vec(),
            b"AN".to_vec(),
            b"AF".to_vec(),
            b"DP".to_vec(),
        ],
        format_list: vec![b"GT".to_vec(), b"AD".to_vec(), b"DP".to_vec()],
        replace_sample_name: Some(vec![b"SAMPLE1".to_vec()]),
        group_names: Some(vec![b"GROUP".to_vec()]),
    };
    let mut vcf_data_reader = BufReader::new(&vcf_data[..]);
    let mut vcf_reader = vcf::VCFReader::new(&mut vcf_data_reader)?;
    let workbook = xlsxwriter::Workbook::new("./target/table-split-multi-with-group.xlsx")?;
    let mut format_manager = FormatManager::new();
    format_manager.add_format(&EMPTY_FORMAT)?;
    let mut sheet = workbook.add_worksheet(None)?;
    let empty_hash = HashSet::new();
    let mut writer =
        tablewriter::XlsxSheetWriter::new(&mut sheet, &format_manager, 0, 0, None, &empty_hash);
    let header_contents = create_header_line(&vcf_reader.header(), &config);
    vcf2table_set_data_type(&header_contents, &mut writer)?;
    let row = vcf2table(
        &mut vcf_reader,
        &header_contents,
        &config,
        None,
        true,
        &mut writer,
    )?;
    sheet.autofilter(0, 0, row, (header_contents.len() - 1) as u16)?;
    workbook.close()?;
    Ok(())
}
