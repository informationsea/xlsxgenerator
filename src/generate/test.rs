use super::*;

#[test]
fn test_generate1() -> anyhow::Result<()> {
    let json_data = include_bytes!("../../examples/test1.json");
    let data: WorkbookDef = serde_json::from_reader(&json_data[..])?;
    generate(
        &data,
        "test1.xlsx",
        "examples",
        Some(load_list("examples/vcf/canonical.txt")?),
    )?;
    Ok(())
}

#[test]
fn test_generate2() -> anyhow::Result<()> {
    let json_data = include_bytes!("../../examples/test2.json");
    let data: WorkbookDef = serde_json::from_reader(&json_data[..])?;
    generate(
        &data,
        "test2.xlsx",
        "examples",
        Some(load_list("examples/real-vcf/canonical.txt.xz")?),
    )?;
    Ok(())
}
