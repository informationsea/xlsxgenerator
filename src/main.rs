pub mod generate;
pub mod jsonmarker;
pub mod model;
use std::path::Path;

use anyhow::Context;
use clap::Parser;
use generate::load_list;
use jsonmarker::render;

use crate::generate::generate;

#[derive(Parser, Debug)]
#[clap(author, version, about, long_about = None)]
struct Cli {
    #[clap(help = "Input excel file definition in JSON/YAML")]
    definition: String,
    #[clap(
        long = "parameter",
        short = 'p',
        help = "Parameter file to render definition template"
    )]
    parameter: Option<String>,
    #[clap(
        long = "output",
        short = 'o',
        required = true,
        help = "Output filename (required)"
    )]
    output_filename: String,
    #[clap(
        long = "vcf-canonical-transcript",
        short = 'c',
        help = "Input canonical transcript list to create VCF/SnpEff table",
        long_help = "Input canonical transcript list to create VCF/SnpEff table. Each line should have one transcript ID."
    )]
    vcf_canonical_transcript: Option<String>,
    #[clap(
        long = "base-path",
        short = 'b',
        help = "Base path to search additional CSV/VCF/Image"
    )]
    base_path: Option<String>,
}

fn main() -> anyhow::Result<()> {
    let cli = Cli::parse();
    let base_path = cli
        .base_path
        .as_deref()
        .map(|x| Path::new(x))
        .or_else(|| {
            cli.parameter
                .as_deref()
                .map(|x| Path::new(x).parent())
                .flatten()
        })
        .unwrap_or_else(|| {
            Path::new(&cli.definition)
                .parent()
                .unwrap_or_else(|| Path::new("/"))
        });

    let canonical_list = if let Some(canonical_transcripts) = cli.vcf_canonical_transcript.as_ref()
    {
        Some(load_list(canonical_transcripts).context(format!(
            "Cannot load canonical transcript: {}",
            canonical_transcripts
        ))?)
    } else {
        None
    };

    let workbook_def = serde_json::from_value(if let Some(parameter) = cli.parameter.as_deref() {
        let parameter_data = jsonmarker::load_data(parameter)
            .with_context(|| format!("Failed to load: {}", parameter))?;
        let template_data = jsonmarker::load_data(&cli.definition)
            .with_context(|| format!("Failed to load: {}", &cli.definition))?;
        render(&template_data, &parameter_data).context("Failed to render handlebars")?
    } else {
        jsonmarker::load_data(&cli.definition)?
    })?;

    generate(
        &workbook_def,
        &cli.output_filename,
        base_path,
        canonical_list,
    )?;
    Ok(())
}
