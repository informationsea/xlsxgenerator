use anyhow::Result;
use handlebars::{handlebars_helper, Handlebars};
use serde_json::Value;
use std::path::Path;

const YAML_SUFFIX: &[&'static str] = &[".yaml", ".yaml.gz", ".yml", ".yml.gz"];

pub fn load_data<P: AsRef<Path>>(path: P) -> Result<Value> {
    let path_str = path.as_ref().to_string_lossy();
    let reader = autocompress::autodetect_open(path.as_ref())?;

    if YAML_SUFFIX.iter().any(|x| path_str.ends_with(x)) {
        Ok(serde_yaml::from_reader(reader)?)
    } else {
        Ok(serde_json::from_reader(reader)?)
    }
}

pub fn save_data<P: AsRef<Path>>(value: &Value, path: P, indent: bool) -> Result<()> {
    let writer = autocompress::autodetect_create(path, autocompress::CompressionLevel::Default)?;
    if indent {
        serde_json::to_writer_pretty(writer, value)?;
    } else {
        serde_json::to_writer(writer, value)?;
    }
    Ok(())
}

fn internal_render(value: &mut Value, parameters: &Value, reg: &mut Handlebars) -> Result<()> {
    match value {
        Value::Array(array) => {
            for one in array.iter_mut() {
                internal_render(one, parameters, reg)?;
            }
        }
        Value::Object(map) => {
            for one in map.iter_mut() {
                internal_render(one.1, parameters, reg)?;
            }
        }
        Value::String(string) => *string = reg.render_template(&string, parameters)?,
        _ => (),
    }
    Ok(())
}

handlebars_helper!(as_percent: |x: f64| format!("{:.2}", x * 100.));
handlebars_helper!(as_ratio: |x: f64|  x / 100.);
handlebars_helper!(div: |x: f64, y: f64|  x / y);
handlebars_helper!(mul: |x: f64, y: f64|  x * y);
handlebars_helper!(add: |x: f64, y: f64|  x + y);
handlebars_helper!(sub: |x: f64, y: f64|  x - y);

pub fn render(template: &Value, parameters: &Value) -> Result<Value> {
    let mut value = template.clone();
    let mut reg = Handlebars::new();
    reg.register_helper("as_percent", Box::new(as_percent));
    reg.register_helper("as_ratio", Box::new(as_ratio));
    reg.register_helper("div", Box::new(div));
    reg.register_helper("mul", Box::new(mul));
    reg.register_helper("add", Box::new(add));
    reg.register_helper("sub", Box::new(sub));

    internal_render(&mut value, parameters, &mut reg)?;
    Ok(value)
}

#[cfg(test)]
mod test {
    use super::*;
    use serde_json::json;

    #[test]
    fn test_render() -> Result<()> {
        let template: Value = json!({
            "foo": "{{foo}}",
            "array": [
                "{{data1}}",
                "data2",
                "{{data3}}"
            ],
            "integer": 10,
            "boolean": false,
            "object": {
                "key": "{{data4}}"
            },
            "float": "{{as_percent f }}"
        });
        let parameters: Value = json!({
            "foo": "FOO",
            "data1": "DATA1",
            "data3": "DATA3",
            "data4": "DATA4",
            "f": 0.12345
        });
        let result = render(&template, &parameters)?;
        //let writer = std::fs::File::create("target/result.json")?;
        //serde_json::to_writer_pretty(writer, &result)?;

        assert_eq!(
            result,
            json!({
                "foo": "FOO",
                "array": [
                    "DATA1",
                    "data2",
                    "DATA3"
                ],
                "integer": 10,
                "boolean": false,
                "object": {
                    "key": "DATA4"
                },
                "float": "12.35"
            })
        );

        Ok(())
    }
}
