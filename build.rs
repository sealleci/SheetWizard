use std::env::var;
use std::fs::copy;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let profile = var("PROFILE")?;

    if profile == "release" {
        copy("path.toml", "./target/release/path.toml")?;
    }

    Ok(())
}
