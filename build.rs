use anyhow::Result;
use std::fs;
use std::path::Path;

fn main() -> Result<()> {
    println!("cargo:rerun-if-changed=build.rs");
    
    // Create assets directory if it doesn't exist
    let fonts_dir = Path::new("assets/fonts");
    fs::create_dir_all(fonts_dir)?;
    
    // Check if fonts exist, if not, create placeholder files
    // In production, you would download actual font files here
    let font_files = vec![
        "LiberationSans-Regular.ttf",
        "LiberationSans-Bold.ttf",
        "LiberationSans-Italic.ttf",
        "LiberationMono-Regular.ttf",
        "NotoSans-Regular.ttf",
        "NotoSans-Bold.ttf",
    ];
    
    for font_file in font_files {
        let font_path = fonts_dir.join(font_file);
        if !font_path.exists() {
            // For now, we'll create empty placeholder files
            // In production, download actual Liberation or Noto fonts (which are open source)
            println!("cargo:warning=Font file {} not found. Please download Liberation fonts from https://github.com/liberationfonts/liberation-fonts", font_file);
            
            // Create a minimal placeholder TTF file (this won't work for actual rendering)
            // You should download the actual fonts
            fs::write(&font_path, &[0u8; 100])?;
        }
    }
    
    Ok(())
}