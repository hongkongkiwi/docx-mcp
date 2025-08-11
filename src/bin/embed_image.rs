use std::fs::{self, File};
use std::path::PathBuf;

use anyhow::Result;
use docx_rs::{Docx, Paragraph, Run, Pic, BreakType};

fn main() -> Result<()> {
    // Generate a simple 100x100 PNG in-memory (red square)
    let width = 100u32;
    let height = 100u32;
    let mut img = ::image::RgbaImage::new(width, height);
    for y in 0..height {
        for x in 0..width {
            img.put_pixel(x, y, ::image::Rgba([255, 0, 0, 255]));
        }
    }
    let mut png_bytes: Vec<u8> = Vec::new();
    let dyn_img = ::image::DynamicImage::ImageRgba8(img);
    dyn_img.write_to(&mut std::io::Cursor::new(&mut png_bytes), ::image::ImageFormat::Png)?;

    // Build a DOCX with an image and a caption
    let mut docx = Docx::new();

    let para = Paragraph::new()
        .add_run(Run::new().add_text("Embedded image demo").bold().size(28))
        .add_run(Run::new().add_break(BreakType::TextWrapping));
    docx = docx.add_paragraph(para);

    let image_para = Paragraph::new().add_run({
        let run = Run::new();
        run.add_image(Pic::new_with_dimensions(png_bytes, width, height))
    });
    docx = docx.add_paragraph(image_para);

    // Ensure output directory exists
    let out_dir = PathBuf::from("example/output");
    fs::create_dir_all(&out_dir)?;
    let out_path = out_dir.join("embed_image.docx");

    let file = File::create(&out_path)?;
    docx.build().pack(file)?;

    println!("Wrote {}", out_path.display());
    Ok(())
}
