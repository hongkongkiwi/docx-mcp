use anyhow::Result;
use docx_mcp::docx_handler::{DocxHandler, ImageData};
use tempfile::TempDir;
use std::fs;
use std::path::PathBuf;
use zip::ZipArchive;

#[test]
fn test_golden_xml_links_images_numbering_header() -> Result<()> {
    let temp_dir = TempDir::new()?;
    let mut handler = DocxHandler::new_with_base_dir(temp_dir.path())?;
    let doc_id = handler.create_document()?;

    // Content: paragraph, hyperlink, image, list with levels, header page numbering
    handler.add_paragraph(&doc_id, "Intro paragraph.", None)?;
    handler.add_hyperlink(&doc_id, "OpenAI", "https://openai.com")?;

    let png_data: Vec<u8> = {
        // Small 1x1 PNG
        let mut img = ::image::RgbaImage::new(1, 1);
        img.put_pixel(0, 0, ::image::Rgba([0, 0, 0, 0]));
        let r#dyn = ::image::DynamicImage::ImageRgba8(img);
        let mut buf = Vec::new();
        r#dyn.write_to(&mut std::io::Cursor::new(&mut buf), ::image::ImageFormat::Png)?;
        buf
    };
    handler.add_image(&doc_id, ImageData { data: png_data, width: Some(10), height: Some(10), alt_text: Some("dot".into()) })?;

    handler.add_list(&doc_id, vec!["Item 1".into(), "Item 2".into()], true)?;
    handler.add_list_item(&doc_id, "Sub 2.1", 1, true)?;

    handler.set_page_numbering(&doc_id, "header", Some("Page {PAGE} of {PAGES}"))?;

    // Save DOCX to disk
    let out_path = temp_dir.path().join("golden_test.docx");
    handler.save_document(&doc_id, &out_path)?;

    // Open as zip and inspect XMLs
    let file = fs::File::open(&out_path)?;
    let mut zip = ZipArchive::new(file)?;

    // document.xml should contain hyperlink and drawing (image) and numPr (list numbering)
    {
        let mut doc_xml = zip.by_name("word/document.xml")?;
        let mut s = String::new();
        use std::io::Read as _;
        doc_xml.read_to_string(&mut s)?;
        assert!(s.contains("w:hyperlink") || s.contains(":hyperlink"), "document.xml missing hyperlink element");
        assert!(s.contains("w:drawing") || s.contains(":drawing"), "document.xml missing drawing element for image");
        assert!(s.contains("w:numPr") || s.contains(":numPr"), "document.xml missing numbering properties for list");
    }

    // numbering.xml should exist
    {
        let mut numbering = zip.by_name("word/numbering.xml")?;
        let mut s = String::new();
        use std::io::Read as _;
        numbering.read_to_string(&mut s)?;
        assert!(s.contains("w:numbering") || s.contains(":numbering"), "numbering.xml missing numbering root");
    }

    // header1.xml should contain our page numbering text template
    {
        let mut header = zip.by_name("word/header1.xml")?;
        let mut s = String::new();
        use std::io::Read as _;
        header.read_to_string(&mut s)?;
        assert!(s.contains("Page {PAGE} of {PAGES}"), "header1.xml missing page numbering text");
    }

    Ok(())
}
