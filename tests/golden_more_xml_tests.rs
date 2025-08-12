use anyhow::Result;
use docx_mcp::docx_handler::{DocxHandler, TableData, TableMerge};
use tempfile::TempDir;
use std::fs;
use zip::ZipArchive;
use docx_mcp::docx_handler::MarginsSpec;

fn open_zip_str(path: &std::path::Path, name: &str) -> Result<String> {
    let file = fs::File::open(path)?;
    let mut zip = ZipArchive::new(file)?;
    let mut f = zip.by_name(name)?;
    let mut s = String::new();
    use std::io::Read as _;
    f.read_to_string(&mut s)?;
    Ok(s)
}

#[test]
fn test_embed_page_number_fields_into_header_xml() -> Result<()> {
    let temp_dir = TempDir::new()?;
    let mut handler = DocxHandler::new_with_base_dir(temp_dir.path())?;
    let doc_id = handler.create_document()?;

    // Add header with placeholder
    handler.set_page_numbering(&doc_id, "header", Some("Page {PAGE} of {PAGES}"))?;

    // Save once to ensure header part exists
    let out_path = temp_dir.path().join("page_fields.docx");
    handler.save_document(&doc_id, &out_path)?;

    // Embed field codes and resave to propagate to out_path
    handler.embed_page_number_fields(&doc_id)?;
    handler.save_document(&doc_id, &out_path)?;

    // Verify header XML has field runs
    let header_xml = open_zip_str(&out_path, "word/header1.xml")?;
    assert!(header_xml.contains("w:fldChar") && header_xml.contains("PAGE") && header_xml.contains("NUMPAGES"),
            "Expected PAGE/NUMPAGES fields in header1.xml, got: {}", header_xml);
    Ok(())
}

#[test]
fn test_section_break_emits_page_break() -> Result<()> {
    let temp_dir = TempDir::new()?;
    let mut handler = DocxHandler::new_with_base_dir(temp_dir.path())?;
    let doc_id = handler.create_document()?;

    handler.add_paragraph(&doc_id, "Before section", None)?;
    handler.add_section_break(&doc_id, Some("A4"), Some("portrait"), None)?;
    handler.add_paragraph(&doc_id, "After section", None)?;

    let out_path = temp_dir.path().join("section_break.docx");
    handler.save_document(&doc_id, &out_path)?;

    // Best-effort placeholder: expect a page break in document.xml
    let doc_xml = open_zip_str(&out_path, "word/document.xml")?;
    assert!(doc_xml.contains("w:br") && doc_xml.contains("w:type=\"page\""),
            "Expected a page break to denote section break");
    Ok(())
}

#[test]
fn test_table_merge_best_effort_xml() -> Result<()> {
    let temp_dir = TempDir::new()?;
    let mut handler = DocxHandler::new_with_base_dir(temp_dir.path())?;
    let doc_id = handler.create_document()?;

    // 2x2 table where first row cells are merged (2 columns)
    let table = TableData {
        rows: vec![
            vec!["TopLeft".into(), "RightMergedShouldBeEmpty".into()],
            vec!["BottomLeft".into(), "BottomRight".into()],
        ],
        headers: None,
        border_style: Some("single".into()),
        col_widths: None,
        merges: Some(vec![TableMerge { row: 0, col: 0, row_span: 1, col_span: 2 }]),
        cell_shading: None,
    };

    handler.add_table(&doc_id, table)?;
    let out_path = temp_dir.path().join("table_merge.docx");
    handler.save_document(&doc_id, &out_path)?;

    let doc_xml = open_zip_str(&out_path, "word/document.xml")?;
    // Expect TopLeft to be present once, and RightMergedShouldBeEmpty to be absent
    assert!(doc_xml.contains("TopLeft"));
    assert!(!doc_xml.contains("RightMergedShouldBeEmpty"));

    // When hi-fidelity-tables is enabled, verify gridSpan
    #[cfg(feature = "hi-fidelity-tables")]
    {
        assert!(doc_xml.contains("w:gridSpan"), "Expected w:gridSpan for horizontal merge");
        // For row_span in this test it's 1, so no vMerge expected
        assert!(!doc_xml.contains("w:vMerge w:val=\"restart\""));
    }
    Ok(())
}

#[test]
fn test_table_vmerge_and_col_widths_injection() -> Result<()> {
    let temp_dir = TempDir::new()?;
    let mut handler = DocxHandler::new_with_base_dir(temp_dir.path())?;
    let doc_id = handler.create_document()?;

    // 3x2 table with a vertical merge on first column (2 rows) and column widths
    let table = TableData {
        rows: vec![
            vec!["A".into(), "B".into()],
            vec!["A2-should-be-empty".into(), "C".into()],
            vec!["D".into(), "E".into()],
        ],
        headers: None,
        border_style: None,
        col_widths: Some(vec![2400, 3600]),
        merges: Some(vec![TableMerge { row: 0, col: 0, row_span: 2, col_span: 1 }]),
        cell_shading: None,
    };

    handler.add_table(&doc_id, table)?;
    let out_path = temp_dir.path().join("table_vmerge.docx");
    handler.save_document(&doc_id, &out_path)?;

    let doc_xml = open_zip_str(&out_path, "word/document.xml")?;
    assert!(!doc_xml.contains("A2-should-be-empty"));

    #[cfg(feature = "hi-fidelity-tables")]
    {
        // Expect vMerge restart and continue
        assert!(doc_xml.contains("<w:vMerge w:val=\"restart\"/>"));
        assert!(doc_xml.contains("<w:vMerge w:val=\"continue\"/>"));

        // Expect tblGrid with specified widths
        assert!(doc_xml.contains("<w:tblGrid>"));
        assert!(doc_xml.contains("<w:gridCol w:w=\"2400\"/>") && doc_xml.contains("<w:gridCol w:w=\"3600\"/>"));
    }

    Ok(())
}

#[test]
fn test_footer_field_embedding() -> Result<()> {
    let temp_dir = TempDir::new()?;
    let mut handler = DocxHandler::new_with_base_dir(temp_dir.path())?;
    let doc_id = handler.create_document()?;
    handler.set_page_numbering(&doc_id, "footer", Some("Page {PAGE} of {PAGES}"))?;
    let out_path = temp_dir.path().join("footer_fields.docx");
    handler.save_document(&doc_id, &out_path)?;
    handler.embed_page_number_fields(&doc_id)?;
    handler.save_document(&doc_id, &out_path)?;
    let footer_xml = open_zip_str(&out_path, "word/footer1.xml")?;
    assert!(footer_xml.contains("w:fldChar") && footer_xml.contains("NUMPAGES"));
    Ok(())
}

#[test]
fn test_styles_and_lists_and_sections_hifi_xml() -> Result<()> {
    let temp_dir = TempDir::new()?;
    let mut handler = DocxHandler::new_with_base_dir(temp_dir.path())?;
    let doc_id = handler.create_document()?;

    // Table with header row to trigger TableHeader style usage
    let table = TableData {
        rows: vec![
            vec!["H1".into(), "H2".into()],
            vec!["x".into(), "y".into()],
        ],
        headers: Some(vec!["H1".into(), "H2".into()]),
        border_style: None,
        col_widths: Some(vec![3000, 3000]),
        merges: None,
        cell_shading: None,
    };
    handler.add_table(&doc_id, table)?;

    // Ordered and unordered lists
    handler.add_list(&doc_id, vec!["one".into(), "two".into()], true)?;
    handler.add_list(&doc_id, vec!["dot".into(), "dash".into()], false)?;

    // Section setup
    handler.add_section_break(&doc_id, Some("Letter"), Some("landscape"), Some(MarginsSpec { top: Some(1.25), bottom: Some(1.25), left: Some(1.0), right: Some(1.0) }))?;

    let out_path = temp_dir.path().join("hifi_bundle.docx");
    handler.save_document(&doc_id, &out_path)?;

    #[cfg(feature = "hi-fidelity-styles")]
    {
        let styles_xml = open_zip_str(&out_path, "word/styles.xml")?;
        assert!(styles_xml.contains("w:styleId=\"TableHeader\""), "Expected TableHeader style defined");
    }
    #[cfg(feature = "hi-fidelity-lists")]
    {
        let numbering_xml = open_zip_str(&out_path, "word/numbering.xml")?;
        assert!(numbering_xml.contains("w:abstractNumId=\"10\""));
        assert!(numbering_xml.contains("w:abstractNumId=\"20\""));
    }
    #[cfg(feature = "hi-fidelity-sections")]
    {
        let doc_xml = open_zip_str(&out_path, "word/document.xml")?;
        assert!(doc_xml.contains("w:sectPr"));
        assert!(doc_xml.contains("w:orient=\"landscape\""));
        assert!(doc_xml.contains("w:pgMar"));
    }

    Ok(())
}

#[test]
fn test_insert_toc_and_bookmark_placeholders() -> Result<()> {
    let temp_dir = TempDir::new()?;
    let mut handler = DocxHandler::new_with_base_dir(temp_dir.path())?;
    let doc_id = handler.create_document()?;

    handler.add_heading(&doc_id, "Intro", 1)?;
    handler.insert_bookmark_after_heading(&doc_id, "Intro", "bm-intro")?;
    handler.insert_toc(&doc_id, 1, 3, true)?;

    let out_path = temp_dir.path().join("toc_bm.docx");
    handler.save_document(&doc_id, &out_path)?;

    let doc_xml = open_zip_str(&out_path, "word/document.xml")?;
    assert!(doc_xml.contains("__TOC__") || cfg!(feature = "hi-fidelity-toc"), "Expect TOC placeholder or transformed field");

    #[cfg(feature = "hi-fidelity-toc")]
    {
        let doc_xml = open_zip_str(&out_path, "word/document.xml")?;
        assert!(doc_xml.contains("w:fldChar") && doc_xml.contains("TOC"));
    }

    #[cfg(feature = "hi-fidelity-bookmarks")]
    {
        let doc_xml = open_zip_str(&out_path, "word/document.xml")?;
        assert!(!doc_xml.contains("__BOOKMARK__"));
    }

    Ok(())
}
