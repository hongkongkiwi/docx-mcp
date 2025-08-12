use anyhow::{Context, Result};
use ::image::{ImageFormat};
use printpdf::*;
use dotext::MsDoc;
use ::lopdf::{dictionary, Object, ObjectId, Document as LoDocument};
use std::fs::{self, File};
use std::io::{BufWriter, Read};
use std::path::{Path, PathBuf};
use std::process::Command;
use tempfile::NamedTempFile;
use tracing::{debug, info};

use crate::pure_converter::PureRustConverter;

pub struct DocumentConverter {
    pure_converter: PureRustConverter,
    prefer_external_tools: bool,
}

impl DocumentConverter {
    pub fn new() -> Self {
        Self {
            pure_converter: PureRustConverter::new(),
            prefer_external_tools: cfg!(feature = "hi-fidelity"), // Prefer external/hi-fi if feature enabled
        }
    }

    pub fn docx_to_pdf(&self, docx_path: &Path, pdf_path: &Path) -> Result<()> {
        if self.prefer_external_tools {
            // Try external tools first if preferred
            // Method 1: Try LibreOffice if available
            if self.try_libreoffice_conversion(docx_path, pdf_path).is_ok() {
                info!("Successfully converted DOCX to PDF using LibreOffice");
                return Ok(());
            }
            
            // Method 2: Try unoconv if available
            if self.try_unoconv_conversion(docx_path, pdf_path).is_ok() {
                info!("Successfully converted DOCX to PDF using unoconv");
                return Ok(());
            }
        }
        
        // Use pure Rust implementation (default)
        self.pure_converter.docx_to_pdf_pure(docx_path, pdf_path)?;
        info!("Successfully converted DOCX to PDF using pure Rust implementation");
        Ok(())
    }

    /// Convert with explicit preference overriding internal default
    pub fn docx_to_pdf_with_preference(&self, docx_path: &Path, pdf_path: &Path, prefer_external: bool) -> Result<()> {
        if prefer_external {
            if self.try_libreoffice_conversion(docx_path, pdf_path).is_ok() {
                info!("Successfully converted DOCX to PDF using LibreOffice (explicit preference)");
                return Ok(());
            }
            if self.try_unoconv_conversion(docx_path, pdf_path).is_ok() {
                info!("Successfully converted DOCX to PDF using unoconv (explicit preference)");
                return Ok(());
            }
        }
        // Fallback to pure implementation
        self.pure_converter.docx_to_pdf_pure(docx_path, pdf_path)?;
        info!("Successfully converted DOCX to PDF using pure Rust implementation (explicit preference)");
        Ok(())
    }

    fn try_libreoffice_conversion(&self, docx_path: &Path, pdf_path: &Path) -> Result<()> {
        let output = Command::new("libreoffice")
            .args(&[
                "--headless",
                "--invisible",
                "--nodefault",
                "--nolockcheck",
                "--nologo",
                "--norestore",
                "--convert-to",
                "pdf",
                "--outdir",
                pdf_path.parent().unwrap().to_str().unwrap(),
                docx_path.to_str().unwrap(),
            ])
            .output();
        
        match output {
            Ok(output) if output.status.success() => {
                // LibreOffice creates the PDF with the same base name
                let temp_pdf = pdf_path.parent().unwrap()
                    .join(docx_path.file_stem().unwrap())
                    .with_extension("pdf");
                
                if temp_pdf != pdf_path {
                    fs::rename(&temp_pdf, pdf_path)?;
                }
                Ok(())
            }
            Ok(output) => {
                let stderr = String::from_utf8_lossy(&output.stderr);
                anyhow::bail!("LibreOffice conversion failed: {}", stderr)
            }
            Err(e) => {
                debug!("LibreOffice not available: {}", e);
                anyhow::bail!("LibreOffice not available")
            }
        }
    }

    fn try_unoconv_conversion(&self, docx_path: &Path, pdf_path: &Path) -> Result<()> {
        let output = Command::new("unoconv")
            .args(&[
                "-f", "pdf",
                "-o", pdf_path.to_str().unwrap(),
                docx_path.to_str().unwrap(),
            ])
            .output();
        
        match output {
            Ok(output) if output.status.success() => Ok(()),
            Ok(output) => {
                let stderr = String::from_utf8_lossy(&output.stderr);
                anyhow::bail!("unoconv conversion failed: {}", stderr)
            }
            Err(e) => {
                debug!("unoconv not available: {}", e);
                anyhow::bail!("unoconv not available")
            }
        }
    }

    fn basic_docx_to_pdf(&self, docx_path: &Path, pdf_path: &Path) -> Result<()> {
        // Extract text from DOCX (fallback using dotext)
        let mut reader = dotext::Docx::open(docx_path)
            .with_context(|| format!("Failed to open DOCX {:?}", docx_path))?;
        let mut data = String::new();
        use std::io::Read as _;
        reader.read_to_string(&mut data)?;
        let text = data;
        
        // Create a basic PDF with the extracted text
        let (doc, page1, layer1) = PdfDocument::new("Document", Mm(210.0), Mm(297.0), "Layer 1");
        let _current_layer = doc.get_page(page1).get_layer(layer1);
        
        // Load a basic font
        let font = doc.add_builtin_font(BuiltinFont::Helvetica)?;
        
        // Split text into lines and add to PDF
        let lines: Vec<&str> = text.lines().collect();
        let mut y_position = Mm(280.0);
        let line_height = Mm(5.0);
        
        let mut current_layer = doc.get_page(page1).get_layer(layer1);
        for line in lines {
            if y_position < Mm(20.0) {
                let (page, layer) = doc.add_page(Mm(210.0), Mm(297.0), "Page layer");
                current_layer = doc.get_page(page).get_layer(layer);
                y_position = Mm(280.0);
            }
            current_layer.use_text(line, 12.0, Mm(10.0), y_position, &font);
            y_position -= line_height;
        }
        
        doc.save(&mut BufWriter::new(File::create(pdf_path)?))?;
        Ok(())
    }

    pub fn pdf_to_images(
        &self,
        pdf_path: &Path,
        output_dir: &Path,
        format: ImageFormat,
        dpi: u32,
    ) -> Result<Vec<PathBuf>> {
        // Try multiple methods for PDF to image conversion
        
        // Method 1: Try pdftoppm if available
        if let Ok(images) = self.try_pdftoppm_conversion(pdf_path, output_dir, format, dpi) {
            info!("Successfully converted PDF to images using pdftoppm");
            return Ok(images);
        }
        
        // Method 2: Try ImageMagick if available
        if let Ok(images) = self.try_imagemagick_conversion(pdf_path, output_dir, format, dpi) {
            info!("Successfully converted PDF to images using ImageMagick");
            return Ok(images);
        }
        
        // Method 3: Try Ghostscript if available
        if let Ok(images) = self.try_ghostscript_conversion(pdf_path, output_dir, format, dpi) {
            info!("Successfully converted PDF to images using Ghostscript");
            return Ok(images);
        }
        
        anyhow::bail!("No PDF to image converter available. Please install pdftoppm, ImageMagick, or Ghostscript")
    }

    fn try_pdftoppm_conversion(
        &self,
        pdf_path: &Path,
        output_dir: &Path,
        format: ImageFormat,
        dpi: u32,
    ) -> Result<Vec<PathBuf>> {
        fs::create_dir_all(output_dir)?;
        
        let output_prefix = output_dir.join("page");
        let format_arg = match format {
            ImageFormat::Png => "-png",
            ImageFormat::Jpeg => "-jpeg",
            _ => "-png",
        };
        
        let output = Command::new("pdftoppm")
            .args(&[
                format_arg,
                "-r", &dpi.to_string(),
                pdf_path.to_str().unwrap(),
                output_prefix.to_str().unwrap(),
            ])
            .output()?;
        
        if !output.status.success() {
            let stderr = String::from_utf8_lossy(&output.stderr);
            anyhow::bail!("pdftoppm failed: {}", stderr);
        }
        
        // Collect generated image files
        let extension = match format {
            ImageFormat::Png => "png",
            ImageFormat::Jpeg => "jpg",
            _ => "png",
        };
        
        let mut images = Vec::new();
        for entry in fs::read_dir(output_dir)? {
            let entry = entry?;
            let path = entry.path();
            if path.extension() == Some(std::ffi::OsStr::new(extension)) {
                images.push(path);
            }
        }
        
        images.sort();
        Ok(images)
    }

    fn try_imagemagick_conversion(
        &self,
        pdf_path: &Path,
        output_dir: &Path,
        format: ImageFormat,
        dpi: u32,
    ) -> Result<Vec<PathBuf>> {
        fs::create_dir_all(output_dir)?;
        
        let extension = match format {
            ImageFormat::Png => "png",
            ImageFormat::Jpeg => "jpg",
            _ => "png",
        };
        
        let output_pattern = output_dir.join(format!("page-%03d.{}", extension));
        
        let output = Command::new("convert")
            .args(&[
                "-density", &dpi.to_string(),
                pdf_path.to_str().unwrap(),
                "-quality", "100",
                output_pattern.to_str().unwrap(),
            ])
            .output()?;
        
        if !output.status.success() {
            let stderr = String::from_utf8_lossy(&output.stderr);
            anyhow::bail!("ImageMagick convert failed: {}", stderr);
        }
        
        // Collect generated image files
        let mut images = Vec::new();
        for entry in fs::read_dir(output_dir)? {
            let entry = entry?;
            let path = entry.path();
            if path.extension() == Some(std::ffi::OsStr::new(extension)) {
                images.push(path);
            }
        }
        
        images.sort();
        Ok(images)
    }

    fn try_ghostscript_conversion(
        &self,
        pdf_path: &Path,
        output_dir: &Path,
        format: ImageFormat,
        dpi: u32,
    ) -> Result<Vec<PathBuf>> {
        fs::create_dir_all(output_dir)?;
        
        let device = match format {
            ImageFormat::Png => "png16m",
            ImageFormat::Jpeg => "jpeg",
            _ => "png16m",
        };
        
        let extension = match format {
            ImageFormat::Png => "png",
            ImageFormat::Jpeg => "jpg",
            _ => "png",
        };
        
        let output_pattern = output_dir.join(format!("page-%03d.{}", extension));
        
        let output = Command::new("gs")
            .args(&[
                "-dNOPAUSE",
                "-dBATCH",
                "-sDEVICE", device,
                &format!("-r{}", dpi),
                "-dTextAlphaBits=4",
                "-dGraphicsAlphaBits=4",
                &format!("-sOutputFile={}", output_pattern.to_str().unwrap()),
                pdf_path.to_str().unwrap(),
            ])
            .output()?;
        
        if !output.status.success() {
            let stderr = String::from_utf8_lossy(&output.stderr);
            anyhow::bail!("Ghostscript failed: {}", stderr);
        }
        
        // Collect generated image files
        let mut images = Vec::new();
        for entry in fs::read_dir(output_dir)? {
            let entry = entry?;
            let path = entry.path();
            if path.extension() == Some(std::ffi::OsStr::new(extension)) {
                images.push(path);
            }
        }
        
        images.sort();
        Ok(images)
    }

    pub fn docx_to_images(
        &self,
        docx_path: &Path,
        output_dir: &Path,
        format: ImageFormat,
        dpi: u32,
    ) -> Result<Vec<PathBuf>> {
        // First convert DOCX to PDF
        let temp_pdf = NamedTempFile::new()?.into_temp_path();
        self.docx_to_pdf(docx_path, &temp_pdf)?;
        
        // Then convert PDF to images
        let images = self.pdf_to_images(&temp_pdf, output_dir, format, dpi)?;
        
        Ok(images)
    }

    pub fn docx_to_images_with_preference(
        &self,
        docx_path: &Path,
        output_dir: &Path,
        format: ImageFormat,
        dpi: u32,
        prefer_external: bool,
    ) -> Result<Vec<PathBuf>> {
        let temp_pdf = NamedTempFile::new()?.into_temp_path();
        self.docx_to_pdf_with_preference(docx_path, &temp_pdf, prefer_external)?;
        let images = self.pdf_to_images(&temp_pdf, output_dir, format, dpi)?;
        Ok(images)
    }

    pub fn create_thumbnail(
        &self,
        image_path: &Path,
        output_path: &Path,
        width: u32,
        height: u32,
    ) -> Result<()> {
        let img = ::image::open(image_path)
            .with_context(|| format!("Failed to open image {:?}", image_path))?;
        
        let thumbnail = img.thumbnail(width, height);
        thumbnail.save(output_path)
            .with_context(|| format!("Failed to save thumbnail to {:?}", output_path))?;
        
        info!("Created thumbnail {}x{} at {:?}", width, height, output_path);
        Ok(())
    }

    pub fn merge_pdfs(&self, pdf_paths: &[PathBuf], output_path: &Path) -> Result<()> {
        // Try using pdftk if available
        if self.try_pdftk_merge(pdf_paths, output_path).is_ok() {
            info!("Successfully merged PDFs using pdftk");
            return Ok(());
        }
        
        // Fallback to lopdf for merging
        self.merge_pdfs_with_lopdf(pdf_paths, output_path)?;
        info!("Successfully merged PDFs using lopdf");
        Ok(())
    }

    fn try_pdftk_merge(&self, pdf_paths: &[PathBuf], output_path: &Path) -> Result<()> {
        let mut args = Vec::new();
        for path in pdf_paths {
            args.push(path.to_str().unwrap());
        }
        args.push("cat");
        args.push("output");
        args.push(output_path.to_str().unwrap());
        
        let output = Command::new("pdftk")
            .args(&args)
            .output()?;
        
        if !output.status.success() {
            let stderr = String::from_utf8_lossy(&output.stderr);
            anyhow::bail!("pdftk merge failed: {}", stderr);
        }
        
        Ok(())
    }

    fn merge_pdfs_with_lopdf(&self, pdf_paths: &[PathBuf], output_path: &Path) -> Result<()> {
        let mut merged = LoDocument::new();
        merged.version = "1.5".to_string();
        
        for pdf_path in pdf_paths {
            let mut doc = LoDocument::load(pdf_path)?;
            
            // Merge pages
            for page_id in doc.get_pages().values() {
                merged.add_object(doc.get_object(*page_id)?.clone());
            }
        }
        
        merged.save(output_path)?;
        Ok(())
    }

    pub fn split_pdf(&self, pdf_path: &Path, output_dir: &Path) -> Result<Vec<PathBuf>> {
        fs::create_dir_all(output_dir)?;
        
        let doc = LoDocument::load(pdf_path)?;
        let pages = doc.get_pages();
        let mut output_paths = Vec::new();
        
        for (i, (_, page_id)) in pages.iter().enumerate() {
            let mut single_page = LoDocument::new();
            single_page.version = doc.version.clone();
            
            // Clone the page to the new document
            single_page.add_object(doc.get_object(*page_id)?.clone());
            
            let output_path = output_dir.join(format!("page_{:03}.pdf", i + 1));
            single_page.save(&output_path)?;
            output_paths.push(output_path);
        }
        
        info!("Split PDF into {} pages", output_paths.len());
        Ok(output_paths)
    }
}