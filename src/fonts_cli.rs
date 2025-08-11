use anyhow::{Context, Result};
use sha2::{Digest, Sha256};
use std::fs;
use std::io::Read;
use std::path::Path;

const FONTS_DIR: &str = "assets/fonts";

// Pin sources and expected checksums
const LIBERATION_VERSION: &str = "2.1.5";
const LIBERATION_TAR_URL: &str = "https://github.com/liberationfonts/liberation-fonts/files/7261482/liberation-fonts-ttf-2.1.5.tar.gz";
const NOTO_BASE_URL: &str = "https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSans";

const FONT_FILES: &[(&str, Option<&str>)] = &[
    ("LiberationSans-Regular.ttf", Some("76d04c18ea243f426b7de1f3ad208e927008f961dc5945e5aad352d0dfde8ee8")),
    ("LiberationSans-Bold.ttf",    Some("788abee4c806d660e8aee46689dd8540cd4bb98da03dcc9d171ce3efd99a9173")),
    ("LiberationSans-Italic.ttf",  Some("e5bae5c4cde31f22142753855f4f8fb86da6ff39955ed3c0a11248b0d16948b0")),
    ("LiberationMono-Regular.ttf", Some("f2b83c763e8afd21709333370bed4774337fae82267937e2b5aea7e2fbd922c1")),
    ("NotoSans-Regular.ttf",       Some("b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5")),
    ("NotoSans-Bold.ttf",          Some("c976e4b1b99edc88775377fcc21692ca4bfa46b6d6ca6522bfda505b28ff9d6a")),
];

pub fn download_fonts_blocking() -> Result<()> {
    fs::create_dir_all(FONTS_DIR).context("create fonts dir")?;

    // Download Liberation tarball
    let tar_bytes = download_bytes(LIBERATION_TAR_URL)?;
    extract_liberation_from_tar(&tar_bytes, Path::new(FONTS_DIR))?;

    // Download Noto fonts
    for name in ["NotoSans-Regular.ttf", "NotoSans-Bold.ttf"] {
        let url = format!("{}/{}", NOTO_BASE_URL, name);
        let bytes = download_bytes(&url)?;
        let out = Path::new(FONTS_DIR).join(name);
        fs::write(&out, bytes).context("write noto font")?;
        // verify immediate
        verify_single(&out, expected_for(name))?;
    }
    // Verify all fonts after extraction
    verify_fonts_blocking()
}

pub fn verify_fonts_blocking() -> Result<()> {
    for (name, expected_opt) in FONT_FILES {
        let path = Path::new(FONTS_DIR).join(name);
        if !path.exists() {
            anyhow::bail!("missing font: {}", name);
        }
        let actual = sha256_file(&path)?;
        if let Some(expected) = expected_opt {
            if !actual.eq_ignore_ascii_case(expected) {
                anyhow::bail!("checksum mismatch for {}: {} != {}", name, actual, expected);
            }
        }
    }
    Ok(())
}

fn download_bytes(url: &str) -> Result<Vec<u8>> {
    let mut res = ureq::get(url).call().context("request failed")?;
    let mut buf = Vec::new();
    res.into_reader().read_to_end(&mut buf).context("read body")?;
    Ok(buf)
}

fn extract_liberation_from_tar(tar_gz: &[u8], out_dir: &Path) -> Result<()> {
    let gz = flate2::read::GzDecoder::new(tar_gz);
    let mut archive = tar::Archive::new(gz);

    for entry in archive.entries().context("iter entries")? {
        let mut entry = entry.context("entry")?;
        // Extract filename into an owned String to avoid borrowing `entry`
        let filename_owned: Option<String> = {
            let path_buf = entry.path().context("entry path")?;
            path_buf
                .file_name()
                .and_then(|s| s.to_str())
                .map(|s| s.to_string())
        };
        let Some(filename) = filename_owned.as_deref() else { continue };
        match filename {
            "LiberationSans-Regular.ttf" |
            "LiberationSans-Bold.ttf" |
            "LiberationSans-Italic.ttf" |
            "LiberationMono-Regular.ttf" => {
                let dest = out_dir.join(filename);
                let context_msg = format!("unpack {}", filename);
                entry.unpack(&dest).context(context_msg)?;
                // verify immediate
                verify_single(&dest, expected_for(filename))?;
            }
            _ => {}
        }
    }

    Ok(())
}

fn expected_for(name: &str) -> Option<&'static str> {
    FONT_FILES.iter().find(|(n, _)| *n == name).and_then(|(_, s)| *s)
}

fn verify_single(path: &Path, expected: Option<&str>) -> Result<()> {
    if let Some(exp) = expected {
        let actual = sha256_file(path)?;
        if !actual.eq_ignore_ascii_case(exp) {
            anyhow::bail!(
                "checksum mismatch for {}: {} != {}",
                path.display(),
                actual,
                exp
            );
        }
    }
    Ok(())
}

fn sha256_file(path: &Path) -> Result<String> {
    let mut file = fs::File::open(path).with_context(|| format!("open {}", path.display()))?;
    let mut hasher = Sha256::new();
    let mut buf = [0u8; 8192];
    loop {
        let n = file.read(&mut buf)?;
        if n == 0 { break; }
        hasher.update(&buf[..n]);
    }
    Ok(format!("{:x}", hasher.finalize()))
}
