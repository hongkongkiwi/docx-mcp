use once_cell::sync::Lazy;

// Conditionally embed fonts if they exist
// If fonts don't exist, we'll use empty placeholders and rely on PDF built-in fonts

#[cfg(all(feature = "embedded-fonts", not(debug_assertions)))]
pub static LIBERATION_SANS_REGULAR: &[u8] = include_bytes!("../assets/fonts/LiberationSans-Regular.ttf");
#[cfg(not(all(feature = "embedded-fonts", not(debug_assertions))))]
pub static LIBERATION_SANS_REGULAR: &[u8] = &[];

#[cfg(all(feature = "embedded-fonts", not(debug_assertions)))]
pub static LIBERATION_SANS_BOLD: &[u8] = include_bytes!("../assets/fonts/LiberationSans-Bold.ttf");
#[cfg(not(all(feature = "embedded-fonts", not(debug_assertions))))]
pub static LIBERATION_SANS_BOLD: &[u8] = &[];

#[cfg(all(feature = "embedded-fonts", not(debug_assertions)))]
pub static LIBERATION_SANS_ITALIC: &[u8] = include_bytes!("../assets/fonts/LiberationSans-Italic.ttf");
#[cfg(not(all(feature = "embedded-fonts", not(debug_assertions))))]
pub static LIBERATION_SANS_ITALIC: &[u8] = &[];

#[cfg(all(feature = "embedded-fonts", not(debug_assertions)))]
pub static LIBERATION_MONO_REGULAR: &[u8] = include_bytes!("../assets/fonts/LiberationMono-Regular.ttf");
#[cfg(not(all(feature = "embedded-fonts", not(debug_assertions))))]
pub static LIBERATION_MONO_REGULAR: &[u8] = &[];

#[cfg(all(feature = "embedded-fonts", not(debug_assertions)))]
pub const EMBEDDED_FONT_REGULAR: &[u8] = include_bytes!("../assets/fonts/NotoSans-Regular.ttf");
#[cfg(not(all(feature = "embedded-fonts", not(debug_assertions))))]
pub const EMBEDDED_FONT_REGULAR: &[u8] = &[];

#[cfg(all(feature = "embedded-fonts", not(debug_assertions)))]
pub const EMBEDDED_FONT_BOLD: &[u8] = include_bytes!("../assets/fonts/NotoSans-Bold.ttf");
#[cfg(not(all(feature = "embedded-fonts", not(debug_assertions))))]
pub const EMBEDDED_FONT_BOLD: &[u8] = &[];

pub struct EmbeddedFonts {
    pub regular: &'static [u8],
    pub bold: &'static [u8],
    pub italic: &'static [u8],
    pub mono: &'static [u8],
}

pub static FONTS: Lazy<EmbeddedFonts> = Lazy::new(|| {
    EmbeddedFonts {
        regular: LIBERATION_SANS_REGULAR,
        bold: LIBERATION_SANS_BOLD,
        italic: LIBERATION_SANS_ITALIC,
        mono: LIBERATION_MONO_REGULAR,
    }
});