//! Rasterize our generated PDF bytes. Preview never reflows or reshapes text.
use hayro::hayro_interpret::InterpreterSettings;
use hayro::hayro_syntax::Pdf;
use hayro::{RenderCache, RenderSettings};
use std::sync::{
    atomic::{AtomicBool, Ordering},
    Arc,
};

pub struct PreviewPage {
    pub width: u32,
    pub height: u32,
    /// Opaque RGBA pixels on white paper. UI adapters may swap channels.
    pub rgba: Vec<u8>,
}

/// Only used with VisiGrid-generated PDFs. One page at a time, bounded to
/// 2400 pixels per axis (~23 MB maximum), without external programs or fonts.
pub fn rasterize<T: AsRef<[u8]> + Send + Sync + 'static>(
    data: Arc<T>,
    page_index: usize,
    long_edge: u16,
) -> Result<PreviewPage, String> {
    if !(64..=2400).contains(&long_edge) {
        return Err("Preview resolution must be between 64 and 2400 pixels.".into());
    }
    std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        let pdf = Pdf::new(data).map_err(|_| "Could not read the generated PDF for preview.")?;
        let page = pdf.pages().get(page_index).ok_or("Preview page is out of range.")?;
        let (width, height) = page.render_dimensions();
        if !width.is_finite() || !height.is_finite() || width <= 0.0 || height <= 0.0 {
            return Err("Invalid PDF page dimensions.".to_string());
        }
        let scale = f32::from(long_edge) / width.max(height);
        let warned = Arc::new(AtomicBool::new(false));
        let warning_flag = warned.clone();
        let settings = InterpreterSettings {
            warning_sink: Arc::new(move |_| { warning_flag.store(true, Ordering::Relaxed); }),
            ..Default::default()
        };
        let pixmap = hayro::render(page, &RenderCache::new(), &settings, &RenderSettings {
            x_scale: scale, y_scale: scale,
            bg_color: hayro::vello_cpu::color::palette::css::WHITE,
            ..Default::default()
        });
        if warned.load(Ordering::Relaxed) {
            return Err("The preview renderer could not display every PDF element. Export PDF is still available.".into());
        }
        Ok(PreviewPage {
            width: u32::from(pixmap.width()), height: u32::from(pixmap.height()),
            rgba: pixmap.data_as_u8_slice().to_vec(),
        })
    })).unwrap_or_else(|_| Err("Could not render this PDF preview. Export PDF is still available.".into()))
}
