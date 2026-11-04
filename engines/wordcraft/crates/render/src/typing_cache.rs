//! Bounded, sharp page rasters for the embedded Word editor.
//!
//! Ordinary typing changes glyph draws in a small band. Reuse the rest of the
//! raster, painting that band with the same renderer and integer pixel origin.
//! Structural/complex drawings deliberately fall back to a complete repaint.
use super::*;
use std::collections::VecDeque;

#[derive(Default)]
pub struct TypingCache {
    pages: VecDeque<Entry>,
}
struct Entry {
    index: usize,
    scale: f32,
    size: (f32, f32),
    paper: Option<Rgb>,
    eligible: bool,
    draws: Vec<Draw>,
    image: Rendered,
}

fn equal(a: &Draw, b: &Draw) -> bool {
    match (a, b) {
        (
            Draw::Glyphs { face: af, size: az, glyphs: ag, color: ac, alpha: aa, synth_bold: ab, synth_italic: ai, .. },
            Draw::Glyphs { face: bf, size: bz, glyphs: bg, color: bc, alpha: ba, synth_bold: bb, synth_italic: bi, .. },
        ) => af == bf && az == bz && ag == bg && ac == bc && aa == ba && ab == bb && ai == bi,
        (Draw::Fill { .. }, Draw::Fill { .. }) | (Draw::Line { .. }, Draw::Line { .. }) => format!("{a:?}") == format!("{b:?}"),
        _ => false,
    }
}

// Reflow can change raster edge coverage across many lines. Keep the fast path
// for stationary text; structural changes get a fresh raster, like native tiles.
fn stationary(a: &Draw, b: &Draw) -> bool {
    match (a, b) {
        (Draw::Glyphs { glyphs: a, .. }, Draw::Glyphs { glyphs: b, .. }) => {
            let band = |g: &[(u32, f32, f32)]| g.iter().fold((f32::INFINITY, f32::NEG_INFINITY), |(lo, hi), (_, _, y)| (lo.min(*y), hi.max(*y)));
            band(a) == band(b)
        }
        (Draw::Fill { rect: a, .. }, Draw::Fill { rect: b, .. }) => a.y == b.y && a.h == b.h,
        (Draw::Line { y0: a0, y1: a1, .. }, Draw::Line { y0: b0, y1: b1, .. }) => a0 == b0 && a1 == b1,
        _ => false,
    }
}

/// Ink bounds, including synthetic bold's miter limit, stem darkening and AA.
/// Use actual glyph outlines, not line height: accents/superscripts can overflow it.
fn ink_band(draw: &Draw, scale: f32) -> Option<(f64, f64)> {
    let padding = 2.0 / scale as f64;
    match draw {
        Draw::Fill { rect, .. } => return Some((rect.y as f64 - padding, rect.bottom() as f64 + padding)),
        Draw::Line { y0, y1, width, .. } => {
            // Covers solid/dashed strokes, double lines and proofing waves.
            let pad = (*width as f64 * 3.0).max(2.0) + padding;
            return Some((y0.min(*y1) as f64 - pad, y0.max(*y1) as f64 + pad));
        }
        _ => {}
    }
    let Draw::Glyphs { face, size, glyphs, synth_bold, .. } = draw else {
        return None;
    };
    let k = *size as f64 / face.upem.max(1.0);
    let pad = if *synth_bold { *size as f64 * 0.06 } else { 0.0 } + 2.0 / scale as f64;
    let mut lo = f64::INFINITY;
    let mut hi = f64::NEG_INFINITY;
    for (gid, _, y) in glyphs {
        let outline = FontDb::global().outline(face, *gid);
        if outline.elements().is_empty() {
            continue;
        }
        let bounds = outline.bounding_box();
        lo = lo.min(*y as f64 + bounds.y0 * k - pad);
        hi = hi.max(*y as f64 + bounds.y1 * k + pad);
    }
    Some((lo, hi))
}

impl TypingCache {
    /// At most two pages and 32 MiB of pixel storage. Never retain whole documents.
    pub fn render(&mut self, doc: &Document, page: &Page, index: usize, scale: f32) -> &Rendered {
        self.render_with_options(doc, page, index, scale, &RenderOptions::default())
    }

    /// Reuse the same bounded raster cache when switching between body and header editing.
    pub fn render_with_options(&mut self, doc: &Document, page: &Page, index: usize, scale: f32, opts: &RenderOptions) -> &Rendered {
        let draws = page_display(doc, page, &opts.display);
        let eligible =
            doc.settings.watermark.is_none() && draws.iter().all(|d| matches!(d, Draw::Glyphs { .. } | Draw::Fill { .. } | Draw::Line { .. }));
        let old = self.pages.iter().position(|e| e.index == index).and_then(|i| self.pages.remove(i));
        let mut image = None;
        if let Some(mut old) = old {
            if eligible
                && old.eligible
                && old.scale == scale
                && old.size == (page.w, page.h)
                && old.paper == doc.settings.page_color
                && old.draws.len() == draws.len()
                && old.draws.iter().zip(&draws).all(|(a, b)| stationary(a, b))
            {
                // Match unchanged prefix/suffix even when a run or line was inserted
                // or removed. Include both old and new ink so deletions get erased.
                let prefix = old.draws.iter().zip(&draws).take_while(|(a, b)| equal(a, b)).count();
                let suffix = old.draws[prefix..].iter().rev().zip(draws[prefix..].iter().rev()).take_while(|(a, b)| equal(a, b)).count();
                let mut band = Some((f64::INFINITY, f64::NEG_INFINITY));
                for draw in old.draws[prefix..old.draws.len() - suffix].iter().chain(&draws[prefix..draws.len() - suffix]) {
                    band = band.and_then(|(lo, hi)| ink_band(draw, scale).map(|(a, b)| (lo.min(a), hi.max(b))));
                }
                if let Some((lo, hi)) = band {
                    let top = (lo * scale as f64).floor().clamp(0.0, old.image.height as f64) as u32;
                    let bottom = (hi * scale as f64).ceil().clamp(0.0, old.image.height as f64) as u32;
                    if bottom > top {
                        // Exact integer translation keeps glyph coverage identical to a full
                        // page render even at fractional zoom/device scales.
                        let view = Affine::translate((0.0, -(top as f64))) * Affine::scale(scale as f64);
                        let visible = kurbo::Rect::new(0.0, 0.0, old.image.width as f64 / scale as f64, old.image.height as f64 / scale as f64)
                            .inflate(40.0, 40.0);
                        let patch = render_draws(doc, page, old.image.width, bottom - top, view, opts, &draws, &visible);
                        let start = top as usize * old.image.width as usize * 4;
                        old.image.pixels[start..start + patch.pixels.len()].copy_from_slice(&patch.pixels);
                    }
                    image = Some(old.image);
                }
            }
        }
        let image = image.unwrap_or_else(|| render_page(doc, page, scale, opts));
        // Ineligible pages keep no draw baseline, so the next request repaints them.
        self.pages.push_front(Entry {
            index,
            scale,
            size: (page.w, page.h),
            paper: doc.settings.page_color,
            eligible,
            draws: if eligible { draws } else { Vec::new() },
            image,
        });
        while self.pages.len() > 1 && (self.pages.len() > 2 || self.pages.iter().map(|e| e.image.pixels.len()).sum::<usize>() > 32 * 1024 * 1024) {
            self.pages.pop_back();
        }
        &self.pages[0].image
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use wordcraft_layout::{LayoutCache, LayoutOptions, layout};

    #[test]
    fn cached_typing_matches_full_raster_at_fractional_and_retina_scales() {
        for scale in [1.0, 1.375, 2.0] {
            let mut cache = TypingCache::default();
            let mut layout_cache = LayoutCache::new();
            for text in [
                "Hello world\nUnchanged second paragraph",
                "Hello world!\nUnchanged second paragraph",
                "Héllo wörld!\nUnchanged second paragraph",
                "\nUnchanged second paragraph",
                "A longer sentence which will wrap across several lines. ".repeat(12).as_str(),
                "Back to one line",
            ] {
                let doc = Document::from_text(text);
                let layout = layout(&doc, &mut layout_cache, &LayoutOptions::default());
                let actual = cache.render(&doc, &layout.pages[0], 0, scale);
                let expected = render_page(&doc, &layout.pages[0], scale, &Default::default());
                assert_eq!(actual.pixels, expected.pixels, "cached pixels at scale {scale}, text {text}");
            }
        }
    }
}
