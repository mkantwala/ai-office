//! Guarded page property patches over native section settings. Omitted values stay intact.
use crate::{CmdError, CmdResult, Session};
use serde_json::{Value, json};
use wordcraft_doc::props::{Border, BorderStyle, Borders, Rgb};
use wordcraft_doc::section::SectionProps;

fn object<'a>(v: &'a Value, keys: &[&str]) -> Result<&'a serde_json::Map<String, Value>, CmdError> {
    let value = v.as_object().ok_or_else(|| CmdError::Params("Expected a page property object".into()))?;
    if let Some(key) = value.keys().find(|key| !keys.contains(&key.as_str())) {
        return Err(CmdError::Params(format!("Unsupported page property: {key}")));
    }
    Ok(value)
}
fn number(v: &Value, lo: f32, hi: f32) -> Result<f32, CmdError> {
    v.as_f64().filter(|n| n.is_finite() && *n >= lo as f64 && *n <= hi as f64).map(|n| n as f32)
        .ok_or_else(|| CmdError::Params(format!("Expected a number between {lo} and {hi}")))
}

pub fn read(section: &SectionProps) -> Value {
    let side = |b: Option<Border>| b.filter(Border::is_visible).map(|b| json!({"style":b.style,"width":b.width,"color":b.color.map(Rgb::hex),"space":b.space}));
    let borders = section.page_borders.unwrap_or_default();
    json!({"width":section.page_w,"height":section.page_h,
        "margins":{"top":section.margin_top,"bottom":section.margin_bottom,"left":section.margin_left,"right":section.margin_right,"gutter":section.gutter},
        "borders":{"top":side(borders.top),"bottom":side(borders.bottom),"left":side(borders.left),"right":side(borders.right)}})
}

pub fn apply(s: &mut Session, v: &Value) -> CmdResult {
    object(v, &["ifRevision","ifSelection","page","applyTo","margins","borders"])?;
    if v.get("ifRevision").is_some_and(|r| r.as_u64() != Some(s.rev()))
        || v.get("ifSelection").is_some_and(|sel| serde_json::to_value(&s.sel).ok().as_ref() != Some(sel)) {
        return Err(CmdError::Params("Document or selection changed. Reopen Page Settings.".into()));
    }
    if v.get("margins").is_none() && v.get("borders").is_none() { return Err(CmdError::Params("A margins or borders patch is required".into())); }
    let layout = s.layout();
    let page = match v.get("page") {
        Some(page) => page.as_u64().and_then(|p| usize::try_from(p).ok()).ok_or_else(|| CmdError::Params("page must be a non-negative integer".into()))?,
        None => layout.caret_on(&s.sel.focus, s.page_hint).map(|c| c.page).unwrap_or(s.page_hint),
    };
    let section = layout.pages.get(page).ok_or_else(|| CmdError::Params("no such page".into()))?.section;
    let sections = s.doc.sections();
    let targets: Vec<_> = match v.get("applyTo").map_or(Some("section"), Value::as_str) {
        Some("section") => sections.get(section).into_iter().map(|(end, props)| (*end, (*props).clone())).collect(),
        Some("document") => sections.iter().map(|(end, props)| (*end, (*props).clone())).collect(),
        _ => return Err(CmdError::Params("applyTo must be section or document".into())),
    };
    let mut updated = Vec::new();
    for (end, mut props) in targets {
        if let Some(margins) = v.get("margins") {
            object(margins, &["top","bottom","left","right","gutter"])?;
            for (name, value) in [("top", &mut props.margin_top),("bottom", &mut props.margin_bottom),("left", &mut props.margin_left),("right", &mut props.margin_right),("gutter", &mut props.gutter)] {
                if let Some(n) = margins.get(name) { *value = number(n, 0.0, 1584.0)?; }
            }
            // Match upstream's margin fit rule; validate every selected section before writing any.
            if props.margin_left + props.margin_right + props.gutter > props.page_w - 36.0
                || props.margin_top + props.margin_bottom > props.page_h - 36.0 {
                return Err(CmdError::Params("Margins leave no room for text".into()));
            }
        }
        if let Some(patch) = v.get("borders") {
            if patch.is_null() { props.page_borders = None; }
            else {
                object(patch, &["top","bottom","left","right"])?;
                let mut borders = props.page_borders.unwrap_or_default();
                for (name, target) in [("top", &mut borders.top),("bottom", &mut borders.bottom),("left", &mut borders.left),("right", &mut borders.right)] {
                    if let Some(value) = patch.get(name) {
                        if value.is_null() { *target = None; continue; }
                        object(value, &["style","width","color","space"])?;
                        let mut border = target.unwrap_or(Border { style: BorderStyle::Single, width: 1.0, color: None, space: 24.0 });
                        if let Some(style) = value.get("style") {
                            border.style = match style.as_str() {
                                Some("single") => BorderStyle::Single, Some("double") => BorderStyle::Double,
                                Some("dotted") => BorderStyle::Dotted, Some("dashed") => BorderStyle::Dashed,
                                Some("thick") => BorderStyle::Thick, Some("triple") => BorderStyle::Triple,
                                Some("dotDash") => BorderStyle::DotDash, Some("wave") => BorderStyle::Wave,
                                _ => return Err(CmdError::Params("Unsupported page border style".into())),
                            };
                        }
                        if let Some(n) = value.get("width") { border.width = number(n, 0.25, 6.0)?; }
                        if let Some(n) = value.get("space") { border.space = number(n, 0.0, 31.0)?; }
                        if let Some(color) = value.get("color") {
                            border.color = if color.is_null() { None } else { Some(color.as_str().filter(|s| s.len() == 6 && s.bytes().all(|c| c.is_ascii_hexdigit())).and_then(Rgb::parse).ok_or_else(|| CmdError::Params("Color must be six hexadecimal digits".into()))?) };
                        }
                        *target = Some(border);
                    }
                }
                props.page_borders = Some(borders).filter(Borders::any_visible);
            }
        }
        updated.push((end, props));
    }
    for (end, props) in updated { *s.doc.section_mut(end) = props; }
    super::sel_result(s)
}
