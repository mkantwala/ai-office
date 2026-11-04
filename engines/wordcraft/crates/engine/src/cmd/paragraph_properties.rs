//! Bounded paragraph-property patches shared by AI Office controls and block edits.
use crate::{CmdError, CmdResult, Session};
use serde_json::{Value, json};
use wordcraft_doc::Document;
use wordcraft_doc::props::{Align, Border, BorderStyle, Borders, LineSpacing, ParaProps, Rgb, TabAlign, TabLeader, TabStop};

fn object<'a>(v: &'a Value, keys: &[&str]) -> Result<&'a serde_json::Map<String, Value>, CmdError> {
    let m = v.as_object().filter(|m| !m.is_empty()).ok_or_else(|| CmdError::Params("Expected a non-empty property object".into()))?;
    if let Some(k) = m.keys().find(|k| !keys.contains(&k.as_str())) {
        return Err(CmdError::Params(format!("Unsupported paragraph property: {k}")));
    }
    Ok(m)
}
fn number(v: &Value, lo: f32, hi: f32) -> Result<f32, CmdError> {
    v.as_f64()
        .filter(|n| n.is_finite() && *n >= lo as f64 && *n <= hi as f64)
        .map(|n| n as f32)
        .ok_or_else(|| CmdError::Params(format!("Expected a number between {lo} and {hi}")))
}
fn color(v: &Value) -> Result<Rgb, CmdError> {
    v.as_str()
        .filter(|s| s.len() == 6 && s.bytes().all(|c| c.is_ascii_hexdigit()))
        .and_then(Rgb::parse)
        .ok_or_else(|| CmdError::Params("Color must be six hexadecimal digits".into()))
}

/// Concrete values for an exact paragraph, including inherited style properties.
pub fn read(doc: &Document, props: &ParaProps) -> Value {
    let r = doc.styles.resolve_para(props);
    let side = |b: Option<Border>| {
        b.filter(Border::is_visible).map(|b| json!({"width":b.width,"color":b.color.unwrap_or(Rgb::BLACK).hex(),"style":b.style}))
    };
    let borders =
        r.borders.map(|b| json!({"top":side(b.top),"left":side(b.left),"bottom":side(b.bottom),"right":side(b.right),"between":side(b.between)}));
    json!({"alignment":r.align.visual(r.bidi),"indents":{"left":r.indent_left,"right":r.indent_right,"firstLine":r.indent_first},
        "spacing":{"before":r.space_before,"after":r.space_after},"lineSpacing":r.line_spacing,
        "tabs":r.tabs,"borders":borders,"shading":r.shading.map(|c| format!("{:02X}{:02X}{:02X}",c.0,c.1,c.2)),
        "keepNext":r.keep_next,"keepLines":r.keep_lines,"pageBreakBefore":r.page_break_before,"widowControl":r.widow_control})
}

pub fn apply(s: &mut Session, v: &Value) -> CmdResult {
    if v.get("ifRevision").is_some_and(|r| r.as_u64() != Some(s.rev()))
        || v.get("ifSelection").is_some_and(|sel| serde_json::to_value(&s.sel).ok().as_ref() != Some(sel))
    {
        return Err(CmdError::Params("Document or selection changed. Reopen Paragraph Properties.".into()));
    }
    object(
        v,
        &[
            "ifRevision",
            "ifSelection",
            "alignment",
            "indents",
            "spacing",
            "lineSpacing",
            "tabs",
            "borders",
            "shading",
            "keepNext",
            "keepLines",
            "pageBreakBefore",
            "widowControl",
        ],
    )?;
    let mut patch = ParaProps::default();
    if let Some(a) = v.get("alignment") {
        patch.align = Some(match a.as_str() {
            Some("left") => Align::Left,
            Some("right") => Align::Right,
            Some("center") => Align::Center,
            Some("justify") => Align::Justify,
            Some("distribute") => Align::Distribute,
            _ => return Err(CmdError::Params("Unsupported paragraph alignment".into())),
        });
    }
    if let Some(x) = v.get("indents") {
        object(x, &["left", "right", "firstLine"])?;
        if let Some(n) = x.get("left") {
            patch.indent_left = Some(number(n, -1584., 1584.)?);
        }
        if let Some(n) = x.get("right") {
            patch.indent_right = Some(number(n, -1584., 1584.)?);
        }
        if let Some(n) = x.get("firstLine") {
            patch.indent_first = Some(number(n, -1584., 1584.)?);
        }
    }
    if let Some(x) = v.get("spacing") {
        object(x, &["before", "after"])?;
        if let Some(n) = x.get("before") {
            patch.space_before = Some(number(n, 0., 1584.)?);
            patch.space_before_auto = Some(false);
        }
        if let Some(n) = x.get("after") {
            patch.space_after = Some(number(n, 0., 1584.)?);
            patch.space_after_auto = Some(false);
        }
    }
    if let Some(x) = v.get("lineSpacing") {
        object(x, &["rule", "value"])?;
        patch.line_spacing = Some(match x["rule"].as_str() {
            Some("multiple") => LineSpacing::Multiple(number(&x["value"], 0.5, 10.)?),
            Some("exactly") => LineSpacing::Exactly(number(&x["value"], 1., 1584.)?),
            Some("atLeast") => LineSpacing::AtLeast(number(&x["value"], 1., 1584.)?),
            _ => return Err(CmdError::Params("Line spacing rule must be multiple, exactly or atLeast".into())),
        });
    }
    let mut tabs = None;
    if let Some(x) = v.get("tabs") {
        let list = x.as_array().filter(|a| a.len() <= 64).ok_or_else(|| CmdError::Params("tabs must contain at most 64 stops".into()))?;
        let mut stops = Vec::new();
        for t in list {
            object(t, &["pos", "align", "leader"])?;
            let pos = number(&t["pos"], 0., 1584.)?;
            if stops.iter().any(|s: &TabStop| (s.pos - pos).abs() < 0.01) {
                return Err(CmdError::Params("Tab positions must be distinct".into()));
            }
            let align = match t.get("align").and_then(Value::as_str).unwrap_or("left") {
                "left" => TabAlign::Left,
                "right" => TabAlign::Right,
                "center" => TabAlign::Center,
                "decimal" => TabAlign::Decimal,
                "bar" => TabAlign::Bar,
                _ => return Err(CmdError::Params("Unsupported tab alignment".into())),
            };
            let leader = match t.get("leader").and_then(Value::as_str).unwrap_or("none") {
                "none" => TabLeader::None,
                "dot" => TabLeader::Dot,
                "hyphen" => TabLeader::Hyphen,
                "underscore" => TabLeader::Underscore,
                "middleDot" => TabLeader::MiddleDot,
                _ => return Err(CmdError::Params("Unsupported tab leader".into())),
            };
            for key in ["align", "leader"] {
                if t.get(key).is_some_and(|v| !v.is_string()) {
                    return Err(CmdError::Params(format!("{key} must be a string")));
                }
            }
            stops.push(TabStop { pos, align, leader });
        }
        stops.sort_by(|a, b| a.pos.total_cmp(&b.pos));
        tabs = Some(stops);
    }
    let mut borders = Vec::new();
    if let Some(x) = v.get("borders") {
        object(x, &["top", "left", "bottom", "right", "between"])?;
        for (side, b) in x.as_object().into_iter().flatten() {
            let border = if b.is_null() {
                Some(Border { style: BorderStyle::None, width: 0., color: None, space: 0. })
            } else {
                object(b, &["width", "color", "style"])?;
                let style = match b.get("style").and_then(Value::as_str).unwrap_or("single") {
                    "single" => BorderStyle::Single,
                    "double" => BorderStyle::Double,
                    "dotted" => BorderStyle::Dotted,
                    "dashed" => BorderStyle::Dashed,
                    _ => return Err(CmdError::Params("Unsupported border style".into())),
                };
                if b.get("style").is_some_and(|v| !v.is_string()) {
                    return Err(CmdError::Params("Border style must be a string".into()));
                }
                Some(Border {
                    style,
                    width: b.get("width").map(|n| number(n, 0.25, 6.)).transpose()?.unwrap_or(0.5),
                    color: b.get("color").map(color).transpose()?,
                    space: 0.,
                })
            };
            borders.push((side.clone(), border));
        }
    }
    if let Some(x) = v.get("shading") {
        patch.shading = Some(color(x)?);
    }
    for (key, slot) in [
        ("keepNext", &mut patch.keep_next),
        ("keepLines", &mut patch.keep_lines),
        ("pageBreakBefore", &mut patch.page_break_before),
        ("widowControl", &mut patch.widow_control),
    ] {
        if let Some(x) = v.get(key) {
            *slot = Some(x.as_bool().ok_or_else(|| CmdError::Params(format!("{key} must be a boolean")))?);
        }
    }
    let styles = s.doc.styles.clone();
    super::para::fmt(s, &|p| {
        let resolved = styles.resolve_para(p);
        p.overlay(&patch);
        if patch.align.is_some() {
            p.align = patch.align.map(|a| a.visual(resolved.bidi));
        }
        p.cancel_char_indents(patch.indent_left.is_some(), patch.indent_right.is_some(), patch.indent_first.is_some());
        if let Some(stops) = &tabs {
            let inherited = styles.resolve_para(&ParaProps { tabs: None, ..p.clone() }).tabs;
            let mut result: Vec<_> = inherited
                .iter()
                .filter(|t| !stops.iter().any(|x| (x.pos - t.pos).abs() < 0.01))
                .map(|t| TabStop { pos: t.pos, align: TabAlign::Clear, leader: TabLeader::None })
                .collect();
            result.extend(stops.iter().copied());
            p.tabs = Some(result);
        }
        if !borders.is_empty() {
            let mut result = resolved.borders.unwrap_or_else(Borders::default);
            for (side, border) in &borders {
                let slot = match side.as_str() {
                    "top" => &mut result.top,
                    "left" => &mut result.left,
                    "bottom" => &mut result.bottom,
                    "right" => &mut result.right,
                    "between" => &mut result.between,
                    _ => continue,
                };
                let mut next = *border;
                if let (Some(old), Some(b)) = ((*slot).filter(Border::is_visible), next.as_mut()) {
                    let input = &v["borders"][side];
                    if !input.is_null() {
                        if input.get("width").is_none() {
                            b.width = old.width;
                        }
                        if input.get("color").is_none() {
                            b.color = old.color;
                        }
                        if input.get("style").is_none() {
                            b.style = old.style;
                        }
                        b.space = old.space;
                    }
                }
                *slot = next;
            }
            p.borders = Some(result);
        }
    })
}
