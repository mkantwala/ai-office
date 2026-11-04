//! Read-only table drag feedback, shared with Wordcraft's native mouse commands.
use serde_json::{Value, json};
use wordcraft_engine::{
    Session,
    cmd::{
        table_mouse,
        table_resize::{self, Edge, Share},
    },
};
use wordcraft_layout::DocLayout;

pub fn preview(s: &Session, layout: &DocLayout, v: &Value) -> Result<Value, String> {
    if v["revision"].as_u64() != Some(s.rev()) {
        return Err("Document changed during the table drag".into());
    }
    let page = v["page"].as_u64().filter(|p| *p < layout.pages.len() as u64).ok_or("Invalid page")? as usize;
    let number =
        |key: &str| v[key].as_f64().filter(|n| n.is_finite() && n.abs() < 100_000.0).map(|n| n as f32).ok_or_else(|| format!("Invalid {key}"));
    let (x, y, to_x, to_y, reach) = (number("x")?, number("y")?, number("toX")?, number("toY")?, number("reach")?.clamp(0.5, 24.0));
    let rect = |x: f32, y: f32, width: f32, height: f32| json!({"x":x,"y":y,"width":width,"height":height});
    if v["command"] == "table.moveDrag" {
        let target = table_mouse::target_at(layout, &s.doc, s.sel.focus.story, page, x, y, reach).ok_or("No table move handle here")?;
        if target.spot != table_mouse::Spot::Handle || !table_mouse::movable(&target.geom) {
            return Err("This table cannot be moved".into());
        }
        let g = target.geom;
        let (left, top) = table_mouse::moved_to(layout, page, &g, to_x - x, to_y - y);
        return Ok(json!({"kind":"move","rects":[rect(left,top,g.right()-g.left(),g.bottom()-g.top())]}));
    }
    if v["command"] != "table.resizeDrag" {
        return Err("Invalid table preview command".into());
    }
    let mut g = table_resize::resize_target(layout, &s.doc, page, x, y, reach).ok_or("No table border here")?;
    g.limit_to_selection(s, y);
    let share = if v["ctrl"] == true {
        Share::Proportional
    } else if v["shift"] == true {
        Share::LeftOnly
    } else {
        Share::Neighbour
    };
    let plan = g.plan(to_x - x, to_y - y, share);
    let (kind, bounds) = match g.edge {
        Edge::Column(line) => {
            let rows: Vec<_> = g.rows.iter().filter(|r| plan.cells.is_empty() || plan.cells.iter().any(|(i, _)| *i == r.row)).collect();
            let top = rows.iter().map(|r| r.top).fold(f32::INFINITY, f32::min);
            let bottom = rows.iter().map(|r| r.bottom).fold(f32::NEG_INFINITY, f32::max);
            ("column", rect(g.planned_x(&plan, plan.cells.first().map(|(r, _)| *r), line), top, 0.0, bottom - top))
        }
        Edge::Row(row) => {
            let r = g.rows.iter().find(|r| r.row == row).ok_or("No row here")?;
            let height = plan.heights.iter().find(|(i, _)| *i == row).map_or(r.bottom - r.top, |(_, h)| *h);
            ("row", rect(g.left(), r.top + height, g.right() - g.left(), 0.0))
        }
        Edge::Corner => {
            let width = plan.widths.as_ref().map_or(g.right() - g.left(), |w| w.iter().sum());
            let height = g.rows.iter().map(|r| plan.heights.iter().find(|(i, _)| *i == r.row).map_or(r.bottom - r.top, |(_, h)| *h)).sum();
            ("corner", rect(g.left(), g.top(), width, height))
        }
    };
    Ok(json!({"kind":kind,"rects":[bounds]}))
}
