//! Content-only Word pagination and rendering for AI Office.
#![cfg(target_arch = "wasm32")]
#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic, clippy::unimplemented, clippy::todo, clippy::unreachable)]

use serde_json::{Value, json};
use std::sync::Arc;
use wasm_bindgen::prelude::*;
use wordcraft_engine::Session;
use wordcraft_layout::DocLayout;
mod agent;
mod review;
mod table;
mod images;

#[wasm_bindgen]
pub struct OfficeDocument {
    session: Session,
    layout: Arc<DocLayout>,
    raster: wordcraft_render::typing_cache::TypingCache,
    width: u32,
    height: u32,
    password: Option<String>,
    id_prefix: String,
    id_counter: u64,
    review: review::Review,
}

#[wasm_bindgen]
impl OfficeDocument {
    #[wasm_bindgen(constructor)]
    pub fn new(name: &str, bytes: &[u8], password: Option<String>) -> Result<OfficeDocument, JsValue> {
        if bytes.len() > 128 * 1024 * 1024 {
            return Err(JsValue::from_str("Document exceeds the 128 MB limit"));
        }
        let (doc, _) = wordcraft_engine::io::open_bytes_with(name, bytes, password.as_deref()).map_err(|e| JsValue::from_str(&e))?;
        let mut session = Session::new(doc);
        let layout = session.layout();
        if layout.pages.len() > 10000 || layout.pages.iter().any(|p| [p.w, p.h].iter().any(|v| !v.is_finite() || *v <= 0.0 || *v > 14400.0)) {
            return Err(JsValue::from_str("Document exceeds the page count or page size limit"));
        }
        Ok(Self {
            session,
            layout,
            raster: Default::default(),
            width: 0,
            height: 0,
            password,
            id_prefix: String::new(),
            id_counter: 0,
            review: review::Review::default(),
        })
    }

    pub fn pages(&self) -> Result<String, JsValue> {
        let sizes: Vec<_> = self.layout.pages.iter().map(display_size).collect();
        serde_json::to_string(&sizes).map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// The same command ids and parameters used by wordcraft-cli and MCP.
    pub fn execute(&mut self, command: &str, params: &str) -> Result<String, JsValue> {
        if params.len() > 24 * 1024 * 1024 {
            return Err(JsValue::from_str("Command is too large"));
        }
        if !allowed(command) {
            return Err(JsValue::from_str("Command is not enabled in AI Office"));
        }
        let args: Value = serde_json::from_str(params).map_err(|e| JsValue::from_str(&e.to_string()))?;
        let revision = self.session.rev();
        if let Some(target) = args.get("_image") {
            let current = images::selected(&self.session, &self.layout);
            if target["revision"].as_u64() != Some(revision) || current.is_null() || target["parentId"] != current["parentId"] || target["offset"] != current["offset"] {
                return Err(JsValue::from_str("Image changed or selection moved. Select the image again."));
            }
        }
        self.session.close_typing();
        let result = self.session.run(command, &args).map_err(|e| JsValue::from_str(&e.to_string()))?;
        if self.session.rev() != revision && !self.id_prefix.is_empty() {
            agent::normalize(&mut self.session.doc, &self.id_prefix, &mut self.id_counter);
        }
        self.layout = self.session.layout();
        let mut state: Value = serde_json::from_str(&self.state()?).map_err(|e| JsValue::from_str(&e.to_string()))?;
        // Commands carry their result to the host the way the native engines do:
        // inside the state, bounded so a huge listing cannot bloat every broadcast.
        if !result.is_null() && result.to_string().len() <= 64 * 1024 {
            state["agentCommandResult"] = result;
        }
        Ok(state.to_string())
    }

    pub fn commands(&self) -> Result<String, JsValue> {
        let list: Vec<_> =
            self.session.registry.all().iter().filter(|c| allowed(c.id)).map(|c| json!({"id": c.id, "label": c.label, "params": c.params})).collect();
        serde_json::to_string(&list).map_err(|e| JsValue::from_str(&e.to_string()))
    }

    pub fn state(&self) -> Result<String, JsValue> {
        let sel = &self.session.sel;
        let caret = self
            .layout
            .caret_on(&sel.focus, self.session.page_hint)
            .map(|c| json!({"page":c.page,"x":c.x,"y":c.top,"height":c.height,"width":c.width}));
        let rects: Vec<_> = self
            .layout
            .selection_rects(&self.session.doc, &sel.anchor, &sel.focus, self.session.page_hint)
            .iter()
            .map(|(page, r)| json!({"page":page,"x":r.x,"y":r.y,"width":r.w,"height":r.h}))
            .collect();
        let comments: Vec<_> = self
            .session
            .doc
            .comments
            .iter()
            .map(|(id,c)| json!({"commentId":id,"parentId":c.parent,"resolved":c.resolved,"author":c.author,"text":self.session.doc.plain_text(wordcraft_doc::StoryRef::Part(c.part))}))
            .collect();
        let paths =
            if sel.is_collapsed() { vec![sel.focus.path.clone()] } else { self.session.doc.para_paths_in(&sel.ordered().0, &sel.ordered().1) };
        let at = paths.first().unwrap_or(&sel.focus.path);
        let paragraph = self.session.doc.para(sel.focus.story, at).map(|p| {
            let properties = wordcraft_engine::cmd::paragraph_properties::read(&self.session.doc, &p.props);
            let mut mixed = std::collections::BTreeSet::new();
            for path in paths.iter().skip(1) {
                if let Some(other) = self.session.doc.para(sel.focus.story, path) {
                    let other = wordcraft_engine::cmd::paragraph_properties::read(&self.session.doc, &other.props);
                    for (key, value) in properties.as_object().into_iter().flatten() {
                        if other.get(key) != Some(value) {
                            mixed.insert(key.clone());
                        }
                    }
                }
            }
            let r = self.session.doc.styles.resolve_para(&p.props);
            let ruler = self
                .layout
                .caret_on(&sel.focus, self.session.page_hint)
                .and_then(|c| self.layout.pages.get(c.page))
                .filter(|page| {
                    sel.focus.story == wordcraft_doc::StoryRef::Body
                        && at.0.len() == 1
                        && !r.bidi
                        && !r.mirror_indents
                        && r.frame.is_none()
                        && r.numbering.is_none()
                        && self.session.doc.sections().get(page.section).is_some_and(|(_, s)| s.columns.count == 1)
                })
                .map(|page| json!({"left":page.body.x,"width":page.body.w}));
            json!({"id":p.props.office_id,"properties":properties,"mixed":mixed,"ruler":ruler})
        });
        // Page measurements stay available even when paragraph handles are unsupported.
        let page_ruler = self.layout.caret_on(&sel.focus, self.session.page_hint)
            .and_then(|c| self.layout.pages.get(c.page))
            .or_else(|| self.layout.pages.get(self.session.page_hint))
            .map(|page| json!({"left":page.body.x,"width":page.body.w}));
        let active_page = self.layout.caret_on(&sel.focus, self.session.page_hint).map(|c| c.page).unwrap_or(self.session.page_hint);
        let page_setup = self.layout.pages.get(active_page).and_then(|page| self.session.doc.sections().get(page.section).map(|(_, props)| {
            let mut value = wordcraft_engine::cmd::page_properties::read(props);
            value["page"] = json!(active_page); value["section"] = json!(page.section); value
        }));
        let regions: Vec<_> = self.layout.pages.iter().map(|p| json!({"top":p.body.y,"bottom":p.body.bottom(),"width":p.w,"height":p.h})).collect();
        let editing_region = match sel.focus.story {
            wordcraft_doc::StoryRef::Part(id) => match self.session.doc.parts.get(&id).map(|p| &p.kind) {
                Some(wordcraft_doc::PartKind::Header) => "header",
                Some(wordcraft_doc::PartKind::Footer) => "footer",
                _ => "body",
            },
            _ => "body",
        };
        let style = self.session.doc.para(sel.focus.story, &sel.focus.path).and_then(|p| p.props.style.as_deref()).unwrap_or("Normal");
        serde_json::to_string(&json!({"image":images::selected(&self.session,&self.layout),"pageSetup":page_setup,"pageRegions":regions,"editingRegion":editing_region,"formatting":wordcraft_engine::cmd::format::state(&self.session),"paragraph":paragraph,"pageRuler":page_ruler,"undoDepth":self.session.undo_depth(),"paragraphStyle":style,"selection":sel,"caret":caret,"rects":rects,"comments":comments,"selectedText":self.session.selected_text(),"revision":self.session.rev(),"undo":self.session.can_undo(),"redo":self.session.can_redo(),"find":if self.session.find.results.is_empty() { None } else { Some(json!({"count":self.session.find.results.len(),"current":self.session.find.current})) },"painter":self.session.painter.as_ref().map(|(_,_,sticky)| json!({"active":true,"sticky":*sticky}))})).map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// Geometry only for the page being rendered; no document-wide scan on typing.
    pub fn tables(&self, page: usize) -> Result<String, JsValue> {
        let tables: Vec<_> = wordcraft_engine::cmd::table_resize::tables_on(&self.layout, &self.session.doc, page)
            .iter().map(|t| json!({"cols":t.cols,"rows":t.rows.iter().map(|r| json!({"row":r.row,"top":r.top,"bottom":r.bottom,"edges":r.edges})).collect::<Vec<_>>(),"endsHere":t.ends_here,"movable":wordcraft_engine::cmd::table_mouse::movable(t)})).collect();
        serde_json::to_string(&tables).map_err(|e| JsValue::from_str(&e.to_string()))
    }

    pub fn images(&self, page: usize) -> Result<String, JsValue> {
        Ok(images::geometry(&self.session, &self.layout, page).to_string())
    }

    /// Pure drag feedback. Uses the same plan as the eventual native command.
    pub fn table_preview(&self, args: &str) -> Result<String, JsValue> {
        let args: Value = serde_json::from_str(args).map_err(|e| JsValue::from_str(&e.to_string()))?;
        table::preview(&self.session, &self.layout, &args).map(|v| v.to_string()).map_err(|e| JsValue::from_str(&e))
    }

    pub fn export(&self) -> Result<Vec<u8>, JsValue> {
        wordcraft_engine::io::save_bytes_with("document.docx", &self.session.doc, self.password.as_deref(), &Default::default())
            .map_err(|e| JsValue::from_str(&e))
    }

    pub fn agent_init(&mut self, prefix: &str, metadata: &str) -> Result<(), JsValue> {
        if prefix.is_empty() || prefix.len() > 60 {
            return Err(JsValue::from_str("Invalid identity prefix"));
        }
        self.id_prefix = prefix.into();
        if let Ok(value) = serde_json::from_str(metadata) {
            let _ = agent::restore_ids(&mut self.session.doc, &value);
        }
        agent::normalize(&mut self.session.doc, prefix, &mut self.id_counter);
        Ok(())
    }

    pub fn agent_metadata(&self) -> String {
        agent::metadata(&self.session.doc).to_string()
    }

    pub fn review_start(&mut self) -> Result<(), JsValue> {
        self.review.start(&self.session.doc).map_err(|e| JsValue::from_str(&e))
    }

    pub fn review_finish(&mut self) {
        self.review.finish(&self.session.doc);
    }

    pub fn review_read(&mut self) -> String {
        self.review.refresh(&self.session.doc);
        self.review.read().to_string()
    }

    pub fn review_export(&self) -> String {
        self.review.export()
    }
    pub fn review_import(&mut self, value: &str) -> Result<(), JsValue> {
        self.review.import(value).map_err(|e| JsValue::from_str(&e))
    }
    pub fn review_projection(&self) -> String {
        review::Review::projection(&self.session.doc)
    }

    pub fn review_decide(&mut self, revision: u64, action: &str, keys: &str) -> Result<String, JsValue> {
        if revision != self.session.rev() {
            return Err(JsValue::from_str("Document changed. Refresh the review before deciding."));
        }
        let keys: Vec<String> = serde_json::from_str(keys).map_err(|e| JsValue::from_str(&e.to_string()))?;
        let result = self.review.decide(&self.session.doc, action, &keys).map_err(|e| JsValue::from_str(&e))?;
        if self.session.doc != result {
            self.session.close_typing();
            self.session.checkpoint("Undo agent change");
            self.session.doc = result;
            self.session.clamp_selection();
            self.session.touch();
            self.session.relayout();
            self.layout = self.session.layout();
        }
        self.state()
    }

    pub fn agent_read(&self, args: &str) -> Result<String, JsValue> {
        let args: Value = serde_json::from_str(args).map_err(|e| JsValue::from_str(&e.to_string()))?;
        if args["operation"] == "comments" {
            return Ok(agent::comments(&self.session.doc, &args).to_string());
        }
        if let Some(id) = args["sectionId"].as_str() {
            let node = agent::target(&self.session.doc, id).map_err(|e| JsValue::from_str(&e))?;
            if node.kind != "heading" {
                return Err(JsValue::from_str("sectionId must identify a heading"));
            }
        }
        if args["operation"] == "styles" {
            let query = args["query"].as_str().unwrap_or("").to_lowercase();
            let builtin = wordcraft_doc::styles::StyleSheet::builtin();
            let mut styles = self.session.doc.styles.styles.clone();
            styles.extend(builtin.styles.iter().filter(|s| self.session.doc.styles.get(&s.id).is_none()).cloned());
            let values: Vec<_> = styles
                .iter()
                .filter(|s| query.is_empty() || s.name.to_lowercase().contains(&query) || s.id.to_lowercase().contains(&query))
                .take(50)
                .map(|s| {
                    let sheet = if self.session.doc.styles.get(&s.id).is_some() { &self.session.doc.styles } else { &builtin };
                    let format = sheet.resolve_char(Some(&s.id), &Default::default());
                    json!({"id":s.id,"name":s.name,"size":format.size,"font":format.font})
                })
                .collect();
            return Ok(json!({"styles":values}).to_string());
        }
        if args["operation"] == "selection" {
            let selection = &self.session.sel;
            let all = agent::nodes(&self.session.doc);
            let focus = all.iter().find(|n| n.story == selection.focus.story && n.path == selection.focus.path);
            let anchor = all.iter().find(|n| n.story == selection.anchor.story && n.path == selection.anchor.path);
            let offset =
                |pos: &wordcraft_doc::Pos| self.session.doc.para_at(pos).map_or(0, |p| p.text.get(..pos.off).unwrap_or(&p.text).chars().count());
            let result = focus.map(|n| agent::read(&self.session.doc, &json!({"id":n.id}))).unwrap_or(json!({"blocks":[]}));
            return Ok(json!({"content":result,"selection":{"anchorId":anchor.map(|n|&n.id),"focusId":focus.map(|n|&n.id),"start":offset(&selection.anchor),"end":offset(&selection.focus)},"selectedText":self.session.selected_text().chars().take(8000).collect::<String>()}).to_string());
        }
        Ok(agent::read(&self.session.doc, &args).to_string())
    }

    pub fn agent_rects(&self, id: &str) -> Result<String, JsValue> {
        agent::rects(&self.session.doc, &self.layout, id).map(|v| v.to_string()).map_err(|e| JsValue::from_str(&e))
    }

    pub fn agent_edit(&mut self, args: &str) -> Result<String, JsValue> {
        if args.len() > 160_000 {
            return Err(JsValue::from_str("Agent edit is too large"));
        }
        let args: Value = serde_json::from_str(args).map_err(|e| JsValue::from_str(&e.to_string()))?;
        let before_ids: std::collections::HashSet<_> = agent::nodes(&self.session.doc).into_iter().map(|n| n.id).collect();
        // Prepare on a private copy: failures leave content, revision, selection and undo intact.
        let mut draft = Session::new(self.session.doc.clone());
        draft.sel = self.session.sel.clone();
        draft.author = self.session.author.clone();
        let edits = if args["operation"] == "batch" {
            let edits = args["edits"].as_array().ok_or_else(|| JsValue::from_str("edits must be an array"))?;
            if edits.is_empty() || edits.len() > 200 {
                return Err(JsValue::from_str("A batch requires 1-200 edits"));
            }
            if edits.iter().any(|edit| edit["operation"] == "batch") {
                return Err(JsValue::from_str("Nested batches are not supported"));
            }
            edits.clone()
        } else {
            vec![args.clone()]
        };
        agent::validate_guards(&self.session.doc, &edits, args["revision"].as_u64() == Some(self.session.rev()))
            .map_err(|error| JsValue::from_str(&error))?;
        let mut applied = Vec::new();
        let mut command_result = Value::Null;
        for edit in &edits {
            command_result = agent::edit(&mut draft, edit).map_err(|e| JsValue::from_str(&e))?;
            applied.push(json!({"id":edit["id"],"operation":edit["operation"],"action":edit["action"],"commentId":edit["commentId"],"command":edit["command"],"params":edit["params"]}));
        }
        agent::normalize(&mut draft.doc, &self.id_prefix, &mut self.id_counter);
        if self.session.doc != draft.doc {
            self.session.close_typing();
            self.session.checkpoint("Agent document edit");
            self.session.doc = draft.doc;
            self.session.sel = draft.sel;
            self.session.touch();
        }
        self.session.relayout();
        self.layout = self.session.layout();
        let nodes = agent::nodes(&self.session.doc);
        let created: std::collections::HashSet<_> = nodes.iter().filter(|n| !before_ids.contains(&n.id)).map(|n| n.id.clone()).collect();
        let roots: Vec<_> = nodes
            .iter()
            .filter(|n| created.contains(&n.id) && !created.contains(&n.parent))
            .take(50)
            .map(|n| json!({"id":n.id,"kind":n.kind,"parentId":n.parent}))
            .collect();
        let mut result: Value = serde_json::from_str(&self.state()?).map_err(|e| JsValue::from_str(&e.to_string()))?;
        result["agentCreatedBlocks"] = json!(roots);
        result["agentCommandResult"] = command_result;
        result["agentAppliedEdits"] = json!(applied);
        Ok(result.to_string())
    }

    pub fn render(&mut self, index: usize, scale: f32) -> Result<Vec<u8>, JsValue> {
        let page = self.layout.pages.get(index).ok_or_else(|| JsValue::from_str("Page does not exist"))?;
        if !scale.is_finite() || scale <= 0.0 {
            return Err(JsValue::from_str("Invalid page scale"));
        }
        let size = display_size(page);
        let expanded;
        let page = if size != [page.w, page.h] {
            expanded = { let mut p = page.clone(); p.w = size[0]; p.h = size[1]; p };
            &expanded
        } else { page };
        let scale = scale.min(4096.0 / page.w).min(4096.0 / page.h).min((8_000_000.0 / (page.w * page.h)).sqrt());
        let mut opts = wordcraft_render::RenderOptions::default();
        let editing_hf = match self.session.sel.focus.story {
            wordcraft_doc::StoryRef::Part(id) => self.session.doc.parts.get(&id).is_some_and(|p| matches!(p.kind, wordcraft_doc::PartKind::Header | wordcraft_doc::PartKind::Footer)),
            _ => false,
        };
        opts.display.dim_header = !editing_hf;
        opts.display.dim_body = editing_hf;
        let result = self.raster.render_with_options(&self.session.doc, page, index, scale, &opts);
        self.width = result.width;
        self.height = result.height;
        Ok(result.to_straight())
    }

    #[wasm_bindgen(getter)]
    pub fn width(&self) -> u32 {
        self.width
    }
    #[wasm_bindgen(getter)]
    pub fn height(&self) -> u32 {
        self.height
    }
}

fn allowed(command: &str) -> bool {
    matches!(
        command,
        "text.insert"
            | "layout.pageProperties"
            | "insert.editHeader"
            | "insert.editFooter"
            | "insert.closeHeader"
            | "edit.pasteText"
            | "text.newParagraph"
            | "text.lineBreak"
            | "text.tab"
            | "text.backTab"
            | "text.backspace"
            | "text.delete"
            | "caret.set"
            | "caret.left"
            | "caret.right"
            | "caret.up"
            | "caret.down"
            | "caret.home"
            | "caret.end"
            | "caret.docStart"
            | "caret.docEnd"
            | "select.all"
            | "select.range"
            | "select.word"
            | "select.text"
            | "select.collapse"
            | "format.bold"
            | "format.strikethrough"
            | "format.color"
            | "format.highlight"
            | "format.clear"
            | "format.italic"
            | "format.underline"
            | "format.font"
            | "format.size"
            | "para.properties"
            | "para.align"
            | "para.style"
            | "para.bullets"
            | "para.numbering"
            | "para.lineSpacing"
            | "table.insertRowAbove"
            | "table.insertRowBelow"
            | "table.insertColumnLeft"
            | "table.insertColumnRight"
            | "table.deleteRow"
            | "table.deleteColumn"
            | "table.deleteTable"
            | "table.columnWidth"
            | "table.moveAxis"
            | "table.sort"
            | "table.shading"
            | "table.cellAlign"
            | "table.selectRow"
            | "table.selectColumn"
            | "table.selectTable"
            | "table.resizeDrag"
            | "table.moveDrag"
            | "table.fitColumn"
            | "table.selectAt"
            | "table.insertAt"
            | "table.autofit"
            | "table.distributeColumns"
            | "table.distributeRows"
            | "table.rowHeight"
            | "table.repeatHeader"
            | "table.merge"
            | "table.split"
            | "table.borders"
            | "table.selectCell"
            | "insert.table"
            | "picture.select" | "picture.delete" | "picture.size" | "picture.crop" | "picture.change" | "picture.altText"
            | "object.altText" | "picture.reset" | "picture.border" | "picture.style"
            | "arrange.wrap" | "arrange.wrapText" | "arrange.moveWithText" | "arrange.position" | "arrange.align"
            | "arrange.bounds" | "arrange.nudge" | "arrange.rotation" | "arrange.rotate"
            | "arrange.bringForward" | "arrange.sendBackward" | "references.caption"
            | "insert.picture"
            | "insert.link"
            | "review.newComment"
            | "review.comments"
            | "review.wordCount"
            | "edit.goto"
            | "edit.find"
            | "edit.findNext"
            | "edit.findPrevious"
            | "edit.replace"
            | "edit.replaceAll"
            | "edit.formatPainter"
            | "edit.pasteFormat"
            | "edit.undo"
            | "edit.redo"
    )
}

/// Grow only the editor surface to include laid-out overflow. The document's page
/// settings and native layout stay unchanged, including table wrapping and export.
fn display_size(page: &wordcraft_layout::Page) -> [f32; 2] {
    use wordcraft_layout::Placed;
    let mut width = page.w;
    let mut height = page.h;
    let mut include = |right: f32, bottom: f32| {
        if right.is_finite() && right > page.w { width = width.max((right + 16.0).min(14400.0)); }
        if bottom.is_finite() && bottom > page.h { height = height.max((bottom + 16.0).min(14400.0)); }
    };
    for item in page.items.iter().chain(&page.header).chain(&page.footer) {
        match item {
            Placed::Fill { rect, .. } | Placed::Image { rect, .. } | Placed::Shape { rect, .. }
            | Placed::Graphic { rect, .. } | Placed::Cell { rect, .. } | Placed::Object { rect, .. }
            | Placed::Clip { rect, .. } => include(rect.right(), rect.bottom()),
            Placed::Rule { x0, y0, x1, y1, .. } => include(x0.max(*x1), y0.max(*y1)),
            Placed::Lines { para, l0, l1, x, y, turn, .. } => {
                if let Some(first) = para.lines.get(*l0) {
                    for line in para.lines.iter().take(*l1).skip(*l0) {
                        let r = wordcraft_layout::turn_rect(*turn, *x, *y, wordcraft_geom::Rect {
                            x: line.left, y: line.top - first.top,
                            w: line.right.max(line.visual_end_x()) - line.left, h: line.height,
                        });
                        include(r.right(), r.bottom());
                    }
                }
            }
        }
    }
    [width, height]
}
