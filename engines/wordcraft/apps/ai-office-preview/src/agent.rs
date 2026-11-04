//! Internal, revision-checked block editing for the AI Office agent.
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use std::collections::{HashMap, HashSet};
use std::sync::Arc;
use wordcraft_doc::para::InlineObject;
use wordcraft_doc::{Block, Blocks, Document, Paragraph, Path, Pos, StoryRef, Table};
use wordcraft_engine::{Selection, Session};
use wordcraft_layout::DocLayout;

#[derive(Clone)]
pub struct Node {
    pub id: String,
    pub kind: &'static str,
    pub parent: String,
    pub story: StoryRef,
    pub path: Path,
    pub text: String,
    pub style: Option<String>,
    pub level: Option<u8>,
}

pub fn normalize(doc: &mut Document, prefix: &str, counter: &mut u64) {
    // Inspect shared blocks before writing: history and the paragraph layout cache
    // retain these Arcs. Making every block mutable invalidates all their caches.
    fn next(id: &str, seen: &mut HashSet<String>, prefix: &str, counter: &mut u64) -> Option<String> {
        if !id.is_empty() && seen.insert(id.to_owned()) {
            return None;
        }
        loop {
            *counter += 1;
            let id = format!("{prefix}-{:x}", *counter);
            if seen.insert(id.clone()) {
                return Some(id);
            }
        }
    }
    fn repair(block: &Arc<Block>, seen: &mut HashSet<String>, prefix: &str, counter: &mut u64) -> Option<Arc<Block>> {
        match &**block {
            Block::Para(p) => next(&p.props.office_id, seen, prefix, counter).map(|id| {
                let mut p = p.clone();
                p.props.office_id = id;
                Arc::new(Block::Para(p))
            }),
            Block::Table(t) => {
                let mut changed = None;
                if let Some(id) = next(&t.props.office_id, seen, prefix, counter) {
                    changed.get_or_insert_with(|| t.clone()).props.office_id = id;
                }
                for (r, row) in t.rows.iter().enumerate() {
                    if let Some(id) = next(&row.props.office_id, seen, prefix, counter) {
                        changed.get_or_insert_with(|| t.clone()).rows[r].props.office_id = id;
                    }
                    for (c, cell) in row.cells.iter().enumerate() {
                        if let Some(id) = next(&cell.props.office_id, seen, prefix, counter) {
                            changed.get_or_insert_with(|| t.clone()).rows[r].cells[c].props.office_id = id;
                        }
                        for (b, child) in cell.blocks.iter().enumerate() {
                            if let Some(child) = repair(child, seen, prefix, counter) {
                                changed.get_or_insert_with(|| t.clone()).rows[r].cells[c].blocks[b] = child;
                            }
                        }
                    }
                }
                changed.map(|t| Arc::new(Block::Table(t)))
            }
        }
    }
    fn walk(blocks: &mut Blocks, seen: &mut HashSet<String>, prefix: &str, counter: &mut u64) {
        for block in blocks {
            if let Some(fixed) = repair(block, seen, prefix, counter) {
                *block = fixed;
            }
        }
    }
    let mut seen = HashSet::new();
    walk(&mut doc.body, &mut seen, prefix, counter);
    for part in doc.parts.values_mut() {
        walk(&mut part.blocks, &mut seen, prefix, counter);
    }
}

pub fn nodes(doc: &Document) -> Vec<Node> {
    fn walk(doc: &Document, blocks: &Blocks, story: StoryRef, prefix: &[u32], parent: &str, out: &mut Vec<Node>) {
        for (index, block) in blocks.iter().enumerate() {
            let mut path = prefix.to_vec();
            path.push(index as u32);
            match &**block {
                Block::Para(p) => {
                    let level = doc.styles.resolve_para(&p.props).outline_level.map(|level| level + 1);
                    let kind = if level.is_some() {
                        "heading"
                    } else if p.props.numbering.is_some() {
                        "listItem"
                    } else {
                        "paragraph"
                    };
                    out.push(Node {
                        id: p.props.office_id.clone(),
                        kind,
                        parent: parent.into(),
                        story,
                        path: Path(path),
                        text: p.text.clone(),
                        style: p.props.style.clone(),
                        level,
                    });
                }
                Block::Table(t) => {
                    out.push(Node {
                        id: t.props.office_id.clone(),
                        kind: "table",
                        parent: parent.into(),
                        story,
                        path: Path(path.clone()),
                        text: String::new(),
                        style: t.props.style.clone(),
                        level: None,
                    });
                    for (r, row) in t.rows.iter().enumerate() {
                        let mut rp = path.clone();
                        rp.push(r as u32);
                        out.push(Node {
                            id: row.props.office_id.clone(),
                            kind: "row",
                            parent: t.props.office_id.clone(),
                            story,
                            path: Path(rp.clone()),
                            text: String::new(),
                            style: None,
                            level: None,
                        });
                        for (c, cell) in row.cells.iter().enumerate() {
                            let mut cp = rp.clone();
                            cp.push(c as u32);
                            out.push(Node {
                                id: cell.props.office_id.clone(),
                                kind: "cell",
                                parent: row.props.office_id.clone(),
                                story,
                                path: Path(cp.clone()),
                                text: String::new(),
                                style: None,
                                level: None,
                            });
                            walk(doc, &cell.blocks, story, &cp, &cell.props.office_id, out);
                        }
                    }
                }
            }
        }
    }
    let mut out = Vec::new();
    walk(doc, &doc.body, StoryRef::Body, &[], "body", &mut out);
    let mut parts: Vec<_> = doc.parts.iter().collect();
    parts.sort_by_key(|(id, _)| story_key(doc, StoryRef::Part(**id)));
    for (id, part) in parts {
        walk(doc, &part.blocks, StoryRef::Part(*id), &[], &format!("part-{}", story_key(doc, StoryRef::Part(*id))), &mut out);
    }
    out
}

// DOCX can renumber parts when serializing default/first/even references.
// Address headers and footers by their section and slot, not a transient part number.
fn story_key(doc: &Document, story: StoryRef) -> String {
    if let StoryRef::Part(id) = story {
        for (section, (_, props)) in doc.sections().iter().enumerate() {
            for (kind, set) in [("header", props.headers), ("footer", props.footers)] {
                for (slot, target) in [("default", set.default), ("first", set.first), ("even", set.even)] {
                    if target == Some(id) { return format!("{kind}-{section}-{slot}"); }
                }
            }
        }
        return id.to_string();
    }
    "body".into()
}

pub fn metadata(doc: &Document) -> Value {
    Value::Array(nodes(doc).iter().map(|n| json!({"id":n.id,"story":n.story,"storyKey":story_key(doc,n.story),"path":n.path,"kind":n.kind})).collect())
}

pub fn restore_ids(doc: &mut Document, values: &Value) -> Result<(), String> {
    let items = values.as_array().ok_or("Invalid identity metadata")?;
    let current = nodes(doc);
    if current.len() != items.len() {
        return Err("Identity metadata no longer matches the document".into());
    }
    let keyed = items.iter().all(|item| item["storyKey"].is_string());
    let key = |story: &str, path: &Value, kind: &str| format!("{story}:{}:{kind}", path);
    let index: HashMap<_, _> = items.iter().map(|item| (key(item["storyKey"].as_str().unwrap_or_default(), &item["path"], item["kind"].as_str().unwrap_or_default()), item)).collect();
    if keyed && index.len() != items.len() { return Err("Duplicate identity location".into()); }
    let matched: Vec<_> = current.iter().enumerate().map(|(i, node)| {
        if keyed { index.get(&key(&story_key(doc, node.story), &json!(node.path), node.kind)).copied().ok_or("Identity location no longer matches") }
        else { Ok(&items[i]) }
    }).collect::<Result<_, _>>()?;
    let mut seen = HashSet::new();
    for (node, item) in current.iter().zip(&matched) {
        let id = item["id"].as_str().ok_or("Missing identity")?;
        if id.len() > 100
            || id.is_empty()
            || !seen.insert(id)
            || (!keyed && item["story"] != json!(node.story))
            || item["path"] != json!(node.path)
            || item["kind"] != node.kind
        {
            return Err("Identity metadata no longer matches the document".into());
        }
    }
    for (node, item) in current.iter().zip(matched) {
        let id = item["id"].as_str().unwrap_or_default().to_string();
        match node.kind {
            "table" => doc.table_mut(node.story, &node.path).map_err(|e| e.to_string())?.props.office_id = id,
            "row" | "cell" => {
                let p = &node.path.0;
                let count = if node.kind == "row" { 1 } else { 2 };
                let t = doc.table_mut(node.story, &Path(p[..p.len() - count].to_vec())).map_err(|e| e.to_string())?;
                let row = &mut t.rows[p[p.len() - count] as usize];
                if count == 1 {
                    row.props.office_id = id;
                } else {
                    row.cells[p[p.len() - 1] as usize].props.office_id = id;
                }
            }
            _ => doc.para_mut(node.story, &node.path).map_err(|e| e.to_string())?.props.office_id = id,
        }
    }
    Ok(())
}

fn heading_level(node: &Node) -> Option<u8> {
    node.level
}

pub fn block_hash(doc: &Document, node: &Node) -> String {
    let value = match node.kind {
        "table" => doc.table(node.story, &node.path).map(|value| json!(value)),
        "row" | "cell" => {
            let path = &node.path.0;
            let count = if node.kind == "row" { 1 } else { 2 };
            doc.table(node.story, &Path(path[..path.len() - count].to_vec())).and_then(|t| {
                let row = t.rows.get(path[path.len() - count] as usize)?;
                if count == 1 { Some(json!(row)) } else { row.cells.get(path[path.len() - 1] as usize).map(|cell| json!(cell)) }
            })
        }
        _ => doc.para(node.story, &node.path).map(|value| json!(value)),
    };
    Sha256::digest(value.unwrap_or(Value::Null).to_string().as_bytes()).iter().map(|byte| format!("{byte:02x}")).collect()
}

pub fn guarded_text_edit(edit: &Value) -> bool {
    edit["ifHash"].as_str().is_some_and(|value| value.len() == 64 && value.bytes().all(|c| c.is_ascii_hexdigit()))
        && matches!(edit["operation"].as_str(), Some("update" | "replaceText"))
}

pub fn validate_guards(doc: &Document, edits: &[Value], current_revision: bool) -> Result<(), String> {
    if !current_revision && !edits.iter().all(guarded_text_edit) {
        return Err("Document changed. Structural edits require a fresh document version.".into());
    }
    for edit in edits {
        if let Some(hash) = edit.get("ifHash") {
            let node = target(doc, edit["id"].as_str().ok_or("A block guard requires a target ID")?)?;
            if hash.as_str() != Some(block_hash(doc, &node).as_str()) {
                return Err("Block changed. Read this target again; your later edits were preserved.".into());
            }
        }
    }
    Ok(())
}

fn contains_comments(doc: &Document, node: &Node) -> bool {
    doc.para_paths(node.story).iter().filter(|path| path.0.starts_with(&node.path.0)).any(|path| {
        doc.para(node.story, path)
            .is_some_and(|p| p.objects.iter().any(|o| matches!(o, InlineObject::CommentStart { .. } | InlineObject::CommentEnd { .. })))
    })
}

fn require_detach(doc: &Document, node: &Node, args: &Value) -> Result<(), String> {
    if args["allowDetach"] != true && contains_comments(doc, node) {
        return Err("This would detach comment anchors. Preserve the commented content, or use allowDetach only when the user asked to remove it. Threads are retained as orphaned comments.".into());
    }
    Ok(())
}

pub fn read(doc: &Document, args: &Value) -> Value {
    if args["operation"] == "tail" {
        return read(doc, &json!({"parentId":"body","offset":doc.body.len().saturating_sub(1),"limit":1}));
    }
    let all = nodes(doc);
    let id = args["id"].as_str();
    let parent = args["parentId"].as_str();
    let query = args["query"].as_str().unwrap_or("").to_lowercase();
    let offset = usize::try_from(args["offset"].as_u64().unwrap_or(0)).unwrap_or(usize::MAX);
    let limit = args["limit"].as_u64().unwrap_or(20).clamp(1, 20) as usize;
    let text_offset = usize::try_from(args["textOffset"].as_u64().unwrap_or(0)).unwrap_or(usize::MAX);
    let text_limit = args["textLimit"].as_u64().unwrap_or(if id.is_some() { 4000 } else { 160 }).clamp(1, 8000) as usize;
    let text_limit = text_limit.min(8000 / if id.is_some() { 1 } else { limit });
    let section = args["sectionId"].as_str().and_then(|id| all.iter().position(|n| n.id == id));
    let section_end = section.map(|start| {
        let heading = &all[start];
        all.iter()
            .enumerate()
            .skip(start + 1)
            .find(|(_, n)| {
                n.story == heading.story
                    && n.parent == heading.parent
                    && heading_level(n).is_some_and(|level| level <= heading_level(heading).unwrap_or(9))
            })
            .map(|(index, _)| index)
            .unwrap_or(all.len())
    });
    let outline = args["operation"] == "outline";
    let matches: Vec<_> = all
        .iter()
        .enumerate()
        .filter(|(index, n)| {
            let in_section = if args.get("sectionId").is_some() {
                section.zip(section_end).is_some_and(|(start, end)| {
                    *index >= start
                        && *index < end
                        && n.story == all[start].story
                        && n.path.0.starts_with(&all[start].path.0[..all[start].path.0.len() - 1])
                })
            } else {
                true
            };
            in_section && (!outline || heading_level(n).is_some() || n.kind == "table")
        })
        .map(|(_, n)| n)
        .filter(|n| {
            id.is_none_or(|id| n.id == id) && parent.is_none_or(|p| n.parent == p) && (query.is_empty() || n.text.to_lowercase().contains(&query))
        })
        .collect();
    let results: Vec<_> = matches.iter().skip(offset).take(limit).map(|n| {
        let length = n.text.chars().count();
        let text: String = n.text.chars().skip(text_offset).take(text_limit).collect();
        let end = text_offset + text.chars().count();
        json!({"id":n.id,"kind":n.kind,"parentId":n.parent,"style":n.style,"text":text,"textLength":length,"textOffset":text_offset,"nextTextOffset":if end < length {Some(end)}else{None},"ifHash":block_hash(doc,n),"headingLevel":heading_level(n),"paragraph":if id.is_some(){doc.para(n.story,&n.path).map(|p|wordcraft_engine::cmd::paragraph_properties::read(doc,&p.props))}else{None},"images":if !outline { doc.para(n.story,&n.path).map(|p|p.object_offsets().into_iter().filter_map(|off|super::images::details(doc,&Pos::new(n.story,n.path.clone(),off),p.object_at(off)?)).collect::<Vec<_>>()) }else{None},"children":all.iter().filter(|c|c.parent==n.id).count()})
    }).collect();
    json!({"blocks":results,"total":matches.len(),"nextOffset":if offset+results.len()<matches.len(){Some(offset+results.len())}else{None},"outline":outline,"summary":if outline {json!({"headings":all.iter().filter(|n|heading_level(n).is_some()).count(),"tables":all.iter().filter(|n|n.kind=="table").count(),"blocks":all.len()})}else{Value::Null}})
}

pub fn target(doc: &Document, id: &str) -> Result<Node, String> {
    nodes(doc).into_iter().find(|n| n.id == id).ok_or_else(|| "Block no longer exists. Read the document again.".into())
}

pub fn rects(doc: &Document, layout: &DocLayout, id: &str) -> Result<Value, String> {
    let n = target(doc, id)?;
    let mut rectangles = Vec::new();
    for path in doc.para_paths(n.story).into_iter().filter(|p| p.0.starts_with(&n.path.0)) {
        let Some(p) = doc.para(n.story, &path) else { continue };
        let a = Pos::new(n.story, path.clone(), 0);
        let b = Pos::new(n.story, path, p.text.len());
        for (page, r) in layout.selection_rects(doc, &a, &b, 0) {
            rectangles.push(json!({"page":page,"x":r.x,"y":r.y,"width":r.w.max(8.0),"height":r.h}));
        }
        if p.text.is_empty() {
            if let Some(c) = layout.caret_on(&a, 0) {
                rectangles.push(json!({"page":c.page,"x":c.x,"y":c.top,"width":80,"height":c.height}));
            }
        }
    }
    Ok(Value::Array(rectangles))
}

fn text(args: &Value) -> Result<&str, String> {
    let value = args["text"].as_str().ok_or("text must be a string")?;
    if value.len() > 100_000 || value.contains('\u{fffc}') {
        return Err("Text is too large or contains an inline-object marker".into());
    }
    Ok(value)
}

fn index(args: &Value, key: &str, default: usize) -> Result<usize, String> {
    match args.get(key) {
        None => Ok(default),
        Some(value) => value.as_u64().and_then(|v| usize::try_from(v).ok()).ok_or_else(|| format!("{key} must be a non-negative character offset")),
    }
}

fn new_block(value: &Value) -> Result<Block, String> {
    match value["kind"].as_str().unwrap_or("paragraph") {
        "paragraph" | "heading" => {
            let content = text(value)?;
            if content.contains(['\n', '\r']) {
                return Err("Use separate blocks for separate paragraphs".into());
            }
            let mut p = Paragraph::with_text(content, Default::default());
            let style =
                if value["kind"] == "heading" { value["style"].as_str().unwrap_or("Heading1") } else { value["style"].as_str().unwrap_or("Normal") };
            p.props.style = Some(style.into());
            Ok(Block::Para(p))
        }
        "table" => {
            let rows = value["rows"].as_array().ok_or("rows must be an array of arrays of strings")?;
            let cols = rows.first().and_then(Value::as_array).map_or(0, Vec::len);
            if rows.is_empty() || rows.len() > 20 || cols == 0 || cols > 10 {
                return Err("Use 1-20 rows and 1-10 columns".into());
            }
            let mut table = Table::new(rows.len(), cols, 432.0);
            table.props.borders = Some(wordcraft_doc::props::Borders::all(wordcraft_doc::props::Border::single(0.5)));
            for (r, row) in rows.iter().enumerate() {
                let cells = row.as_array().ok_or("Invalid row")?;
                if cells.len() != cols {
                    return Err("Rows must have equal column counts".into());
                }
                for (c, cell) in cells.iter().enumerate() {
                    let value = cell.as_str().ok_or("Cells must be strings")?;
                    if value.len() > 10_000 || value.contains('\u{fffc}') {
                        return Err("Invalid cell text".into());
                    }
                    table.rows[r].cells[c].blocks = vec![Arc::new(Block::Para(Paragraph::with_text(value, Default::default())))];
                }
            }
            Ok(Block::Table(table))
        }
        _ => Err("Supported new blocks: paragraph, heading, table".into()),
    }
}

pub fn edit(session: &mut Session, args: &Value) -> Result<Value, String> {
    if session.doc.settings.protection.is_some() || session.doc.settings.track_changes {
        return Err("Agent block editing is unavailable while document protection or native Track Changes is enabled".into());
    }
    if args["operation"] == "comment" {
        return edit_comment(session, args);
    }
    let append = args["operation"] == "append";
    let node = if append {
        if args.get("id").is_some() || args.get("position").is_some() || args.get("ifHash").is_some() {
            return Err("Append resolves the document end internally. Omit id, position and ifHash.".into());
        }
        Node {id:String::new(),kind:"paragraph",parent:"body".into(),story:StoryRef::Body,path:Path(vec![session.doc.body.len() as u32]),text:String::new(),style:None,level:None}
    } else {
        target(&session.doc, args["id"].as_str().ok_or("A target block ID is required")?)?
    };
    let id = node.id.as_str();
    let mut command_result = Value::Null;
    match args["operation"].as_str().ok_or("operation is required")? {
        "insert" | "append" => {
            if matches!(node.kind, "row" | "cell") {
                return Err("Insert relative to a paragraph or table, including paragraphs inside cells".into());
            }
            let values = args["blocks"].as_array().ok_or("blocks must be an array")?;
            if values.is_empty() || values.len() > 50 {
                return Err("Insert 1-50 blocks at a time".into());
            }
            let blocks: Vec<_> = values.iter().map(new_block).collect::<Result<_, _>>()?;
            let after = if append { 0 } else { match args["position"].as_str().unwrap_or("after") {
                "after" => 1,
                "before" => 0,
                _ => return Err("position must be before or after".into()),
            }};
            for (i, block) in blocks.into_iter().enumerate() {
                let style = match &block {
                    Block::Para(p) => p.props.style.clone(),
                    _ => None,
                };
                let path = node.path.with_last(node.path.last() + after + i);
                session.doc.insert_block(node.story, &path, block).map_err(|e| e.to_string())?;
                if let Some(style) = style {
                    // Reuse the editor's imported-style handling inside the private transaction.
                    session.sel = Selection::caret(Pos::new(node.story, path, 0));
                    session.run("para.style", &json!({"style":style})).map_err(|e| e.to_string())?;
                }
            }
            session.doc.ensure_nonempty();
        }
        "replaceBlocks" => {
            require_detach(&session.doc, &node, args)?;
            edit(session, &json!({"operation":"insert","id":id,"position":"before","blocks":args["blocks"]}))?;
            edit(session, &json!({"operation":"delete","id":id,"allowDetach":args["allowDetach"]}))?;
        }
        "update" | "replaceText" => {
            if matches!(node.kind, "table" | "row" | "cell") {
                return Err("Update text using the paragraph IDs inside a cell; use command for table formatting".into());
            }
            let value = text(args)?;
            if value.contains(['\n', '\r']) {
                return Err("Use insert for multiple paragraphs".into());
            }
            let p = session.doc.para(node.story, &node.path).ok_or("Paragraph missing")?;
            let (start, end) = if args["operation"] == "replaceText" {
                find_range(p, args)?
            } else {
                (index(args, "start", 0)?, index(args, "end", p.text.chars().count())?)
            };
            let offsets: Vec<_> = p.text.char_indices().map(|(i, _)| i).chain(std::iter::once(p.text.len())).collect();
            if start > end || end >= offsets.len() {
                return Err("Invalid character range; read the block again".into());
            }
            if args["allowDetach"] != true && p.text[offsets[start]..offsets[end]].contains('\u{fffc}') {
                return Err("This text edit would remove inline objects or comment anchors. Use a smaller range or explicit replacement with allowDetach when requested.".into());
            }
            session.sel = Selection {
                anchor: Pos::new(node.story, node.path.clone(), offsets[start]),
                focus: Pos::new(node.story, node.path.clone(), offsets[end]),
            };
            session.run("text.insert", &json!({"text":value})).map_err(|e| e.to_string())?;
        }
        "delete" => {
            require_detach(&session.doc, &node, args)?;
            if matches!(node.kind, "row" | "cell") {
                return Err("Delete a paragraph or table; row/cell deletion is not exposed".into());
            }
            // Empty terminal paragraphs after tables (and in empty containers) are structural.
            // Removing them only regenerates a fresh ID and invites an agent deletion loop.
            if session.doc.para(node.story, &node.path).is_some_and(|p| p.text.is_empty()) {
                if let Some(container) = session.doc.container(node.story, &node.path) {
                    let index = node.path.last();
                    if index + 1 == container.len() && (index == 0 || matches!(&*container[index - 1], Block::Table(_))) {
                        return Err(
                            "This empty paragraph is required after a table or in an empty container. Keep it; removing it would only recreate it."
                                .into(),
                        );
                    }
                }
            }
            session.doc.remove_block(node.story, &node.path).map_err(|e| e.to_string())?;
        }
        "command" => {
            let command = args["command"].as_str().ok_or("command is required")?;
            if !super::allowed(command) || command.starts_with("edit.") || command.starts_with("caret.") || command.starts_with("select.") || command == "picture.select" {
                return Err("Command cannot be used as a targeted agent edit; use help for supported commands".into());
            }
            if command == "para.properties" && session.doc.para(node.story, &node.path).is_none() {
                return Err("Paragraph properties require a paragraph, heading or list-item ID".into());
            }
            let paragraphs = session.doc.para_paths(node.story);
            let paths: Vec<_> = paragraphs.iter().filter(|p| p.0.starts_with(&node.path.0)).collect();
            let first = paths.first().ok_or("Target has no editable paragraph")?;
            let last = paths.last().ok_or("Target has no editable paragraph")?;
            let len = session.doc.para(node.story, last).map_or(0, |p| p.text.len());
            session.sel = Selection { anchor: Pos::new(node.story, (*first).clone(), 0), focus: Pos::new(node.story, (*last).clone(), len) };
            if args.get("start").is_some() || args.get("end").is_some() {
                if first != last || !(command.starts_with("format.") || command.starts_with("picture.") || command.starts_with("arrange.") || command == "object.altText" || command == "references.caption") {
                    return Err("Ranges require one paragraph and a formatting or image command".into());
                }
                let paragraph = session.doc.para(node.story, first).ok_or("Paragraph missing")?;
                let offsets: Vec<_> = paragraph.text.char_indices().map(|(i, _)| i).chain(std::iter::once(paragraph.text.len())).collect();
                let start = index(args, "start", 0)?;
                let end = index(args, "end", paragraph.text.chars().count())?;
                let (Some(&start), Some(&end)) = (offsets.get(start), offsets.get(end)) else {
                    return Err("Invalid character range".into());
                };
                if start > end {
                    return Err("Invalid character range".into());
                }
                if !command.starts_with("format.") && (end != start + wordcraft_doc::para::OBJ.len_utf8() || !matches!(paragraph.object_at(start), Some(InlineObject::Image { .. }))) {
                    return Err("Image commands require the exact image start/end from a fresh read".into());
                }
                session.sel = Selection { anchor: Pos::new(node.story, (*first).clone(), start), focus: Pos::new(node.story, (*last).clone(), end) };
            }
            if command.starts_with("insert.") || command.starts_with("text.") {
                session.sel = Selection::caret(session.sel.focus.clone());
            }
            if (command.starts_with("picture.") || command.starts_with("arrange.") || command == "object.altText") && args.get("start").is_none() {
                return Err("Image commands require start/end from the target paragraph images in a fresh read".into());
            }
            command_result = session.run(command, args.get("params").unwrap_or(&json!({}))).map_err(|e| e.to_string())?;
        }
        _ => return Err("Unknown operation".into()),
    }
    session.clamp_selection();
    Ok(command_result)
}

fn find_range(paragraph: &Paragraph, args: &Value) -> Result<(usize, usize), String> {
    let find = args["find"]
        .as_str()
        .filter(|s| !s.is_empty() && s.len() <= 2000 && !s.contains('\u{fffc}'))
        .ok_or("find requires 1-2000 literal characters")?;
    let matches: Vec<_> = paragraph.text.match_indices(find).collect();
    if args.get("nth").is_none() && matches.len() > 1 {
        return Err("The quoted text occurs more than once. Specify nth after reading the target.".into());
    }
    let nth = index(args, "nth", 1)?;
    let (offset, _) = matches.get(nth.checked_sub(1).ok_or("nth starts at 1")?).ok_or("Quoted text was not found. Read the target again.")?;
    let start = paragraph.text[..*offset].chars().count();
    Ok((start, start + find.chars().count()))
}

pub fn comments(doc: &Document, args: &Value) -> Value {
    let all = nodes(doc);
    let offset = args["offset"].as_u64().unwrap_or(0) as usize;
    let limit = args["limit"].as_u64().unwrap_or(20).clamp(1, 20) as usize;
    let text_limit = args["textLimit"].as_u64().unwrap_or(160).clamp(1, 4000) as usize;
    let text_offset = args["textOffset"].as_u64().unwrap_or(0) as usize;
    let selected: Vec<_> = doc.comments.iter().filter(|(id, _)| args["commentId"].as_u64().is_none_or(|value| value == u64::from(**id))).collect();
    let values: Vec<_> = selected.iter().skip(offset).take(limit).map(|(id,comment)| {
        let mut root = **id;
        for _ in 0..doc.comments.len() { if let Some(parent)=doc.comments.get(&root).and_then(|c|c.parent) {root=parent;} else {break;} }
        let mut start: Option<Value> = None;
        let mut end: Option<Value> = None;
        for node in &all {
            if matches!(node.kind,"table"|"row"|"cell") {continue;}
            if let Some(p)=doc.para(node.story,&node.path) {
                for (off, _) in p.text.char_indices().filter(|(_,c)|*c=='\u{fffc}') {
                    match p.object_at(off) {
                        Some(InlineObject::CommentStart{id}) if *id==root => start=Some(json!({"id":node.id,"offset":p.text[..off].chars().count()+1})),
                        Some(InlineObject::CommentEnd{id}) if *id==root => end=Some(json!({"id":node.id,"offset":p.text[..off].chars().count()})),
                        _=>{}
                    }
                }
            }
        }
        let text=doc.plain_text(StoryRef::Part(comment.part));
        let length=text.chars().count();
        let slice:String=text.chars().skip(text_offset).take(text_limit.min(8000/limit)).collect();
        let next=text_offset+slice.chars().count();
        json!({"commentId":id,"parentId":comment.parent,"threadId":root,"author":comment.author,"resolved":comment.resolved,"text":slice,"textLength":length,"nextTextOffset":if next<length{Some(next)}else{None},"anchor":{"start":start,"end":end},"orphaned":start.is_none()||end.is_none()})
    }).collect();
    json!({"comments":values,"total":selected.len(),"nextOffset":if offset+values.len()<selected.len(){Some(offset+values.len())}else{None}})
}

fn edit_comment(session: &mut Session, args: &Value) -> Result<Value, String> {
    let action = args["action"].as_str().ok_or("Comment action is required")?;
    if !matches!(action, "add" | "reply" | "update" | "resolve" | "delete") {
        return Err("Unknown comment action".into());
    }
    if matches!(action, "add" | "reply" | "update") && args["text"].as_str().is_none_or(|text| text.is_empty() || text.len() > 10000) {
        return Err("Comment text requires 1-10000 characters".into());
    }
    let params = if action == "add" {
        let node = target(&session.doc, args["id"].as_str().ok_or("Read the comment's target paragraph first")?)?;
        let paragraph = session.doc.para(node.story, &node.path).ok_or("Comments target one paragraph")?;
        let (start, end) = if args.get("find").is_some() {
            find_range(paragraph, args)?
        } else {
            (index(args, "start", 0)?, index(args, "end", paragraph.text.chars().count())?)
        };
        let offsets: Vec<_> = paragraph.text.char_indices().map(|(i, _)| i).chain(std::iter::once(paragraph.text.len())).collect();
        let (Some(&start), Some(&end)) = (offsets.get(start), offsets.get(end)) else {
            return Err("Invalid comment anchor range".into());
        };
        if start > end {
            return Err("Invalid comment anchor range".into());
        }
        session.sel = Selection { anchor: Pos::new(node.story, node.path.clone(), start), focus: Pos::new(node.story, node.path, end) };
        json!({"text":args["text"]})
    } else {
        let id = args["commentId"].as_u64().and_then(|id| u32::try_from(id).ok()).ok_or("A commentId from the latest comment read is required")?;
        let comment = session.doc.comments.get(&id).ok_or("Comment no longer exists")?;
        if action == "reply" && comment.parent.is_some() {
            return Err("Reply to the root threadId, not to a reply".into());
        }
        if action == "resolve" && args["resolved"].as_bool().is_none() {
            return Err("resolved must be an explicit boolean".into());
        }
        json!({"id":id,"text":args["text"],"value":args["resolved"]})
    };
    let command = match action {
        "add" => "review.newComment",
        "reply" => "review.reply",
        "update" => "review.editComment",
        "resolve" => "review.resolveComment",
        _ => "review.deleteComment",
    };
    session.run(command, &params).map_err(|error| error.to_string())
}
