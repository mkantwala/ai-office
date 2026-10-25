//! Document review over private native snapshots; exported only to local version storage.
use serde_json::{Value, json};
use std::collections::BTreeMap;
use std::sync::Arc;
use wordcraft_doc::{Block, Blocks, Document};

#[derive(Default)]
pub struct Review {
    baseline: Option<Document>,
    proposed: Option<Document>,
}

fn id(block: &Block) -> &str {
    match block {
        Block::Para(p) => &p.props.office_id,
        Block::Table(t) => &t.props.office_id,
    }
}

fn roots(doc: &Document) -> BTreeMap<String, Arc<Block>> {
    let mut result = BTreeMap::new();
    for block in &doc.body {
        result.insert(format!("body:{}", id(block)), block.clone());
    }
    for (part, value) in &doc.parts {
        for block in &value.blocks {
            result.insert(format!("part{part}:{}", id(block)), block.clone());
        }
    }
    result
}

fn scrub(value: &mut Value) {
    match value {
        Value::Object(map) => {
            map.remove("officeId");
            for item in map.values_mut() {
                scrub(item);
            }
        }
        Value::Array(items) => {
            for item in items {
                scrub(item);
            }
        }
        _ => {}
    }
}

fn metadata(doc: &Document) -> Value {
    let mut value = json!(doc);
    if let Some(map) = value.as_object_mut() {
        map.remove("body");
        map.remove("core");
        if let Some(parts) = map.get_mut("parts").and_then(Value::as_object_mut) {
            for part in parts.values_mut() {
                if let Some(part) = part.as_object_mut() {
                    part.remove("blocks");
                }
            }
        }
    }
    scrub(&mut value);
    value
}

fn properties(value: Value) -> String {
    let mut value = value;
    scrub(&mut value);
    value
        .as_object()
        .map(|map| {
            map.iter()
                .filter(|(_, value)| !value.is_null())
                .map(|(key, value)| {
                    let name =
                        key.chars().flat_map(|c| if c.is_uppercase() { vec![' ', c.to_ascii_lowercase()] } else { vec![c] }).collect::<String>();
                    format!("{name}: {}", value.as_str().map(str::to_string).unwrap_or_else(|| value.to_string()))
                })
                .collect::<Vec<_>>()
                .join("; ")
        })
        .unwrap_or_default()
}

fn display(block: &Block) -> String {
    match block {
        Block::Para(p) => {
            let mut text = format!("{}\nParagraph formatting: {}\n", p.text, properties(json!(p.props)));
            let mut offset = 0;
            for run in &p.runs {
                let end = (offset + run.len).min(p.text.len());
                let format = properties(json!(run.props));
                if !format.is_empty() {
                    let content = p.text.get(offset..end).unwrap_or("");
                    text.push_str(&format!("Character formatting for {content:?}: {format}\n"));
                }
                offset = end;
            }
            if !p.objects.is_empty() {
                let mut objects = json!(p.objects);
                scrub(&mut objects);
                text.push_str(&format!("Inline objects, links and annotations: {objects}\n"));
            }
            text
        }
        Block::Table(table) => {
            let mut text = format!("Table formatting: {}\nColumn widths: {:?}\n", properties(json!(table.props)), table.grid);
            for (row_index, row) in table.rows.iter().enumerate() {
                text.push_str(&format!("Row {} formatting: {}\n", row_index + 1, properties(json!(row.props))));
                for (cell_index, cell) in row.cells.iter().enumerate() {
                    text.push_str(&format!("Row {}, column {}\n", row_index + 1, cell_index + 1));
                    text.push_str(&format!("Cell formatting: {}\n", properties(json!(cell.props))));
                    for block in &cell.blocks {
                        text.push_str(&display(block));
                    }
                }
            }
            text
        }
    }
}

fn plain_text(block: &Block) -> String {
    match block {
        Block::Para(p) => p.text.clone(),
        Block::Table(t) => t
            .rows
            .iter()
            .map(|r| r.cells.iter().map(|c| c.blocks.iter().map(|b| plain_text(b)).collect::<Vec<_>>().join("\n")).collect::<Vec<_>>().join(" | "))
            .collect::<Vec<_>>()
            .join("\n"),
    }
}

fn ordered_roots(doc: &Document) -> Vec<String> {
    doc.body
        .iter()
        .map(|b| format!("body:{}", id(b)))
        .chain(doc.parts.iter().flat_map(|(part, value)| value.blocks.iter().map(move |b| format!("part{part}:{}", id(b)))))
        .collect()
}

fn anchor(base: &Document, after: &Document, key: &str) -> Value {
    let current = roots(after);
    if let Some(block) = current.get(key) {
        return json!({"id":id(block),"after":false});
    }
    let ordered = ordered_roots(base);
    if let Some(index) = ordered.iter().position(|k| k == key) {
        let story = key.split(':').next();
        for next in &ordered[index + 1..] {
            if next.split(':').next() == story {
                if let Some(block) = current.get(next) {
                    return json!({"id":id(block),"after":false});
                }
            }
        }
        for previous in ordered[..index].iter().rev() {
            if previous.split(':').next() == story {
                if let Some(block) = current.get(previous) {
                    return json!({"id":id(block),"after":true});
                }
            }
        }
    }
    json!(null)
}

fn document_text(doc: &Document) -> String {
    let blocks = roots(doc);
    let mut text = ordered_roots(doc).iter().filter_map(|key| blocks.get(key).map(|block| display(block))).collect::<Vec<_>>().join("\n");
    text.push_str(&format!("\nDocument formatting and annotations:\n{}\n", metadata(doc)));
    text
}

fn copy_root(target: &mut Document, source: &Document, key: &str) -> Result<(), String> {
    let (story, block_id) = key.split_once(':').ok_or("Invalid review change")?;
    let (target_blocks, source_blocks): (&mut Blocks, &Blocks) = if story == "body" {
        (&mut target.body, &source.body)
    } else {
        let part: u32 = story.strip_prefix("part").ok_or("Invalid story")?.parse().map_err(|_| "Invalid story")?;
        (
            &mut target.parts.get_mut(&part).ok_or("Document structure changed")?.blocks,
            &source.parts.get(&part).ok_or("Document structure changed")?.blocks,
        )
    };
    let old_position = target_blocks.iter().position(|block| id(block) == block_id);
    let replacement = source_blocks.iter().position(|block| id(block) == block_id);
    match (old_position, replacement) {
        (Some(position), Some(index)) => target_blocks[position] = source_blocks[index].clone(),
        (Some(position), None) => {
            target_blocks.remove(position);
        }
        (None, Some(index)) => {
            let next = source_blocks[index + 1..].iter().find_map(|block| target_blocks.iter().position(|candidate| id(candidate) == id(block)));
            let position = next.unwrap_or(target_blocks.len());
            target_blocks.insert(position, source_blocks[index].clone());
        }
        (None, None) => {}
    }
    Ok(())
}

impl Review {
    pub fn refresh(&mut self, doc: &Document) {
        let (Some(base), Some(after)) = (&self.baseline, &self.proposed) else {
            return;
        };
        if metadata(doc) == metadata(base) && roots(doc) == roots(base) {
            self.baseline = None;
            self.proposed = None;
            return;
        }
        if metadata(base) != metadata(after) {
            return;
        }
        let old = roots(base);
        let proposed = roots(after);
        let current = roots(doc);
        let mut next = after.clone();
        for key in old.keys().chain(proposed.keys()).collect::<std::collections::BTreeSet<_>>() {
            if old.get(key) != proposed.get(key) && current.get(key) == old.get(key) {
                let _ = copy_root(&mut next, base, key);
            }
        }
        self.proposed = Some(next);
    }
    pub fn start(&mut self, doc: &Document) -> Result<(), String> {
        self.refresh(doc);
        if let (Some(base), Some(proposed)) = (&mut self.baseline, &self.proposed) {
            let before = roots(base);
            let after = roots(proposed);
            let current = roots(doc);
            if metadata(proposed) != metadata(doc) {
                return Err("Document changed since the pending review. Keep or resolve its changes first.".into());
            }
            for key in before.keys().chain(after.keys()).chain(current.keys()).collect::<std::collections::BTreeSet<_>>() {
                if before.get(key) != after.get(key) && current.get(key) != after.get(key) {
                    return Err("A pending agent change was edited manually. Keep it before asking the agent to edit again.".into());
                }
                if before.get(key) == after.get(key) && current.get(key) != after.get(key) {
                    copy_root(base, doc, key)?;
                }
            }
        } else {
            self.baseline = Some(doc.clone());
        }
        self.proposed = Some(doc.clone());
        Ok(())
    }

    pub fn finish(&mut self, doc: &Document) {
        self.proposed = Some(doc.clone());
    }

    pub fn read(&self) -> Value {
        let (Some(base), Some(after)) = (&self.baseline, &self.proposed) else {
            return json!({"changes":[]});
        };
        if metadata(base) != metadata(after) {
            return json!({"changes":[{"id":"document","label":"Document structure, formatting and annotations","before":document_text(base),"after":document_text(after),"beforeText":"","afterText":"","anchor":ordered_roots(after).first().map(|key|anchor(base,after,key))}]});
        }
        let original = roots(base);
        let proposed = roots(after);
        // Group only uninterrupted inserted/deleted runs in a single story. Existing
        // blocks are semantic boundaries; paragraph IDs and native decisions stay intact.
        let mut groups = BTreeMap::new();
        for (doc, other, kind) in [(after, &original, "insert"), (base, &proposed, "delete")] {
            let mut run: Option<(String, String)> = None;
            for key in ordered_roots(doc) {
                if other.contains_key(&key) {
                    run = None;
                    continue;
                }
                let story = key.split(':').next().unwrap_or("").to_string();
                if run.as_ref().is_none_or(|(previous, _)| previous != &story) {
                    run = Some((story, format!("{kind}:{key}")));
                }
                groups.insert(key, run.as_ref().unwrap().1.clone());
            }
        }
        let mut keys = ordered_roots(base);
        for key in ordered_roots(after) {
            if !keys.contains(&key) {
                keys.push(key);
            }
        }
        let changes: Vec<_> = keys.iter().filter(|key| original.get(*key) != proposed.get(*key)).map(|key| json!({
            "id":key,"groupId":groups.get(key),"label":if proposed.get(key).or_else(||original.get(key)).is_some_and(|block|matches!(block.as_ref(),Block::Table(_))) {"Table change"} else {"Paragraph change"},
            "before":original.get(key).map(|block|display(block)).unwrap_or_default(),"after":proposed.get(key).map(|block|display(block)).unwrap_or_default(),
            "beforeText":original.get(key).map(|block|plain_text(block)).unwrap_or_default(),"afterText":proposed.get(key).map(|block|plain_text(block)).unwrap_or_default(),"anchor":anchor(base,after,key)
        })).collect();
        json!({"changes":changes})
    }

    pub fn export(&self) -> String {
        json!({"schema":1,"baseline":self.baseline,"proposed":self.proposed}).to_string()
    }
    pub fn import(&mut self, value: &str) -> Result<(), String> {
        let data: Value = serde_json::from_str(value).map_err(|e| e.to_string())?;
        if data["schema"] != 1 {
            return Err("Unsupported document review version".into());
        }
        let baseline: Option<Document> = serde_json::from_value(data["baseline"].clone()).map_err(|e| e.to_string())?;
        let proposed: Option<Document> = serde_json::from_value(data["proposed"].clone()).map_err(|e| e.to_string())?;
        if baseline.is_some() != proposed.is_some() {
            return Err("Incomplete document review version".into());
        }
        self.baseline = baseline;
        self.proposed = proposed;
        Ok(())
    }
    pub fn projection(doc: &Document) -> String {
        document_text(doc)
    }

    pub fn decide(&mut self, doc: &Document, action: &str, keys: &[String]) -> Result<Document, String> {
        if action != "keep" && action != "undo" {
            return Err("Unknown review decision".into());
        }
        let (Some(base), Some(after)) = (&self.baseline, &self.proposed) else {
            return Err("No pending changes".into());
        };
        let available = self.read();
        let changes = available["changes"].as_array().ok_or("Review unavailable")?;
        if keys.is_empty() || keys.iter().any(|key| !changes.iter().any(|change| change["id"].as_str() == Some(key))) {
            return Err("Review changed. Refresh the differences before deciding.".into());
        }
        let mut next_base = base.clone();
        let mut next_after = after.clone();
        let mut result = doc.clone();
        for key in keys {
            if key == "document" {
                if action == "undo" {
                    if doc != after {
                        return Err("Document changed since review. Undo would overwrite later edits. Keep it or resolve the conflict first.".into());
                    }
                    result = base.clone();
                    result.core = doc.core.clone();
                    next_after = result.clone();
                } else {
                    next_base = after.clone();
                }
            } else {
                if action == "undo" {
                    if roots(doc).get(key) != roots(after).get(key) {
                        return Err("This paragraph or table changed since review. Undo would overwrite your edits.".into());
                    }
                    copy_root(&mut result, base, key)?;
                    copy_root(&mut next_after, base, key)?;
                } else {
                    copy_root(&mut next_base, after, key)?;
                }
            }
        }
        self.baseline = Some(next_base);
        self.proposed = Some(next_after);
        if self.read()["changes"].as_array().is_some_and(Vec::is_empty) {
            self.baseline = None;
            self.proposed = None;
        }
        Ok(result)
    }
}
