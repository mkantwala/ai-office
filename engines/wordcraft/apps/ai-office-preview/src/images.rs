//! Bounded, read-only image geometry. Native commands own edits and undo.
use serde_json::{Value,json};
use wordcraft_doc::{Pos,para::InlineObject};
use wordcraft_engine::{Session,cmd::objects};
use wordcraft_layout::{DocLayout,hit::ObjectHit};

pub fn details(doc: &wordcraft_doc::Document, pos: &Pos, o: &InlineObject) -> Option<Value> {
    let InlineObject::Image {w,h,alt,float,crop,..} = o else {return None};
    let p=doc.para_at(pos)?;
    let offset=p.text.get(..pos.off)?.chars().count();
    Some(json!({"parentId":p.props.office_id,"offset":offset,"width":w,"height":h,"alt":alt,"title":float.alt.title,"decorative":float.alt.decorative,"crop":crop,"wrap":float.wrap,"rotation":float.rot,"flipH":float.flip_h,"flipV":float.flip_v}))
}
fn geometry_one(s:&Session,hit:&ObjectHit)->Option<Value>{
    let p=s.doc.para_at(&hit.pos())?;
    let mut v=details(&s.doc,&hit.pos(),p.object_at(hit.off)?)?;
    let r=hit.rect;
    v["page"]=json!(hit.page);v["x"]=json!(r.x);v["y"]=json!(r.y);
    Some(v)
}
pub fn selected(s:&Session,l:&DocLayout)->Value{
    let Some((pos,_))=objects::object_selection(s) else{return Value::Null};
    l.object(&pos,s.page_hint).and_then(|hit|geometry_one(s,&hit)).unwrap_or(Value::Null)
}
pub fn geometry(s:&Session,l:&DocLayout,page:usize)->Value{
    let Some(p)=l.pages.get(page) else{return json!([])};
    json!(wordcraft_layout::hit::all_objects(p,page).filter(|hit|hit.story==s.sel.focus.story).filter_map(|hit|geometry_one(s,&hit)).take(200).collect::<Vec<_>>())
}
