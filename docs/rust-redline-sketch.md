# Rust Redliner Sketch

This document sketches how a Rust version of `docx-mcp` can be structured:
- preserve all package parts,
- mutate only WordprocessingML parts,
- generate tracked changes + comments,
- skip unsupported targets safely.

It focuses on critical patterns and intentionally omits full implementation detail.

## Dependency profile

Two realistic Rust stacks:

1. `libxml`-backed (closest to `lxml` behavior)
- `zip` (OPC package read/write)
- `libxml` (DOM + XPath over libxml2)

2. pure Rust (minimal native deps)
- `zip`
- `quick-xml` (stream parse/write) and/or `xmltree` (editable tree)

If XPath parity is important, the `libxml` path is usually the closest translation.

## Mutability and borrow checker (critical)

This is the main Rust-specific design concern for this project.

Short answer to your question:
- You usually hold `&mut DocxDocument` while applying a batch of changes.
- You do **not** safely hold many long-lived `&mut` references to multiple XML nodes at once in pure Rust trees.
- You should use a two-phase approach: discover targets first, mutate second.

### Why this matters

In this codebase, handlers often need parent/child/sibling edits (`insert before/after`, move runs, wrap runs).
In Rust tree APIs, a mutable borrow into one node commonly blocks other borrows of the same tree until that borrow ends.

### Recommended pattern: two-phase mutation

1. Walk tree with shared/immutable access and collect stable target descriptors (IDs, paths, indexes).
2. Re-enter with mutable access and apply one mutation at a time.

```rust
pub struct ParaTarget {
    pub fragment_id: usize,
    pub child_index_in_body: usize,
}

pub fn collect_targets(doc: &DocxDocument) -> Vec<ParaTarget> {
    // Immutable read pass only.
    // Build stable references by index/path, not by storing &mut node refs.
    vec![]
}

pub fn apply_changes(doc: &mut DocxDocument, targets: &[ParaTarget]) -> anyhow::Result<()> {
    for t in targets {
        // Borrow mutably only for this mutation scope.
        apply_one_change(doc, t)?;
    }
    Ok(())
}
```

### Practical patterns that avoid borrow fights

Pattern A: index/path addressing
- Store `Vec<usize>` path from root to node, or body child index.
- Re-resolve node each time right before mutation.

Pattern B: detach-modify-reattach
- Remove target subtree from parent.
- Mutate detached subtree freely.
- Insert back at known index.

Pattern C: operation log
- Convert user changes into an ordered list of concrete edit ops.
- Execute ops sequentially against `&mut DocxDocument`.

### `libxml` backend nuance

With `libxml` bindings, node handles are often pointer-like wrappers over C structures.
Borrow checker pressure can feel lighter than pure Rust trees because mutability is largely managed by the C library model.
But then correctness is more runtime-discipline-driven:
- avoid keeping stale node handles after structural edits,
- avoid cross-thread mutation of the same document tree.

### Pure Rust backend nuance (`xmltree` style)

Borrow checker is stricter and explicit:
- you cannot keep a mutable borrow to one child while also mutating siblings/parent,
- you should keep mutable borrows short and local,
- path/index re-resolution is the most reliable pattern.

### Minimal shape for redliner API

```rust
pub fn apply_redlines(
    doc: &mut DocxDocument,
    changes: &[Change],
) -> anyhow::Result<ApplySummary> {
    let plan = plan_operations(doc, changes)?; // immutable pass
    execute_operations(doc, &plan)?;           // mutable pass
    validate_document(doc)?;                   // immutable pass
    Ok(ApplySummary::from(plan))
}
```

This keeps ownership and mutability explicit and predictable.

## 1. Core document model: preserve unknown ZIP entries

```rust
use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};
use zip::{ZipArchive, ZipWriter};

pub struct DocxDocument {
    pub parts: BTreeMap<String, Vec<u8>>,
    pub document_xml: XmlPart,
    pub comments_xml: Option<XmlPart>,
    pub content_types_xml: Option<XmlPart>,
    pub rels_xml: Option<XmlPart>,
}

pub enum XmlPart {
    // Pick one representation depending on stack:
    Libxml(libxml::tree::Document),
    XmlTree(xmltree::Element),
}

pub fn parse_docx(raw: &[u8]) -> anyhow::Result<DocxDocument> {
    let mut archive = ZipArchive::new(Cursor::new(raw))?;
    let mut parts = BTreeMap::new();

    for i in 0..archive.len() {
        let mut f = archive.by_index(i)?;
        let mut buf = Vec::new();
        f.read_to_end(&mut buf)?;
        parts.insert(f.name().to_string(), buf);
    }

    let doc = parts
        .get("word/document.xml")
        .ok_or_else(|| anyhow::anyhow!("missing word/document.xml"))?;

    // Parse XML part via selected backend.
    let document_xml = parse_xml_part(doc)?;
    let comments_xml = parts.get("word/comments.xml").map(parse_xml_part).transpose()?;
    let content_types_xml = parts
        .get("[Content_Types].xml")
        .map(parse_xml_part)
        .transpose()?;
    let rels_xml = parts
        .get("word/_rels/document.xml.rels")
        .map(parse_xml_part)
        .transpose()?;

    Ok(DocxDocument {
        parts,
        document_xml,
        comments_xml,
        content_types_xml,
        rels_xml,
    })
}
```

Pattern:
- package fidelity first,
- explicit editable XML part set,
- everything else is pass-through bytes.

## 2. Namespace and query layer

Create one small internal query API and keep all namespace details there.

### `libxml`-style (schematic)

```rust
pub const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn find_body(doc: &libxml::tree::Document) -> anyhow::Result<libxml::tree::Node> {
    let mut ctx = libxml::xpath::Context::new(doc)?;
    ctx.register_namespace("w", W_NS)?;
    let nodes = ctx.findnodes("//w:body", None)?;
    nodes
        .into_iter()
        .next()
        .ok_or_else(|| anyhow::anyhow!("missing w:body"))
}
```

### pure-Rust traversal helper

```rust
fn child_by_local_name<'a>(el: &'a xmltree::Element, local: &str) -> Option<&'a xmltree::Element> {
    el.children.iter().find_map(|n| match n {
        xmltree::XMLNode::Element(child) if child.name.ends_with(local) => Some(child),
        _ => None,
    })
}
```

Important design choice:
- avoid sprinkling XPath strings all over handlers,
- map high-level operations to typed helper functions.

## 3. Interleaved fragment ID map

Same pattern as current Python behavior:

```rust
pub struct FragmentRef<'a> {
    pub id: usize,
    pub node: &'a mut XmlNodeRef,
}

pub fn interleaved_element_ids(body_children: &[XmlNodeRef]) -> Vec<usize> {
    body_children
        .iter()
        .enumerate()
        .filter_map(|(i, n)| {
            let local = n.local_name();
            if local == "p" || local == "tbl" {
                Some(i + 1)
            } else {
                None
            }
        })
        .collect()
}
```

Keep IDs 1-based and stable in document order.

## 4. Tracked-change element builders

Always create wrappers through one constructor so required attrs are never forgotten.

```rust
pub struct TcMeta<'a> {
    pub id: u32,
    pub author: &'a str,
    pub date_iso: &'a str,
}

pub fn make_tracked_change(local: &str, meta: &TcMeta) -> XmlNodeRef {
    let mut el = XmlNodeRef::new_ns("w", local, W_NS);
    el.set_attr_ns("w", "id", W_NS, meta.id.to_string());
    el.set_attr_ns("w", "author", W_NS, meta.author.to_string());
    el.set_attr_ns("w", "date", W_NS, meta.date_iso.to_string());
    el
}
```

Usage pattern:
- `modify`: diff paragraph text, emit equal/insert/delete runs with formatting inheritance.
- `delete`: wrap existing run content in `w:del` and convert `w:t` to `w:delText`.
- `append_after`: construct new `w:p` with `w:ins` plus inherited `w:rPr`.

## 5. Comment part creation + package link updates

Replicate your existing behavior in a single helper:

```rust
pub fn ensure_comments_part(doc: &mut DocxDocument) -> anyhow::Result<()> {
    if doc.comments_xml.is_none() {
        doc.comments_xml = Some(new_comments_root()?);          // word/comments.xml
        add_comments_relationship(doc.rels_xml.as_mut())?;      // document.xml.rels
        add_comments_content_type(doc.content_types_xml.as_mut())?; // [Content_Types].xml
    }
    Ok(())
}
```

This is one of the most failure-prone areas; keep it centralized.

## 6. Safe skip behavior for unsupported targets

Treat unsupported shapes as explicit non-fatal outcomes.

```rust
pub enum ApplyOutcome {
    Applied,
    Skipped { reason: String },
}

pub fn apply_table_change(cell: &XmlNodeRef, change: &TableChange) -> ApplyOutcome {
    if has_grid_span(cell) || has_vmerge(cell) || has_nested_table(cell) {
        return ApplyOutcome::Skipped {
            reason: "unsupported merged/nested table structure".to_string(),
        };
    }
    // normal mutation path...
    ApplyOutcome::Applied
}
```

Policy:
- skip only targeted unsupported elements,
- keep the rest of the document untouched and serializable,
- report skip reasons back to caller/tool layer.

## 7. Repack document bytes

```rust
pub fn to_bytes(doc: &mut DocxDocument) -> anyhow::Result<Vec<u8>> {
    let mut out = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(&mut out);
    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    // Clone original parts, overwrite edited XML parts.
    let mut parts = doc.parts.clone();
    parts.insert("word/document.xml".into(), serialize_xml(&doc.document_xml)?);
    if let Some(ref comments) = doc.comments_xml {
        parts.insert("word/comments.xml".into(), serialize_xml(comments)?);
    }
    if let Some(ref ct) = doc.content_types_xml {
        parts.insert("[Content_Types].xml".into(), serialize_xml(ct)?);
    }
    if let Some(ref rels) = doc.rels_xml {
        parts.insert("word/_rels/document.xml.rels".into(), serialize_xml(rels)?);
    }

    for (name, payload) in parts {
        zip.start_file(name, options)?;
        zip.write_all(&payload)?;
    }
    zip.finish()?;
    Ok(out.into_inner())
}
```

## 8. Parallel document processing

A per-document job model is the safest concurrency boundary.

### Rayon worker pattern

```rust
use rayon::prelude::*;

pub fn process_batch(inputs: Vec<Job>) -> Vec<Result<Vec<u8>, anyhow::Error>> {
    inputs
        .into_par_iter()
        .map(|job| apply_redlines(&job.doc_bytes, &job.changes))
        .collect()
}
```

### Tokio service pattern

```rust
// Keep CPU-heavy XML work on a bounded blocking pool.
let out = tokio::task::spawn_blocking(move || apply_redlines(&doc, &changes)).await??;
```

Guidance:
- bound concurrency by memory budget,
- per-job timeout and max input size,
- never share mutable DOM nodes across threads.

## 9. Validation gates before output

Before returning bytes, enforce:
- all `w:ins`/`w:del` have `w:id`, `w:author`, `w:date`,
- no cross-group ID collisions (comments vs tracked changes),
- if comments part exists, content type + rel entries exist.

If validation fails, return structured error instead of emitting potentially repair-triggering DOCX.
