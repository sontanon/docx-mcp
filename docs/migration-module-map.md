# Migration Module Map: Python -> Go and Rust

This is a planning map from current `src/docx_mcp/` modules to likely Go and Rust module boundaries.
It is intentionally high-level and focused on maintainable architecture, not one-to-one code translation.

## Mapping table

| Current Python module | Current responsibility | Go target module(s) | Rust target module(s) | Notes and migration risk |
| --- | --- | --- | --- | --- |
| `namespaces.py` | Namespace constants, `qn`, XPath helpers | `internal/ooxml/ns.go`, `internal/ooxml/query.go` | `src/ooxml/ns.rs`, `src/ooxml/query.rs` | Low risk. Keep all QName and namespace lookup logic centralized. |
| `document.py` | DOCX ZIP parse/write, XML part ownership | `internal/docx/package.go` | `src/docx/package.rs` | Medium risk. Must preserve unknown parts and only rewrite edited XML parts. |
| `models.py` | Change models, enums, validation | `internal/model/types.go` | `src/model/types.rs` | Low risk. Rust enums are a good fit; Go uses tagged structs/interfaces. |
| `id_manager.py` | Monotonic annotation IDs | `internal/redline/id_allocator.go` | `src/redline/id_allocator.rs` | Low risk. Straightforward port. |
| `tokenizer.py` | Word tokenization | `internal/diff/tokenize.go` | `src/diff/tokenize.rs` | Low risk; keep behavior identical to preserve diff stability. |
| `differ.py` | Word-level diff wrappers | `internal/diff/engine.go` | `src/diff/engine.rs` | Medium risk. Choose and pin a diff library, then lock behavior via tests. |
| `run_ops.py` | Run extraction/splitting/building, formatting preservation | `internal/ooxml/runops.go` | `src/ooxml/runops.rs` | High risk. This is one of the most correctness-sensitive modules. |
| `converter.py` | OOXML paragraph/table to pseudo-Markdown | `internal/convert/markdown.go` | `src/convert/markdown.rs` | Medium risk. Formatting and whitespace parity are key. |
| `comments.py` | Comment insertion, comments part/rels/content type wiring | `internal/redline/comments.go` | `src/redline/comments.rs` | High risk. Easy to produce files that trigger Word repair if incorrect. |
| `handlers/modify.py` | Paragraph modify redline | `internal/redline/modify.go` | `src/redline/modify.rs` | High risk. Depends on `run_ops` parity and whitespace behavior. |
| `handlers/delete.py` | Paragraph delete redline | `internal/redline/delete.go` | `src/redline/delete.rs` | Medium/high risk. Must preserve paragraph mark semantics and `w:delText` conversion. |
| `handlers/append.py` | Paragraph append with inherited formatting | `internal/redline/append.go` | `src/redline/append.rs` | High risk. Inheritance and insertion order are easy to get subtly wrong. |
| `table_utils.py` | Table shape validation and cell lookup | `internal/table/utils.go` | `src/table/utils.rs` | Medium risk. Keep safe-skip behavior explicit and deterministic. |
| `table_redliner.py` | Table cell redlining orchestration | `internal/table/redliner.go` | `src/table/redliner.rs` | Medium/high risk. Needs precise mutation ordering and validation. |
| `redliner.py` | Main orchestration and validation of changes | `internal/redline/pipeline.go` | `src/redline/pipeline.rs` | High risk. This is cross-cutting coordination; migrate after primitives are stable. |
| `validator.py` | Structural OOXML validation | `internal/validate/validate.go` | `src/validate/mod.rs` | High risk in impact, low complexity per rule. Keep as hard output gate. |
| `cli.py` | CLI commands and file I/O | `cmd/docx-mcp/main.go` + `internal/cli/*.go` | `src/bin/docx_mcp.rs` + `src/cli/*.rs` | Low/medium risk. Mostly integration and UX parity. |
| `server.py` | MCP server tools/resources | `internal/mcp/server.go` | `src/mcp/server.rs` | Medium risk. Protocol glue; easiest to migrate last. |

## Suggested migration order

1. `document` + `namespaces` + `models`
2. `tokenizer` + `differ`
3. `run_ops`
4. paragraph handlers (`modify`, `delete`, `append`)
5. `comments` + `validator`
6. table modules (`table_utils`, `table_redliner`)
7. `redliner` orchestration
8. CLI and MCP server

This order maximizes early testable parity for core XML mutation behavior.

## Rust-specific ownership mapping for core XML edits

This is the practical ownership plan for the Rust port:

- `DocxDocument` is owned by one worker and passed as `&mut DocxDocument` through apply pipeline.
- Avoid persistent `&mut` node references across function boundaries.
- Store stable target descriptors (fragment ID, path, body index), then re-resolve mutable nodes at execution time.
- Keep mutation scopes short: mutate one target, release borrow, then continue.
- Never mutate the same document tree from multiple threads.

Representative shape:

```rust
pub fn apply_redlines(doc: &mut DocxDocument, changes: &[Change]) -> anyhow::Result<()> {
    let plan = plan_operations(doc, changes)?; // immutable read pass
    for op in plan {
        execute_op(doc, &op)?; // short mutable borrow per op
    }
    validate_document(doc)?;
    Ok(())
}
```

## Go-specific concurrency mapping

Go keeps per-document concurrency simple:
- parse/mutate/write one document per goroutine,
- avoid shared mutable state except read-only config,
- bound workers by memory budget and document size.

Representative shape:

```go
func ProcessBatch(jobs []Job, workers int) []Result {
    // worker pool, one doc per worker task
    // each task owns its DocxDocument instance end-to-end
    return nil
}
```

## Coverage checklist for parity

Before considering migration complete, ensure parity in:
- fragment ID mapping for mixed paragraph/table bodies,
- run-level whitespace and formatting preservation,
- tracked change attrs (`w:id`, `w:author`, `w:date`) everywhere,
- comments.xml relationship and content-type wiring,
- safe skip handling for unsupported tables,
- unchanged ZIP entry pass-through fidelity.

