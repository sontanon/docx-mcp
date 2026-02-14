# lxml API Usage in docx-mcp

This document catalogs the lxml APIs used in this codebase, with concrete examples
and file references. lxml is a Python binding for libxml2/libxslt (C libraries),
providing fast XML parsing and manipulation.

## 1. Namespace Handling

OOXML (Office Open XML) uses XML namespaces extensively. This codebase abstracts
namespace handling through `src/docx_mcp/namespaces.py`.

### Namespace Constants

Namespace URIs are defined as module-level constants:

```python
# src/docx_mcp/namespaces.py:9-35
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML = "http://www.w3.org/XML/1998/namespace"
# ... etc
```

### NSMAP - Namespace Map for XPath

A dictionary mapping prefixes to URIs, required for XPath queries:

```python
# src/docx_mcp/namespaces.py:38-44
NSMAP: dict[str, str] = {
    "w": W,
    "r": R,
    "wp": WP,
    "mc": MC,
}
```

### qn() - Clark Notation Qualified Names

Builds qualified names in Clark notation: `{namespace-uri}localname`.

```python
# src/docx_mcp/namespaces.py:47-57
def qn(ns_prefix: str, local_name: str) -> str:
    ns_uri = NSMAP.get(ns_prefix)
    if ns_uri is None:
        msg = f"Unknown namespace prefix: {ns_prefix!r}"
        raise ValueError(msg)
    return f"{{{ns_uri}}}{local_name}"
```

**Usage examples:**

```python
# src/docx_mcp/handlers/delete.py:50
del_el = etree.Element(qn("w", "del"))

# src/docx_mcp/document.py:153
el.get(qn("w", "id"))

# src/docx_mcp/run_ops.py:486
text_tag = qn("w", "delText") if is_delete else qn("w", "t")
```

### make_element() - Element Creation Helper

Creates elements with namespaced attributes using double-underscore syntax:

```python
# src/docx_mcp/namespaces.py:60-85
def make_element(ns_prefix: str, local_name: str, **attribs: str) -> etree._Element:
    tag = qn(ns_prefix, local_name)
    el = etree.Element(tag)
    
    for key, value in attribs.items():
        if "__" in key:
            # "w__id" -> "{W}id"
            prefix, attr_local = key.split("__", 1)
            attr_qn = qn(prefix, attr_local)
            el.set(attr_qn, value)
        elif key == "xml_space":
            el.set(f"{{{XML}}}space", value)
        else:
            el.set(key, value)
    return el
```

**Usage:**

```python
# src/docx_mcp/handlers/delete.py:50-53
del_el = make_element(
    "w", "del",
    w__id=str(del_id),
    w__author=config.author,
    w__date=config.date,
)
```

---

## 2. Element Creation

### etree.Element() - Create Standalone Element

Creates a new element not attached to any tree:

```python
# src/docx_mcp/handlers/delete.py:50
del_el = etree.Element(qn("w", "del"))

# src/docx_mcp/handlers/append.py:238
new_p = etree.Element(qn("w", "p"))

# src/docx_mcp/comments.py:52
comments_root = etree.Element(qn("w", "comments"), nsmap={"w": W})
```

The `nsmap` parameter adds namespace declarations to the serialized output.

### etree.SubElement() - Create and Attach in One Call

Creates an element and immediately attaches it to a parent:

```python
# src/docx_mcp/handlers/append.py:182
etree.SubElement(result, qn("w", "b"))      # <w:b/>

# src/docx_mcp/run_ops.py:487
t_el = etree.SubElement(r_el, text_tag)

# src/docx_mcp/comments.py:97
comment_el = etree.SubElement(comments_root, qn("w", "comment"))
```

---

## 3. Tree Manipulation

### .append() - Add as Last Child

Adds an element as the last child of a parent. Importantly, if the element
already has a parent, it's automatically removed from that parent first (move semantics).

```python
# src/docx_mcp/handlers/modify.py:142, 146
paragraph.append(new_child)

# src/docx_mcp/handlers/delete.py:63
for run in runs:
    del_el.append(run)  # Moves runs into <w:del> wrapper

# src/docx_mcp/comments.py:169
paragraph.append(comment_ref_run)
```

### .insert() - Insert at Specific Index

Inserts an element at a specific position among children:

```python
# src/docx_mcp/handlers/delete.py:106
paragraph.insert(0, ppr)  # Insert pPr at the start
```

### .addprevious() - Insert as Preceding Sibling

Inserts an element immediately before the current element:

```python
# src/docx_mcp/handlers/delete.py:57
first_run.addprevious(del_el)  # Insert <w:del> before first run

# src/docx_mcp/handlers/delete.py:84
t_el.addprevious(dt)  # Replace <w:t> with <w:delText> by inserting before

# src/docx_mcp/comments.py:145
first.addprevious(comment_range_start)  # Insert comment marker
```

### .addnext() - Insert as Following Sibling

Inserts an element immediately after the current element:

```python
# src/docx_mcp/handlers/append.py:91
anchor.addnext(new_p)  # Insert new paragraph after anchor

# src/docx_mcp/comments.py:149
last.addnext(comment_ref_run)  # Insert comment reference after last char
```

### .remove() - Remove Child Element

Removes a child from its parent:

```python
# src/docx_mcp/handlers/modify.py:138
paragraph.remove(child)

# src/docx_mcp/handlers/delete.py:85
t_el.getparent().remove(t_el)

# src/docx_mcp/run_ops.py:526
parent.remove(element)
```

---

## 4. XPath Queries

### xpath() Helper

The codebase wraps lxml's xpath with the namespace map:

```python
# src/docx_mcp/namespaces.py:88-90
def xpath(element: etree._Element, expr: str) -> list[etree._Element]:
    return element.xpath(expr, namespaces=NSMAP)
```

### Common XPath Patterns

| Pattern | File | Purpose |
|---------|------|---------|
| `.//w:body` | `document.py:97` | Find body anywhere in tree |
| `./w:p` | `document.py:110` | Direct child paragraphs |
| `w:r` | `run_ops.py:81` | Child runs (shorthand) |
| `w:pPr` | `modify.py:121` | Paragraph properties |
| `w:rPr` | `run_ops.py:91` | Run properties |
| `w:pPr/w:rPr` | `append.py:145` | Nested property lookup |
| `w:t` | `delete.py:76` | Text elements |

**Example usage:**

```python
# src/docx_mcp/document.py:97
body = xpath(doc_root, ".//w:body")

# src/docx_mcp/document.py:110
paragraphs = xpath(body, "./w:p")

# src/docx_mcp/handlers/append.py:61
ref_ppr = xpath(ref_p, "w:pPr")
```

---

## 5. Tree Iteration

### .iter() - Walk the Tree

Iterates over all elements (optionally filtered by tag):

```python
# src/docx_mcp/document.py:152 - Scan all elements for w:id
for el in self.document_tree.iter():
    aid = el.get(qn("w", "id"))

# src/docx_mcp/validator.py:151 - Filter by tag
for comment_el in doc.comments_tree.iter(qn("w", "comment")):
    comment_id = comment_el.get(qn("w", "id"))
```

---

## 6. Element Attributes & Text

### .get() / .set() - Attribute Access

```python
# Getting attributes
# src/docx_mcp/document.py:153
aid = el.get(qn("w", "id"))

# src/docx_mcp/converter.py:33
val = el.get(qn("w", "val"))

# src/docx_mcp/comments.py:192 - Non-namespaced attribute
rel_type = rel.get("Type")

# Setting attributes
# src/docx_mcp/handlers/delete.py:51-53
del_el.set(qn("w", "id"), str(del_id))
del_el.set(qn("w", "author"), config.author)

# src/docx_mcp/run_ops.py:492 - xml:space attribute
t_el.set(f"{{{XML}}}space", "preserve")
```

### .text - Element Text Content

```python
# Getting text
# src/docx_mcp/converter.py:155
if child.text:
    result.append(child.text)

# Setting text
# src/docx_mcp/run_ops.py:488
t_el.text = text

# src/docx_mcp/handlers/delete.py:79 - Copy text between elements
del_text_el.text = original_t_el.text
```

---

## 7. Tag Inspection

### .tag - Qualified Tag Name

Returns the Clark-notation tag (`{namespace}localname`):

```python
# src/docx_mcp/validator.py:223
if el.tag not in tracked_change_tags:
    continue
```

### etree.QName() - Extract Local Name

Common pattern when iterating children:

```python
# src/docx_mcp/handlers/modify.py:129
tag_local = etree.QName(child.tag).localname if isinstance(child.tag, str) else ""

# src/docx_mcp/converter.py:154
tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else None
```

The `isinstance(child.tag, str)` check is necessary because comments and
processing instructions have non-string tags (tuples in lxml).

---

## 8. Navigation

### .getparent() - Get Parent Element

```python
# src/docx_mcp/run_ops.py:524
parent = element.getparent()

# src/docx_mcp/handlers/delete.py:85
t_el.getparent().remove(t_el)
```

### .getnext() - Get Next Sibling

```python
# src/docx_mcp/redliner.py:293
next_el = current.getnext()
```

---

## 9. Parsing & Serialization

### etree.fromstring() - Parse XML Bytes

```python
# src/docx_mcp/document.py:71
self._document_tree = etree.fromstring(doc_xml)

# src/docx_mcp/document.py:76
self._comments_tree = etree.fromstring(comments_xml)
```

### etree.tostring() - Serialize to XML Bytes

```python
# src/docx_mcp/document.py:185-187
etree.tostring(
    self._document_tree,
    xml_declaration=True,
    encoding="UTF-8",
    standalone=True,
)

# src/docx_mcp/run_ops.py:438 - Structural comparison
if etree.tostring(a) == etree.tostring(b):
    return True
```

---

## 10. Deep Copying

### copy.deepcopy() - Clone Elements

Elements must be deep-copied to duplicate them independently:

```python
# src/docx_mcp/run_ops.py:111
return copy.deepcopy(rpr)

# src/docx_mcp/handlers/append.py:62
base_ppr = copy.deepcopy(ref_ppr[0])
```

Standard Python `copy` module works because lxml elements implement
`__deepcopy__` hooks.

---

## 11. DOCX/ZIP Handling

The `DocxDocument` class (`src/docx_mcp/document.py`) handles .docx files:

**Parsing:**

```python
# src/docx_mcp/document.py:60-86
def _parse_zip(self, raw: bytes) -> None:
    with zipfile.ZipFile(io.BytesIO(raw)) as zf:
        for entry in zf.namelist():
            self._zip_entries[entry] = zf.read(entry)
    
    doc_xml = self._zip_entries.get("word/document.xml")
    self._document_tree = etree.fromstring(doc_xml)
```

**Serialization:**

```python
# src/docx_mcp/document.py:172-208
def to_bytes(self) -> bytes:
    entries = dict(self._zip_entries)
    
    # Re-serialize modified XML trees
    entries["word/document.xml"] = etree.tostring(
        self._document_tree, xml_declaration=True, encoding="UTF-8"
    )
    
    # Write back to ZIP
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for path, data in entries.items():
            zf.writestr(path, data)
    return buf.getvalue()
```

---

## Quick Reference

| Operation | API | Example Location |
|-----------|-----|------------------|
| Create element | `etree.Element(qn("w", "p"))` | `append.py:238` |
| Create & attach | `etree.SubElement(parent, qn("w", "r"))` | `run_ops.py:487` |
| Get attribute | `el.get(qn("w", "id"))` | `document.py:153` |
| Set attribute | `el.set(qn("w", "author"), "AI")` | `delete.py:52` |
| Get text | `el.text` | `converter.py:155` |
| Set text | `el.text = "Hello"` | `run_ops.py:488` |
| Append child | `parent.append(child)` | `modify.py:142` |
| Insert at index | `parent.insert(0, child)` | `delete.py:106` |
| Remove child | `parent.remove(child)` | `modify.py:138` |
| Insert before | `el.addprevious(new)` | `delete.py:57` |
| Insert after | `el.addnext(new)` | `append.py:91` |
| Get parent | `el.getparent()` | `run_ops.py:524` |
| Get next sibling | `el.getnext()` | `redliner.py:293` |
| XPath query | `el.xpath("w:r", namespaces=NSMAP)` | `namespaces.py:88` |
| Iterate tree | `el.iter()` | `document.py:152` |
| Parse XML | `etree.fromstring(bytes)` | `document.py:71` |
| Serialize XML | `etree.tostring(el)` | `document.py:185` |
| Get local tag | `etree.QName(el.tag).localname` | `modify.py:129` |
| Deep copy | `copy.deepcopy(el)` | `run_ops.py:111` |