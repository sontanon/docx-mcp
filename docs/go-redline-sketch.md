# Go Redliner Sketch

This document shows a practical Go shape for `docx-mcp` style redlining:
- edit OOXML XML directly inside a `.docx` ZIP package,
- preserve unknown parts untouched,
- apply tracked changes with comments,
- skip unsupported structures safely.

The snippets are intentionally partial and focus on critical patterns.

## Dependency profile

Conservative, well-known choices:
- stdlib: `archive/zip`, `encoding/xml`, `bytes`, `io`, `sync`
- XML tree: `github.com/beevik/etree`
- optional XPath convenience: `github.com/antchfx/xmlquery` + `github.com/antchfx/xpath`

If you want the smallest dependency surface, use only `etree` and avoid full XPath.

## 1. Core document model: keep all ZIP parts

```go
package docx

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"

	"github.com/beevik/etree"
)

type PartMap map[string][]byte

type DocxDocument struct {
	Parts       PartMap
	DocumentXML *etree.Document // word/document.xml
	CommentsXML *etree.Document // optional: word/comments.xml
	ContentXML  *etree.Document // [Content_Types].xml
	RelsXML     *etree.Document // word/_rels/document.xml.rels
}

func ParseDocx(raw []byte) (*DocxDocument, error) {
	r, err := zip.NewReader(bytes.NewReader(raw), int64(len(raw)))
	if err != nil {
		return nil, err
	}
	parts := make(PartMap, len(r.File))
	for _, f := range r.File {
		rc, err := f.Open()
		if err != nil {
			return nil, err
		}
		b, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			return nil, err
		}
		parts[f.Name] = b
	}

	docXML, ok := parts["word/document.xml"]
	if !ok {
		return nil, fmt.Errorf("missing word/document.xml")
	}

	doc := &DocxDocument{Parts: parts}
	doc.DocumentXML = etree.NewDocument()
	if err := doc.DocumentXML.ReadFromBytes(docXML); err != nil {
		return nil, err
	}

	// Optional parts if present.
	if b, ok := parts["word/comments.xml"]; ok {
		doc.CommentsXML = etree.NewDocument()
		if err := doc.CommentsXML.ReadFromBytes(b); err != nil {
			return nil, err
		}
	}
	if b, ok := parts["[Content_Types].xml"]; ok {
		doc.ContentXML = etree.NewDocument()
		if err := doc.ContentXML.ReadFromBytes(b); err != nil {
			return nil, err
		}
	}
	if b, ok := parts["word/_rels/document.xml.rels"]; ok {
		doc.RelsXML = etree.NewDocument()
		if err := doc.RelsXML.ReadFromBytes(b); err != nil {
			return nil, err
		}
	}

	return doc, nil
}
```

Pattern:
- read every ZIP entry once,
- parse only editable XML parts into trees,
- leave all other entries untouched for round-trip safety.

## 2. Namespace helpers and "XPath-lite" traversal

In OOXML-heavy code, explicit namespace handling keeps behavior predictable.

```go
const (
	WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
)

func findBody(doc *etree.Document) (*etree.Element, error) {
	root := doc.Root()
	if root == nil {
		return nil, fmt.Errorf("empty document.xml")
	}
	body := findFirstByTag(root, WNS, "body")
	if body == nil {
		return nil, fmt.Errorf("missing w:body")
	}
	return body, nil
}

func findFirstByTag(el *etree.Element, space, local string) *etree.Element {
	if el.Space == "w" && el.Tag == local {
		// For etree, prefix is in Space, URI comes from Attr declarations.
		// Keep this helper centralized and strict in production.
	}
	for _, c := range el.ChildElements() {
		if c.Tag == local {
			// Production code should verify namespace mapping to avoid false positives.
			return c
		}
		if hit := findFirstByTag(c, space, local); hit != nil {
			return hit
		}
	}
	return nil
}
```

In production, prefer one internal query layer:
- either strict helpers over `etree`,
- or `xmlquery` with registered namespaces for complex selectors.

## 3. Interleaved fragment IDs (paragraph/table)

Mirror your current map strategy so change IDs stay stable:

```go
type FragmentRef struct {
	ID int
	El *etree.Element // w:p or w:tbl
}

func InterleavedMap(body *etree.Element) map[int]*etree.Element {
	out := map[int]*etree.Element{}
	next := 1
	for _, child := range body.ChildElements() {
		if child.Tag == "p" || child.Tag == "tbl" {
			out[next] = child
			next++
		}
	}
	return out
}
```

## 4. Build tracked-change wrappers (`w:ins` / `w:del`)

Centralize wrapper construction so required attrs are always set.

```go
func makeTrackedChange(local string, id int, author, dateISO string) *etree.Element {
	tc := etree.NewElement("w:" + local)
	tc.CreateAttr("w:id", fmt.Sprintf("%d", id))
	tc.CreateAttr("w:author", author)
	tc.CreateAttr("w:date", dateISO)
	return tc
}
```

Usage pattern:
- for modify: split runs into equal/insert/delete segments, then emit `w:r` under `w:ins` and `w:del`.
- for delete paragraph: wrap existing runs in `w:del` and convert `w:t` to `w:delText`.
- for append: create new `w:p` and insert `w:ins` with inherited formatting.

## 5. Safe handling of unsupported structures

Return structured skip reasons instead of partial edits:

```go
func IsSimpleTable(tbl *etree.Element) (bool, string) {
	rows := tbl.SelectElements("w:tr")
	if len(rows) == 0 {
		return false, "table has no rows"
	}
	expected := -1
	for r, row := range rows {
		cells := row.SelectElements("w:tc")
		if expected == -1 {
			expected = len(cells)
		} else if len(cells) != expected {
			return false, fmt.Sprintf("row %d has %d cells, expected %d", r+1, len(cells), expected)
		}
		for c, tc := range cells {
			if hasPath(tc, "w:tcPr/w:gridSpan") {
				return false, fmt.Sprintf("cell %d.%d has gridSpan", r+1, c+1)
			}
			if hasPath(tc, "w:tcPr/w:vMerge") {
				return false, fmt.Sprintf("cell %d.%d has vMerge", r+1, c+1)
			}
		}
	}
	return true, ""
}
```

Policy:
- reject unsupported target elements up front,
- do not mutate those targets,
- keep document writable by preserving untouched XML.

## 6. Repack unchanged + modified parts

```go
func (d *DocxDocument) ToBytes() ([]byte, error) {
	parts := make(PartMap, len(d.Parts))
	for k, v := range d.Parts {
		parts[k] = v
	}

	parts["word/document.xml"], _ = d.DocumentXML.WriteToBytes()
	if d.CommentsXML != nil {
		parts["word/comments.xml"], _ = d.CommentsXML.WriteToBytes()
	}
	if d.ContentXML != nil {
		parts["[Content_Types].xml"], _ = d.ContentXML.WriteToBytes()
	}
	if d.RelsXML != nil {
		parts["word/_rels/document.xml.rels"], _ = d.RelsXML.WriteToBytes()
	}

	buf := new(bytes.Buffer)
	zw := zip.NewWriter(buf)
	for name, payload := range parts {
		w, err := zw.Create(name)
		if err != nil {
			return nil, err
		}
		if _, err := w.Write(payload); err != nil {
			return nil, err
		}
	}
	if err := zw.Close(); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}
```

## 7. Concurrency model for multi-doc throughput

Per-document isolation is the clean unit of concurrency.

```go
type Job struct {
	Doc []byte
	Req ApplyRequest
}
type Result struct {
	Out []byte
	Err error
}

func RunPool(jobs <-chan Job, workers int) <-chan Result {
	out := make(chan Result)
	var wg sync.WaitGroup
	for i := 0; i < workers; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for j := range jobs {
				edited, err := ApplyRedlines(j.Doc, j.Req)
				out <- Result{Out: edited, Err: err}
			}
		}()
	}
	go func() {
		wg.Wait()
		close(out)
	}()
	return out
}
```

Operational guidance:
- cap worker count by memory budget, not just CPU,
- enforce per-doc max size and timeout,
- fail one job without affecting others.

## 8. Validation gates before returning output

Keep these mandatory checks:
- every `w:ins` / `w:del` has `w:id`, `w:author`, `w:date`,
- no ID collisions across comment/tracked-change groups,
- if `comments.xml` exists, relationship + content type entries exist.

This mirrors your existing validator contract and is the strongest defense against Word repair dialogs.

