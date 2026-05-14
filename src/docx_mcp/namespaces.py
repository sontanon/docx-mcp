"""OOXML namespace constants and helpers for working with .docx XML."""

from lxml import etree

# --- Core OOXML Namespaces ---

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XML = "http://www.w3.org/XML/1998/namespace"

# --- Package/OPC Namespaces ---

CT = "http://schemas.openxmlformats.org/package/2006/content-types"
RELS = "http://schemas.openxmlformats.org/package/2006/relationships"

# --- Relationship Types ---

REL_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
REL_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"

# --- Content Types ---

CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"

# --- Extended comment namespaces (for modern Word compat) ---

W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"

# --- Namespace map for XPath and element creation ---

NSMAP: dict[str, str] = {
    "w": W,
    "r": R,
    "wp": WP,
    "mc": MC,
    "w15": W15,
}


def qn(ns_prefix: str, local_name: str) -> str:
    """Build a Clark-notation qualified name: {namespace}localname.

    Usage:
        qn("w", "p") -> "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
    """
    ns_uri = NSMAP.get(ns_prefix)
    if ns_uri is None:
        msg = f"Unknown namespace prefix: {ns_prefix!r}. Known: {list(NSMAP.keys())}"
        raise ValueError(msg)
    return f"{{{ns_uri}}}{local_name}"


def make_element(ns_prefix: str, local_name: str, **attribs: str) -> etree._Element:
    """Create an lxml element with a qualified name and optional attributes.

    Attribute names can use double-underscore for namespaced attributes:
        make_element("w", "ins", w__id="1", w__author="AI")
        -> <w:ins w:id="1" w:author="AI"/>

    Plain attribute names are used as-is (no namespace):
        make_element("w", "t", **{"xml:space": "preserve"})
    """
    tag = qn(ns_prefix, local_name)
    el = etree.Element(tag)

    for key, value in attribs.items():
        if "__" in key:
            # Namespaced attribute: "w__id" -> "{W}id"
            prefix, attr_local = key.split("__", 1)
            attr_qn = qn(prefix, attr_local)
            el.set(attr_qn, value)
        elif key == "xml_space":
            # Special case for xml:space="preserve"
            el.set(f"{{{XML}}}space", value)
        else:
            el.set(key, value)

    return el


def xpath(element: etree._Element, expr: str) -> list[etree._Element]:
    """Run an XPath expression with the OOXML namespace map."""
    return element.xpath(expr, namespaces=NSMAP)
