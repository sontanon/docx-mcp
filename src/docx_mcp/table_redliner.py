"""Apply tracked changes to table cells.

This module handles table-specific change operations:
- modify_cell: Word-level diff on cell paragraphs (multi-para alignment)
- clear_cell: Mark all cell paragraphs as deleted

Cell operations delegate to existing paragraph handlers (handle_modify,
handle_delete, handle_append_after) to ensure consistent tracked-change
formatting.
"""

from lxml import etree

from docx_mcp.comments import add_comment
from docx_mcp.document import DocxDocument
from docx_mcp.handlers.append import handle_append_after
from docx_mcp.handlers.delete import handle_delete
from docx_mcp.handlers.modify import handle_modify
from docx_mcp.id_manager import IdManager
from docx_mcp.models import RedlineConfig, TableChange, TableChangeType
from docx_mcp.table_utils import get_cell_element, get_cell_paragraphs


def apply_table_changes(
    doc: DocxDocument,
    table_changes: list[TableChange],
    element_map: dict[str, etree._Element],
    *,
    id_manager: IdManager,
    config: RedlineConfig,
) -> None:
    """Apply tracked changes to table cells.

    Args:
        doc: The document being modified.
        table_changes: List of table cell changes to apply.
        element_map: Mapping of fragment_id → <w:p> or <w:tbl> elements.
        id_manager: ID allocator for annotation IDs.
        config: Author / date configuration.

    Raises:
        ValueError: If table_id does not point to a <w:tbl> element, or if
            the cell coordinates are out of range.
    """
    for change in table_changes:
        # Look up the table element
        tbl = element_map.get(str(change.table_id))
        if tbl is None:
            msg = (
                f"Table change references table_id={change.table_id}, "
                f"but no element exists at that position"
            )
            raise ValueError(msg)

        # Verify it's a table
        tag_local = tbl.tag.split("}")[-1] if "}" in tbl.tag else tbl.tag
        if tag_local != "tbl":
            msg = (
                f"Table change references table_id={change.table_id}, "
                f"but the element at that position is <w:{tag_local}>, not <w:tbl>"
            )
            raise ValueError(msg)

        # Get the cell
        try:
            tc = get_cell_element(tbl, change.row, change.col)
        except ValueError as exc:
            msg = f"Table change for cell_id={change.cell_id}: {exc}"
            raise ValueError(msg) from exc

        # Get cell paragraphs
        paras = get_cell_paragraphs(tc)

        if not paras:
            msg = f"Table change for cell_id={change.cell_id}: cell has no paragraphs"
            raise ValueError(msg)

        # Dispatch by change type
        if change.change_type == TableChangeType.MODIFY_CELL:
            _modify_cell(paras, change, id_manager=id_manager, config=config)
            # Attach comment to first paragraph
            add_comment(
                doc,
                paras[0],
                change.justification,
                id_manager=id_manager,
                config=config,
            )

        elif change.change_type == TableChangeType.CLEAR_CELL:
            _clear_cell(paras, id_manager=id_manager, config=config)
            # Attach comment to first paragraph
            add_comment(
                doc,
                paras[0],
                change.justification,
                id_manager=id_manager,
                config=config,
            )

        else:
            msg = f"Unsupported change_type: {change.change_type.value}"
            raise ValueError(msg)


def _modify_cell(
    paras: list[etree._Element],
    change: TableChange,
    *,
    id_manager: IdManager,
    config: RedlineConfig,
) -> None:
    """Apply modify_cell change: align old/new paragraphs positionally.

    Args:
        paras: List of <w:p> elements in the cell.
        change: The table change (must have new_text).
        id_manager: ID allocator.
        config: Author / date configuration.

    Raises:
        ValueError: If new_text is None.
    """
    if change.new_text is None:
        msg = f"modify_cell on cell_id={change.cell_id} requires new_text"
        raise ValueError(msg)

    # Split new text by \n to get new paragraph texts
    new_texts = change.new_text.split("\n")

    n_old = len(paras)
    n_new = len(new_texts)

    # Paired: modify existing paragraphs
    for i in range(min(n_old, n_new)):
        handle_modify(
            paras[i],
            new_texts[i],
            id_manager=id_manager,
            config=config,
        )

    # Excess old: delete with preserve_paragraph_mark=True
    if n_old > n_new:
        for i in range(n_new, n_old):
            handle_delete(
                paras[i],
                id_manager=id_manager,
                config=config,
                preserve_paragraph_mark=True,
            )

    # Excess new: append after last old paragraph
    # addnext() inside <w:tc> keeps new <w:p> in the same cell
    if n_new > n_old:
        last_para = paras[-1]
        for i in range(n_old, n_new):
            new_p, _ = handle_append_after(
                last_para,
                new_texts[i],
                id_manager=id_manager,
                config=config,
            )
            last_para = new_p


def _clear_cell(
    paras: list[etree._Element],
    *,
    id_manager: IdManager,
    config: RedlineConfig,
) -> None:
    """Mark all paragraphs in the cell as deleted.

    Args:
        paras: List of <w:p> elements in the cell.
        id_manager: ID allocator.
        config: Author / date configuration.
    """
    for para in paras:
        handle_delete(
            para,
            id_manager=id_manager,
            config=config,
            preserve_paragraph_mark=True,
        )
