"""Change handlers for the three supported change types."""

from docx_mcp.handlers.append import handle_append_after
from docx_mcp.handlers.delete import handle_delete
from docx_mcp.handlers.modify import handle_modify

__all__ = ["handle_append_after", "handle_delete", "handle_modify"]
