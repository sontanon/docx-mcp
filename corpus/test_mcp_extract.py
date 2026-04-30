from __future__ import annotations

import asyncio
import json
from pathlib import Path

from fastmcp import Client

from docx_mcp.server import mcp


def _text(result) -> str:
    """Extract the text content from a CallToolResult."""
    return result.content[0].text


async def main() -> None:
    doc_path = (
        Path("/home/santiago/Code/docx-mcp/corpus/downloaded")
        / "05fbf2a0d628f10fde61e6f5a67d365920459e6903733543426ad2027ecb594c.docx"
    )

    async with Client(mcp) as client:
        # Tagged format
        print("=== extract_fragments (tagged) ===")
        try:
            result_tagged = await client.call_tool(
                "extract_fragments",
                {"document_path": str(doc_path), "format": "tagged"},
            )
            tagged_text = _text(result_tagged)
            print("SUCCESS")
            for line in tagged_text.splitlines()[:20]:
                print(line)
        except Exception as exc:
            print(f"FAILED: {exc}")

        print()

        # JSON format
        print("=== extract_fragments (json) ===")
        try:
            result_json = await client.call_tool(
                "extract_fragments",
                {"document_path": str(doc_path), "format": "json"},
            )
            json_text = _text(result_json)
            data = json.loads(json_text)
            fragments = data.get("fragments", [])
            print("SUCCESS")
            for item in fragments[:3]:
                print(json.dumps(item, indent=2))
        except Exception as exc:
            print(f"FAILED: {exc}")


if __name__ == "__main__":
    asyncio.run(main())
