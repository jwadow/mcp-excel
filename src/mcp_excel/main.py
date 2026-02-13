"""MCP Excel Server - Main entry point."""

import asyncio
import logging
from typing import Any

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

from .core.file_loader import FileLoader
from .models.requests import (
    GetColumnNamesRequest,
    GetSheetInfoRequest,
    InspectFileRequest,
)
from .operations.inspection import InspectionOperations

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger("mcp_excel")


class MCPExcelServer:
    """MCP server for Excel operations."""

    def __init__(self) -> None:
        """Initialize MCP Excel server."""
        self.server = Server("mcp-excel")
        self.file_loader = FileLoader()
        self.inspection_ops = InspectionOperations(self.file_loader)

        # Register handlers
        self._register_handlers()

    def _register_handlers(self) -> None:
        """Register MCP handlers."""

        @self.server.list_tools()
        async def list_tools() -> list[Tool]:
            """List available tools."""
            return [
                Tool(
                    name="inspect_file",
                    description="Inspect Excel file structure and get basic information about all sheets",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            }
                        },
                        "required": ["file_path"],
                    },
                ),
                Tool(
                    name="get_sheet_info",
                    description="Get detailed information about a specific sheet including column names, types, and sample data",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet to inspect",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name"],
                    },
                ),
                Tool(
                    name="get_column_names",
                    description="Get list of column names from a sheet",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name"],
                    },
                ),
            ]

        @self.server.call_tool()
        async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
            """Handle tool calls."""
            try:
                logger.info(f"Tool called: {name} with arguments: {arguments}")

                if name == "inspect_file":
                    request = InspectFileRequest(**arguments)
                    response = self.inspection_ops.inspect_file(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "get_sheet_info":
                    request = GetSheetInfoRequest(**arguments)
                    response = self.inspection_ops.get_sheet_info(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "get_column_names":
                    request = GetColumnNamesRequest(**arguments)
                    response = self.inspection_ops.get_column_names(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                else:
                    raise ValueError(f"Unknown tool: {name}")

            except Exception as e:
                logger.error(f"Error executing tool {name}: {e}", exc_info=True)
                error_response = {
                    "error": type(e).__name__,
                    "message": str(e),
                    "recoverable": True,
                }
                return [TextContent(type="text", text=str(error_response))]

    async def run(self) -> None:
        """Run the MCP server."""
        logger.info("Starting MCP Excel Server...")
        async with stdio_server() as (read_stream, write_stream):
            await self.server.run(
                read_stream,
                write_stream,
                self.server.create_initialization_options(),
            )


def main() -> None:
    """Main entry point."""
    server = MCPExcelServer()
    asyncio.run(server.run())


if __name__ == "__main__":
    main()
