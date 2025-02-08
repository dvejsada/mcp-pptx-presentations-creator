import mcp.types as types
from mcp.server import Server, NotificationOptions
from mcp.server.models import InitializationOptions
from create_pptx import create_presentation
import logging


def create_server():
    logging.basicConfig(level=logging.DEBUG)
    logger = logging.getLogger("mcp-pptxcreator")
    logger.setLevel(logging.DEBUG)
    logger.info("Starting MCP Presentation")

    # Initialize base MCP server
    server = Server("pptx_presentation")

    init_options = InitializationOptions(
        server_name="mcp-pptxcreator",
        server_version="0.7",
        capabilities=server.get_capabilities(
            notification_options=NotificationOptions(),
            experimental_capabilities={},
        ),
    )

    @server.list_tools()
    async def handle_list_tools() -> list[types.Tool]:
        """
        List available tools.
        Each tool specifies its arguments using JSON Schema validation.
        Name must be maximum of 64 characters
        """
        return [
            types.Tool(
                name="create-powerpoint-presentation",
                description="Creates powerpoint presentation as pptx file.",
                inputSchema={
                    "$schema": "http://json-schema.org/draft-07/schema#",
                    "title": "Input schema for presentation creation tool. Provide only JSON in single line.",
                    "type": "object",
                    "properties": {
                        "format": {
                            "type": "string",
                            "enum": ["4:3", "16:9"],
                            "default": "4:3",
                            "description": "Format of the presentation, either 4:3 or 16:9. Defaults to 4:3 if not specified."
                        },
                        "slides": {
                            "type": "array",
                            "description": "Individual slides content. One slide per list item.",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "slide_type": {
                                        "type": "string",
                                        "description": "Type of slide layout. Title slide can be used only as a first slide.",
                                        "enum": ["title", "section", "content"]
                                    },
                                    "slide_title": {
                                        "type": "string",
                                        "description": "Title of the slide."
                                    },
                                    "author": {
                                        "type": "string",
                                        "description": "Name of the author. Required for title slide."
                                    },
                                    "slide_text": {
                                        "type": "array",
                                        "description": "An array of text items with indentation levels. Required for content slides.",
                                        "items": {
                                            "type": "object",
                                            "properties": {
                                                "text": {
                                                    "type": "string",
                                                    "description": "Text bullet point content."
                                                },
                                                "indentation_level": {
                                                    "type": "integer",
                                                    "description": "Indentation level for the bullet point (1 for no indentation).",
                                                    "minimum": 1,
                                                    "maximum": 3,
                                                    "default": 1
                                                }
                                            },
                                            "required": ["text","indentation_level"]
                                        }
                                    }
                                },
                                "required": ["slide_type"],
                                "oneOf": [
                                    {
                                        "properties": {"slide_type": {"const": "title"}},
                                        "required": ["slide_title", "author"]
                                    },
                                    {
                                        "properties": {"slide_type": {"const": "section"}},
                                        "required": ["slide_title"]
                                    },
                                    {
                                        "properties": {"slide_type": {"const": "content"}},
                                        "required": ["slide_title", "slide_text"]
                                    }
                                ]
                            }
                        }
                    },
                    "required": ["slides"]
                }

            ),
        ]

    @server.call_tool()
    async def handle_call_tool(
            name: str, arguments: dict | None
    ) -> list[types.TextContent | types.ImageContent | types.EmbeddedResource]:
        """
        Handle tool execution requests.
        """
        if not arguments:
            raise ValueError("Missing arguments")

        if name == "create-powerpoint-presentation":

            pptx_format: str = arguments.get("format")
            slides: list = arguments.get("slides")

            if not slides:
                raise ValueError("Missing slides")

            if not pptx_format:
                pptx_format = "4:3"

            result_text = create_presentation(slides, pptx_format)

            return [
                types.TextContent(
                    type="text",
                    text=result_text
                )
            ]

        else:
            raise ValueError(f"Unknown tool: {name}")

    return server, init_options
