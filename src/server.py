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
        server_version="0.6",
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
                description="Creates powerpoint presentation and returns a link for the created file.",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "author": {
                            "type": "string",
                            "description": "Name of the author."
                        },
                        "format": {
                            "type": "string",
                            "enum": ["4:3", "16:9"],
                            "default":"4:3",
                            "description": "Format of the presentation, either 4:3 or 16:9. Will default to 4:3 if not specified."
                        },
                        "slides": {
                            "type": "array",
                            "description": "Individual slides content. One slide per list item.",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "slide_type": {
                                        "type": "integer",
                                        "description": "Type of slide layout - use 0 for title slide, 1 for section divider or 2 for text content slide.",
                                    },
                                    "slide_title": {
                                        "type": "string",
                                        "description": "Title of the slide."
                                    },
                                    "slide_text": {
                                        "type": "string",
                                        "description": "Text of the slide in paragraphs. Each paragraph shall be separated by newline. Text of each paragraph must be prefixed by %1 or %2 for indentation level. Do not include for title and section slide layouts."
                                    }
                                },
                                "required": ["slide_type","title"]
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

            author: str = arguments.get("author")
            pptx_format: str = arguments.get("format")
            slides: list = arguments.get("slides")

            if not slides:
                raise ValueError("Missing slides")

            if not author:
                author = "[ADD YOUR NAME]"

            if not pptx_format:
                pptx_format = "4:3"

            result_text = create_presentation(author,slides, pptx_format)

            return [
                types.TextContent(
                    type="text",
                    text=result_text
                )
            ]

        else:
            raise ValueError(f"Unknown tool: {name}")

    return server, init_options


