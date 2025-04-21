import mcp.types as types
from mcp.server import Server, NotificationOptions
from mcp.server.models import InitializationOptions
from create_pptx import create_presentation
import logging
from create_docx import html_to_word
from create_msg import create_eml


def create_server():
    logging.basicConfig(level=logging.DEBUG)
    logger = logging.getLogger("mcp-office-documents")
    logger.setLevel(logging.DEBUG)
    logger.info("Starting MCP Office Docs")

    # Initialize base MCP server
    server = Server("office_documents")

    init_options = InitializationOptions(
        server_name="mcp-office-documents",
        server_version="0.2",
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
            types.Tool(
                name="create-word-document",
                description="Creates a Word document (docx) from HTML input.",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "html_content": {
                            "type": "string",
                            "description": "Content in HTML 5 format. Do not include any numbering in text. Supports: <h1>-<h3> for headings (h1 for title, h2 for chapter/clause titles, h3 for subchapter titles, <strong> for bold, <em> for italic, <a href='url'> for links, <table> with <tr>/<td> for tables, <ol> for ordered lists and <ul> for unordered lists. For list nesting, ensure each <li> element contains nested <ul> or <ol> as direct children, e.g. <ul><li>Item 1<ul><li>Subitem</li></ul></li></ul>. For legal document (contract etc.), each clause, except for definition of parties, must be part of ordered list, e.g. <h2>Final Provisions</h2><ol><li>This contract is valid from 1 January 2026</li><li>Contract is drown in 2 originals</li></ol>."
                        }
                    },
                    "required": ["html_content"]
                }
            ),
            types.Tool(
                name="create-email-draft",
                description="Creates an email draft.",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "to": {
                            "type": "array",
                            "items": {
                                "type": "string"
                            },
                            "description": "List of recipient email addresses."
                        },
                        "cc": {
                            "type": "array",
                            "items": {
                                "type": "string"
                            },
                            "description": "List of carbon copy recipient email addresses."
                        },
                        "bcc": {
                            "type": "array",
                            "items": {
                                "type": "string"
                            },
                            "description": "List of blind carbon copy recipient email addresses."
                        },
                        "re": {
                            "type": "string",
                            "description": "Subject of the email."
                        },
                        "content": {
                            "type": "string",
                            "description": "HTML content for body of the email. Must contain valid HTML markup (paragraphs, lists, etc.) without the enclosing <html>, <head>, or <body> tags. Do not use header tags. Example: <p>Hello</p><p>This is <b>valid</b> point for our <u>discussion</u>.</p><p>We consider <ol><li>Point 1</li><li>Point 2</li></ol></p><p><b>Subheader</b></p><p>Also note this: <ul><li>Note 1</li><li>Note 2</li><ul></p><p>Kind regards</p>"
                        }
                    },
                    "required": ["re", "content"]
                }
            )
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

        elif name == "create-word-document":
            html: str = arguments.get("html_content")

            url = html_to_word(html)

            return [
                types.TextContent(
                    type="text",
                    text=url
                )
            ]

        elif name == "create-email-draft":

            to: list = list(arguments.get("to", []))
            cc: list = list(arguments.get("cc", []))
            bcc: list = list(arguments.get("bcc", []))
            re: str = arguments.get("re")
            content: str = arguments.get("content")

            if not re or not content:
                raise ValueError("Missing argument re or content")

            url = create_eml(to, cc, bcc, re, content)

            return [
                types.TextContent(
                    type="text",
                    text=url
                )
            ]

        else:
            raise ValueError(f"Unknown tool: {name}")

    return server, init_options

