from fastmcp import FastMCP
from pydantic import Field
from typing import Annotated, List, Dict, Any, Optional
import io
from create_xlsx import markdown_to_excel
from create_docx import markdown_to_word
from create_pptx import create_presentation
from create_msg import create_eml
from upload_file import upload_file

mcp = FastMCP("MCP Office Documents")

@mcp.tool(
    name="create_excel_from_markdown",
    description="Converts markdown content with tables and formulas to Excel (.xlsx) format.",
    tags={"excel", "spreadsheet", "data"},
    annotations={"title": "Markdown to Excel Converter"}
)
async def create_excel_document(
    markdown_content: Annotated[str, Field(description="Markdown content containing tables, headers, and formulas. Use T1.B[0] for cross-table references and B[0] for current row references. ALWAYS use [0], [1], [2] notation, NEVER use absolute row numbers like B2, B3. Do NOT count table header as first row, first row has index [0]. Supports cell formatting: **bold**, *italic*.")]
) -> str:
    """
    Converts markdown to Excel with advanced formula support.

    CRITICAL - Formula Syntax (USE ONLY THESE FORMATS):
    - Within table: =B[0]+C[0] (current row), =SUM(B[0]:E[0]) (range)
    - Cross-table: =T1.B[0] (Table 1, first data row), =T1.SUM(C[0]:F[2]) (Table 1, range)
    - Functions: SUM, AVERAGE, MAX, MIN with ranges using [offset] notation

    IMPORTANT - Row Indexing Rules:
    - ALWAYS use bracket notation: [0], [1], [2], etc.
    - NEVER use absolute row numbers like B2, B3, C4
    - [0] = FIRST DATA ROW (after header)
    - [1] = SECOND DATA ROW
    - [2] = THIRD DATA ROW, and so on
    - Headers are automatically styled and excluded from indexing

    Examples (CORRECT):
    - =T1.B[0]+T1.B[1]+T1.B[2] (sum first 3 data rows from Table 1, column B)
    - =B[0]+C[0]+D[0] (sum current row across columns B, C, D)
    - =T2.SUM(B[0]:E[3]) (sum range in Table 2 from first to fourth data row)

    Examples (WRONG - DO NOT USE):
    - =T1.B2+T1.B3+T1.B4 (absolute row numbers not supported)
    - =B2+C2+D2 (absolute row numbers not supported)

    Features:
    - Auto-formats numbers, percentages, and currencies
    - Professional styling with borders and headers
    - Cross-table formula resolution
    - Position-independent references
    """

    print(f"Converting markdown to Excel document")

    try:
        # markdown_to_excel now handles upload internally and returns URL
        result = markdown_to_excel(markdown_content)
        print(f"Excel document uploaded successfully")
        return result
    except Exception as e:
        print(f"Error creating Excel document: {e}")
        return f"Error creating Excel document: {str(e)}"

@mcp.tool(
    name="create_word_from_markdown",
    description="Converts markdown content to Word (.docx) format. Supports headers, tables, lists, formatting, hyperlinks, and block quotes.",
    tags={"word", "document", "text", "legal", "contract"},
    annotations={"title": "Markdown to Word Converter"}
)
async def create_word_document(
    markdown_content: Annotated[str, Field(description="Markdown content. For LEGAL CONTRACTS use numbered lists (1., 2., 3.) for sections and nested lists for provisions - DO NOT use headers (except for contract title). For other documents use headers (# ## ###).")]
) -> str:
    """
    Converts markdown to professionally formatted Word document.

    DOCUMENT STRUCTURE GUIDELINES:

    FOR LEGAL CONTRACTS:
    - Use numbered lists for sections: 1. Článek I - Předmět smlouvy
    - Use nested numbered lists for provisions:
      1. Section heading
         1. First provision
         2. Second provision
      2. Next section heading
         1. First provision (automatically restarts at 1)
         2. Second provision
    - DO NOT use headers (# ## ###) in contracts

    FOR LETTERS, MEMOS, REPORTS:
    - Use headers for sections: # Main Title, ## Section, ### Subsection
    - Use lists for bullet points or numbered items within sections
    - Headers provide proper document outline structure

    Supported Markdown:
    - Headers: # ## ### (for letters/memos/reports)
    - Tables: | Column | Column | (with borders)
    - Lists: - bullet or 1. numbered (with automatic nesting)
    - Formatting: **bold**, *italic*, `code`, [links](url)
    - Block quotes: > quoted text
    - Line breaks: two spaces at end of line

    Features:
    - Professional styling and fonts
    - Proper table formatting
    - Word's automatic list numbering with restart
    - Hyperlink creation
    - Template support
    """

    print(f"Converting markdown to Word document")

    try:
        # markdown_to_word now handles upload internally and returns URL
        result = markdown_to_word(markdown_content)
        print(f"Word document uploaded successfully")
        return result
    except Exception as e:
        print(f"Error creating Word document: {e}")
        return f"Error creating Word document: {str(e)}"

@mcp.tool(
    name="create_powerpoint_presentation",
    description="Creates PowerPoint (.pptx) presentations with multiple slide types.",
    tags={"powerpoint", "presentation", "slides"},
    annotations={"title": "PowerPoint Presentation Creator"}
)
async def create_powerpoint_presentation(
    slides: Annotated[List[Dict[str, Any]], Field(description="List of slide dictionaries. Each slide must have 'slide_type' (title/section/content), 'slide_title', and content based on type.")],
    format: Annotated[str, Field(description="Presentation format: '4:3' for traditional or '16:9' for widescreen", default="16:9")]
) -> str:
    """
    Creates PowerPoint presentations with professional templates.

    Slide Types:
    - title: {"slide_type": "title", "slide_title": "Title", "author": "Author"}
    - section: {"slide_type": "section", "slide_title": "Section Title"}
    - content: {"slide_type": "content", "slide_title": "Title", "slide_text": [{"text": "Bullet point", "indentation_level": 1}]}

    Features:
    - Professional templates (4:3 and 16:9 formats)
    - Multi-level bullet points
    - Consistent styling
    - Custom layouts for different slide types
    """

    print(f"Creating PowerPoint presentation with {len(slides)} slides in {format} format")

    try:
        # create_presentation already handles upload internally and returns URL
        result = create_presentation(slides, format)
        print(f"PowerPoint presentation created: {result}")
        return result
    except Exception as e:
        print(f"Error creating PowerPoint presentation: {e}")
        return f"Error creating PowerPoint presentation: {str(e)}"

@mcp.tool(
    name="create_email_draft",
    description="Creates an email draft in EML format with HTML content using preset professional styling.",
    tags={"email", "eml", "communication"},
    annotations={"title": "Email Draft Creator"}
)
async def create_email_draft(
    content: Annotated[str, Field(description="BODY CONTENT ONLY - Do NOT include HTML structure tags like <html>, <head>, <body>, or <style>. Do NOT include any CSS styling. Use <p> for greetings and for signatures, never headers. Use <h2> for section headers (will be bold), <h3> for subsection headers (will be underlined). HTML tags allowed: <p>, <h2>, <h3>, <ul>, <li>, <strong>, <em>, <div>.")],
    subject: Annotated[str, Field(description="Email subject line")],
    to: Annotated[Optional[List[str]], Field(description="List of recipient email addresses", default=None)],
    cc: Annotated[Optional[List[str]], Field(description="List of CC recipient email addresses", default=None)],
    bcc: Annotated[Optional[List[str]], Field(description="List of BCC recipient email addresses", default=None)],
    priority: Annotated[str, Field(description="Email priority: 'low', 'normal', or 'high'", default="normal")],
    language: Annotated[str, Field(description="Language code for proofreading in Outlook (e.g., 'cs-CZ' for Czech, 'en-US' for English, 'de-DE' for German, 'sk-SK' for Slovak)", default="cs-CZ")]
) -> str:
    """
    Creates professional email drafts in EML format with preset styling and language settings.

    IMPORTANT - Content Guidelines:
    - Provide ONLY the body content of the email
    - Do NOT include <html>, <head>, <body>, or <style> tags
    - Do NOT include any CSS styling or inline styles

    PROPER USAGE OF TAGS:
    - Greetings & Signatures: Use <p>Vážený klientě,</p> and <p>S pozdravem,<br>Váš tým</p>
    - Main sections: Use <h2>Kontrola a doplnění údajů</h2> (will appear bold)
    - Subsections: Use <h3>Mzdové podmínky</h3> (will appear underlined)
    - Regular text: Use <p> for paragraphs
    - Lists: <ul>, <ol>, <li>
    - Emphasis: <strong>, <em>

    Features:
    - Consistent formatting (Arial font, same color for all text)
    - Simple styling: H2 = bold, H3 = underlined
    - Language settings for Outlook proofreading/spell-check
    - Multiple recipient types (To, CC, BCC)
    - Priority settings (low/normal/high)
    - UTF-8 encoding for international characters

    Common Language Codes:
    - cs-CZ: Czech (Czech Republic)
    - en-US: English (United States)
    - en-GB: English (United Kingdom)
    - de-DE: German (Germany)
    - sk-SK: Slovak (Slovakia)
    - fr-FR: French (France)
    """

    print(f"Creating email draft with subject: {subject}")

    try:
        # create_eml already handles upload internally and returns URL
        result = create_eml(
            to=to,
            cc=cc,
            bcc=bcc,
            re=subject,
            content=content,
            priority=priority,
            language=language
        )
        print(f"Email draft created: {result}")
        return result
    except Exception as e:
        print(f"Error creating email draft: {e}")
        return f"Error creating email draft: {str(e)}"

if __name__ == "__main__":
    mcp.run(
        transport="streamable-http",
        host="0.0.0.0",
        port=8958,
        log_level="info",
        path="/mcp"
    )
