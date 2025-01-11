from pptx import Presentation
from pptx.enum.text import PP_ALIGN
import re
import unicodedata

def sanitize_filename(filename, replacement=''):
    """
    Remove or replace invalid characters from a filename string.

    Parameters:
        filename (str): The original filename string.
        replacement (str): The character to replace invalid characters with.

    Returns:
        str: The sanitized filename.
    """
    filename = unicodedata.normalize('NFKD', filename)
    # Remove diacritical marks (accents)
    filename = ''.join(c for c in filename if not unicodedata.combining(c))
    # Encode to ASCII (ignore errors for characters that can't be encoded)
    filename = filename.encode('ascii', 'ignore').decode('ascii')
    # Define a list of invalid characters for filenames on Windows and Unix system
    invalid_chars = r'<>:"/\\|!?* '
    # Remove control characters (ASCII codes 0-31)
    filename = re.sub(r'[\0-\31]', '', filename)
    # Replace invalid characters with the specified replacement character
    sanitized = re.sub(r'[{}]+'.format(re.escape(invalid_chars)), replacement, filename)
    # Remove any leading or trailing whitespace
    sanitized = sanitized.strip()
    # Optionally, limit the filename length to a maximum number of characters (e.g., 255)
    sanitized = sanitized[:255]
    return sanitized

TEMPLATE_REGULAR = "template_pptx.pptx"
TEMPLATE_WIDE = ""

TITLE_SLIDE_LAYOUT = 0
SECTION_SLIDE_LAYOUT = 1
TITLE_AND_CONTENT_LAYOUT = 2

class PowerpointPresentation:

    def __init__(self, title: str, author: str, slides: list, format: str):

        self.title = title
        self.author = author
        self.slides: list = slides

        if format == "4:3":
            self.presentation = Presentation(TEMPLATE_REGULAR)
        elif format == "16:9":
            self.presentation = Presentation(TEMPLATE_REGULAR)
        else:
            self.presentation = Presentation(TEMPLATE_REGULAR)

        self.create_title_slide()

        for slide in self.slides:
            self.create_content_slide(slide)


    def create_title_slide(self):
        title_layout = self.presentation.slide_layouts[TITLE_SLIDE_LAYOUT]
        title_slide = self.presentation.slides.add_slide(title_layout)
        title_slide.placeholders[0].text = self.title
        title_slide.placeholders[1].text = self.author

    def create_section_header_slide(self, slide: dict):
        section_layout = self.presentation.slide_layouts[SECTION_SLIDE_LAYOUT]
        section_slide = self.presentation.slides.add_slide(section_layout)
        section_slide.placeholders[0].text = slide["title"]

    def create_content_slide(self, slide: dict[str]):
        content_layout = self.presentation.slide_layouts[TITLE_AND_CONTENT_LAYOUT]
        content_slide = self.presentation.slides.add_slide(content_layout)
        content_slide.placeholders[0].text = slide["title"]

        slide_content: str = slide["content"]
        paragraphs: list = slide_content.split("\n")

        content_slide.placeholders[1].text = ""
        content_slide.placeholders[1].text_frame.paragraphs[0].text = paragraphs[0][2:]
        content_slide.placeholders[1].text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        content_slide.placeholders[1].text_frame.paragraphs[0].level = int(paragraphs[0][1])

        for paragraph in paragraphs[1:]:
            p = content_slide.placeholders[1].text_frame.add_paragraph()
            p.text = paragraph[2:]
            p.alignment = PP_ALIGN.LEFT
            p.level = int(paragraph[1])

    def save(self, filename):
        self.presentation.save(filename)

def create_presentation(title: str, author: str, slides: list, format: str) -> str:
    """Creates new presentation."""

    # Create presentation
    presentation = PowerpointPresentation(title, author, slides, format)
    filename = sanitize_filename(title) + ".pptx"

    # Save presentation
    presentation.save(filename)

    # Upload presentation (to be implemented).


    # Return presentation link
    return f"Link to created presentation to be shared with user: 'file:///C:/Users/Daniel/PycharmProjects/MCP-PRESENTATION/src/{filename}'. Share this exact link with user."

