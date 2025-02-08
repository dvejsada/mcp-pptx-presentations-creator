from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from upload_file import upload_file_to_s3
from pathlib import Path
import io
import logging

TITLE_SLIDE_LAYOUT = 2
SECTION_SLIDE_LAYOUT = 7
TITLE_AND_CONTENT_LAYOUT = 4

# Create a logger
logger = logging.getLogger(__name__)

def load_templates():
    """Loads presentation teplates"""

    custom_template_4_3 = Path("../templates/template_4_3.pptx")
    custom_template_16_9 = Path("../templates/template_4_3.pptx")

    if custom_template_4_3.exists():
        template_4_3 = custom_template_4_3
        logger.info("Custom 4:3 template loaded.")
    else:
        template_4_3 = Path("general_template_4_3.pptx")
        logger.info("General 4:3 template loaded.")

    if custom_template_4_3.exists():
        template_16_9 = custom_template_16_9
        logger.info("Custom 16:9 template loaded.")
    else:
        template_16_9 = Path("general_template_16_9.pptx")
        logger.info("General 16:9 template loaded.")

    return str(template_4_3), str(template_16_9)


class PowerpointPresentation:

    def __init__(self, author: str, slides: list, format: str):

        self.author = author
        self.slides: list = slides
        self.template_regular, self.template_wide = load_templates()

        if format == "4:3":
            self.presentation = Presentation(self.template_regular)
        elif format == "16:9":
            self.presentation = Presentation(self.template_wide)
        else:
            self.presentation = Presentation(self.template_regular)

        for slide in self.slides:
            if slide["slide_type"] == 2:
                self.create_content_slide(slide)
            elif slide["slide_type"] == 1:
                self.create_section_header_slide(slide)
            elif slide["slide_type"] == 0:
                self.create_title_slide(slide)


    def create_title_slide(self, slide: dict):
        title_layout = self.presentation.slide_layouts[TITLE_SLIDE_LAYOUT]
        title_slide = self.presentation.slides.add_slide(title_layout)
        title_slide.placeholders[0].text = slide["slide_title"]
        title_slide.placeholders[1].text = self.author

    def create_section_header_slide(self, slide: dict):
        section_layout = self.presentation.slide_layouts[SECTION_SLIDE_LAYOUT]
        section_slide = self.presentation.slides.add_slide(section_layout)
        section_slide.placeholders[0].text = slide["slide_title"]

    def create_content_slide(self, slide: dict[str]):
        content_layout = self.presentation.slide_layouts[TITLE_AND_CONTENT_LAYOUT]
        content_slide = self.presentation.slides.add_slide(content_layout)
        content_slide.placeholders[0].text = slide["slide_title"]

        slide_content: str = slide["slide_text"]
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

    def save(self):
        file_like_object = io.BytesIO()
        self.presentation.save(file_like_object)
        file_like_object.seek(0)
        return file_like_object

def create_presentation(author: str, slides: list, format: str) -> str:
    """Creates new presentation."""

    # Create presentation
    presentation = PowerpointPresentation(author, slides, format)

    # Save presentation
    file_object = presentation.save()

    # Upload presentation.
    url = upload_file_to_s3(file_object)
    file_object.close()

    # Return presentation link
    return f"Link to created presentation to be shared with user: {url} . Link is valid for 1 hour."

