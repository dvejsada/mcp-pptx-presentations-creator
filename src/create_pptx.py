from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from upload_file import upload_file_to_s3
import io

TEMPLATE_REGULAR = "/app/src/prk_template_4_3.pptx"
TEMPLATE_WIDE = "/app/src/prk_template_16_9.pptx"
# TEMPLATE_REGULAR = "prk_template_4_3.pptx"
# TEMPLATE_WIDE = "prk_template_16_9.pptx"

TITLE_SLIDE_LAYOUT = 2
SECTION_SLIDE_LAYOUT = 7
TITLE_AND_CONTENT_LAYOUT = 4

class PowerpointPresentation:

    def __init__(self, author: str, slides: list, format: str):

        self.author = author
        self.slides: list = slides

        if format == "4:3":
            self.presentation = Presentation(TEMPLATE_REGULAR)
        elif format == "16:9":
            self.presentation = Presentation(TEMPLATE_WIDE)
        else:
            self.presentation = Presentation(TEMPLATE_REGULAR)


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

    # Upload presentation (to be implemented).

    url = upload_file_to_s3(file_object)

    file_object.close()

    # Return presentation link
    return f"Link to created presentation to be shared with user: {url} . Link is valid for 1 hour."

