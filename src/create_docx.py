from os.path import exists
import io
from docx import Document
from bs4 import BeautifulSoup
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE
from upload_file import upload_file
from pathlib import Path

def load_templates():
    """Loads presentation templates"""
    custom_template = Path("/app/templates/template.docx")
    """ custom_template = Path("template.docx")"""
    if exists(custom_template):
        template = custom_template
    else:
        template = Path("/app/src/template.docx")
    return str(template)

def add_heading(doc, element, level=1):
    """Add a heading"""
    heading = doc.add_heading('', level=level)
    child_elements(element, heading)

def add_hyperlink(paragraph, text, url, color="0000FF", underline=True):
    """Adds a hyperlink to a paragraph"""
    part = paragraph.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    if underline:
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)

    if color:
        c = OxmlElement('w:color')
        c.set(qn('w:val'), color)
        rPr.append(c)

    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)

    paragraph._p.append(hyperlink)

def add_table(html_element, document):
    rows = html_element.find_all('tr')
    max_cols = max(len(row.find_all(['th', 'td'])) for row in rows)

    word_table = document.add_table(rows=len(rows), cols=max_cols)
    word_table.style = 'Table Grid'

    for i, row in enumerate(rows):
        cells = row.find_all(['th', 'td'])
        for j, cell in enumerate(cells):
            if j < max_cols:
                style = cell.get('style', '')
                word_cell = word_table.cell(i, j)

                if word_cell.paragraphs:
                    word_cell.paragraphs[0].clear()

                cell_paragraph = word_cell.paragraphs[0]
                child_elements(cell, cell_paragraph)

                if 'text-align: right' in style:
                    for paragraph in word_cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                elif 'text-align: center' in style:
                    for paragraph in word_cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif 'text-align: left' in style:
                    for paragraph in word_cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

def child_elements(element, paragraph):
    """Process child elements and add them to the paragraph."""
    for child in element.children:
        if isinstance(child, str):
            if child.strip():
                paragraph.add_run(child)
        elif child.name == 'strong':
            paragraph.add_run(child.text).bold = True
        elif child.name == 'em':
            paragraph.add_run(child.text).italic = True
        elif child.name == 'a':
            add_hyperlink(paragraph, child.text, child.get('href'))
        elif child.name in ['ul', 'ol']:
            # Skip lists - they're handled separately
            continue
        elif hasattr(child, 'children') and list(child.children):
            child_elements(child, paragraph)
        elif hasattr(child, 'text'):
            paragraph.add_run(child.text)

def process_list(list_element, doc, level=0):
    """Process a list element and all its items with proper nesting."""
    # Different types of list styles in Word based on level
    bullet_styles = ['List Bullet', 'List Bullet 2', 'List Bullet 3']
    number_styles = ['List Number', 'List Number 2', 'List Number 3']

    # Determine which style array to use based on the list type
    style_array = number_styles if list_element.name == 'ol' else bullet_styles

    # Get the appropriate style for this level (capped at highest available level)
    style = style_array[min(level, len(style_array) - 1)]

    # Process each list item at this level
    for li in list_element.find_all('li', recursive=False):
        # Create paragraph with appropriate list style
        paragraph = doc.add_paragraph(style=style)

        # Process direct content of the list item (excluding nested lists)
        for child in li.children:
            if isinstance(child, str):
                # Add text content
                if child.strip():
                    paragraph.add_run(child.strip())
            elif child.name in ['ul', 'ol']:
                # Skip nested lists - they'll be processed separately
                continue
            elif child.name == 'strong':
                paragraph.add_run(child.text).bold = True
            elif child.name == 'em':
                paragraph.add_run(child.text).italic = True
            elif child.name == 'a':
                add_hyperlink(paragraph, child.text, child.get('href'))
            elif hasattr(child, 'text'):
                paragraph.add_run(child.text)

        # Now process any nested lists within this list item
        for nested_list in li.find_all(['ul', 'ol'], recursive=False):
            process_list(nested_list, doc, level + 1)


def html_to_word(html_content):
    """Convert HTML directly to Word document."""
    path = load_templates()
    doc = Document(path)

    # Parse the HTML content
    soup = BeautifulSoup(html_content, 'html.parser')
    body = soup.body or soup  # Use body if available, otherwise use the soup object

    try:
        # Process each element in the body
        for element in body.children:
            if hasattr(element, 'name'):
                if element.name == 'h1':
                    add_heading(doc, element, level=1)
                elif element.name == 'h2':
                    add_heading(doc, element, level=2)
                elif element.name == 'h3':
                    add_heading(doc, element, level=3)
                elif element.name == 'h4':
                    add_heading(doc, element, level=4)
                elif element.name == 'p':
                    paragraph = doc.add_paragraph()
                    child_elements(element, paragraph)
                elif element.name == 'ul' or element.name == 'ol':
                    process_list(element, doc, 0)
                elif element.name == 'table':
                    add_table(element, doc)
    except Exception as e:
        return f"Error in parsing html: {e}"

    # Save the document
    file_like_object = io.BytesIO()
    doc.save(file_like_object)
    file_like_object.seek(0)

    url = upload_file(file_like_object, "docx")
    file_like_object.close()

    return url


if __name__ == "__main__":
    html = "<!DOCTYPE html><html><body><h1>Document Testing Lists and Tables</h1><h2>Introduction</h2><p>This document tests the formatting of various HTML elements, with focus on nested lists and tables.</p><h2>List Examples</h2><h3>Simple Lists</h3><h4>Bullet List:</h4><ul><li>Item 1</li><li>Item 2</li><li>Item 3</li></ul><h4>Numbered List:</h4><ol><li>First item</li><li>Second item</li><li>Third item</li></ol><h3>Nested Lists</h3><h4>Nested Bullet Lists:</h4><ul><li>Main item 1<ul><li>Sub-item 1.1</li><li>Sub-item 1.2<ul><li>Sub-sub-item 1.2.1</li><li>Sub-sub-item 1.2.2</li></ul></li></ul></li><li>Main item 2<ul><li>Sub-item 2.1</li></ul></li></ul><h4>Nested Numbered Lists:</h4><ol><li>First level item 1<ol><li>Second level item 1.1</li><li>Second level item 1.2<ol><li>Third level item 1.2.1</li><li>Third level item 1.2.2</li></ol></li></ol></li><li>First level item 2<ol><li>Second level item 2.1</li></ol></li></ol><h4>Mixed Nested Lists:</h4><ul><li>Bullet main item<ol><li>Numbered sub-item 1</li><li>Numbered sub-item 2<ul><li>Bullet sub-sub-item</li><li>Another bullet sub-sub-item</li></ul></li></ol></li></ul><ol><li>Numbered main item<ul><li>Bullet sub-item 1</li><li>Bullet sub-item 2<ol><li>Numbered sub-sub-item</li><li>Another numbered sub-sub-item</li></ol></li></ul></li></ol><h3>Multiple Separate Lists</h3><p>Here we test if consecutive lists restart numbering properly:</p><ol><li>List 1, Item 1</li><li>List 1, Item 2</li></ol><p>Some text between lists.</p><ol><li>List 2, Item 1 (should start at 1 again)</li><li>List 2, Item 2</li></ol><h2>Table Examples</h2><h3>Simple Table</h3><table border=\"1\"><tr><th>Header 1</th><th>Header 2</th><th>Header 3</th></tr><tr><td>Cell 1</td><td>Cell 2</td><td>Cell 3</td></tr><tr><td>Cell 4</td><td>Cell 5</td><td>Cell 6</td></tr></table><h3>Table with Formatting</h3><table border=\"1\"><tr><th>Name</th><th>Description</th><th>Status</th></tr><tr><td><strong>Project A</strong></td><td><em>An important initiative</em></td><td>Active</td></tr><tr><td><strong>Project B</strong></td><td><em>Secondary project</em></td><td>Pending</td></tr></table><h2>Text Formatting and Links</h2><p>This paragraph has <strong>bold text</strong>, <em>italic text</em>, and <a href=\"https://www.example.com\">a hyperlink</a>.</p></body></html>"

    html_to_word(html)


