import markdown
import io
from docx import Document
from bs4 import BeautifulSoup
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE
from upload_file import upload_file

def add_heading(doc, element, level=1):
    """
    Add a heading

    Parameters:
    doc: documents to which the heading should be added
    element: element with heading
    level: The heading level (1-9)

    """
    # Add a heading first (this will be empty initially)
    heading = doc.add_heading('', level=level)

    child_elements(element, heading)

def add_hyperlink(paragraph, text, url, color="0000FF", underline=True):
    """
    Adds a hyperlink to a paragraph at the current run position.

    :param paragraph: The paragraph to add the hyperlink to.
    :param text: The display text for the hyperlink.
    :param url: The URL the hyperlink points to.
    :param color: Hex code for the hyperlink text color.
    :param underline: Whether to underline the hyperlink text.
    """

    # This gets access to the document.xml.rels file and gets a new relation id
    part = paragraph.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Create the hyperlink element
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    # Create a new run element
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    # Set run properties
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


def add_table(html_element, document: Document):
    # Find all rows
    rows = html_element.find_all('tr')

    # Determine number of columns (use the row with the most cells)
    max_cols = max(len(row.find_all(['th', 'td'])) for row in rows)

    # Create a Word table
    word_table = document.add_table(rows=len(rows), cols=max_cols)
    word_table.style = 'Table Grid'  # Apply a table style

    # Process each row
    for i, row in enumerate(rows):
        cells = row.find_all(['th', 'td'])

        # Process each cell
        for j, cell in enumerate(cells):
            if j < max_cols:  # Ensure we don't exceed the number of columns
                # Get style attributes if any
                style = cell.get('style', '')

                # Add content to Word table cell
                word_cell = word_table.cell(i, j)

                # Clear the default paragraph text that's automatically created
                # when a cell is created
                if word_cell.paragraphs:
                    word_cell.paragraphs[0].clear()

                # Use the cell's first paragraph to add styled content
                cell_paragraph = word_cell.paragraphs[0]

                # Process the cell content while preserving formatting
                child_elements(cell, cell_paragraph)

                # Apply alignment if specified in style
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
    """
    Process child elements and add them to the paragraph.
    """
    for child in element.children:
        if isinstance(child, str):
            # If it's plain text, add it as a run
            if child.strip():
                paragraph.add_run(child)
        elif child.name == 'strong':
            paragraph.add_run(child.text).bold = True
        elif child.name == 'em':
            paragraph.add_run(child.text).italic = True
        elif child.name == 'a':
            add_hyperlink(paragraph, child.text, child.get('href'))
        elif child.name in ['ul', 'ol']:
            # Skip lists - they're handled in process_list
            continue
        elif hasattr(child, 'children') and list(child.children):
            # Recursively process other elements with children
            child_elements(child, paragraph)
        elif hasattr(child, 'text'):
            # Handle any other element with text
            paragraph.add_run(child.text)

def process_list(list_element, doc, base_style, level=0):
    """
    Process a list element (ul or ol) and handle nested lists of any type.

    Parameters:
    list_element: The list element to process (ul or ol)
    doc: The document to add the list to
    base_style: The base style ('List Bullet' or 'List Number')
    level: The nesting level (0 for top level, 1+ for nested levels)
    """
    # Different types of list styles in Word
    bullet_styles = ['List Bullet', 'List Bullet 2', 'List Bullet 3']
    number_styles = ['List Number', 'List Number 2', 'List Number 3']

    # Determine which style to use based on the list type and level
    if base_style.startswith('List Bullet'):
        style = bullet_styles[min(level, len(bullet_styles) - 1)]
    else:  # List Number
        style = number_styles[min(level, len(number_styles) - 1)]

    # Process each list item
    for li in list_element.find_all('li', recursive=False):
        # Create paragraph with appropriate list style
        par = doc.add_paragraph(style=style)

        # Process the direct content of the list item (excluding nested lists)
        process_list_item_content(li, par)

        # Process any nested lists regardless of type
        for nested_list in li.find_all(['ul', 'ol'], recursive=False):
            if nested_list.name == 'ul':
                process_list(nested_list, doc, 'List Bullet', level + 1)
            else:  # ol
                process_list(nested_list, doc, 'List Number', level + 1)

def process_list_item_content(li, paragraph):
    """
    Process only the direct content of a list item, excluding nested lists.
    """
    # Create a copy of the list item to work with
    li_copy = BeautifulSoup(str(li), 'html.parser').li

    # Remove any nested lists from the copy
    for nested_list in li_copy.find_all(['ul', 'ol']):
        nested_list.extract()

    # Now process the remaining content
    for child in li_copy.children:
        if isinstance(child, str):
            # If it's plain text, add it as a run
            if child.strip():
                paragraph.add_run(child)
        elif child.name == 'strong':
            paragraph.add_run(child.text).bold = True
        elif child.name == 'em':
            paragraph.add_run(child.text).italic = True
        elif child.name == 'a':
            add_hyperlink(paragraph, child.text, child.get('href'))
        elif hasattr(child, 'children') and list(child.children):
            # Recursively process other elements with children (except ul/ol)
            if child.name not in ['ul', 'ol']:
                child_elements(child, paragraph)
        elif hasattr(child, 'text') and child.name not in ['ul', 'ol']:
            # Handle any other element with text
            paragraph.add_run(child.text)

def markdown_to_word(markdown_content, word_file):
    # Converting Markdown to HTML
    try:
        html_content = markdown.markdown(markdown_content, extensions=['tables', 'sane_lists'])

    except Exception as e:
        return f"Error in markdown: {e}"

    # Creating a new Word Document
    doc = Document("template.docx")

    # Converting HTML to text and add it to the Word Document
    soup = BeautifulSoup(html_content, 'html.parser')

    # Adding content to the Word Document

    try:
        for element in soup:
            if element.name == 'h1':
                add_heading(doc, element, level=1)
            elif element.name == 'h2':
                add_heading(doc, element, level=2)
            elif element.name == 'h3':
                add_heading(doc, element, level=3)
            elif element.name == 'p':
                paragraph = doc.add_paragraph()
                child_elements(element, paragraph)
            elif element.name == 'ul':
                process_list(element, doc, 'List Bullet', 0)
            elif element.name == 'ol':
                process_list(element, doc, 'List Number', 0)
            elif element.name == 'table':
                add_table(element, doc)
    except Exception as e:
        return f"Error in parsing html: {e}"

    file_like_object = io.BytesIO()
    doc.save(file_like_object)
    file_like_object.seek(0)

    url = upload_file(file_like_object, "docx")

    file_like_object.close()

    return url

if __name__ == "__main__":
    md = """# Project Title

Simple overview of use/purpose.

## Description

An in-depth paragraph about your project and overview of use.


## Help

Any advise for common problems or issues.

## Authors

Contributors names and contact info

ex. Dominique Pizzie  
ex. [@DomPizzie](https://twitter.com/dompizzie)

## Version History

1. 0.2
    * See **commit change** or See **release history
    * See other fings
2. 0.2
    1. See **commit change** or See **release history**
    2. see other things
        * Third level list

## License

This project is licensed under the [NAME HERE] License - see the LICENSE.md file for details

| Feature | Description | Example |
|---------|-------------|---------|
| **Formatting** | Support for *text* styles | **Bold** and *italic* |
| [Links](#) | Hyperlinks to resources | [Documentation](https://example.com) |
| Lists | Ordered and unordered | See the list section |
| Alignment | Text alignment options | Right-aligned text |

## Acknowledgments

Inspiration, code snippets, etc.
* [awesome-readme](https://github.com/matiassingers/awesome-readme)
* [PurpleBooth](https://gist.github.com/PurpleBooth/109311bb0361f32d87a2)
* [dbader](https://github.com/dbader/readme-template)
* [zenorocha](https://gist.github.com/zenorocha/4526327)
* [fvcproductions](https://gist.github.com/fvcproductions/1bfc2d4aecb01a834b46)
"""
    print(markdown_to_word(md, "test.docx"))