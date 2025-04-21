import io
from email.mime.text import MIMEText
from email.utils import formatdate
from email.header import Header
from upload_file import upload_file


def create_eml(to=None, cc=None, bcc=None, re=None, content=None):
    """
    Creates an unsent email draft in EML format with HTML content and specific formatting.

    Args:
        to (list): List of recipient email addresses
        cc (list): List of carbon copy recipient email addresses
        bcc (list): List of blind carbon copy recipient email addresses
        re (str): Subject of the email
        content (str): HTML content to go inside the body tags

    Returns:
        io.BytesIO: A file-like object containing the EML formatted unsent draft
    """

    # Create the complete HTML document with the provided content in the body
    complete_html = f"""
    <html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <style>
            body {{
                font-family: Arial, sans-serif;
                font-size: 10pt;
                color: rgb(0, 20, 137);
            }}
        </style>
    </head>
    <body>
        {content}
    </body>
    </html>
    """

    # Create MIME text with explicit 8bit encoding to prevent "=" characters
    msg = MIMEText(complete_html, 'html', 'utf-8')
    msg.replace_header('Content-Transfer-Encoding', 'base64')

    # Set email headers
    if to:
        msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    if bcc:
        msg["Bcc"] = ", ".join(bcc)

    # Use Header object for proper UTF-8 encoding of the subject
    msg["Subject"] = Header(re, 'utf-8')
    msg["Date"] = formatdate(localtime=True)

    # Add headers to indicate this is an unsent draft
    msg["X-Unsent"] = "1"
    msg["Status"] = "RO"  # Read-Only status
    msg["X-Mozilla-Draft-Info"] = "internal/draft; vcard=0"

    # Convert message to file-like object
    buffer = io.BytesIO()
    buffer.write(msg.as_bytes())
    buffer.seek(0)

    url = upload_file(buffer, "eml")
    buffer.close()

    return url
