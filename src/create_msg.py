import io
from email.mime.text import MIMEText
from email.utils import formatdate
from email.header import Header
from upload_file import upload_file


def create_eml(to=None, cc=None, bcc=None, re=None, content=None, priority="normal", language="cs-CZ"):
    """
    Creates an unsent email draft in EML format with HTML content and specific formatting.

    Args:
        to (list): List of recipient email addresses
        cc (list): List of carbon copy recipient email addresses
        bcc (list): List of blind carbon copy recipient email addresses
        re (str): Subject of the email
        content (str): HTML content to go inside the body tags
        priority (str): Email priority ("low", "normal", "high")
        language (str): Language code for proofreading (e.g., "cs-CZ", "en-US", "de-DE")

    Returns:
        str: URL to the uploaded EML file
        
    Raises:
        ValueError: If priority is not valid or required parameters are missing
        Exception: If file upload fails
    """
    
    # Validate priority
    if priority.lower() not in ["low", "normal", "high"]:
        raise ValueError("Priority must be 'low', 'normal', or 'high'")
    
    # Validate required parameters
    if not content:
        raise ValueError("Email content is required")
    if not re:
        raise ValueError("Email subject is required")

    # Create the complete HTML document with the provided content in the body
    complete_html = f"""
    <html lang="{language}">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <meta http-equiv="Content-Language" content="{language}">
        <style>
            body {{
                font-family: Arial, sans-serif;
                font-size: 10pt;
                color: rgb(0, 20, 137);
                line-height: 1.4;
            }}
            h2 {{
                font-weight: bold;
                font-size: 10pt;
                margin: 8px 0;
            }}
            h3 {{
                text-decoration: underline;
                font-size: 10pt;
                font-weight: normal;
                margin: 8px 0;
            }}
            p {{
                margin: 8px 0;
            }}
            ul, ol {{
                margin: 8px 0;
                padding-left: 20px;
            }}
            li {{
                margin: 3px 0;
            }}
        </style>
    </head>
    <body lang="{language}">
        {content}
    </body>
    </html>
    """

    buffer = None
    try:
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

        # Set language headers for email clients
        msg["Content-Language"] = language
        msg["Accept-Language"] = language

        # Set priority headers
        if priority.lower() == "high":
            msg["X-Priority"] = "1 (Highest)"
            msg["X-MSMail-Priority"] = "High"
            msg["Importance"] = "High"
        elif priority.lower() == "low":
            msg["X-Priority"] = "5 (Lowest)"
            msg["X-MSMail-Priority"] = "Low"
            msg["Importance"] = "Low"

        # Add headers to indicate this is an unsent draft
        msg["X-Unsent"] = "1"

        # Convert message to file-like object
        buffer = io.BytesIO()
        buffer.write(msg.as_bytes())
        buffer.seek(0)

        url = upload_file(buffer, "eml")
        return url
        
    except Exception as e:
        raise Exception(f"Failed to create email draft: {str(e)}")
    finally:
        if buffer:
            buffer.close()
