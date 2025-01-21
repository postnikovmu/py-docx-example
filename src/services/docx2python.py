from docx2python import docx2python


def get_file_content_docx2python(file_path: str) -> None:
    """
    This function based on the docx2python library.

    !limitations
    Can't get text formatting
    """


    # Extract *.docx content
    doc = docx2python(file_path)

    # Initialize an empty dictionary to hold page content
    pages_content = doc.document_runs

    for i in range(len(pages_content)):  # Accessing body text
        print(pages_content[i])
