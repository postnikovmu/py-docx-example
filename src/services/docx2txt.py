import docx2txt


def get_file_content_docx2txt(file_path: str) -> None:
    """
    This function based on the docx2txt library.

    !limitations
    Can't get text formatting
    """

    # Extract text from the document
    text = docx2txt.process(file_path)

    # Print extracted text
    print(text)
