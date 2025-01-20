import aspose.words as aw


def get_file_content_aspose_words(file_path: str) -> None:
    """
    This function based on the python-docx library.
    It can print paragraphs one by one.
    To extract text along with its formatting,
    you will iterate through the paragraphs and runs within those paragraphs.
    Each run contains text and its associated formatting properties.

    ! limitation
    Initialize the license to avoid trial version limitations
    while reading the word file in python

    :param file_path:
    :return:
    """

    try:

        # Load the document
        doc = aw.Document(file_path)

        pass

        # Iterate through each paragraph in the document
        for para in doc.get_child_nodes(aw.NodeType.PARAGRAPH, True):
            para = para.as_paragraph()
            print(para.to_string(aw.SaveFormat.TEXT))

            # Print the paragraph text
            print("Paragraph Text:", text := para.get_text().strip())

            # can't iterate, licence limitations started to work
            # Iterate through each run (a segment of text with consistent formatting)
            # for run in para.runs:
            #     # Get text from the run
            #     run_text = run.get_text()
            #
            #     # Get formatting properties
            #     is_bold = run.font.bold
            #     font_color = run.font.color.to_string() if run.font.color else "Default"
            #
            #     # Print extracted information
            #     print(f"Run Text: {run_text.strip()}, Bold: {is_bold}, Font Color: {font_color}")

    except IOError:
        print('There was an error opening the file!')
