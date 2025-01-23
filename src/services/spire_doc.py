from spire.doc import *


def get_file_content_spire_docs(file_path: str) -> None:
    """
    This function based on the Spire.Doc library.

    ! limitation
    Initialize the license to avoid trial version limitations
    while reading the word file in python
    Can get only plain text without licence propably
    There is an exception in case of attempt to access document sections in my case

    :param file_path:
    :return:
    """
    # Create a Document object
    document = Document()
    # Load a Word document
    document.LoadFromFile(file_path)

    # Get text from the entire document
    print(document.GetText())

    # Close the document after processing
    document.Close()
