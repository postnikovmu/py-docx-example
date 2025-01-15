import docx


def get_file_content_python_docx(file_path: str) -> None:
    """
    This function based on the python-docx library.
    It can print paragraphs one by one.
    The paragraph can contain more then one 'Run' object,
    in case there are several different formattings inside the paragraph.


    :param file_path:
    :return:
    """
    try:
        # Open the document
        doc = docx.Document(file_path)


        # Iterate through paragraphs and print their text
        for para in doc.paragraphs:
            for run in para.runs:
                text = run.text
                is_bold = run.bold if run.bold is not None else False  # Default to False if None
                is_italic = run.italic if run.italic is not None else False  # Default to False if None
                font_size = str(run.font.size) if run.font.size else 'Not Set'
                font_color = str(run.font.color.rgb) if run.font.color and run.font.color.rgb else 'Not Set'
                font_name = str(run.font.name) if run.font.name else 'Not Set'

                print(f'Text: {text}')
                print(f'Bold: {is_bold}')
                print(f'Italic: {is_italic}')
                print(f'Font Size: {font_size}')
                print(f'Font Color: {font_color}')
                print(f'Font Name: {font_name}\n')

    except IOError:
        print('There was an error opening the file!')

    print(file_path)
