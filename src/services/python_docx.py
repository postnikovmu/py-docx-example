import docx
from lxml import etree
import zipfile


def get_word_xml(docx_filename):
    with zipfile.ZipFile(docx_filename) as zip_ref:
        xml_content = zip_ref.read('word/document.xml')
    return xml_content


def parse_xml(xml_string):
    return etree.fromstring(xml_string)


def print_text_elements(root):
    for elem in root.iter():
        if elem.tag.endswith('t'):
            if elem.text is not None:
                print("Текст из XML:", elem.text)


def get_file_content_python_docx(file_path: str) -> None:
    """
    This function based on the python-docx library.
    It can print paragraphs one by one.
    The paragraph can contain more then one 'Run' object,
    in case there are several different formattings inside the paragraph.

    ! limitation
    Can't get formatting for old *.doc files
    The primary reason python-docx cannot retrieve
    formatting from old .doc files is due to the inherent differences between the two formats.
    The .doc format uses a binary structure that encapsulates text and formatting in a way that is not compatible
    with the XML-based approach of .docx. This means that any code written for python-docx,
    which relies on parsing XML elements, will not function correctly when applied to .doc files

    :param file_path:
    :return:
    """
    try:
        # Open the document
        doc = docx.Document(file_path)

        default_style = doc.styles['Normal']
        print('Имя шрифта по умолчанию:', default_style.font.name)
        print('Размер шрифта по умолчанию:', default_style.font.size)

        for style in doc.styles:
            print(style.name, style.style_id)
            print(dir(style))
            #font = style.font

            # Получение параметров шрифта
            #font_name = font.name  # Название шрифта
            #font_size = font.size  # Размер шрифта (в пунктах)
            #font_bold = font.bold  # Жирный текст (True/False)
            #font_italic = font.italic  # Курсив (True/False)
            #font_underline = font.underline  # Подчеркивание (None/Single/Double)
            #print(font_name, font_size, font_bold, font_italic, font_underline)

        # Iterate through all elements and print their text
        index_for_paragraphs = None
        index_for_tables = None

        for element_num in range(len(doc.element.body.inner_content_elements)):

            element = doc.element.body.inner_content_elements[element_num]
            element_type = None
            if element.tag.endswith('p'):  #check for a paragraph
                element_type = 'paragraph'
                if index_for_paragraphs is None:
                    index_for_paragraphs = 0
                else:
                    index_for_paragraphs += 1

            elif element.tag.endswith('tbl'):  #check for a paragraph
                element_type = 'table'
                if index_for_tables is None:
                    index_for_tables = 0
                else:
                    index_for_tables += 1

            if element_type == 'paragraph':
                element_content = doc.paragraphs[index_for_paragraphs]
                print(element_content.text)

                for run in element_content.runs:
                    text = run.text
                    is_bold = run.bold if run.bold is not None else False  # Default to False if None
                    is_italic = run.italic if run.italic is not None else False  # Default to False if None
                    font_size = str(run.font.size) if run.font.size else 'Not Set'
                    font_color = str(run.font.color.rgb) if run.font.color and run.font.color.rgb else 'Not Set'
                    font_name = str(run.font.name) if run.font.name else 'Not Set'

                    print(f'Run Text: {text}')
                    print(f'Run Bold: {is_bold}')
                    print(f'Run Italic: {is_italic}')
                    print(f'Run Font Size: {font_size}')
                    print(f'Run Font Color: {font_color}')
                    print(f'Run Font Name: {font_name}\n')

            if element_type == 'table':
                print('>>>>>>Table')
                element_content = doc.tables[index_for_tables]
                for row in element_content.rows:
                    for cell in row.cells:
                        # print data from the cell
                        for paragraph in cell.paragraphs:
                            print(paragraph.text)
                print('<<<<<<Table')

        xml_content = get_word_xml(file_path)
        root = parse_xml(xml_content)
        print_text_elements(root)

    except IOError:
        print('There was an error opening the file!')

