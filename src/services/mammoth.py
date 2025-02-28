import mammoth


def get_file_content_mammoth_docx(file_path: str) -> None:
    with open(file_path, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html = result.value # The generated HTML
        messages = result.messages # Any messages, such as warnings during conversion
        with open('C:/temp/example_1_result_mammoth.html', "w", encoding='utf-8') as mammoth_file:
            mammoth_file.write(html)
