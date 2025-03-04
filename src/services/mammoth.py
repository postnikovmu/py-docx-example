import mammoth

# Определение пользовательской карты стилей
style_map = "u => text-decoration: underline"

def get_file_content_mammoth_docx(file_path: str) -> None:
    with open(file_path, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file, style_map=style_map)
        html = result.value # The generated HTML
        messages = result.messages # Any messages, such as warnings during conversion
        with open('C:/temp/example_3_result_mammoth.html', "w", encoding='utf-8') as mammoth_file:
            mammoth_file.write(html)
