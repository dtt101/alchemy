import mammoth
from pathlib import Path

HERE = Path(__file__).resolve().parent
DOCX = HERE / "test.docx"
with open(DOCX, "rb") as test_doc:
    result = mammoth.convert_to_html(test_doc)
    html = result.value
    messages = result.messages

print('MAMMOTH')
print(html)
print(messages)

from markitdown import MarkItDown

md = MarkItDown(enable_plugins=False) # Set to True to enable plugins
result = md.convert(DOCX)
print('MARKITDOWN')
print(result.text_content)
print(result.markdown)
print(result.title)
