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

# TODO 
# python-pptx for ppt
#
# TODO 
# pdfminer.six for pdf
#
#
# TODO - convert HTML to markdown with beautifulsoup
