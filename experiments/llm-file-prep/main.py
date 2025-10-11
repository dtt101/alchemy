import mammoth
from pptx import Presentation
from pathlib import Path

HERE = Path(__file__).resolve().parent
DOCX = HERE / "test.docx"
PPTX = HERE / "test.pptx"
with open(DOCX, "rb") as test_doc:
    result = mammoth.convert_to_html(test_doc)
    html = result.value
    messages = result.messages

print('MAMMOTH')
print(html)
print(messages)

# TODO 
# python-pptx for ppt
text_runs = []
prs = Presentation(PPTX)
for slide in prs.slides:
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                text_runs.append(run.text)

print("PYTHON-PPTX")
print(text_runs)
# TODO 
# pdfminer.six for pdf
#
#
# TODO - convert HTML to markdown with beautifulsoup
