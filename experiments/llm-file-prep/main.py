import mammoth

with open("test.docx", "rb") as test_doc:
    result = mammoth.convert_to_html(test_doc)
    html = result.value
    messages = result.messages

print(html)
print(messages)
