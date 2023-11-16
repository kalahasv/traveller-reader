from PyPDF2 import PdfReader

reader = PdfReader('test document.pdf')

print(len(reader.pages))

page = reader.pages[0]

text = page.extract_text()
print(text)