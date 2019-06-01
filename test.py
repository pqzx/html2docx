from docx import Document
from h2d import HtmlToDocx

document = Document()
table = document.add_table(3,3, style='Table Grid')
cell = table.cell(1,1)
parser = HtmlToDocx()

with open('text1.html', 'r') as tb:
    text1 = tb.read()

parser.add_html_to_document(text1, document)
parser.add_html_to_document(text1, cell)

with open('table.html', 'r') as tb:
    table_html = tb.read()

parser.add_html_to_document(table_html, document)
document.save('test.docx')