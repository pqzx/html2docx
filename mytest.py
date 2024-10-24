from docx import Document
from htmldocx import HtmlToDocx

document = Document()
new_parser = HtmlToDocx()

html = '''
<p>there are some words：<img src="testimg.png"D:/kaiyuan/html2docx/testimg.png width="100" height="50" alt="测试图片"/>，there are some words in the same line。</p>
'''

new_parser.add_html_to_document(html, document)
document.save('test_output.docx')
