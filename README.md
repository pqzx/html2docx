# h2d
Converts html to docx

Dependencies: `python-docx` & `bs4` (optional, if you want to fix html before converting)

Usage: Add strings of html to an existing docx.Document object

```from parser import HtmlToDocx
from docx import Document

document = Document()
new_parser = HtmlToDocx()
# do stuff to document

html = '<h1>Hello world</h1>'
new_parser.add_html_to_document(html, document)

# do more stuff to document
document.save('your_file_name')
```

Also can use from command line to convert files.

`python parse.py html-file [docx-savefile-name]`
