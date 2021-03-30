# htmldocx
Convert html to docx

Dependencies: `python-docx` & `bs4`

### To install

`pip install htmldocx`

### Usage

Add strings of html to an existing docx.Document object

```
from docx import Document
from htmldocx import HtmlToDocx

document = Document()
new_parser = HtmlToDocx()
# do stuff to document

html = '<h1>Hello world</h1>'
new_parser.add_html_to_document(html, document)

# do more stuff to document
document.save('your_file_name')
```

Convert files directly

```
from htmldocx import HtmlToDocx

new_parser = HtmlToDocx()
new_parser.parse_html_file(input_html_file_path, output_docx_file_path)
```

Convert files from a string

```
from htmldocx import HtmlToDocx

new_parser = HtmlToDocx()
docx = new_parser.parse_html_string(input_html_file_string)
```
