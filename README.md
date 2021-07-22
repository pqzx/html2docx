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

Change table styles

Tables are not styled by default. Use the `table_style` attribute on the parser to set a table
style. The style is used for all tables.

```
from htmldocx import HtmlToDocx

new_parser = HtmlToDocx()
new_parser.table_style = 'Light Shading Accent 4'
```

To add borders to tables, use the `TableGrid` style:

```
new_parser.table_style = 'TableGrid'
```

Default table styles can be found here: https://python-docx.readthedocs.io/en/latest/user/styles-understanding.html#table-styles-in-default-template
