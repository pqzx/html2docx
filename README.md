# html3docx
A fork of https://github.com/pqzx/html2docx.  This version will focus on expedient changes for our particular use case.

Dependencies: `python-docx` & `bs4`

### To install

`pip install html3docx`

### Improvements

- Fix for KeyError when handling an img tag without a src attribute.
- Images with a width attribute will be scaled according to that width.
- Fix for AttributeError when handling a leading br tag, either at the top of the HTML snippet, or within a td or th cell.

## Original README

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

Default table styles can be found
here: https://python-docx.readthedocs.io/en/latest/user/styles-understanding.html#table-styles-in-default-template

Change default paragraph style

No style is applied to the paragraphs by default. Use the `paragraph_style` attribute on the parser
to set a default paragraph style. The style is used for all paragraphs. If additional styling (
color, background color, alignment...) is defined in the HTML, it will be applied after the
paragraph style.

```
from htmldocx import HtmlToDocx

new_parser = HtmlToDocx()
new_parser.paragraph_style = 'Quote'
```

Default paragraph styles can be found
here: https://python-docx.readthedocs.io/en/latest/user/styles-understanding.html#paragraph-styles-in-default-template
