import os
from .context import HtmlToDocx, test_dir

# Manual test (requires inspection of result) for converting html with nested tables

filename = os.path.join(test_dir, 'tables2.html')
d = HtmlToDocx()

d.parse_html_file(filename)
