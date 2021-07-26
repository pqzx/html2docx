import os
from .context import HtmlToDocx, test_dir

# Manual test (requires inspection of result) for converting code and pre blocks.

filename = os.path.join(test_dir, 'code.html')
d = HtmlToDocx()

d.parse_html_file(filename)
