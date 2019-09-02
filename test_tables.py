from h2d import HtmlToDocx

filename = 'tables.html'
d = HtmlToDocx()

d.parse_html_file(filename)
