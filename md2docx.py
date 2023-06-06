from bs4 import BeautifulSoup as bs
from htmldocx import HtmlToDocx
from docx import Document
import argparse
import mistletoe
import glob

parser = argparse.ArgumentParser(description='Generate Docx reports using a Docx reference template and Markdown files')
parser.add_argument('output', default=None, help='Output file')
parser.add_argument('--files', default="*.md", help='Regex for Markdown files')
args = parser.parse_args()

file_data = []
for part in sorted(glob.glob(args.files)):
    with open(part, 'r', encoding="utf-8") as f:
        file_data.append(f.read())

tmp_html = ''
for data in file_data:
    html = mistletoe.markdown(data)
    tmp_html += html
html = tmp_html


soup = bs(html, 'html.parser')


img_tags = soup.find_all('img')
for img_tag in img_tags:
    img_tag['style'] = 'width: 6.27in; height: auto;'






html = soup.prettify("utf-8").decode()
open('test.html','w').write(html)
print(html)

document = Document()
new_parser = HtmlToDocx()

# Specify style=True to include styles
new_parser.add_html_to_document(html, document)

# do more stuff to document
document.save(args.output)
