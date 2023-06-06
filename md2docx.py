from htmldocx import HtmlToDocx
from docx import Document
import argparse
import mistletoe
import glob

parser = argparse.ArgumentParser(description='Generate Docx reports using a Docx reference template and Markdown files')
parser.add_argument('output', default=None, help='Output file')
parser.add_argument('--files', default="*.md", help='Regex for Markdown files')
args = parser.parse_args()
print(args.files)

file_data = []
for part in sorted(glob.glob(args.files)):
    with open(part, 'r', encoding="utf-8") as f:
        file_data.append(f.read())

final_html = ''
for data in file_data:
    html = mistletoe.markdown(data)
    print(html)
    final_html += html

document = Document()
new_parser = HtmlToDocx()
# do stuff to document

new_parser.add_html_to_document(final_html, document)

# do more stuff to document
document.save(args.output)
