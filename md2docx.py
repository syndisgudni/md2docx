from bs4 import BeautifulSoup as bs
from h2d import HtmlToDocx

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, Inches, RGBColor

import argparse
import mistletoe
import glob

arg_parser = argparse.ArgumentParser(description='Generate Docx reports using a Docx reference template and Markdown files')
arg_parser.add_argument('output', default=None, help='Output file')
arg_parser.add_argument('--files', default="*.md", help='Regex for Markdown files')
args = arg_parser.parse_args()

def apply_style(document):
    ## Page setup
    section = document.sections[-1]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.59)
    section.bottom_margin = Inches(0.59)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)

    ## Normal
    style = document.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    ## Code
    style = document.styles.add_style('Code', WD_STYLE_TYPE.CHARACTER)
    style.font.name = 'Roboto Mono'
    style.font.size = Pt(9)

    ## Headings
    style = document.styles['Heading 1']
    style.font.name = 'Calibri'
    style.font.size = Pt(20)
    style.font.bold = True
    style.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    style = document.styles['Heading 2']
    style.font.name = 'Calibri'
    style.font.size = Pt(18)
    style.font.bold = True
    style.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    style = document.styles['Heading 3']
    style.font.name = 'Calibri'
    style.font.size = Pt(14)
    style.font.bold = True
    style.font.color.rgb = RGBColor(0x43, 0x43, 0x43)
    style = document.styles['Heading 4']
    style.font.name = 'Calibri'
    style.font.size = Pt(12)
    style.font.bold = False
    style.font.italic = False
    style.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    # TODO: Remove empty first lines from all tables

file_data = []
for part in sorted(glob.glob(args.files)):
    with open(part, 'r', encoding="utf-8") as f:
        file_data.append(f.read())

html = ''.join([mistletoe.markdown(data) for data in file_data])

soup = bs(html, 'html.parser')

html = soup.prettify("utf-8").decode()
open('test.html','w').write(html)
print(html)

document = Document()
parser = HtmlToDocx()

parser.add_html_to_document(html, document)

apply_style(document)

# do more stuff to document
document.save(args.output)
