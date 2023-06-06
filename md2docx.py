import os
import shutil
import glob
import re
import itertools
import argparse

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, Inches, RGBColor
import mistune


class DocxRenderer(mistune.renderers.BaseRenderer):
    # utility
    def text_strip(self, text):
        text = text.rstrip()

        while text.startswith("p.add_run(\""):
            text = text.removeprefix("p.add_run(\"")
            text = text.removesuffix("\").style = document.styles['Code']")
            text = text.removesuffix("\").italic = True")
            text = text.removesuffix("\").bold = True")
            text = text.removesuffix("\")\n")
            text = text.removesuffix("\")")
        return text

    # inline level
    def text(self, text):
        text = text.replace('\n', ' ')
        return "p.add_run(\"%s\")\n" % text

    def emphasis(self, text):
        return "%s.italic = True\n" % text.rstrip()

    def strong(self, text):
        return "%s.bold = True\n" % text.rstrip()

    def codespan(self, text):
        text = text.replace('\n', ' ')
        return "p.add_run(\"%s\").style = document.styles['Code']\n" % text

    # block level
    def paragraph(self, text):
        if 'add_picture' in text:
            return text
        add_break = '' if text.endswith(':")\n') else 'p.add_run().add_break()'
        return '\n'.join(('p = document.add_paragraph()', text, add_break)) + '\n'

    def heading(self, text, level):
        return "p = document.add_heading('', %d)\n" % level + text

    def block_text(self, text):
        return "%s\n" % text # TODO

    def list(self, text, ordered, level, start=None):
        output = ''
        list_style = 'List '
        if ordered: list_style += 'Number'
        else: list_style += 'Bullet'
        for item in text.split('\n'):
            if item == '': break
            output += "p = document.add_paragraph('%s', style = '%s')\n" % (item, list_style)
        # return text % list_style
        return output + '\n'

    def list_item(self, text, level):
        text = text.replace('\n', ' ')
        # return "p = document.add_paragraph('%s', style = '%s')\n" % self.text_strip(text)
        return self.text_strip(text) + '\n'

    # provide by table plugin
    def table(self, text):
        return "t = document.add_table(rows=0, cols=2)\n" + text + "\n" # TODO: Generalize

    def table_head(self, text):
        return ""

    def table_body(self, text):
        return text + "\n"

    def table_row(self, text):
        return "col = 0\nr = t.add_row()\n" + text + "\n"

    def table_cell(self, text, align=None, is_head=False):
        italic = "italic = True" in text
        bold = "bold = True" in text
        return "c = r.cells[col]\np = c.paragraphs[0]\nrun = p.add_run(\"" \
            + self.text_strip(text) + "\")\nrun.bold = " + str(bold) + "\nrun.italic = " \
            + str(italic) + "\ncol+=1\n"
        # return "c = r.cells[col]\nc.italic = " + str(italic) + "\nc.bold = " + str(bold) + \
        #     "\nc.text = '" + self.text_strip(text) + "'\ncol+=1\n"


    # Finalize rendered content (define output)
    def finalize(self, data):
        return ''.join(list(data))


def style_setup():
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


parser = argparse.ArgumentParser(description='Generate Docx reports using a Docx reference template and Markdown files')
parser.add_argument('output', default=None, help='Output file')
parser.add_argument('--files', default="*.md", help='Regex for Markdown files')
args = parser.parse_args()

document = Document()

style_setup()

T = []

for part in sorted(glob.glob(args.files)):
    with open(part, 'r', encoding="utf-8") as f:
        T.append(f.read())

print(mistune.Markdown(renderer=DocxRenderer(), plugins=[mistune.plugins.plugin_table])('\n'.join(T)))
exec(mistune.Markdown(renderer=DocxRenderer(), plugins=[mistune.plugins.plugin_table])('\n'.join(T)))

document.save(os.path.abspath(args.output))