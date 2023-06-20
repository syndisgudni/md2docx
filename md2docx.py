#! /usr/bin/python

# MD > HTML
import mistletoe
# HTML parser
from bs4 import BeautifulSoup as bs
from bs4.element import Tag, NavigableString
from prism import highlight
# DOCX
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
# Utils
import argparse
import glob
import os
from PIL import Image, ImageOps
from colour import Color

DOC_TEXT_WIDTH = "450pt"
COLOR_SYN_BLUE = "#057d9f"

def font_size(sz):
    return Pt(float(sz.split('pt')[0]))

def font_color(c):
    color = Color(c.strip())
    rgb = color.rgb
    return RGBColor(int(rgb[0]*255), int(rgb[1]*255), int(rgb[2]*255))

def par_align(j):
    match j:
        case 'justify':
            return WD_ALIGN_PARAGRAPH.JUSTIFY
        case 'center':
            return WD_ALIGN_PARAGRAPH.CENTER
        case 'left':
            return WD_ALIGN_PARAGRAPH.LEFT
        case 'right':
            return WD_ALIGN_PARAGRAPH.RIGHT

# TODO: Replace this hacky mess with another, hackier mess that does not require saving a bordered copy to disk
def set_image_border(input_image):
    img = Image.open(input_image)
    bimg = ImageOps.expand(img, border=1)
    bimg.save('bdr-' + input_image)

def set_paragraph_border(par, **kwargs):
    """
    Set paragraph`s border
    Usage:

    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        start={"sz": 24, "val": "dashed", "shadow": "true"},
        end={"sz": 12, "val": "dashed"},
    )
    """
    p = par._p
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    pPr.insert_element_before(pBdr,
        'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
        'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
        'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
        'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
        'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
        'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr',
        'w:pPrChange'
    )
    
    # list over all available tags
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            if 'color' in edge_data:
                color = edge_data['color'][1:]
                edge_data['color'] = color
            if 'sz' in edge_data:
                sz = int(edge_data['sz'][:-2])
                edge_data['sz'] = str(sz*2)

            tag = 'w:{}'.format(edge)

            # check for tag existnace, if none found, then create one
            element = pBdr.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                pBdr.append(element)
            # looks like order of attributes is important
            for key in ["color", "space", "sz", "val", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

def set_cell_border(c, **kwargs):
    """
    Set cell`s border
    Usage:

    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        start={"sz": 24, "val": "dashed", "shadow": "true"},
        end={"sz": 12, "val": "dashed"},
    )
    """
    tc = c._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existance, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags
    for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            if 'color' in edge_data:
                color = edge_data['color'][1:]
                edge_data['color'] = color
            if 'sz' in edge_data:
                sz = int(edge_data['sz'][:-2])
                edge_data['sz'] = str(sz*2)

            tag = 'w:{}'.format(edge)

            # check for tag existnace, if none found, then create one
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            # looks like order of attributes is important
            for key in ["color", "space", "sz", "val", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

def set_cell_bg_color(c, color):
    color = color[1:]
    tblCell = c._tc
    tblCellProperties = tblCell.get_or_add_tcPr()
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    tblCellProperties.append(shading)

def dict_to_style(dict):
    return ';'.join(['%s:%s' % (key, value) for (key, value) in dict.items()])

def style_to_dict(style):
    d = {}
    for s in style.split(';'):
        if s =='': continue
        k,v = s.split(':')
        d[k] = v
    return d

def get_style(tag):
    style = {}
    tags = []
    # Build tag stack
    while tag.parent:
        tags.append(tag)
        tag = tag.parent
    # Traverse tag stack and build style object
    while len(tags) > 0:
        tag = tags.pop()
        if not tag.name is None and 'style' in tag.attrs:
            style.update(style_to_dict(tag['style']))
    return style

def table_dimensions(table):
    rows = len(table.find_all('tr'))
    cols = len(table.tr.select('th,td'))
    return rows, cols

def apply_html_style(soup):
    # Global
    base_style = {
        'font-family': 'Calibri',
        'font-size': '11pt'
    }

    # Text
    for t in soup.select('p, th, td, li'):
        style = base_style.copy()
        style['text-align'] = 'justify'
        t['style'] = dict_to_style(style)

    # Inline Code
    for t in soup.find_all('code'):
        style = base_style.copy()
        style['font-family'] = 'Roboto Mono'
        style['font-size'] = '9pt'
        t['style'] = dict_to_style(style)

    # Code blocks
    for t in soup.select('pre > code'):
        style = base_style.copy()
        style['font-family'] = 'Roboto Mono'
        style['font-size'] = '9pt'
        style['text-align'] = 'left'
        style['border'] = '1px solid #000000'
        t['style'] = dict_to_style(style)

        # Apply Prism syntax highlighting
        lang = ''
        if 'class' in t.attrs:
            lang = t['class'][0].removeprefix('language-')
        code = t.string
        t.clear()
        new_soup = bs(highlight(code, lang), 'html.parser')
        t.append(new_soup)

    # Image
    for i in soup.find_all('img'):
        style = base_style.copy()
        style['width'] = DOC_TEXT_WIDTH
        i['style'] = dict_to_style(style)

    # Captions
    for p in soup.select('p:has(img), pre + p, table + p'):
        style = base_style.copy()
        style['font-size'] = '10pt'
        style['text-align'] = 'center'
        p['style'] = dict_to_style(style)
        for t in p.find_all('code'):
            style = base_style.copy()
            if 'style' in t.attrs:
                style = style_to_dict(t['style'])
            style['font-size'] = '8pt'
            t['style'] = dict_to_style(style)


    # Headings
    heading_style = base_style.copy()
    heading_style['color'] = '#666666'
    heading_style['text-align'] = 'left'

    for t in soup.find_all('h1'):
        style = heading_style.copy()
        style['font-size'] = '20pt'
        style['font-weight'] = 'bold'
        t['style'] = dict_to_style(style)

    for t in soup.find_all('h2'):
        style = heading_style.copy()
        style['font-size'] = '18pt'
        style['font-weight'] = 'bold'
        t['style'] = dict_to_style(style)
        
    for t in soup.find_all('h3'):
        style = heading_style.copy()
        style['font-size'] = '14pt'
        style['font-weight'] = 'bold'
        style['color'] = '#434343'
        t['style'] = dict_to_style(style)

    for t in soup.find_all('h4'):
        style = heading_style.copy()
        style['font-size'] = '12pt'
        t['style'] = dict_to_style(style)

    # Tables
    table_style = base_style.copy()
    table_style['border-collapse'] = 'collapse'

    ## Remove empty table heads
    for t in soup.find_all('table'):
        header_empty = True
        for th in t.thead.tr:
            if not th.text.rstrip() == '':
                header_empty = False
                break
        if header_empty:
            t.thead.decompose()
    ## Remove those weird alignment attributes on every table cell
    for t in soup.select('tr > *'):
        if 'align' in t.attrs: del t['align']
    ## All tables
    for t in soup.find_all('table'):
        t['style'] = dict_to_style(table_style)
    ## Summary tables
    for t in soup.select("h2 + table"):
        ### Top row
        for c in t.tr.find_all('td'):
            if 'style' in c.attrs: style = style_to_dict(c['style'])
            else: style = base_style.copy()
            style['border-top'] = '6px solid %s' % COLOR_SYN_BLUE
            c['style'] = dict_to_style(style)
        ### First column
        for r in t.tbody.find_all('tr'):
            c = r.td
            if 'style' in c.attrs: style = style_to_dict(c['style'])
            else: style = base_style.copy()
            style['text-align'] = 'right'
            style['color'] = '#ffffff'
            style['background-color'] = COLOR_SYN_BLUE
            style['width'] = '82.512pt'
            c['style'] = dict_to_style(style)
        ### Second column
        for r in t.tbody.find_all('tr'):
            c = r.select('td:nth-child(2)')[0]
            if 'style' in c.attrs: style = style_to_dict(c['style'])
            else: style = base_style.copy()
            style['width'] = '367.488pt'
            c['style'] = dict_to_style(style)
    ## Other tables
    for t in soup.select("*:not(h2) + table"):
        ### Top row
        for c in t.thead.find_all('th'):
            if 'style' in c.attrs: style = style_to_dict(c['style'])
            else: style = base_style.copy()
            style['color'] = '#ffffff'
            style['background-color'] = COLOR_SYN_BLUE
            c['style'] = dict_to_style(style)
        ### Every other row
        for r in t.tbody.select('tr:nth-child(even)'):
            for c in r.find_all('td'):
                if 'style' in c.attrs: style = style_to_dict(c['style'])
                else: style = base_style.copy()
                style['background-color'] = '#efefef'
                c['style'] = dict_to_style(style)


class HtmlToDocx:
    def __init__(self, soup: bs):
        self.soup = soup
        self.doc = self.setup_document()
        self.stack = []
    
    def setup_document(self):
        doc = Document()

        # Page setup
        section = doc.sections[-1]
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.59)
        section.bottom_margin = Inches(0.59)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

        p_format = doc.styles['Normal'].paragraph_format
        p_format.space_before = Pt(1)

        return doc

    def apply_inline_style(self, r, style):
        if 'font-family' in style:
            r.font.name = style['font-family']
        if 'font-size' in style:
            r.font.size = font_size(style['font-size'])
        if 'font-weight' in style:
            r.bold = (style['font-weight'] == 'bold')
        if 'color' in style:
            r.font.color.rgb = font_color(style['color'])
        scope = [t.name for t in self.stack]
        r.bold = r.bold or 'strong' in scope
        r.italic = r.italic or 'em' in scope

    def apply_block_style(self, p, style):
        if 'text-align' in style:
            p.alignment = par_align(style['text-align'])
        if 'border' in style:
            # p.paragraph_format.line_spacing = Pt(32)
            # p.paragraph_format.left_indent = Pt(32)
            sz,_,color = style['border'].split(' ')
            set_paragraph_border(p,
                top={'sz': sz, 'val': 'single', 'color': color, 'space': '4'},
                bottom={'sz': sz, 'val': 'single', 'color': color, 'space': '4'},
                left={'sz': sz, 'val': 'single', 'color': color, 'space': '4'},
                right={'sz': sz, 'val': 'single', 'color': color, 'space': '4'}
            )

    def apply_cell_style(self, c, style):
        if 'background-color' in style:
            set_cell_bg_color(c, style['background-color'])
        if 'border-top' in style:
            sz,_,color = style['border-top'].split(' ')
            set_cell_border(c, top={'sz': sz, 'val': 'single', 'color': color, 'space': '0'})

    def apply_column_style(self, c, style):
        if 'width' in style:
            c.width = font_size(style['width'])
    
    def apply_table_style(self, t, style):
        t.autofit = True

    def apply_image_style(self, i, style):
        if 'width' in style:
            i.width = font_size(style['width'])

    def render_text(self, tag, p):
        if not isinstance(tag, Tag):
            r = p.add_run(tag.text)
            self.apply_inline_style(r, get_style(tag))
            return
        self.stack.append(tag)
        for t in tag:
            self.render_text(t, p)
        self.stack.pop()
        return tag.text

    def render_paragraph(self, tag, p=None):
        if p == None:
            if 'img' in [t.name for t in tag.children]:
                self.render_image(tag.img)
                p = self.doc.paragraphs[-1]
            else:
                p = self.doc.add_paragraph()
        self.apply_block_style(p, get_style(tag))
        for t in tag:
            if t.name == 'img': continue
            self.render_text(t, p)
        return p

    def render_table(self, tag):
        rows, cols = table_dimensions(tag)
        t = self.doc.add_table(rows=rows, cols=cols)
        self.apply_table_style(t, get_style(tag))

        soup_rows = tag.find_all('tr')
        for r, row in enumerate(t.rows):
            soup_row = soup_rows[r]
            soup_cells = soup_row.select('th,td')
            for c, cell in enumerate(row.cells):
                soup_cell = soup_cells[c]
                self.render_paragraph(soup_cell, cell.paragraphs[0])
                self.apply_cell_style(cell, get_style(soup_cell))
        
        for c, col in enumerate(t.columns):
            soup_cell = tag.select('td:nth-child(%d)' % (c+1))[0]
            self.apply_column_style(col, get_style(soup_cell))
        return t

    def render_heading(self, tag):
        level = int(tag.name[1])
        h = self.doc.add_heading('', level)
        self.apply_block_style(h, get_style(tag))
        self.render_text(tag, h)
        return h

    def render_image(self, tag):
        filename = 'bdr-' + tag['src']
        if not os.path.exists(filename):
            set_image_border(tag['src'])
        i = self.doc.add_picture(filename, width=font_size(DOC_TEXT_WIDTH))
        self.apply_image_style(i, get_style(tag))
        return i

    def render_code(self, tag):
        # Remove trailing newline
        if len(tag.code.contents) > 1:
            tag.code.contents[-1].replace_with('')
        else:
            tag.string.replace_with(tag.string.rstrip())
        self.render_paragraph(tag.code)

    def render_tag(self, tag, level=0):
        self.stack.append(tag)

        rendered = True

        # Parse block-level tags
        match tag.name:
            case 'h1' | 'h2' | 'h3' | 'h4':
                self.render_heading(tag)
            case 'p':
                self.render_paragraph(tag)
            case 'table':
                self.render_table(tag)
            case 'pre':
                self.render_code(tag)
            case _:
                rendered = False
        
        # Iterate
        if not rendered:
            for t in tag.contents:
                if t.name is None: continue
                self.render_tag(t, level+1)

        self.stack.pop()

    def render(self):
        self.render_tag(self.soup)
        return self.doc


arg_parser = argparse.ArgumentParser(description='Generate Docx reports using a Docx reference template and Markdown files')
arg_parser.add_argument('output', default=None, help='Output file')
arg_parser.add_argument('--files', default="*.md", help='Regex for Markdown files')
args = arg_parser.parse_args()

file_data = []
for part in sorted(glob.glob(args.files)):
    with open(part, 'r', encoding="utf-8") as f:
        file_data.append(f.read())

html = ''.join([mistletoe.markdown(data) for data in file_data])

soup = bs(html, 'html.parser')
html = soup.decode()

apply_html_style(soup)

# TODO: For testing
# print("->BEGIN HTML INPUT\n%s<-END HTML INPUT" % soup)
open('test.html','w').write(str(soup))

document = HtmlToDocx(soup).render()

document.save(args.output)
