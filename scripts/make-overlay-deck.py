"""A deck for testing transparent-background export.

Slide 1: dark master background, white title, black text box, a 50%
         see-through red box, a soft shadow, a rounded lower-third bar.
Slide 2: a full-slide picture at the back with text over it.
Slide 3: its own gradient background and a semi-transparent circle.
Slide 4: hidden, a green star.
The master carries a "LOGO" box (a background graphic) top right.
"""
import copy
import sys

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Pt

out = sys.argv[1]
pic_path = sys.argv[2]

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

prs = Presentation()
prs.slide_width = Emu(12192000)
prs.slide_height = Emu(6858000)
blank = prs.slide_layouts[6]
title_only = prs.slide_layouts[5]

# Master background: dark navy.
mbg = prs.slide_master.background.fill
mbg.solid()
mbg.fore_color.rgb = RGBColor(0x14, 0x22, 0x44)


def alpha(shape, pct):
    """Make a shape's solid fill see-through (pct 0..100 opaque)."""
    srgb = shape._element.spPr.find(f'{{{A}}}solidFill/{{{A}}}srgbClr')
    a = etree.SubElement(srgb, f'{{{A}}}alpha')
    a.set('val', str(int(pct * 1000)))


def shadow(shape):
    spPr = shape._element.spPr
    eff = etree.SubElement(spPr, f'{{{A}}}effectLst')
    sh = etree.SubElement(eff, f'{{{A}}}outerShdw', blurRad='254000', dist='76200', dir='2700000', algn='tl', rotWithShape='0')
    clr = etree.SubElement(sh, f'{{{A}}}srgbClr', val='000000')
    etree.SubElement(clr, f'{{{A}}}alpha', val='60000')


def no_line(shape):
    shape.line.fill.background()


# Slide 1
s1 = prs.slides.add_slide(title_only)
s1.shapes.title.text = 'WHITE TITLE'
for p in s1.shapes.title.text_frame.paragraphs:
    for r in p.runs:
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        r.font.bold = True
        r.font.size = Pt(54)

panel = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(700000), Emu(2000000), Emu(4200000), Emu(2400000))
panel.fill.solid()
panel.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
no_line(panel)
shadow(panel)
tb = s1.shapes.add_textbox(Emu(900000), Emu(2300000), Emu(3800000), Emu(800000))
tb.text_frame.text = 'black text on white'
tb.text_frame.paragraphs[0].runs[0].font.size = Pt(28)
tb.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0, 0, 0)

red = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(5600000), Emu(1900000), Emu(3000000), Emu(2200000))
red.fill.solid()
red.fill.fore_color.rgb = RGBColor(0xE0, 0x20, 0x30)
alpha(red, 50)
no_line(red)

bar = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Emu(600000), Emu(5300000), Emu(10900000), Emu(900000))
bar.fill.solid()
bar.fill.fore_color.rgb = RGBColor(0xF0, 0xA0, 0x20)
no_line(bar)
bar.text_frame.text = 'Lower third — Name Surname, Job Title'
bar.text_frame.paragraphs[0].runs[0].font.size = Pt(28)
bar.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0x10, 0x10, 0x10)

# Master "logo": make it on slide 1, then move it to the master's shape tree.
logo = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(10400000), Emu(250000), Emu(1500000), Emu(600000))
logo.fill.solid()
logo.fill.fore_color.rgb = RGBColor(0x30, 0xC0, 0xF0)
no_line(logo)
logo.text_frame.text = 'LOGO'
el = logo._element
el.getparent().remove(el)
prs.slide_master.shapes._spTree.append(el)

# Slide 2: picture at the back filling the slide.
w, h = 1920, 1080
img = Image.new('RGB', (w, h))
px = img.load()
for y in range(h):
    for x in range(w):
        px[x, y] = (int(255 * x / w), int(120 + 100 * y / h), int(255 - 255 * x / w))
img.save(pic_path)
s2 = prs.slides.add_slide(blank)
s2.shapes.add_picture(pic_path, 0, 0, prs.slide_width, prs.slide_height)
t2 = s2.shapes.add_textbox(Emu(1000000), Emu(2800000), Emu(10000000), Emu(1200000))
t2.text_frame.text = 'Text over a picture background'
t2.text_frame.paragraphs[0].runs[0].font.size = Pt(48)
t2.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

# Slide 3: own gradient background, see-through circle.
s3 = prs.slides.add_slide(blank)
g = s3.background.fill
g.gradient()
g.gradient_stops[0].color.rgb = RGBColor(0x80, 0x00, 0x80)
g.gradient_stops[1].color.rgb = RGBColor(0xFF, 0xC0, 0x00)
c = s3.shapes.add_shape(MSO_SHAPE.OVAL, Emu(4000000), Emu(1400000), Emu(4000000), Emu(4000000))
c.fill.solid()
c.fill.fore_color.rgb = RGBColor(0x00, 0x60, 0xFF)
alpha(c, 40)
no_line(c)

# Slide 4: hidden.
s4 = prs.slides.add_slide(blank)
st = s4.shapes.add_shape(MSO_SHAPE.STAR_5_POINT, Emu(4500000), Emu(1800000), Emu(3000000), Emu(3000000))
st.fill.solid()
st.fill.fore_color.rgb = RGBColor(0x20, 0xC0, 0x40)
s4._element.set('show', '0')

prs.save(out)
print('wrote', out)
