"""Slides JSON in, .pptx out.

    pptx-codec generate < slides.json > deck.pptx

Input:
    {"font": "Arial", "slides": [{"shapes": [shape, ...]}, ...]}

Every shape has x, y, w, h as fractions of a 20 x 11.25 in (16:9) slide.

    text_box (default)  text, size (pt), bold, caps, align (r | ctr), valign (ctr), line_spacing, inset (pt),
                        fill, line (RRGGBB | tx1 | bg1), line_width (pt),
                        geometry (rounded_rectangle | ellipse), adjustment (corner radius, 0..0.5)
    picture             kind: "picture", data (base64 PNG or JPEG), drawn into x, y, w, h as is
"""
import base64
import io
import json
import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE
from pptx.util import Emu, Pt

SLIDE_WIDTH = Emu(18288000)
SLIDE_HEIGHT = Emu(10287000)
DRAWING = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

ALIGN = {"l": PP_ALIGN.LEFT, "r": PP_ALIGN.RIGHT, "ctr": PP_ALIGN.CENTER}
VALIGN = {"t": MSO_ANCHOR.TOP, "ctr": MSO_ANCHOR.MIDDLE}
GEOMETRY = {"rounded_rectangle": "roundRect", "ellipse": "ellipse"}
THEME_COLORS = {"tx1": MSO_THEME_COLOR.TEXT_1, "bg1": MSO_THEME_COLOR.BACKGROUND_1}


def position(shape):
    return (
        Emu(round(shape["x"] * SLIDE_WIDTH)),
        Emu(round(shape["y"] * SLIDE_HEIGHT)),
        Emu(round(shape["w"] * SLIDE_WIDTH)),
        Emu(round(shape["h"] * SLIDE_HEIGHT)),
    )


def paint(color, value):
    if value in THEME_COLORS:
        color.theme_color = THEME_COLORS[value]
    else:
        color.rgb = RGBColor.from_string(value)


def use_font(presentation, font):
    theme = presentation.slide_master.part.part_related_by(RELATIONSHIP_TYPE.THEME)
    root = etree.fromstring(theme.blob)
    for latin in root.iter(f"{DRAWING}latin"):
        if latin.getparent().tag in (f"{DRAWING}majorFont", f"{DRAWING}minorFont"):
            latin.set("typeface", font)
    theme._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def write_text(frame, shape):
    inset = Pt(shape.get("inset", 2))
    frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = inset
    frame.word_wrap = True
    frame.auto_size = MSO_AUTO_SIZE.NONE
    frame.vertical_anchor = VALIGN[shape.get("valign", "t")]
    frame._bodyPr.set("spcFirstLastPara", "1")

    for number, line in enumerate(shape["text"].split("\n")):
        paragraph = frame.paragraphs[0] if number == 0 else frame.add_paragraph()
        paragraph.alignment = ALIGN[shape.get("align", "l")]
        paragraph.line_spacing = shape.get("line_spacing", 1.0)

        run = paragraph.add_run()
        run.text = line
        run.font.size = Pt(shape["size"])
        run.font.bold = shape.get("bold", False)
        run.font._rPr.set("kern", "0")
        if shape.get("caps"):
            run.font._rPr.set("cap", "all")


def draw_text_box(slide, shape):
    box = slide.shapes.add_textbox(*position(shape))

    if "geometry" in shape:
        geometry = box._element.spPr.prstGeom
        geometry.set("prst", GEOMETRY[shape["geometry"]])
        if "adjustment" in shape:
            corner = f"val {round(shape['adjustment'] * 100000)}"
            etree.SubElement(geometry.find(f"{DRAWING}avLst"), f"{DRAWING}gd", name="adj", fmla=corner)

    if "fill" in shape:
        box.fill.solid()
        paint(box.fill.fore_color, shape["fill"])

    if "line" in shape:
        paint(box.line.color, shape["line"])
        box.line.width = Pt(shape.get("line_width", 0.75))
        etree.SubElement(box.line._get_or_add_ln(), f"{DRAWING}round")

    if "text" in shape:
        write_text(box.text_frame, shape)


def draw_picture(slide, shape):
    image = io.BytesIO(base64.b64decode(shape["data"]))
    slide.shapes.add_picture(image, *position(shape))


def build(deck, output):
    presentation = Presentation()
    presentation.slide_width = SLIDE_WIDTH
    presentation.slide_height = SLIDE_HEIGHT
    if "font" in deck:
        use_font(presentation, deck["font"])

    blank = presentation.slide_layouts.get_by_name("Blank")
    for slide_data in deck["slides"]:
        slide = presentation.slides.add_slide(blank)
        for shape in slide_data["shapes"]:
            if shape.get("kind") == "picture":
                draw_picture(slide, shape)
            else:
                draw_text_box(slide, shape)

    presentation.save(output)

