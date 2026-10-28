""".pptx in, slides JSON out. Inverse of generator.py: parse(generate(deck)) == deck.

    pptx-codec parse < deck.pptx > slides.json

Only keys generator.py would have been given are emitted; defaults are left out.
Shapes that are not text boxes or pictures are skipped. Fractions are rounded to 6 places.
"""
import base64
import json
import sys

from lxml import etree
from pptx import Presentation
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE

from .generator import ALIGN, DRAWING, GEOMETRY, THEME_COLORS, VALIGN

ALIGN_NAMES = {value: key for key, value in ALIGN.items()}
VALIGN_NAMES = {value: key for key, value in VALIGN.items()}
GEOMETRY_NAMES = {value: key for key, value in GEOMETRY.items()}
THEME_COLOR_NAMES = {value: key for key, value in THEME_COLORS.items()}


def fraction(length, total):
    return round(length / total, 6)


def put(data, key, value, default=None):
    """Store value unless it is the generator's default (left out of the JSON)."""
    if value not in (None, default):
        data[key] = value


def position(shape, presentation):
    width, height = presentation.slide_width, presentation.slide_height
    return {
        "x": fraction(shape.left, width),
        "y": fraction(shape.top, height),
        "w": fraction(shape.width, width),
        "h": fraction(shape.height, height),
    }


def color_name(color):
    if color.type == MSO_COLOR_TYPE.SCHEME:
        return THEME_COLOR_NAMES[color.theme_color]
    return str(color.rgb)


def read_font(presentation):
    theme = presentation.slide_master.part.part_related_by(RELATIONSHIP_TYPE.THEME)
    root = etree.fromstring(theme.blob)
    return root.find(f".//{DRAWING}minorFont/{DRAWING}latin").get("typeface")


def read_text(frame, data):
    runs = [run for paragraph in frame.paragraphs for run in paragraph.runs]
    if not runs:
        return

    first = frame.paragraphs[0]
    data["text"] = "\n".join(paragraph.text for paragraph in frame.paragraphs)
    data["size"] = runs[0].font.size.pt
    put(data, "bold", runs[0].font.bold, False)
    put(data, "caps", runs[0].font._rPr.get("cap") == "all", False)
    put(data, "align", ALIGN_NAMES.get(first.alignment), "l")
    put(data, "valign", VALIGN_NAMES.get(frame.vertical_anchor), "t")
    put(data, "line_spacing", first.line_spacing, 1.0)
    put(data, "inset", frame.margin_left.pt, 2)


def read_text_box(shape, presentation):
    data = position(shape, presentation)

    geometry = shape._element.spPr.prstGeom
    if geometry is not None and geometry.get("prst") in GEOMETRY_NAMES:
        data["geometry"] = GEOMETRY_NAMES[geometry.get("prst")]
        corner = geometry.find(f"{DRAWING}avLst/{DRAWING}gd[@name='adj']")
        if corner is not None:
            data["adjustment"] = int(corner.get("fmla").split()[1]) / 100000

    if shape.fill.type == MSO_FILL.SOLID:
        data["fill"] = color_name(shape.fill.fore_color)

    if shape.line.fill.type == MSO_FILL.SOLID:
        data["line"] = color_name(shape.line.color)
        put(data, "line_width", shape.line.width.pt, 0.75)

    read_text(shape.text_frame, data)
    return data


def read_picture(shape, presentation):
    return {"kind": "picture", **position(shape, presentation), "data": base64.b64encode(shape.image.blob).decode()}


def parse(source):
    presentation = Presentation(source)
    slides = []
    for slide in presentation.slides:
        shapes = []
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                shapes.append(read_picture(shape, presentation))
            elif shape.has_text_frame:
                shapes.append(read_text_box(shape, presentation))
        slides.append({"shapes": shapes})

    return {"font": read_font(presentation), "slides": slides}

