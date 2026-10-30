""".pptx in, slides JSON out. Inverse of generator.py: parse(generate(deck)) == deck.

    pptx-codec parse < deck.pptx > slides.json

Only keys generator.py would have been given are emitted; defaults are left out.
Shapes that are not text boxes, pictures, charts, tables or groups are skipped. Fractions are rounded to 6 places.
"""
import base64
import json
import sys

from lxml import etree
from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE
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
    if read_runs(frame, data):
        put(data, "valign", VALIGN_NAMES.get(frame.vertical_anchor), "t")
        put(data, "inset", frame.margin_left.pt, 2)


def read_runs(frame, data):
    runs = [run for paragraph in frame.paragraphs for run in paragraph.runs]
    if not runs:
        return False

    first = frame.paragraphs[0]
    data["text"] = "\n".join(paragraph.text for paragraph in frame.paragraphs)
    data["size"] = runs[0].font.size.pt
    if runs[0].font.color.type == MSO_COLOR_TYPE.RGB:
        data["color"] = str(runs[0].font.color.rgb)
    put(data, "bold", runs[0].font.bold, False)
    put(data, "caps", runs[0].font._rPr.get("cap") == "all", False)
    put(data, "align", ALIGN_NAMES.get(first.alignment), "l")
    put(data, "line_spacing", first.line_spacing, 1.0)
    return True


def read_text_box(shape, presentation):
    data = position(shape, presentation)

    geometry = shape._element.spPr.prstGeom
    if geometry is not None and geometry.get("prst") in GEOMETRY_NAMES:
        data["geometry"] = GEOMETRY_NAMES[geometry.get("prst")]
        guides = [int(guide.get("fmla").split()[1]) for guide in geometry.iterfind(f"{DRAWING}avLst/{DRAWING}gd")]
        if guides and data["geometry"] == "block_arc":
            data["adjustment"] = [guides[0] / 60000, guides[1] / 60000, guides[2] / 100000]
        elif guides:
            data["adjustment"] = guides[0] / 100000

    if shape.fill.type == MSO_FILL.SOLID:
        data["fill"] = color_name(shape.fill.fore_color)

    if shape.line.fill.type == MSO_FILL.SOLID:
        data["line"] = color_name(shape.line.color)
        put(data, "line_width", shape.line.width.pt, 0.75)

    read_text(shape.text_frame, data)
    return data


def read_picture(shape, presentation):
    return {"kind": "picture", **position(shape, presentation), "data": base64.b64encode(shape.image.blob).decode()}


def read_table(shape, presentation):
    table = shape.table
    first = table.cell(0, 0)
    data = {"kind": "table", **position(shape, presentation)}
    data["columns"] = [fraction(column.width, shape.width) for column in table.columns]
    data["rows"] = [[read_cell(cell) for cell in row.cells] for row in table.rows]
    put(data, "valign", VALIGN_NAMES.get(first.vertical_anchor), "t")
    rule = first._tc.tcPr.find(f"{DRAWING}lnB/{DRAWING}solidFill/{DRAWING}srgbClr") if first._tc.tcPr is not None else None
    if rule is not None:
        data["rule"] = rule.get("val")
    return data


def read_cell(cell):
    data = {}
    read_runs(cell.text_frame, data)
    return data


def read_chart(shape, presentation):
    plot = shape.chart.plots[0]
    if shape.chart.chart_type == XL_CHART_TYPE.DOUGHNUT:
        kind = "doughnut"
        series = plot.series[0]
        data = [{"label": label, "value": value, "color": str(series.points[index].format.fill.fore_color.rgb)}
                for index, (label, value) in enumerate(zip(plot.categories, series.values))]
    else:
        kind = "bar"
        data = [{"label": series.name, "value": series.values[0], "color": str(series.format.fill.fore_color.rgb)}
                for series in plot.series]
    return {"kind": "chart", "type": kind, **position(shape, presentation), "data": data}


def read(shape, presentation):
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return read_picture(shape, presentation)
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        children = [read(child, presentation) for child in shape.shapes]
        return {"kind": "group", "shapes": [child for child in children if child is not None]}
    if shape.has_chart:
        return read_chart(shape, presentation)
    if shape.has_table:
        return read_table(shape, presentation)
    if shape.has_text_frame:
        return read_text_box(shape, presentation)
    return None


def parse(source):
    presentation = Presentation(source)
    slides = []
    for slide in presentation.slides:
        shapes = [read(shape, presentation) for shape in slide.shapes]
        slides.append({"shapes": [shape for shape in shapes if shape is not None]})

    return {"font": read_font(presentation), "slides": slides}
