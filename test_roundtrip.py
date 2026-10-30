"""python -m unittest"""
import io
import unittest

from pptx_codec import build
from pptx_codec import parse

# 1x1 transparent PNG
PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="

# One text box with every option, one with none (defaults must not be emitted), one arc, one picture, two charts, one table, one group.
DECK = {
    "font": "Arial",
    "slides": [
        {
            "shapes": [
                {
                    "x": 0.05, "y": 0.1, "w": 0.9, "h": 0.2,
                    "text": "Title\nSecond line", "size": 40, "color": "404040", "bold": True, "caps": True,
                    "align": "ctr", "valign": "ctr", "line_spacing": 1.2, "inset": 6,
                    "fill": "1F6F8B", "line": "tx1", "line_width": 2,
                    "geometry": "rounded_rectangle", "adjustment": 0.1,
                },
                {"x": 0.1, "y": 0.4, "w": 0.3, "h": 0.1, "text": "plain", "size": 14},
                {"x": 0.5, "y": 0.4, "w": 0.2, "h": 0.3, "fill": "57A92B", "geometry": "block_arc", "adjustment": [180, 300, 0.15]},
            ]
        },
        {"shapes": [{"kind": "picture", "x": 0.1, "y": 0.3, "w": 0.5, "h": 0.5, "data": PNG}]},
        {
            "shapes": [
                {"kind": "chart", "type": "doughnut", "x": 0.1, "y": 0.1, "w": 0.3, "h": 0.5,
                 "data": [{"label": "On Track", "value": 3, "color": "57A92B"}, {"label": "Completed", "value": 1, "color": "1565C0"}]},
                {"kind": "chart", "type": "bar", "x": 0.5, "y": 0.1, "w": 0.4, "h": 0.05,
                 "data": [{"label": "On Track", "value": 3, "color": "57A92B"}, {"label": "Completed", "value": 1, "color": "1565C0"}]},
                {"kind": "table", "x": 0.5, "y": 0.3, "w": 0.4, "h": 0.2, "columns": [0.25, 0.75], "valign": "ctr", "rule": "D1D4E4",
                 "rows": [[{"text": "■", "size": 14, "color": "57A92B"}, {"text": "On Track", "size": 14, "bold": True, "align": "r"}],
                          [{}, {"text": "Completed", "size": 14}]]},
                {"kind": "group", "shapes": [
                    {"x": 0.1, "y": 0.7, "w": 0.3, "h": 0.05, "fill": "979DAA", "geometry": "rounded_rectangle", "adjustment": 0.1},
                    {"x": 0.1, "y": 0.7, "w": 0.3, "h": 0.05, "text": "37%", "size": 14, "color": "FFFFFF", "align": "r"},
                ]},
            ]
        },
    ],
}


class RoundTrip(unittest.TestCase):
    def test_parse_inverts_generate(self):
        output = io.BytesIO()
        build(DECK, output)
        self.assertEqual(parse(io.BytesIO(output.getvalue())), DECK)

    def test_missing_size_raises(self):
        with self.assertRaises(KeyError):
            build({"slides": [{"shapes": [{"x": 0, "y": 0, "w": 1, "h": 1, "text": "no size"}]}]}, io.BytesIO())
