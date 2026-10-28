"""python -m unittest"""
import io
import unittest

from pptx_codec import build
from pptx_codec import parse

# 1x1 transparent PNG
PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="

# One text box with every option, one with none (defaults must not be emitted), one picture.
DECK = {
    "font": "Arial",
    "slides": [
        {
            "shapes": [
                {
                    "x": 0.05, "y": 0.1, "w": 0.9, "h": 0.2,
                    "text": "Title\nSecond line", "size": 40, "bold": True, "caps": True,
                    "align": "ctr", "valign": "ctr", "line_spacing": 1.2, "inset": 6,
                    "fill": "1F6F8B", "line": "tx1", "line_width": 2,
                    "geometry": "rounded_rectangle", "adjustment": 0.1,
                },
                {"x": 0.1, "y": 0.4, "w": 0.3, "h": 0.1, "text": "plain", "size": 14},
            ]
        },
        {"shapes": [{"kind": "picture", "x": 0.1, "y": 0.3, "w": 0.5, "h": 0.5, "data": PNG}]},
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
