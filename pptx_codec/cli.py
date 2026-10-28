"""pptx-codec generate < slides.json > deck.pptx | pptx-codec parse < deck.pptx > slides.json"""
import io
import json
import sys

from .generator import build
from .parser import parse


def main():
    command = sys.argv[1:]
    data = sys.stdin.buffer.read()

    if command == ["generate"]:
        output = io.BytesIO()
        build(json.loads(data), output)
        sys.stdout.buffer.write(output.getvalue())
    elif command == ["parse"]:
        json.dump(parse(io.BytesIO(data)), sys.stdout)
    else:
        sys.exit(__doc__)
