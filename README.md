# pptx-codec

Slides JSON to `.pptx` and back. A python-pptx library plus a CLI.

```
python3 -m venv .venv && .venv/bin/pip install -e .   # python 3.13+
.venv/bin/python -m unittest -v                      # round trip: parse(generate(deck)) == deck
.venv/bin/pptx-codec generate < slides.json > deck.pptx
.venv/bin/pptx-codec parse < deck.pptx > slides.json
```

The slides contract is documented at the top of `pptx_codec/generator.py`. `parser.py` is its inverse and only emits keys the generator accepts.

Consumers install it from git:

```
pptx-codec @ git+https://github.com/Envisio/pptx.git@v0.1.0
```
