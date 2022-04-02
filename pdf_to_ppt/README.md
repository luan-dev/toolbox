# PDF to PPT

Converts PDF files to PowerPoint presentations with smaller file sizes

## Getting Started

Create a virtual environment:

```
python -m venv .venv
activate .venv/bin/activate
pip install -r requirements.txt
```

Run script:

```
python convert.py {input_file}

optional arguments:
  -o, --output
    output file name (must include extension)
  -l, --legacy_res
    legacy resolution to support 4:3
```
