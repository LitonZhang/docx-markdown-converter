# Document Converter (Windows Local)

A local-first converter that supports both `.doc/.docx -> .md` and `.md -> .docx` with a desktop UI shell.

- Desktop shell: Tauri + TypeScript
- Conversion engine: Python CLI
- Privacy: local processing only

## Features

- `doc/docx -> md` single-file conversion
- Preserves core structure: headings, paragraphs, lists, quotes, formulas (LaTeX), and tables
- Optional image export to `<input_stem>_images` beside the source file
- Optional chapter split after markdown conversion
- `md -> docx` single-file export with configurable style rules
- Style preset save/load (`.json`) for repeatable paper formatting
- Auto title mapping for `md -> docx`: first non-numbered and non-bold H1 is treated as paper title
- Auto semantic styling for abstract blocks and figure/table captions (`Paper*` styles)
- Optional body layout for `md -> docx`: single-column or double-column (title/abstract kept single-column when possible)

## Project Layout

- `src/`: desktop frontend UI (Vite + TS)
- `src-tauri/`: desktop Rust bridge that invokes Python CLI
- `converter/`: Python conversion engine
- `tests/`: Python unit tests
- `release-packages/`: prebuilt Windows installer packages

## Windows Installer Packages

Prebuilt installers are included in this repository under `release-packages/`:

- `release-packages/DOCX Markdown Converter_0.1.1_x64-setup.exe` (recommended for most users)
- `release-packages/DOCX Markdown Converter_0.1.1_x64_en-US.msi` (better for enterprise/IT deployment)

Direct links:

- `https://github.com/LitonZhang/docx-markdown-converter/tree/main/release-packages`

Install steps:

1. Download one installer from `release-packages`.
2. Run installer as normal user (or admin if your system policy requires it).
3. Launch app from Start Menu after installation.

Notes:

- If Windows SmartScreen appears, choose "More info" -> "Run anyway" after verifying source.
- Prefer `setup.exe` for personal installation; use `msi` when your IT tools require MSI.

## Run Converter CLI

`docx -> md`

```bash
python converter/convert_docx_to_md.py \
  --input ./input.docx \
  --output ./output.md \
  --math latex
```

With image export:

```bash
python converter/convert_docx_to_md.py \
  --input ./input.docx \
  --output ./output.md \
  --math latex \
  --extract-images \
  --image-dir ./assets
```

`md -> docx`

```bash
python converter/convert_md_to_docx.py \
  --input ./paper.md \
  --output ./paper.docx \
  --style ./style_preset.json
```

## Run Frontend Only

```bash
npm install
npm run dev
```

## Run Desktop App (Tauri)

Prerequisites:

- Node.js + npm
- Python 3.10+
- Pandoc (required for `md -> docx`)
- Rust toolchain (`cargo`) for Tauri backend build

Commands:

```bash
npm install
npm run tauri:dev:win
```

If Python is not in PATH, set:

```bash
set PYTHON_EXECUTABLE=C:\Path\To\python.exe
```

Build installer packages:

```bash
npm run tauri:build:win
```

Artifacts:

- `src-tauri/target/release/bundle/msi/*.msi`
- `src-tauri/target/release/bundle/nsis/*-setup.exe`

## Tests

```bash
python -m unittest discover -s tests -v
```

## Notes

- Tauri build and bundling are validated locally.
- Current scope is single-file conversion in both directions. Batch conversion can be added later.
