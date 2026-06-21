# PyPDF Studio

A professional desktop **PDF application** built with **PySide6 (Qt6)** and a full
Python PDF stack (PyMuPDF, pypdf, Pillow, ReportLab, pytesseract). Inspired by
Adobe Acrobat / Foxit / Nitro, with a sleek dark/light theme — all in a single,
clean, well-commented `main.py`.

![Theme](https://img.shields.io/badge/theme-dark%20%7C%20light-3b82f6) ![Python](https://img.shields.io/badge/python-3.10%2B-blue) ![Qt](https://img.shields.io/badge/UI-PySide6-41cd52)

---

## Features

- **Viewer** — smooth lazy-rendering pages, single / continuous / two-page modes,
  zoom (Ctrl+wheel), fit width/page, rotate, jump-to-page, fullscreen, and
  reading modes (Normal / Sepia / Eye-Comfort / Night) with auto-scroll.
- **Documents** — multiple tabs, drag & drop, recent files, favorites, properties.
- **Page management** — insert, delete, duplicate, reverse, extract, **split**,
  **merge**, drag-to-reorder thumbnails, rotate.
- **Editing & annotation** — add text, insert images, highlight / underline /
  strikeout, sticky notes.
- **Tools** — text/image **watermark**, **header & footer**, **compression**
  (4 levels with size estimate), **OCR** (language select + export), **security**
  (AES-256 encrypt, remove, disable print/copy/edit), and **batch** operations.
- **Converters** — PDF → Word / Excel / PowerPoint / Text / HTML / Image, and
  Images / Text → PDF (run on background threads with progress).
- **Infrastructure** — SQLite settings/recent/history, theming via Qt
  stylesheets, autosave recovery, file logging, and friendly error dialogs.

## Requirements

- Python **3.10+** (developed/tested on 3.14)
- See [`requirements.txt`](requirements.txt)

```bash
pip install -r requirements.txt
```

> **OCR** also needs the [Tesseract](https://github.com/UB-Mannheim/tesseract/wiki)
> engine installed on your system (the `pytesseract` package is only a wrapper).

## Run

```bash
python main.py
```

On Windows you can also double-click **`run.bat`**, which pins a known-good
interpreter.

## Build a portable app (Windows)

Produce a self-contained app folder that needs no Python install:

```bat
build_portable.bat
```

This builds a slimmed **onedir** bundle (~245 MB) at `dist/PyPDFStudio/`.
Run `dist/PyPDFStudio/PyPDFStudio.exe` — it starts instantly. The script
bundles everything (so Qt plugins are found) and then prunes the large Qt
modules the app never uses (WebEngine, Quick/QML, 3D, Charts, Multimedia, …).

The app stores its settings and logs in a `PyPDFStudio-Data` folder **next to
the executable**, so the whole `PyPDFStudio` folder is portable (copy it to a
USB stick and run anywhere) and leaves no traces on the host.

## Project layout

| File | Purpose |
|------|---------|
| `main.py` | The entire application (UI, logic, helpers) in one file |
| `run.bat` | Convenience launcher (Windows) |
| `build_portable.bat` | PyInstaller build script for the portable exe |
| `requirements.txt` | Python dependencies |

## License

MIT — see [`LICENSE`](LICENSE).
