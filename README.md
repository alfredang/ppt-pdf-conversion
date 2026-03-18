# PPT / PDF Conversion

A web app that translates Chinese presentations and documents to English.

## Features

- **PDF to PPTX** -- Extract text and images from a Chinese PDF, translate to English, and generate a PowerPoint file
- **PPTX to PPTX** -- Translate Chinese text in an existing PowerPoint to English while preserving all formatting, layouts, images, and animations
- **Real-time progress** -- SSE-based progress updates during conversion
- **Drag & drop** -- Simple upload interface with drag-and-drop support

## Tech Stack

- **Backend**: FastAPI + Uvicorn
- **PDF Extraction**: PyMuPDF (fitz)
- **PPTX Generation / Editing**: python-pptx
- **Translation**: Google Translate via deep-translator
- **Frontend**: Vanilla HTML/CSS/JS

## Getting Started

```bash
# Install dependencies
pip install -r requirements.txt

# Run the app
python3 run.py
```

Open http://localhost:8000 in your browser.

## How It Works

### PDF to PPTX
1. Extract text spans and images from each PDF page using PyMuPDF
2. Group text spans into lines by y-coordinate proximity
3. Translate Chinese text to English via Google Translate
4. Generate a new PPTX with positioned text boxes and images

### PPTX to PPTX
1. Open the original PPTX with python-pptx
2. Walk all shapes (text boxes, tables, groups, notes)
3. Translate only runs containing Chinese characters
4. Save -- all formatting, styling, and layout stays intact
