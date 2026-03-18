"""PDF extraction module using PyMuPDF (fitz)."""

import io
import fitz  # PyMuPDF


def extract_page(doc, page):
    """Extract text elements and images from a single PDF page.

    Returns a dict with page dimensions, text groups (lines), and images.
    """
    data = page.get_text('dict')
    width = page.rect.width
    height = page.rect.height

    spans = []
    for block in data.get('blocks', []):
        if block['type'] == 0:  # text block
            for line in block['lines']:
                for span in line['spans']:
                    text = span['text'].strip()
                    if not text:
                        continue
                    spans.append({
                        'text': span['text'],
                        'bbox': list(span['bbox']),
                        'font_size': span['size'],
                        'font_name': span['font'],
                        'color': span['color'],
                        'is_bold': 'Bold' in span['font'] or 'bold' in span['font'],
                        'is_italic': 'Italic' in span['font'] or 'italic' in span['font'],
                    })

    # Group spans into lines by y-coordinate proximity
    text_groups = _group_spans_into_lines(spans)

    # Extract images
    images = _extract_images(doc, page)

    # If page has no text and no images, render as full-page image
    if not text_groups and not images:
        images = [_render_page_as_image(page, width, height)]

    return {
        'width': width,
        'height': height,
        'text_groups': text_groups,
        'images': images,
    }


def _group_spans_into_lines(spans, y_tolerance=3):
    """Group spans that share similar y-coordinates into lines."""
    if not spans:
        return []

    # Sort by y then x
    sorted_spans = sorted(spans, key=lambda s: (s['bbox'][1], s['bbox'][0]))

    lines = []
    current_line = [sorted_spans[0]]
    current_y = sorted_spans[0]['bbox'][1]

    for span in sorted_spans[1:]:
        if abs(span['bbox'][1] - current_y) <= y_tolerance:
            current_line.append(span)
        else:
            lines.append(_merge_line(current_line))
            current_line = [span]
            current_y = span['bbox'][1]

    if current_line:
        lines.append(_merge_line(current_line))

    return lines


def _merge_line(spans):
    """Merge spans in a line into a single text group with combined bbox."""
    # Sort by x position
    spans = sorted(spans, key=lambda s: s['bbox'][0])

    x0 = min(s['bbox'][0] for s in spans)
    y0 = min(s['bbox'][1] for s in spans)
    x1 = max(s['bbox'][2] for s in spans)
    y1 = max(s['bbox'][3] for s in spans)

    # Use the dominant span's formatting (largest font or first span)
    dominant = max(spans, key=lambda s: s['font_size'])

    # Join text with spaces
    full_text = ' '.join(s['text'] for s in spans)

    return {
        'text': full_text,
        'bbox': [x0, y0, x1, y1],
        'font_size': dominant['font_size'],
        'font_name': dominant['font_name'],
        'color': dominant['color'],
        'is_bold': dominant['is_bold'],
        'is_italic': dominant['is_italic'],
        'spans': spans,  # Keep original spans for reference
    }


def _extract_images(doc, page):
    """Extract images from a page with their positions."""
    images = []
    image_list = page.get_images(full=True)

    for img_info in image_list:
        xref = img_info[0]
        try:
            base_image = doc.extract_image(xref)
            if not base_image:
                continue

            image_bytes = base_image['image']
            image_ext = base_image['ext']

            # Get image position on page
            rects = page.get_image_rects(xref)
            if not rects:
                continue

            for rect in rects:
                images.append({
                    'stream': io.BytesIO(image_bytes),
                    'ext': image_ext,
                    'bbox': [rect.x0, rect.y0, rect.x1, rect.y1],
                })
        except Exception:
            continue

    return images


def _render_page_as_image(page, width, height):
    """Render entire page as an image (for image-only pages)."""
    pix = page.get_pixmap(dpi=150)
    image_bytes = pix.tobytes('png')
    return {
        'stream': io.BytesIO(image_bytes),
        'ext': 'png',
        'bbox': [0, 0, width, height],
    }


def extract_pdf(pdf_path, progress_callback=None):
    """Extract all pages from a PDF file.

    Returns a list of page data dicts.
    """
    doc = fitz.open(pdf_path)
    pages = []

    for i, page in enumerate(doc):
        if progress_callback:
            progress_callback(f"Extracting page {i + 1}/{len(doc)}")
        pages.append(extract_page(doc, page))

    doc.close()
    return pages
