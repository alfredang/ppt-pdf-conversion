"""PPTX generation module using python-pptx."""

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor


# Conversion: PDF points to EMU (1 point = 12700 EMU)
PT_TO_EMU = 12700


def _color_from_int(color_int):
    """Convert a PDF integer color to pptx RGBColor."""
    if color_int is None or color_int == 0:
        return RGBColor(0, 0, 0)  # Default black
    r = (color_int >> 16) & 0xFF
    g = (color_int >> 8) & 0xFF
    b = color_int & 0xFF
    return RGBColor(r, g, b)


def create_presentation(pages_data, progress_callback=None):
    """Create a PPTX presentation from extracted and translated page data."""
    prs = Presentation()

    # Set slide size to match PDF (960x540 points = 13.333" x 7.5")
    prs.slide_width = Emu(int(960 * PT_TO_EMU))
    prs.slide_height = Emu(int(540 * PT_TO_EMU))

    for i, page_data in enumerate(pages_data):
        if progress_callback:
            progress_callback(f"Generating slide {i + 1}/{len(pages_data)}")

        slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
        _add_images(slide, page_data)
        _add_text_groups(slide, page_data)

    return prs


def _add_images(slide, page_data):
    """Add images to a slide."""
    for img in page_data.get('images', []):
        bbox = img['bbox']
        left = Emu(int(bbox[0] * PT_TO_EMU))
        top = Emu(int(bbox[1] * PT_TO_EMU))
        width = Emu(int((bbox[2] - bbox[0]) * PT_TO_EMU))
        height = Emu(int((bbox[3] - bbox[1]) * PT_TO_EMU))

        stream = img['stream']
        stream.seek(0)

        try:
            slide.shapes.add_picture(stream, left, top, width, height)
        except Exception as e:
            print(f"Error adding image: {e}")


def _add_text_groups(slide, page_data):
    """Add text groups (lines) as text boxes on the slide."""
    for group in page_data.get('text_groups', []):
        text = group.get('translated_text', group['text'])
        bbox = group['bbox']

        left = Emu(int(bbox[0] * PT_TO_EMU))
        top = Emu(int(bbox[1] * PT_TO_EMU))
        width = Emu(int((bbox[2] - bbox[0]) * PT_TO_EMU))
        height = Emu(int((bbox[3] - bbox[1]) * PT_TO_EMU))

        # Ensure minimum dimensions
        if width < Emu(50000):
            width = Emu(500000)
        if height < Emu(50000):
            height = Emu(300000)

        try:
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            tf.word_wrap = False

            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = text

            # Apply formatting
            font = run.font
            font.size = Pt(group['font_size'])
            font.color.rgb = _color_from_int(group['color'])
            font.bold = group.get('is_bold', False)
            font.italic = group.get('is_italic', False)
            font.name = 'Arial'

            # Remove default margins for tighter positioning
            tf.margin_left = Emu(0)
            tf.margin_right = Emu(0)
            tf.margin_top = Emu(0)
            tf.margin_bottom = Emu(0)

        except Exception as e:
            print(f"Error adding text '{text[:30]}...': {e}")
