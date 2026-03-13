import io
from pathlib import Path

import fitz
from PIL import Image, ImageDraw


OUTPUT_PATH = Path(__file__).with_name("native_pdf_demo.pdf")


def _build_demo_image():
    image = Image.new("RGB", (320, 140), color=(236, 248, 255))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((12, 12, 308, 128), radius=18, fill=(14, 116, 144))
    draw.text((28, 48), "Embedded image block", fill=(255, 255, 255))
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


def build_demo_pdf():
    doc = fitz.open()
    page = doc.new_page(width=720, height=540)

    page.insert_text((40, 45), "Native PDF Elements Demo", fontsize=24, fontname="helv")
    page.insert_text(
        (40, 85),
        "This PDF contains a real text layer plus an embedded image.",
        fontsize=16,
        fontname="helv",
    )
    page.insert_text(
        (40, 140),
        "Expected extractor behavior:",
        fontsize=18,
        fontname="helv",
    )
    page.insert_text((60, 175), "- Keep the text editable in PowerPoint", fontsize=15, fontname="helv")
    page.insert_text((60, 200), "- Place the image separately", fontsize=15, fontname="helv")

    image_rect = fitz.Rect(360, 120, 660, 255)
    page.insert_image(image_rect, stream=_build_demo_image())

    doc.save(OUTPUT_PATH)
    doc.close()
    return OUTPUT_PATH


if __name__ == "__main__":
    output = build_demo_pdf()
    print(output)
