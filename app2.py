import io
import os
from pathlib import Path

import cv2
import easyocr
import numpy as np
import streamlit as st
from PIL import Image
from pptx.chart.data import CategoryChartData
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.util import Pt


OCR_LANGUAGES = ["en"]
MIN_OCR_CONFIDENCE = 0.2
DEFAULT_FONT_SIZE_PT = 14


def _easyocr_cache_dir() -> str:
    # Streamlit deployments can fail if EasyOCR tries to write under HOME.
    cache_dir = Path(os.getenv("EASYOCR_MODULE_STORAGE", "/tmp/easyocr"))
    cache_dir.mkdir(parents=True, exist_ok=True)
    return str(cache_dir)


@st.cache_resource(show_spinner=False)
def load_ocr():
    return easyocr.Reader(
        OCR_LANGUAGES,
        model_storage_directory=_easyocr_cache_dir(),
    )


def _shape_bounds_to_points(shape):
    return (
        shape.left.pt,
        shape.top.pt,
        shape.width.pt,
        shape.height.pt,
    )


def _copy_color(source_color, target_color):
    try:
        if source_color.rgb is not None:
            target_color.rgb = RGBColor(*source_color.rgb)
    except Exception:
        return


def _copy_fill(source_fill, target_fill):
    try:
        if source_fill.type is None:
            return
        if source_fill.fore_color.rgb is not None:
            target_fill.solid()
            _copy_color(source_fill.fore_color, target_fill.fore_color)
    except Exception:
        return


def _copy_line(source_line, target_line):
    try:
        if source_line.color.rgb is not None:
            _copy_color(source_line.color, target_line.color)
        if source_line.width is not None:
            target_line.width = source_line.width
    except Exception:
        return


def _copy_font(source_font, target_font):
    for attr in ("bold", "italic", "underline", "name", "size"):
        try:
            value = getattr(source_font, attr)
        except Exception:
            continue
        if value is not None:
            setattr(target_font, attr, value)

    try:
        if source_font.color.rgb is not None:
            _copy_color(source_font.color, target_font.color)
    except Exception:
        return


def _copy_text_frame(source_text_frame, target_text_frame):
    target_text_frame.clear()
    target_text_frame.word_wrap = source_text_frame.word_wrap

    for index, source_paragraph in enumerate(source_text_frame.paragraphs):
        paragraph = (
            target_text_frame.paragraphs[0]
            if index == 0
            else target_text_frame.add_paragraph()
        )
        paragraph.alignment = source_paragraph.alignment
        paragraph.level = source_paragraph.level

        if source_paragraph.runs:
            for run_index, source_run in enumerate(source_paragraph.runs):
                run = paragraph.add_run()
                run.text = source_run.text
                _copy_font(source_run.font, run.font)
        else:
            paragraph.text = source_paragraph.text
            _copy_font(source_paragraph.font, paragraph.font)


def _clone_text_like_shape(shape, new_slide):
    if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
        cloned_shape = new_slide.shapes.add_shape(
            shape.auto_shape_type,
            shape.left,
            shape.top,
            shape.width,
            shape.height,
        )
        _copy_fill(shape.fill, cloned_shape.fill)
        _copy_line(shape.line, cloned_shape.line)
    else:
        cloned_shape = new_slide.shapes.add_textbox(
            shape.left,
            shape.top,
            shape.width,
            shape.height,
        )

    if shape.has_text_frame:
        _copy_text_frame(shape.text_frame, cloned_shape.text_frame)
    return True


def _clone_table_shape(shape, new_slide):
    source_table = shape.table
    rows = len(source_table.rows)
    cols = len(source_table.columns)
    table_shape = new_slide.shapes.add_table(
        rows,
        cols,
        shape.left,
        shape.top,
        shape.width,
        shape.height,
    )
    target_table = table_shape.table

    for col_index, column in enumerate(source_table.columns):
        target_table.columns[col_index].width = column.width

    for row_index, row in enumerate(source_table.rows):
        target_table.rows[row_index].height = row.height

    for row_index in range(rows):
        for col_index in range(cols):
            source_cell = source_table.cell(row_index, col_index)
            target_cell = target_table.cell(row_index, col_index)
            target_cell.text = source_cell.text
            _copy_fill(source_cell.fill, target_cell.fill)
            _copy_text_frame(source_cell.text_frame, target_cell.text_frame)

    return True


def _clone_category_chart(shape, new_slide):
    chart = shape.chart
    plot = chart.plots[0]
    categories = [category.label for category in plot.categories]
    chart_data = CategoryChartData()
    chart_data.categories = categories

    for series in chart.series:
        chart_data.add_series(series.name, tuple(series.values))

    chart_shape = new_slide.shapes.add_chart(
        chart.chart_type,
        shape.left,
        shape.top,
        shape.width,
        shape.height,
        chart_data,
    )
    target_chart = chart_shape.chart

    try:
        if chart.has_title:
            target_chart.has_title = True
            target_chart.chart_title.text_frame.text = chart.chart_title.text_frame.text
    except Exception:
        pass

    return True


def _clone_chart_shape(shape, new_slide):
    try:
        return _clone_category_chart(shape, new_slide)
    except Exception:
        placeholder = new_slide.shapes.add_textbox(
            shape.left,
            shape.top,
            shape.width,
            shape.height,
        )
        placeholder.text_frame.text = "Chart detected. Auto-rebuild is not yet supported for this chart type."
        return False


def _clone_native_shape(shape, new_slide):
    if getattr(shape, "has_table", False):
        return _clone_table_shape(shape, new_slide), "tables"

    if shape.shape_type == MSO_SHAPE_TYPE.CHART:
        cloned = _clone_chart_shape(shape, new_slide)
        return cloned, "charts" if cloned else "unsupported"

    if shape.has_text_frame or shape.shape_type in {
        MSO_SHAPE_TYPE.AUTO_SHAPE,
        MSO_SHAPE_TYPE.CALLOUT,
        MSO_SHAPE_TYPE.FREEFORM,
        MSO_SHAPE_TYPE.PLACEHOLDER,
        MSO_SHAPE_TYPE.TEXT_BOX,
    }:
        return _clone_text_like_shape(shape, new_slide), "text_shapes"

    return False, "unsupported"


def _extract_text_blocks(reader, image_np: np.ndarray):
    results = reader.readtext(image_np)
    img_h, img_w = image_np.shape[:2]
    mask = np.zeros((img_h, img_w), dtype=np.uint8)
    blocks = []

    for bbox, text, prob in results:
        if prob < MIN_OCR_CONFIDENCE:
            continue

        points = np.array(bbox, dtype=np.int32)
        if points.shape != (4, 2):
            continue

        cv2.fillPoly(mask, [points], 255)
        x_coords = points[:, 0]
        y_coords = points[:, 1]
        left = int(x_coords.min())
        top = int(y_coords.min())
        right = int(x_coords.max())
        bottom = int(y_coords.max())

        if right <= left or bottom <= top:
            continue

        blocks.append(
            {
                "text": text,
                "left_px": left,
                "top_px": top,
                "width_px": right - left,
                "height_px": bottom - top,
                "confidence": float(prob),
            }
        )

    return blocks, mask


def _clean_image(image_np: np.ndarray, mask: np.ndarray) -> bytes:
    if np.count_nonzero(mask) == 0:
        cleaned_img_np = image_np
    else:
        cleaned_img_np = cv2.inpaint(image_np, mask, 3, cv2.INPAINT_TELEA)

    cleaned_pil = Image.fromarray(cleaned_img_np)
    img_bytes = io.BytesIO()
    cleaned_pil.save(img_bytes, format="PNG")
    img_bytes.seek(0)
    return img_bytes.getvalue()


def _add_text_boxes(new_slide, blocks, img_w, img_h, shape):
    shape_left, shape_top, shape_width, shape_height = _shape_bounds_to_points(shape)

    for block in blocks:
        left = shape_left + (block["left_px"] / img_w) * shape_width
        top = shape_top + (block["top_px"] / img_h) * shape_height
        width = (block["width_px"] / img_w) * shape_width
        height = (block["height_px"] / img_h) * shape_height

        textbox = new_slide.shapes.add_textbox(
            Pt(left),
            Pt(top),
            Pt(max(width, 1)),
            Pt(max(height, 1)),
        )
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        paragraph = text_frame.paragraphs[0]
        paragraph.text = block["text"]
        paragraph.font.size = Pt(DEFAULT_FONT_SIZE_PT)


def process_pptx_advanced(input_file):
    prs = Presentation(input_file)
    output_prs = Presentation()
    output_prs.slide_width = prs.slide_width
    output_prs.slide_height = prs.slide_height
    reader = load_ocr()
    report = {
        "slides_processed": len(prs.slides),
        "pictures_ocr": 0,
        "text_shapes": 0,
        "tables": 0,
        "charts": 0,
        "unsupported": 0,
    }

    for slide in prs.slides:
        new_slide = output_prs.slides.add_slide(output_prs.slide_layouts[6])

        for shape in slide.shapes:
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                _, bucket = _clone_native_shape(shape, new_slide)
                report[bucket] += 1
                continue

            image_stream = io.BytesIO(shape.image.blob)
            image = Image.open(image_stream).convert("RGB")
            image_np = np.array(image)
            img_h, img_w = image_np.shape[:2]

            blocks, mask = _extract_text_blocks(reader, image_np)
            cleaned_image_bytes = _clean_image(image_np, mask)

            new_slide.shapes.add_picture(
                io.BytesIO(cleaned_image_bytes),
                shape.left,
                shape.top,
                width=shape.width,
                height=shape.height,
            )
            _add_text_boxes(new_slide, blocks, img_w, img_h, shape)
            report["pictures_ocr"] += 1

    output_stream = io.BytesIO()
    output_prs.save(output_stream)
    output_stream.seek(0)
    return output_stream.getvalue(), report


st.title("Full PPTX Reconstructor")
st.write("Converts flat slide screenshots into editable text overlays and cleaned images.")

uploaded_file = st.file_uploader("Upload a PowerPoint deck", type="pptx")

if uploaded_file and st.button("Extract Everything"):
    with st.spinner("Decomposing slide images..."):
        try:
            result, report = process_pptx_advanced(uploaded_file)
        except Exception as exc:
            st.exception(exc)
        else:
            st.success("Reconstruction complete.")
            st.json(report)
            st.download_button(
                "Download reconstructed PPTX",
                result,
                "fully_editable.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
