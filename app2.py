import streamlit as st
from pptx import Presentation 
from pptx.util import Inches, Pt
import easyocr
import cv2
import numpy as np
from PIL import Image
import io
import os

# 1. Initialize the OCR Reader (English)
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['en'])

reader = load_ocr()

def process_pptx(input_file):
    prs = Presentation(input_file)
    output_prs = Presentation()
 
    # Get slide dimensions for coordinate mapping
    slide_width_pts = output_prs.slide_width.pt
    slide_height_pts = output_prs.slide_height.pt

    st.info(f"Processing {len(prs.slides)} slides...")

    for i, slide in enumerate(prs.slides):
        new_slide = output_prs.slides.add_slide(output_prs.slide_layouts[6]) # Blank layout
        
        for shape in slide.shapes:
            if shape.shape_type == 13:  # 13 is a Picture
                # Convert PPTX image to OpenCV format
                image_stream = io.BytesIO(shape.image.blob)
                img = Image.open(image_stream).convert('RGB')
                img_np = np.array(img)
                
                # Perform OCR
                results = reader.readtext(img_np)
                
                # Get Image dimensions in pixels
                img_h, img_w, _ = img_np.shape

                for (bbox, text, prob) in results:
                    if prob < 0.2: continue # Skip low confidence
                    
                    # Calculate relative position within the image
                    # bbox is: [[x_topleft, y_topleft], [x_tr, y_tr], [x_br, y_br], [x_bl, y_bl]]
                    x_px, y_px = bbox[0][0], bbox[0][1]
                    w_px = bbox[1][0] - bbox[0][0]
                    h_px = bbox[2][1] - bbox[1][1]

                    # Map image pixels to PPTX points
                    # (Note: This assumes the image fills the slide; 
                    # for perfection, we'd scale by the shape's actual size on slide)
                    left = (x_px / img_w) * slide_width_pts
                    top = (y_px / img_h) * slide_height_pts
                    width = (w_px / img_w) * slide_width_pts
                    height = (h_px / img_h) * slide_height_pts

                    # Add text box to the NEW slide
                    txBox = new_slide.shapes.add_textbox(Pt(left), Pt(top), Pt(width), Pt(height))
                    tf = txBox.text_frame
                    tf.text = text
                    
                    # Basic styling
                    for paragraph in tf.paragraphs:
                        paragraph.font.size = Pt(14)

    # Save to buffer
    output_stream = io.BytesIO()
    output_prs.save(output_stream)
    return output_stream.getvalue()

# --- Streamlit UI ---
st.title("PPTX Un-Flattener 🛠️")
st.write("Upload a PPT with flat images to convert them into editable text.")

uploaded_file = st.file_uploader("Choose a PowerPoint file", type="pptx")

if uploaded_file:
    if st.button("Convert to Editable"):
        with st.spinner("Analyzing images... this may take a minute."):
            try:
                processed_data = process_pptx(uploaded_file)
                st.success("Conversion Complete!")
                st.download_button(
                    label="Download Editable PPTX",
                    data=processed_data,
                    file_name="editable_output.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error(f"An error occurred: {e}")