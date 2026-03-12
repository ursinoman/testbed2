import streamlit as st
from pptx import Presentation 
from pptx.util import Inches, Pt
import easyocr
import cv2
import numpy as np
from PIL import Image
import io
import os

# Initialize OCR
@st.cache_resource
def load_ocr():
    return easyocr.Reader(['en'])

reader = load_ocr()

def process_pptx_advanced(input_file):
    prs = Presentation(input_file)
    output_prs = Presentation()
 
    # Standard PPTX Slide Dimensions (960 x 540 in Points)
    slide_width_pts = output_prs.slide_width.pt
    slide_height_pts = output_prs.slide_height.pt

    for slide in prs.slides:
        new_slide = output_prs.slides.add_slide(output_prs.slide_layouts[6])
        
        for shape in slide.shapes:
            if shape.shape_type == 13:  # Picture
               image_stream = io.BytesIO(shape.image.blob)
               img_pil = Image.open(image_stream).convert('RGB')
               img_np = np.array(img_pil)
               img_h, img_w, _ = img_np.shape
                
               # 1. OCR to find Text Bounding Boxes
               results = reader.readtext(img_np)
                
               # 2. Create a Mask for Inpainting (to remove text from image)
               mask = np.zeros((img_h, img_w), dtype=np.uint8)

               # List to store text data for later placement
               extracted_text_blocks = []
                    
               for (bbox, text, prob) in results:
                    if prob < 0.2: continue

                    # Convert bbox to integer coordinates
                    points = np.array(bbox).astype(int)
                    cv2.fillPoly(mask, [points], 255)

                    extracted_text_blocks.append({
                        'text': text,
                        'bbox': bbox,
                        'prob': prob
                    })
                    
                # 3. Perform Inpainting (Remove text, keep background images)
                # This makes the "image" part clean so text isn't "baked in"
                    cleaned_img_np = cv2.inpaint(img_np, mask, 3, cv2.INPAINT_TELEA)

                # 4. Save the "cleaned" image (images/graphics) back to the slide
                    cleaned_pil = Image.fromarray(cleaned_img_np)
                    img_byte_arr = io.BytesIO()
                    cleaned_pil.save(img_byte_arr, format='PNG')
               
                # Place the background image first
                    new_slide.shapes.add_picture(img_byte_arr, 0, 0, Pt(slide_width_pts), Pt(slide_height_pts))

                # 5. Place Editable Text Boxes on top
                    for block in extracted_text_blocks:
                        bbox = np.block('bbox')

                        # Before the crashing line, add a check:
                        if bbox is None or (isinstance(bbox, (list, tuple, dict)) and len(bbox) == 0):      
                            # Handle the empty/None case, e.g., continue or skip this element
                            print("No bounding box found, skipping.")
                            continue 

                        # Now it is safe to access
                        if bbox and len(bbox) > 0:
                             x_px, y_px = bbox[0][0], bbox[0][1]
                        else:
                            # Handle the empty case, e.g., skip it or set default coordinates
                            x_px, y_px = 0, 0
                            # print("Skipping empty box") # Optional debugging   
                        w_px = bbox[1][0] - bbox[0][0]
                        h_px = bbox[2][1] - bbox[1][1]

                    left = (x_px / img_w) * slide_width_pts
                    top = (y_px / img_h) * slide_height_pts
                    width = (w_px / img_w) * slide_width_pts
                    height = (h_px / img_h) * slide_height_pts

                    txBox = new_slide.shapes.add_textbox(Pt(left), Pt(top), Pt(width), Pt(height))
                    tf = txBox.text_frame
                    tf.word_wrap = True
                    p = tf.paragraphs[0]
                    p.text = block['text']
                    p.font.size = Pt(12)

    output_stream = io.BytesIO()
    output_prs.save(output_stream)
    return output_stream.getvalue()

# --- Streamlit UI ---
st.title("Full PPTX Reconstructor 💎")
st.write("Converts flat images into editable text + cleaned background graphics.")

file = st.file_uploader("Upload PPTX", type="pptx")
if file and st.button("Extract Everything"):
    with st.spinner("Decomposing images..."):
        result = process_pptx_advanced(file)
        st.download_button("Download Reconstructed PPTX", result, "fully_editable.pptx")