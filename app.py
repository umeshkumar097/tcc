import streamlit as st
import pandas as pd
import json
import os
import zipfile
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
from streamlit_drawable_canvas import st_canvas
import shutil
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
import numpy as np

# --- All Helper Functions are included in this single file ---
def unzip_and_organize_files(zip_file_path: str, destination_dir: str):
    os.makedirs(destination_dir, exist_ok=True)
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(destination_dir)
    return destination_dir

def create_output_zip(source_dir: str, output_zip_path: str):
    shutil.make_archive(os.path.splitext(output_zip_path)[0], 'zip', source_dir)
    return output_zip_path

def clean_temp_dirs(directory: str):
    if os.path.exists(directory) and os.path.isdir(directory):
        try: shutil.rmtree(directory)
        except Exception as e: print(f"Error cleaning directory: {e}")

def get_excel_df(excel_file_buffer) -> pd.DataFrame:
    return pd.read_excel(excel_file_buffer)

class ImageFormFiller:
    def __init__(self, template_image: Image.Image, mapping_data: dict, font_path: str, font_size: int = 48):
        self.template_image = template_image.convert('RGB')
        self.mapping_data = mapping_data
        self.font_path = font_path
        self.font_size = font_size
        self.dpi = 300
        self.page_width_pts = (self.template_image.width / self.dpi) * inch
        self.page_height_pts = (self.template_image.height / self.dpi) * inch
        try: self.pil_font = ImageFont.truetype(self.font_path, self.font_size)
        except IOError: self.pil_font = ImageFont.load_default()

    def _draw_text_on_image(self, draw: ImageDraw.Draw, text: str, x: int, y: int, w: int, h: int):
        draw.text((x, y), text, font=self.pil_font, fill=(0, 0, 0))

    def fill_and_save_pdf(self, output_folder: str, candidate_data: dict, srno: str, name: str, photo_path: str = None):
        filled_image = self.template_image.copy()
        draw = ImageDraw.Draw(filled_image)
        for field, coords in self.mapping_data["fields"].items():
            x, y, w, h = coords["x"], coords["y"], coords["w"], coords["h"]
            if field.lower() == "photo" and photo_path:
                try:
                    with Image.open(photo_path) as photo:
                        photo = photo.resize((w, h), Image.Resampling.LANCZOS)
                        filled_image.paste(photo, (x, y))
                except Exception as e: print(f"Error with photo for {name}: {e}")
            else:
                value = next((str(v) for k, v in candidate_data.items() if k.lower().replace("_", " ") == field.lower()), None)
                if value: self._draw_text_on_image(draw, value, x, y, w, h)
        
        pdf_path = os.path.join(output_folder, f"{srno}_{name.replace(' ', '_')}.pdf")
        with BytesIO() as img_buffer:
            filled_image.save(img_buffer, format='PNG', dpi=(self.dpi, self.dpi))
            img_buffer.seek(0)
            c = canvas.Canvas(pdf_path, pagesize=(self.page_width_pts, self.page_height_pts))
            c.drawImage(ImageReader(img_buffer), 0, 0, width=self.page_width_pts, height=self.page_height_pts)
            c.save()

# --- Streamlit App UI ---
st.set_page_config(page_title="Aiclex Bulk Form Filler", layout="wide")
TEMP_DIR, OUTPUT_DIR = "temp", "output"
FONT_PATH = "assets/DejaVuSans.ttf"
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

st.sidebar.markdown('<h1 style="color:#1E3A8A;">Aiclex Technologies</h1>', unsafe_allow_html=True)
st.sidebar.markdown('<h3>Bulk Form Filler</h3>', unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs(["🚀 Overview", "✍️ Template Mapping (Drawing Mode)", "🔄 Process Forms"])

with tab1:
    st.header("Welcome!")
    st.write("Use the 'Template Mapping' tab to draw on your form, then use 'Process Forms' to generate the documents.")

with tab2:
    st.header("✍️ Template Mapping (Drawing Mode)")
    st.info("Draw rectangles on the image below and name them. The app will calculate the coordinates for you.")

    if 'mapping_data' not in st.session_state:
        st.session_state.mapping_data = {"image_size": [0, 0], "fields": {}}

    uploaded_template = st.file_uploader("1. Upload Your Blank Form Image", type=["png", "jpg"])
    
    if uploaded_template:
        template_image = Image.open(uploaded_template).convert("RGBA")
        
        original_w, original_h = template_image.size
        st.session_state.mapping_data["image_size"] = [original_w, original_h]
        
        st.subheader("2. Draw Boxes on the Image and Name Them")
        
        display_width = 800
        display_height = int(original_h * (display_width / original_w))
        
        canvas_result = st_canvas(
            fill_color="rgba(255, 165, 0, 0.3)",
            stroke_width=2,
            background_image=template_image,
            update_streamlit=True,
            height=display_height,
            width=display_width,
            drawing_mode="rect",
            key="canvas",
        )

        if canvas_result.json_data is not None and canvas_result.json_data["objects"]:
            st.subheader("3. Name the Boxes You Drew")
            field_names = {}
            for i, obj in enumerate(canvas_result.json_data["objects"]):
                field_names[i] = st.text_input(f"Name for Box {i+1}", key=f"field_name_{i}")
            
            if st.button("Confirm Field Names"):
                st.session_state.mapping_data["fields"] = {} 
                for i, obj in enumerate(canvas_result.json_data["objects"]):
                    field_name = field_names.get(i)
                    if field_name:
                        scale_w = original_w / display_width
                        scale_h = original_h / display_height
                        st.session_state.mapping_data["fields"][field_name] = {
                            "x": int(obj['left'] * scale_w),
                            "y": int(obj['top'] * scale_h),
                            "w": int(obj['width'] * scale_w),
                            "h": int(obj['height'] * scale_h)
                        }
                st.success("Field names confirmed and coordinates saved!")
                st.experimental_rerun()
        
        st.subheader("4. Download Your Mapping File")
        if st.session_state.mapping_data["fields"]:
            st.json(st.session_state.mapping_data["fields"])
            st.download_button(
                label="Download Mapping JSON",
                data=json.dumps(st.session_state.mapping_data, indent=2),
                file_name="drawable_mapping.json",
                mime="application/json"
            )

with tab3:
    st.header("🔄 Process Forms")
    uploaded_template_for_processing = st.file_uploader("1. Upload the Same Blank Form Image Again", type=["png", "jpg"])
    mapping_file = st.file_uploader("2. Upload Your Saved Mapping JSON", type=["json"])
    excel_file = st.file_uploader("3. Upload Candidate Excel File", type=["xlsx"])
    zip_file = st.file_uploader("4. Upload Candidate Photos ZIP", type=["zip"])

    if st.button("🚀 Start Processing", type="primary"):
        if all([uploaded_template_for_processing, mapping_file, excel_file, zip_file]):
            with st.spinner("Processing..."):
                template_image_process = Image.open(uploaded_template_for_processing)
                mapping = json.load(mapping_file)
                df = get_excel_df(excel_file)
                
                zip_path = os.path.join(TEMP_DIR, zip_file.name)
                with open(zip_path, "wb") as f: f.write(zip_file.getbuffer())
                
                photo_dir = unzip_and_organize_files(zip_path, os.path.join(TEMP_DIR, "photos"))
                output_run_dir = os.path.join(OUTPUT_DIR, "run")
                if os.path.exists(output_run_dir): shutil.rmtree(output_run_dir)
                os.makedirs(output_run_dir)

                filler = ImageFormFiller(template_image_process, mapping, FONT_PATH, font_size=24)
                
                progress_bar = st.progress(0)
                total_rows = len(df)
                for i, row in df.iterrows():
                    sr_col = next((c for c in ['SrNo', 'Sl No.', 'SNo', 'Serial'] if c in df.columns), None)
                    if not sr_col:
                        st.error("Error: 'SrNo' or 'Sl No.' column not found in Excel.")
                        break
                    
                    srno = str(row[sr_col]).split('.')[0]
                    name = row.get('Name', f"Candidate_{srno}")
                    
                    photo_path = None
                    for root, _, files in os.walk(photo_dir):
                        found = False
                        for file in files:
                            if srno in file or srno in os.path.basename(root):
                                photo_path = os.path.join(root, file)
                                found = True
                                break
                        if found: break
                    
                    candidate_folder = os.path.join(output_run_dir, f"{srno} {name}")
                    os.makedirs(candidate_folder, exist_ok=True)
                    if photo_path: shutil.copy(photo_path, candidate_folder)
                    
                    filler.fill_and_save_pdf(candidate_folder, row.to_dict(), srno, name, photo_path)
                    progress_bar.progress((i + 1) / total_rows)

                final_zip = os.path.join(OUTPUT_DIR, "final_results.zip")
                create_output_zip(output_run_dir, final_zip)
                with open(final_zip, "rb") as fp:
                    st.download_button("✅ Download Final ZIP", fp, "final_results.zip", "application/zip")
                clean_temp_dirs(TEMP_DIR)
        else:
            st.error("Please upload all four files to start processing.")
