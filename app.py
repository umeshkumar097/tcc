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

#======================================================================
# HELPER FUNCTIONS (from utils.py)
#======================================================================

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
        try:
            shutil.rmtree(directory)
            print(f"Cleaned up temporary directory: {directory}")
        except Exception as e:
            print(f"Error cleaning up directory {directory}: {e}")

def get_excel_df(excel_file_buffer) -> pd.DataFrame:
    return pd.read_excel(excel_file_buffer)

#======================================================================
# IMAGE AND PDF PROCESSING CLASS (from ImageFormFiller.py)
#======================================================================

class ImageFormFiller:
    def __init__(self, template_image: Image.Image, mapping_data: dict, font_path: str, font_size: int = 48):
        self.template_image = template_image
        self.mapping_data = mapping_data
        self.output_font_path = font_path
        self.output_font_size = font_size
        
        if self.template_image.mode != 'RGB':
            self.template_image = self.template_image.convert('RGB')
        
        self.dpi = 300
        self.page_width_pts = (self.template_image.width / self.dpi) * inch
        self.page_height_pts = (self.template_image.height / self.dpi) * inch
        
        try:
            self.pil_font = ImageFont.truetype(self.output_font_path, self.output_font_size)
        except IOError:
            print(f"FATAL: Font file not found at '{self.output_font_path}'. Please ensure it is in the assets folder.")
            self.pil_font = ImageFont.load_default()

    def _draw_text_on_image(self, draw: ImageDraw.Draw, text: str, x: int, y: int, w: int, h: int, color=(0, 0, 0)):
        text_bbox = draw.textbbox((0,0), text, font=self.pil_font)
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]
        draw_y = y + (h - text_height) // 2
        
        if text_width > w:
            words = text.split(' ')
            lines = []
            current_line = ""
            for word in words:
                test_line = f"{current_line} {word}".strip()
                test_bbox = draw.textbbox((0,0), test_line, font=self.pil_font)
                test_width = test_bbox[2] - test_bbox[0]
                if test_width <= w:
                    current_line = test_line
                else:
                    lines.append(current_line)
                    current_line = word
            lines.append(current_line)
            line_height = text_height
            start_y = y + (h - len(lines) * line_height) // 2
            for i, line in enumerate(lines):
                line_bbox = draw.textbbox((0,0), line, font=self.pil_font)
                line_width = line_bbox[2] - line_bbox[0]
                draw_x_line = x + (w - line_width) // 2
                draw_y_line = start_y + i * line_height
                if draw_y_line + line_height <= y + h:
                    draw.text((draw_x_line, draw_y_line), line, font=self.pil_font, fill=color)
        else:
            draw_x = x + (w - text_width) // 2
            draw.text((draw_x, draw_y), text, font=self.pil_font, fill=color)

    def fill_and_save_pdf(self, output_folder: str, candidate_data: dict, candidate_srno: str, candidate_name: str, photo_path: str = None):
        filled_image = self.template_image.copy()
        draw = ImageDraw.Draw(filled_image)

        for field_name, coords in self.mapping_data["fields"].items():
            x, y, w, h = coords["x"], coords["y"], coords["w"], coords["h"]
            if field_name.lower() == "photo" and photo_path and os.path.exists(photo_path):
                try:
                    photo = Image.open(photo_path)
                    photo = photo.resize((w, h), Image.Resampling.LANCZOS)
                    filled_image.paste(photo, (x, y))
                except Exception as e:
                    print(f"Error placing photo for {candidate_name}: {e}")
            else:
                field_value = None
                for key, value in candidate_data.items():
                    if key.lower().replace("_", " ") == field_name.lower().replace("_", " "):
                        field_value = str(value)
                        break
                if field_value:
                    self._draw_text_on_image(draw, field_value, x, y, w, h, color=(0, 0, 0))

        img_byte_arr = BytesIO()
        filled_image.save(img_byte_arr, format='PNG', dpi=(self.dpi, self.dpi))
        img_byte_arr.seek(0)
        
        output_filename = f"{candidate_srno}_{candidate_name.replace(' ', '_')}_filled.pdf"
        output_pdf_path = os.path.join(output_folder, output_filename)
        
        c = canvas.Canvas(output_pdf_path, pagesize=(self.page_width_pts, self.page_height_pts))
        c.drawImage(ImageReader(img_byte_arr), 0, 0, width=self.page_width_pts, height=self.page_height_pts)
        c.save()
        
        return output_pdf_path

#======================================================================
# STREAMLIT APP UI AND MAIN LOGIC
#======================================================================

st.set_page_config(page_title="Aiclex Bulk Form Filler", page_icon="✍️", layout="wide", initial_sidebar_state="expanded")

# --- Constants & Setup ---
TEMP_DIR = "temp_uploads"
OUTPUT_DIR = "final_output"
MAPPING_JSON_FILENAME = "template_mapping.json"
FONT_PATH = "assets/DejaVuSans.ttf"

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- Sidebar ---
st.sidebar.image('assets/aiclex_logo.png', use_column_width=True)
st.sidebar.markdown('<h1 style="color:#1E3A8A; font-size: 28px;">Aiclex Bulk Form Filler</h1>', unsafe_allow_html=True)
st.sidebar.markdown("---")
st.sidebar.write("Automate filling forms with candidate data & photos.")
st.sidebar.markdown("---")

# --- Main App Tabs ---
tab1, tab2, tab3 = st.tabs(["🚀 Overview & Setup", "✍️ Template Mapping", "🔄 Process Forms"])

with tab1:
    st.header("Welcome to Aiclex Bulk Form Filler!")
    st.markdown("""
        This application automates filling out form templates with candidate data and photos.
        
        ### How to Use
        1.  **Go to the `Template Mapping` tab:** Here, you'll upload your blank form and draw boxes on it to tell the app where each piece of information (like Name, Address, Photo) should go.
        2.  **Go to the `Process Forms` tab:** Once mapping is done, come here to upload your Excel file and photos to generate all the filled forms at once.
    """)
    
with tab2:
    st.header("✍️ Template Mapping")
    st.info("Upload your blank form image and draw boxes where text and photos should be placed.")

    if 'mapping_data' not in st.session_state:
        st.session_state.mapping_data = {"image_size": [0, 0], "fields": {}}
    
    uploaded_template_file = st.file_uploader("**1. Upload Blank Form Image (PNG/JPG)**", type=["png", "jpg", "jpeg"], key="mapping_uploader")

    if uploaded_template_file:
        template_image_pil = Image.open(uploaded_template_file)
        st.session_state.mapping_data["image_size"] = template_image_pil.size
        
        st.markdown("---")
        st.subheader("**2. Draw Fields on the Template**")
        
        display_width = 800
        display_height = int(template_image_pil.height * (display_width / template_image_pil.width))
        
        canvas_result = st_canvas(
            fill_color="rgba(255, 165, 0, 0.3)", stroke_width=2, stroke_color="#EE6002",
            background_image=template_image_pil, update_streamlit=True,
            height=display_height, width=display_width, drawing_mode="rect", key="canvas"
        )
        
        st.markdown("---")
        st.subheader("**3. Name the Fields You Drew**")
        
        if canvas_result.json_data and canvas_result.json_data["objects"]:
            for i, obj in enumerate(canvas_result.json_data["objects"]):
                if obj['type'] == 'rect':
                    scale_w = template_image_pil.width / display_width
                    scale_h = template_image_pil.height / display_height
                    x = int(obj['left'] * scale_w)
                    y = int(obj['top'] * scale_h)
                    w = int(obj['width'] * scale_w)
                    h = int(obj['height'] * scale_h)

                    field_name = st.text_input(f"Name for Box {i+1} (e.g., Name, Address, Photo)", key=f"field_name_{i}")
                    if field_name:
                        st.session_state.mapping_data["fields"][field_name] = {"x": x, "y": y, "w": w, "h": h}
        
        st.markdown("---")
        st.subheader("**4. Download Your Mapping File**")
        if st.session_state.mapping_data["fields"]:
            st.json(st.session_state.mapping_data["fields"])
            st.download_button(
                label="Download Mapping JSON",
                data=json.dumps(st.session_state.mapping_data, indent=2),
                file_name=MAPPING_JSON_FILENAME,
                mime="application/json"
            )

with tab3:
    st.header("🔄 Process Forms")
    st.info("Upload all your files here to begin generating the documents.")

    uploaded_mapping_file = st.file_uploader("**1. Upload Mapping JSON File**", type=["json"])
    uploaded_template_for_processing = st.file_uploader("**2. Upload Blank Form Image**", type=["png", "jpg", "jpeg"])
    uploaded_excel_file = st.file_uploader("**3. Upload Candidate Data Excel File (.xlsx)**", type=["xlsx"])
    uploaded_photos_zip = st.file_uploader("**4. Upload Candidate Photos ZIP File**", type=["zip"])

    if st.button("🚀 Start Processing", type="primary", use_container_width=True):
        if uploaded_mapping_file and uploaded_template_for_processing and uploaded_excel_file and uploaded_photos_zip:
            with st.spinner("Processing forms... this might take a moment."):
                mapping_data = json.load(uploaded_mapping_file)
                template_image = Image.open(uploaded_template_for_processing)
                candidate_df = get_excel_df(uploaded_excel_file)
                
                zip_path = os.path.join(TEMP_DIR, uploaded_photos_zip.name)
                with open(zip_path, "wb") as f:
                    f.write(uploaded_photos_zip.getbuffer())
                
                run_temp_dir = os.path.join(TEMP_DIR, "current_run")
                if os.path.exists(run_temp_dir): shutil.rmtree(run_temp_dir)
                photo_base_dir = unzip_and_organize_files(zip_path, os.path.join(run_temp_dir, "photos"))
                
                run_output_dir = os.path.join(OUTPUT_DIR, "results")
                if os.path.exists(run_output_dir): shutil.rmtree(run_output_dir)
                os.makedirs(run_output_dir, exist_ok=True)

                filler = ImageFormFiller(template_image, mapping_data, font_path=FONT_PATH, font_size=48)
                progress_bar = st.progress(0)
                
                for index, row in candidate_df.iterrows():
                    sr_no_col = next((c for c in ['SrNo', 'Sl No.', 'SNo', 'Serial'] if c in row.index), None)
                    if not sr_no_col: continue

                    candidate_srno = str(row[sr_no_col]).split('.')[0]
                    candidate_name = row.get('Name', f"Candidate_{candidate_srno}")
                    output_candidate_folder = os.path.join(run_output_dir, f"{candidate_srno} {candidate_name}")
                    os.makedirs(output_candidate_folder, exist_ok=True)
                    
                    candidate_photo_path = None
                    for root, _, files in os.walk(photo_base_dir):
                        photo_found = False
                        for file in files:
                            if candidate_srno in os.path.basename(root) or candidate_srno in file:
                                candidate_photo_path = os.path.join(root, file)
                                shutil.copy(candidate_photo_path, output_candidate_folder)
                                photo_found = True
                                break
                        if photo_found: break
                    
                    filler.fill_and_save_pdf(output_candidate_folder, row.to_dict(), candidate_srno, candidate_name, candidate_photo_path)
                    progress_bar.progress((index + 1) / len(candidate_df))

                final_zip_path = os.path.join(OUTPUT_DIR, "final_filled_results.zip")
                if os.path.exists(final_zip_path): os.remove(final_zip_path)
                create_output_zip(run_output_dir, final_zip_path)

                with open(final_zip_path, "rb") as fp:
                    st.download_button("✅ Success! Download Final ZIP", fp, "final_filled_results.zip", "application/zip", type="primary")

                clean_temp_dirs(run_temp_dir)
        else:
            st.error("Please upload all four required files before starting.")
