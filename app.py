import streamlit as st
import pandas as pd
import json
import os
import zipfile
from PIL import Image
from io import BytesIO
from streamlit_drawable_canvas import st_canvas
import shutil
from ImageFormFiller import ImageFormFiller
from utils import unzip_and_organize_files, create_output_zip, clean_temp_dirs, get_image_from_bytes, get_excel_df

st.set_page_config(page_title="Aiclex Bulk Form Filler", page_icon="✍️", layout="wide", initial_sidebar_state="expanded")

TEMP_DIR = "temp_uploads"
OUTPUT_DIR = "final_output"
MAPPING_JSON_FILENAME = "template_mapping.json"
FONT_PATH = "assets/DejaVuSans.ttf"

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

st.sidebar.image('assets/aiclex_logo.png', use_column_width=True)
st.sidebar.markdown('<h1 style="color:#1E3A8A; font-size: 28px;">Aiclex Bulk Form Filler</h1>', unsafe_allow_html=True)
st.sidebar.markdown("---")
st.sidebar.write("Automate filling forms with candidate data & photos.")
st.sidebar.markdown("---")

tab1, tab2, tab3 = st.tabs(["🚀 Overview & Setup", "✍️ Template Mapping", "🔄 Process Forms"])

with tab1:
    st.header("Welcome to Aiclex Bulk Form Filler!")
    st.markdown("""
        This application helps you automate filling out form templates with candidate data and photos.
        
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
