import streamlit as st
import pandas as pd
import json
import os
import shutil
from PIL import Image
from io import BytesIO

# Import from your other python files
from ImageFormFiller import ImageFormFiller
from utils import unzip_and_organize_files, create_output_zip, clean_temp_dirs, get_excel_df

#======================================================================
# STREAMLIT APP UI AND MAIN LOGIC
#======================================================================
st.set_page_config(page_title="Aiclex Bulk Form Filler", page_icon="✍️", layout="wide")

# --- Constants & Setup ---
TEMP_DIR = "temp_uploads"
OUTPUT_DIR = "final_output"
FONT_PATH = "assets/DejaVuSans.ttf"
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- Sidebar ---
st.sidebar.markdown('<h1 style="color:#1E3A8A;">Aiclex Technologies</h1>', unsafe_allow_html=True)
st.sidebar.markdown('<h3>Bulk Form Filler</h3>', unsafe_allow_html=True)
st.sidebar.markdown("---")

# --- Main App Tabs ---
tab1, tab2, tab3 = st.tabs(["🚀 Overview", "✍️ Template Mapping", "🔄 Process Forms"])

with tab1:
    st.header("Welcome to Aiclex Bulk Form Filler!")
    st.markdown("This application automates filling out forms with data from Excel and candidate photos.")

with tab2:
    st.header("✍️ Template Mapping (Manual Mode)")
    st.info("This is a reliable manual method to define where fields go on your form.")

    if 'mapping_data' not in st.session_state:
        st.session_state.mapping_data = {"image_size": [0, 0], "fields": {}}

    uploaded_template_file = st.file_uploader("**1. Upload Blank Form Image**", type=["png", "jpg", "jpeg"])

    if uploaded_template_file:
        template_image_pil = Image.open(uploaded_template_file)
        w, h = template_image_pil.size
        st.session_state.mapping_data["image_size"] = [w, h]

        st.markdown("---")
        st.image(template_image_pil, caption="Your Uploaded Template")
        st.success(f"**Image Dimensions:** Width = {w}px, Height = {h}px")
        st.warning("Use an image editor (like MS Paint or Mac Preview) to find the pixel coordinates for each field.")

        st.markdown("---")
        st.subheader("2. Add Fields Manually")

        with st.form("mapping_form"):
            cols = st.columns([2, 1, 1, 1, 1])
            field_name = cols[0].text_input("Field Name (e.g., Name, Photo)")
            x_coord = cols[1].number_input("X", min_value=0, step=10)
            y_coord = cols[2].number_input("Y", min_value=0, step=10)
            width = cols[3].number_input("Width", min_value=10, step=10, value=300)
            height = cols[4].number_input("Height", min_value=10, step=10, value=50)
            submitted = st.form_submit_button("Add Field")

            if submitted and field_name:
                st.session_state.mapping_data["fields"][field_name] = {"x": x_coord, "y": y_coord, "w": width, "h": height}
                st.success(f"Added field '{field_name}'")

        st.markdown("---")
        st.subheader("3. Review and Download Mapping JSON")
        if st.session_state.mapping_data["fields"]:
            st.json(st.session_state.mapping_data["fields"])
            st.download_button(
                label="Download Mapping JSON",
                data=json.dumps(st.session_state.mapping_data, indent=2),
                file_name="manual_mapping.json",
                mime="application/json"
            )

with tab3:
    st.header("🔄 Process Forms")
    st.info("Upload all your files here to begin generating the documents.")

    uploaded_mapping_file = st.file_uploader("**1. Upload Mapping JSON File**", type=["json"])
    uploaded_template_for_processing = st.file_uploader("**2. Upload Blank Form Image**", type=["png", "jpg", "jpeg"])
    uploaded_excel_file = st.file_uploader("**3. Upload Candidate Data Excel File (.xlsx)**", type=["xlsx"])
    uploaded_photos_zip = st.file_uploader("**4. Upload Candidate Photos ZIP File**", type=["zip"])

    if st.button("🚀 Start Processing", type="primary"):
        if all([uploaded_mapping_file, uploaded_template_for_processing, uploaded_excel_file, uploaded_photos_zip]):
            with st.spinner("Processing..."):
                mapping_data = json.load(uploaded_mapping_file)
                template_image = Image.open(uploaded_template_for_processing)
                candidate_df = get_excel_df(uploaded_excel_file)
                
                zip_path = os.path.join(TEMP_DIR, uploaded_photos_zip.name)
                with open(zip_path, "wb") as f: f.write(uploaded_photos_zip.getbuffer())
                
                run_temp_dir = os.path.join(TEMP_DIR, "run")
                if os.path.exists(run_temp_dir): shutil.rmtree(run_temp_dir)
                photo_base_dir = unzip_and_organize_files(zip_path, run_temp_dir)
                
                run_output_dir = os.path.join(OUTPUT_DIR, "results")
                if os.path.exists(run_output_dir): shutil.rmtree(run_output_dir)
                os.makedirs(run_output_dir)

                filler = ImageFormFiller(template_image, mapping_data, font_path=FONT_PATH, font_size=24)
                
                progress_bar = st.progress(0)
                for i, row in candidate_df.iterrows():
                    sr_no_col = next((c for c in ['SrNo', 'Sl No.', 'SNo', 'Serial'] if c in row.index), None)
                    if not sr_no_col: continue
                    
                    candidate_srno = str(row[sr_no_col]).split('.')[0]
                    candidate_name = row.get('Name', f"Candidate_{candidate_srno}")
                    output_folder = os.path.join(run_output_dir, f"{candidate_srno} {candidate_name}")
                    os.makedirs(output_folder)
                    
                    photo_path = None
                    for root, _, files in os.walk(photo_base_dir):
                        found = False
                        for file in files:
                            if candidate_srno in os.path.basename(root) or candidate_srno in file:
                                photo_path = os.path.join(root, file)
                                shutil.copy(photo_path, output_folder)
                                found = True
                                break
                        if found: break
                    
                    filler.fill_and_save_pdf(output_folder, row.to_dict(), candidate_srno, candidate_name, photo_path)
                    progress_bar.progress((i + 1) / len(candidate_df))

                final_zip_path = os.path.join(OUTPUT_DIR, "final_results.zip")
                create_output_zip(run_output_dir, final_zip_path)

                with open(final_zip_path, "rb") as fp:
                    st.download_button("✅ Success! Download Final ZIP", fp, "final_results.zip", "application/zip", type="primary")

                clean_temp_dirs(run_temp_dir)
        else:
            st.error("Please upload all four required files.")
