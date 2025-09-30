import streamlit as st
import pandas as pd
import json
import os
import zipfile
from PIL import Image
from io import BytesIO
from streamlit_drawable_canvas import st_canvas
import shutil

# Assume these will be created in ImageFormFiller.py and utils.py
from ImageFormFiller import ImageFormFiller
from utils import unzip_and_organize_files, create_output_zip, clean_temp_dirs, get_image_from_bytes, get_excel_df

# --- Page Configuration ---
st.set_page_config(
    page_title="Aiclex Bulk Form Filler",
    page_icon="✍️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Constants ---
TEMP_DIR = "temp_uploads"
OUTPUT_DIR = "final_output"
MAPPING_JSON_FILENAME = "template_mapping.json"

# Ensure temp and output directories exist
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- Sidebar ---
st.sidebar.image('assets/aiclex_logo.png', use_column_width=True) # Make sure aiclex_logo.png is in 'assets' folder
st.sidebar.markdown('<h1 style="color:#1E3A8A; font-size: 28px;">Aiclex Bulk Form Filler</h1>', unsafe_allow_html=True)
st.sidebar.markdown("---")
st.sidebar.write("Automate filling forms with candidate data & photos.")
st.sidebar.markdown("---")

# --- Tabs ---
tab1, tab2, tab3 = st.tabs(["🚀 Overview & Setup", "✍️ Template Mapping", "🔄 Process Forms"])

with tab1:
    st.header("Welcome to Aiclex Bulk Form Filler!")
    st.markdown("""
        This application helps you automate the process of filling out form templates with candidate-specific data from an Excel sheet and passport photos from a ZIP archive.
        
        **Key Features:**
        * **Upload Excel**: Provide candidate data including serial numbers and other fields.
        * **Upload Photos ZIP**: A compressed archive containing per-candidate folders with their passport photos.
        * **Upload Template Image**: Use a blank form template in PNG/JPG format.
        * **Interactive Mapping**: Visually define where each field should be placed on your template.
        * **Generate Filled PDFs**: Produces a personalized PDF for each candidate with their data and photo.
        * **Organized Output**: Bundles all filled PDFs and original photos into a final ZIP, preserving structure.
        
        Navigate to the **Template Mapping** tab to define your form fields, and then to the **Process Forms** tab to generate your filled documents!
        
        ---
        ### Quick Start
        1.  **Template Mapping**:
            * Upload your blank form template image.
            * Use the interactive canvas to draw boxes for text fields and the photo.
            * Assign field names (matching your Excel columns) to these boxes.
            * Download the generated mapping JSON.
        2.  **Process Forms**:
            * Upload your saved mapping JSON.
            * Upload your template image (again, for verification).
            * Upload your Excel file with candidate data.
            * Upload the ZIP file containing candidate photos.
            * Click "Start Processing" and download the final ZIP.
        ---
        **Aiclex Technologies** | [Website](https://aiclex.in) | info@aiclex.in
    """)
    
with tab2:
    st.header("✍️ Template Mapping")
    st.info("Upload your blank form template image and define where each field should be placed.")

    # State variables for mapping
    if 'mapping_data' not in st.session_state:
        st.session_state.mapping_data = {
            "image_size": [0, 0],
            "fields": {}
        }
    if 'template_img_bytes' not in st.session_state:
        st.session_state.template_img_bytes = None
    if 'image_display_width' not in st.session_state:
        st.session_state.image_display_width = 800 # Default display width for canvas

    uploaded_template_file = st.file_uploader(
        "**1. Upload Blank Form Template Image (PNG/JPG)**",
        type=["png", "jpg", "jpeg"],
        key="mapping_template_uploader"
    )

    template_image_pil = None
    if uploaded_template_file is not None:
        st.session_state.template_img_bytes = uploaded_template_file.getvalue()
        template_image_pil = Image.open(BytesIO(st.session_state.template_img_bytes))
        st.session_state.mapping_data["image_size"] = [template_image_pil.width, template_image_pil.height]
        
        st.write(f"Template Image Size: {template_image_pil.width} x {template_image_pil.height} pixels")
        
        st.markdown("---")
        st.subheader("**2. Draw Fields on Template (Click-to-Map)**")
        st.info("Draw rectangles on the image below to define text fields or photo boxes. After drawing, select a field name from the dropdown for each box.")

        # Streamlit Drawable Canvas setup
        canvas_result = st_canvas(
            fill_color="rgba(255, 165, 0, 0.3)",  # Orange color with transparency
            stroke_width=2,
            stroke_color="#EE6002", # Aiclex orange
            background_image=template_image_pil if template_image_pil else None,
            update_streamlit=True,
            height=int(template_image_pil.height * (st.session_state.image_display_width / template_image_pil.width)) if template_image_pil else 400,
            width=st.session_state.image_display_width,
            drawing_mode="rect",
            point_display_radius=0,
            key="canvas_for_mapping"
        )
        
        # Mapping logic
        st.markdown("---")
        st.subheader("**3. Review and Name Mapped Fields**")
        
        # Display existing mappings from session state
        if st.session_state.mapping_data["fields"]:
            st.json(st.session_state.mapping_data["fields"])
            
        # Common fields for suggestion
        common_fields = ["Name of the Candidate", "Address of Candidate", "PIN Code", "Photo", "Date", "Signature", "Designation", "Training Institute"]

        if canvas_result.json_data is not None and len(canvas_result.json_data["objects"]) > 0:
            st.write("Newly drawn shapes:")
            
            # Use columns for better layout of shape naming
            cols = st.columns(3)
            col_idx = 0
            
            for i, obj in enumerate(canvas_result.json_data["objects"]):
                if obj['type'] == 'rect':
                    # Scale coordinates from display size back to original image size
                    scale_factor_w = st.session_state.mapping_data["image_size"][0] / st.session_state.image_display_width
                    scale_factor_h = st.session_state.mapping_data["image_size"][1] / (int(template_image_pil.height * (st.session_state.image_display_width / template_image_pil.width)) if template_image_pil else 400)
                    
                    x = int(obj['left'] * scale_factor_w)
                    y = int(obj['top'] * scale_factor_h)
                    w = int(obj['width'] * obj['scaleX'] * scale_factor_w)
                    h = int(obj['height'] * obj['scaleY'] * scale_factor_h)

                    unique_key = f"field_{i}"
                    
                    with cols[col_idx]:
                        st.markdown(f"**Shape {i+1}:** (x:{x}, y:{y}, w:{w}, h:{h})")
                        field_name = st.selectbox(
                            f"Select Field Name for Shape {i+1}",
                            [""] + common_fields + list(st.session_state.mapping_data["fields"].keys()),
                            key=f"field_name_select_{unique_key}"
                        )
                        custom_field_name = st.text_input(f"Or enter custom name:", key=f"custom_field_name_{unique_key}")
                        
                        final_field_name = custom_field_name if custom_field_name else field_name
                        
                        if st.button(f"Add/Update '{final_field_name}'", key=f"add_update_button_{unique_key}") and final_field_name:
                            st.session_state.mapping_data["fields"][final_field_name] = {"x": x, "y": y, "w": w, "h": h}
                            st.success(f"Field '{final_field_name}' mapped successfully!")
                            st.experimental_rerun() # Rerun to update the displayed mapping_data
                    
                    col_idx = (col_idx + 1) % 3 # Move to next column

        st.markdown("---")
        st.subheader("**4. Download Mapping JSON**")
        if st.session_state.mapping_data["fields"]:
            st.download_button(
                label="Download Mapping JSON",
                data=json.dumps(st.session_state.mapping_data, indent=2),
                file_name=MAPPING_JSON_FILENAME,
                mime="application/json"
            )
            st.success("Mapping ready. You can now proceed to the 'Process Forms' tab.")
        else:
            st.warning("Please draw and name fields on the template to generate the mapping JSON.")
            
with tab3:
    st.header("🔄 Process Forms")
    st.info("Upload your mapping, template, Excel data, and candidate photos to generate filled PDFs.")

    # Upload Mapping JSON
    uploaded_mapping_file = st.file_uploader(
        "**1. Upload Mapping JSON File**",
        type=["json"],
        key="processing_mapping_uploader"
    )
    mapping_data = None
    if uploaded_mapping_file:
        try:
            mapping_data = json.load(uploaded_mapping_file)
            st.success("Mapping JSON loaded successfully!")
            st.json(mapping_data["fields"])
        except Exception as e:
            st.error(f"Error loading mapping JSON: {e}")

    st.markdown("---")

    # Upload Template Image (for verification)
    uploaded_processing_template_file = st.file_uploader(
        "**2. Upload Blank Form Template Image (PNG/JPG) - _Must be the same one used for mapping_**",
        type=["png", "jpg", "jpeg"],
        key="processing_template_uploader"
    )
    template_image_for_processing = None
    if uploaded_processing_template_file:
        template_image_for_processing = Image.open(BytesIO(uploaded_processing_template_file.getvalue()))
        if mapping_data and (template_image_for_processing.width != mapping_data["image_size"][0] or 
                            template_image_for_processing.height != mapping_data["image_size"][1]):
            st.warning(f"Template dimensions mismatch! Mapping was for {mapping_data['image_size'][0]}x{mapping_data['image_size'][1]}, but uploaded image is {template_image_for_processing.width}x{template_image_for_processing.height}. This may lead to misalignment.")
        else:
            st.success("Template image loaded and dimensions match mapping.")
    
    st.markdown("---")

    # Upload Excel File
    uploaded_excel_file = st.file_uploader(
        "**3. Upload Candidate Data Excel File (.xlsx)**",
        type=["xlsx"],
        key="excel_uploader"
    )
    candidate_df = None
    if uploaded_excel_file:
        try:
            candidate_df = get_excel_df(uploaded_excel_file)
            st.success("Excel data loaded successfully!")
            st.dataframe(candidate_df.head())
        except Exception as e:
            st.error(f"Error loading Excel file: {e}")

    st.markdown("---")

    # Upload Photos ZIP
    uploaded_photos_zip = st.file_uploader(
        "**4. Upload Candidate Photos ZIP File**",
        type=["zip"],
        key="photos_zip_uploader"
    )
    zip_path = None
    if uploaded_photos_zip:
        zip_path = os.path.join(TEMP_DIR, uploaded_photos_zip.name)
        with open(zip_path, "wb") as f:
            f.write(uploaded_photos_zip.getbuffer())
        st.success("Photos ZIP uploaded successfully!")

    st.markdown("---")

    # Start Processing Button
    if st.button("🚀 Start Processing", type="primary", use_container_width=True):
        if mapping_data and template_image_for_processing and candidate_df is not None and zip_path:
            with st.spinner("Processing forms... This may take a while depending on the number of candidates."):
                processing_successful = True
                
                # Create a specific temp dir for this run
                run_temp_dir = os.path.join(TEMP_DIR, f"run_{pd.Timestamp.now().strftime('%Y%m%d%H%M%S')}")
                os.makedirs(run_temp_dir, exist_ok=True)
                
                # Unzip photos
                photo_base_dir = os.path.join(run_temp_dir, "unzipped_photos")
                try:
                    photo_folder_paths = unzip_and_organize_files(zip_path, photo_base_dir)
                    st.success(f"Unzipped photos to: {photo_base_dir}")
                except Exception as e:
                    st.error(f"Error unzipping photos: {e}")
                    processing_successful = False

                if processing_successful:
                    # Prepare output directory for this run
                    run_output_dir = os.path.join(OUTPUT_DIR, f"results_{pd.Timestamp.now().strftime('%Y%m%d%H%M%S')}")
                    os.makedirs(run_output_dir, exist_ok=True)

                    processed_count = 0
                    report_data = []

                    # Initialize ImageFormFiller outside the loop if template is constant
                    filler = ImageFormFiller(
                        template_image=template_image_for_processing,
                        mapping_data=mapping_data
                    )

                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    for index, row in candidate_df.iterrows():
                        candidate_srno = None
                        # Auto-detect SrNo column
                        for col in ['SrNo', 'Sl No.', 'SNo', 'Serial']:
                            if col in row.index:
                                candidate_srno = str(row[col]).split('.')[0] # Handle float-like SrNo from excel
                                break
                        
                        if not candidate_srno:
                            st.warning(f"Could not find 'SrNo' column for row {index}. Skipping candidate.")
                            report_data.append({
                                "SrNo": "N/A",
                                "Name": row.get('Name', 'N/A'),
                                "Status": "Skipped (No SrNo)",
                                "Output File": ""
                            })
                            continue

                        candidate_name_for_folder = row.get('Name', f"Candidate_{candidate_srno}") # Fallback name
                        output_candidate_folder = os.path.join(run_output_dir, f"{candidate_srno} {candidate_name_for_folder}")
                        os.makedirs(output_candidate_folder, exist_ok=True)

                        # Find photo for candidate
                        photo_found = False
                        photo_filename = "N/A"
                        candidate_photo_path = None
                        
                        # Search logic based on SrNo and name
                        for root, _, files in os.walk(photo_base_dir):
                            for file in files:
                                if file.lower().endswith(('.jpg', '.jpeg', '.png')):
                                    # Simple matching: check SrNo or name in folder/file
                                    if candidate_srno in root or candidate_srno in file or \
                                       candidate_name_for_folder.lower() in root.lower() or \
                                       candidate_name_for_folder.lower() in file.lower():
                                        candidate_photo_path = os.path.join(root, file)
                                        photo_filename = file
                                        photo_found = True
                                        break
                            if photo_found:
                                break

                        if not photo_found:
                            st.warning(f"No photo found for SrNo: {candidate_srno}, Name: {candidate_name_for_folder}")
                            report_data.append({
                                "SrNo": candidate_srno,
                                "Name": candidate_name_for_folder,
                                "Status": "Processed (No Photo)",
                                "Output File": f"{candidate_srno}_{candidate_name_for_folder}_filled.pdf"
                            })
                            # Still proceed to fill text, just without photo
                            final_pdf_path = filler.fill_and_save_pdf(
                                output_folder=output_candidate_folder,
                                candidate_data=row.to_dict(),
                                candidate_srno=candidate_srno,
                                candidate_name=candidate_name_for_folder,
                                photo_path=None
                            )
                            # Copy the default empty photo or just skip copying
                        else:
                            # Copy original photo to output folder
                            shutil.copy(candidate_photo_path, output_candidate_folder)
                            
                            # Fill form with photo
                            final_pdf_path = filler.fill_and_save_pdf(
                                output_folder=output_candidate_folder,
                                candidate_data=row.to_dict(),
                                candidate_srno=candidate_srno,
                                candidate_name=candidate_name_for_folder,
                                photo_path=candidate_photo_path
                            )
                            st.write(f"Generated PDF for {candidate_srno} {candidate_name_for_folder}")
                            report_data.append({
                                "SrNo": candidate_srno,
                                "Name": candidate_name_for_folder,
                                "Status": "Processed Successfully",
                                "Photo Found": photo_filename,
                                "Output File": os.path.basename(final_pdf_path)
                            })


                        processed_count += 1
                        progress_bar.progress(processed_count / len(candidate_df))
                        status_text.text(f"Processing candidate {processed_count}/{len(candidate_df)}: {candidate_name_for_folder}")

                    st.success(f"Successfully processed {processed_count} out of {len(candidate_df)} candidates!")
                    
                    # Generate report.csv
                    report_df = pd.DataFrame(report_data)
                    report_csv_path = os.path.join(run_output_dir, "processing_report.csv")
                    report_df.to_csv(report_csv_path, index=False)
                    st.download_button(
                        label="Download Processing Report (CSV)",
                        data=report_df.to_csv(index=False).encode('utf-8'),
                        file_name="processing_report.csv",
                        mime="text/csv"
                    )

                    # Create final ZIP
                    final_zip_output_path = os.path.join(OUTPUT_DIR, f"final_filled_results_{pd.Timestamp.now().strftime('%Y%m%d%H%M%S')}.zip")
                    create_output_zip(run_output_dir, final_zip_output_path)
                    
                    st.success("All forms processed and bundled into a final ZIP!")
                    with open(final_zip_output_path, "rb") as fp:
                        st.download_button(
                            label="Download Final Filled Results ZIP",
                            data=fp.read(),
                            file_name=os.path.basename(final_zip_output_path),
                            mime="application/zip",
                            type="primary"
                        )
                
                # Clean up temporary directories
                clean_temp_dirs(run_temp_dir)
                st.info("Temporary files cleaned up.")

        else:
            st.warning("Please upload all required files (Mapping JSON, Template Image, Excel, Photos ZIP) before starting the process.")

# --- Cleanup on script end (optional, but good for local dev) ---
# This part is generally for development; in deployment, use scheduled cleanup
# clean_temp_dirs(TEMP_DIR) # You might not want this to run every time Streamlit refreshes
