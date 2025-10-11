import streamlit as st
import pandas as pd
import json
import os
import zipfile
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
import shutil
import stat

# --- Helper Functions ---
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
        except Exception as e:
            print(f"Error cleaning directory: {e}")

def get_excel_df(excel_file_buffer) -> pd.DataFrame:
    return pd.read_excel(excel_file_buffer)

# --- ImageFormFiller Class ---
class ImageFormFiller:
    def __init__(self, template_image: Image.Image, mapping_data: dict, font_path: str, font_size: int = 24):
        self.template_image = template_image.convert('RGB')
        self.mapping_data = mapping_data
        self.font_path = font_path
        self.font_size = font_size
        try:
            self.pil_font = ImageFont.truetype(self.font_path, self.font_size)
        except IOError:
            self.pil_font = ImageFont.load_default()

    def _draw_text_on_image(self, draw: ImageDraw.Draw, text: str, x: int, y: int, w: int, h: int):
        # Wrap text to fit inside box
        text_bbox = draw.textbbox((0, 0), text, font=self.pil_font)
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]
        draw_y = y
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
                    if current_line:
                        lines.append(current_line)
                    current_line = word
            if current_line:
                lines.append(current_line)
            line_height = text_height + 5
            start_y = y + (h - min(len(lines) * line_height, h)) // 2
            for i, line in enumerate(lines):
                draw.text((x, start_y + i * line_height), line, font=self.pil_font, fill=(0,0,0))
        else:
            draw.text((x, draw_y), text, font=self.pil_font, fill=(0,0,0))

    def fill_and_save_jpg(self, output_folder: str, candidate_data: dict, srno: str, name: str, photo_path: str = None):
        filled_image = self.template_image.copy()
        draw = ImageDraw.Draw(filled_image)
        for field, coords in self.mapping_data["fields"].items():
            x, y, w, h = coords["x"], coords["y"], coords["w"], coords["h"]

            # --- Photo handling ---
            if "photo" in field.lower() and photo_path:
                try:
                    with Image.open(photo_path) as photo:
                        photo = photo.resize((w, h), Image.Resampling.LANCZOS)
                        filled_image.paste(photo, (x, y))
                except Exception as e:
                    print(f"Error with photo for {name}: {e}")
                continue

            # --- Field value handling ---
            field_lower = field.lower()
            value = ""
            try:
                if "ted" in field_lower:
                    value = str(pd.to_datetime(candidate_data.get("ted_copy", candidate_data.get("ted",""))).strftime("%d/%m/%Y")) \
                        if pd.notna(candidate_data.get("ted_copy", candidate_data.get("ted",""))) else ""
                elif "tsd" in field_lower:
                    value = str(pd.to_datetime(candidate_data.get("tsd","")).strftime("%d/%m/%Y")) \
                        if pd.notna(candidate_data.get("tsd","")) else ""
                elif "date of birth" in field_lower or "dob" in field_lower:
                    value = str(pd.to_datetime(candidate_data.get("date_of_birth","")).strftime("%d/%m/%Y")) \
                        if pd.notna(candidate_data.get("date_of_birth","")) else ""
                elif "address" in field_lower:
                    parts = []
                    for col in ["address_line1","address_line2","city","district","state"]:
                        val = candidate_data.get(col.lower(),"")
                        if pd.notna(val) and str(val).strip() != "":
                            parts.append(str(val).strip())
                    value = ", ".join(parts)
                else:
                    value = str(candidate_data.get(field_lower.replace(" ","_"),""))
            except Exception as e:
                print(f"Error processing field {field} for {name}: {e}")
                value = ""

            if value:
                # --- Font size handling ---
                original_font = self.pil_font
                if field_lower in ["name"]:
                    self.pil_font = ImageFont.truetype(self.font_path, 30)
                elif field_lower in ["ted","tsd","dob","date of birth","qualification"]:
                    self.pil_font = ImageFont.truetype(self.font_path, 28)
                else:
                    self.pil_font = ImageFont.truetype(self.font_path, self.font_size)

                self._draw_text_on_image(draw, value, x, y, w, h)
                self.pil_font = original_font

        jpg_path = os.path.join(output_folder, f"{srno}_{name.replace(' ','_')}.jpg")
        filled_image.save(jpg_path, "JPEG", quality=95)

# --- Streamlit App UI ---
st.set_page_config(page_title="Aiclex Bulk Form Filler", layout="wide")
TEMP_DIR, OUTPUT_DIR = "temp", "output"
FONT_PATH = "assets/DejaVuSans.ttf/DejaVuSans.ttf"
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

st.sidebar.markdown('<h1 style="color:#1E3A8A;">Aiclex Technologies</h1>', unsafe_allow_html=True)
st.sidebar.markdown('<h3>Bulk Form Filler</h3>', unsafe_allow_html=True)
tab1, tab2, tab3 = st.tabs(["🚀 Overview", "✍️ Template Mapping (Manual)", "🔄 Process Forms"])

# --- Tab 1: Overview ---
with tab1:
    st.header("Welcome!")
    st.write("Use the 'Template Mapping' tab to create your mapping file, then use 'Process Forms' to generate documents.")

# --- Tab 2: Template Mapping ---
with tab2:
    st.header("✍️ Template Mapping (Manual Mode)")
    st.info("Enter coordinates from an image editor like MS Paint.")
    if 'mapping_data' not in st.session_state:
        st.session_state.mapping_data = {"image_size": [0,0], "fields":{}}
    uploaded_template_file = st.file_uploader("**1. Upload Blank Form Image**", type=["png","jpg","jpeg"])
    if uploaded_template_file:
        template_image_pil = Image.open(uploaded_template_file)
        w, h = template_image_pil.size
        st.session_state.mapping_data["image_size"] = [w,h]
        st.image(template_image_pil, caption="Your Template")
        st.success(f"**Image Dimensions:** Width={w}px, Height={h}px")
        st.warning("Use MS Paint or another image editor to find the pixel coordinates for each field.")
    st.subheader("2. Add Fields Manually")
    with st.form("mapping_form"):
        cols = st.columns([2,1,1,1,1])
        field_name = cols[0].text_input("Field Name (e.g., Name, Photo, Address, TED)")
        x_coord = cols[1].number_input("X (from left)", min_value=0, step=10)
        y_coord = cols[2].number_input("Y (from top)", min_value=0, step=10)
        width = cols[3].number_input("Width", min_value=10, step=10, value=300)
        height = cols[4].number_input("Height", min_value=10, step=10, value=50)
        submitted = st.form_submit_button("Add Field")
        if submitted and field_name:
            st.session_state.mapping_data["fields"][field_name] = {"x":x_coord,"y":y_coord,"w":width,"h":height}
            st.success(f"Added field '{field_name}'")
    st.subheader("3. Review and Download Mapping JSON")
    if st.session_state.mapping_data["fields"]:
        st.json(st.session_state.mapping_data["fields"])
        st.download_button(
            label="Download Mapping JSON File",
            data=json.dumps(st.session_state.mapping_data, indent=2),
            file_name="mapping.json",
            mime="application/json"
        )

# --- Tab 3: Process Forms ---
with tab3:
    st.header("🔄 Process Forms")
    uploaded_template_for_processing = st.file_uploader("1. Upload Blank Form Image", type=["png","jpg"])
    mapping_file = st.file_uploader("2. Upload Your Saved Mapping JSON", type=["json"])
    excel_file = st.file_uploader("3. Upload Candidate Excel File", type=["xlsx"])
    zip_file = st.file_uploader("4. Upload Candidate Photos ZIP", type=["zip"])

    if st.button("🚀 Start Processing"):
        if all([uploaded_template_for_processing, mapping_file, excel_file, zip_file]):
            with st.spinner("Processing..."):
                template_image_process = Image.open(uploaded_template_for_processing)
                mapping = json.load(mapping_file)
                df = get_excel_df(excel_file)

                zip_path = os.path.join(TEMP_DIR, zip_file.name)
                with open(zip_path,"wb") as f:
                    f.write(zip_file.getbuffer())
                photo_dir = unzip_and_organize_files(zip_path, os.path.join(TEMP_DIR,"photos"))

                # Output folder
                output_run_dir = os.path.join(OUTPUT_DIR,"run")
                if os.path.exists(output_run_dir):
                    shutil.rmtree(output_run_dir, onerror=lambda func, path, exc: os.chmod(path, stat.S_IWRITE) or func(path))
                os.makedirs(output_run_dir, exist_ok=True)

                filler = ImageFormFiller(template_image_process, mapping, FONT_PATH)
                progress_bar = st.progress(0)
                total_rows = len(df)

                for i, row in df.iterrows():
                    sr_col = next((c for c in ['SrNo','Sl No.','SNo','Serial'] if c in df.columns), None)
                    if not sr_col:
                        st.error("Error: 'SrNo' or 'Sl No.' column not found in Excel.")
                        break
                    srno = str(row[sr_col]).split('.')[0]
                    name = row.get('Name', f"Candidate_{srno}")
                    candidate_data = {k.lower().replace(" ","_"): v for k,v in row.to_dict().items()}

                    # Find candidate photo
                    photo_path = None
                    for root, dirs, _ in os.walk(photo_dir):
                        for d in dirs:
                            if d.strip().startswith(str(srno)) and name.lower() in d.lower():
                                folder_path = os.path.join(root,d)
                                for fname in os.listdir(folder_path):
                                    if fname.lower().startswith("photo") and fname.lower().endswith((".jpg",".jpeg",".png")):
                                        photo_path = os.path.join(folder_path,fname)
                                        break
                                break
                        if photo_path:
                            break

                    filler.fill_and_save_jpg(output_run_dir, candidate_data, srno, name, photo_path)
                    progress_bar.progress((i+1)/total_rows)

                # ZIP JPGs
                final_zip = os.path.join(OUTPUT_DIR,"final_results_jpg.zip")
                create_output_zip(output_run_dir, final_zip)
                with open(final_zip,"rb") as fp:
                    st.download_button("✅ Download Final ZIP of JPGs", fp, "final_results_jpg.zip", "application/zip")
                clean_temp_dirs(TEMP_DIR)
        else:
            st.error("Please upload all four files to start processing.")
