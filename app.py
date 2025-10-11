import streamlit as st
import pandas as pd
import json
import os
import zipfile
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
import shutil
from reportlab.lib.units import inch
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

# --- PDF/JPG Generation Class ---
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
                test_bbox = draw.textbbox((0, 0), test_line, font=self.pil_font)
                test_width = test_bbox[2] - test_bbox[0]
                if test_width <= w:
                    current_line = test_line
                else:
                    if current_line:
                        lines.append(current_line)
                    current_line = word
            if current_line:
                lines.append(current_line)
            line_height = text_height + 14
            start_y = y + (h - min(len(lines) * line_height, h)) // 2 - 2
            for i, line in enumerate(lines):
                line_bbox = draw.textbbox((0, 0), line, font=self.pil_font)
                draw_x_line = x
                draw_y_line = start_y + i * line_height
                if draw_y_line + text_height <= y + h:
                    draw.text((draw_x_line, draw_y_line), line, font=self.pil_font, fill=(0,0,0))
        else:
            draw_x = x
            draw.text((draw_x, draw_y), text, font=self.pil_font, fill=(0,0,0))

    def fill_and_save_jpg(self, output_folder: str, candidate_data: dict, srno: str, name: str, photo_path: str = None):
        filled_image = self.template_image.copy()
        draw = ImageDraw.Draw(filled_image)
        for field, coords in self.mapping_data["fields"].items():
            x, y, w, h = coords.get("x",0), coords.get("y",0), coords.get("w",200), coords.get("h",50)
            if "photo" in field.lower() and photo_path:
                try:
                    with Image.open(photo_path) as photo:
                        photo = photo.resize((w, h), Image.Resampling.LANCZOS)
                        filled_image.paste(photo, (x, y))
                except Exception as e:
                    print(f"Error with photo for {name}: {e}")
            else:
                # Handle special fields
                value = ""
                field_lower = field.lower()
                if "ted" in field_lower:
                    ted_value = candidate_data.get("ted", "")
                    value = str(pd.to_datetime(ted_value).strftime("%d/%m/%Y")) if pd.notna(ted_value) and ted_value != "" else ""
                elif "tsd" in field_lower:
                    tsd_value = candidate_data.get("tsd", "")
                    value = str(pd.to_datetime(tsd_value).strftime("%d/%m/%Y")) if pd.notna(tsd_value) and tsd_value != "" else ""
                elif "date of birth" in field_lower or "dob" in field_lower:
                    dob_value = candidate_data.get("date_of_birth", "")
                    value = str(pd.to_datetime(dob_value).strftime("%d/%m/%Y")) if pd.notna(dob_value) and dob_value != "" else ""
                elif "address" in field_lower:
                    parts = []
                    for col in ["address_line1","address_line2","city","district","state"]:
                        val = candidate_data.get(col.lower(), "")
                        if pd.notna(val) and str(val).strip() != "":
                            parts.append(str(val).strip())
                    value = ", ".join(parts) if parts else ""
                else:
                    value = str(candidate_data.get(field_lower, ""))

                # Adjust font sizes for special fields
                original_font = self.pil_font
                if field_lower in ["name"]:
                    self.pil_font = ImageFont.truetype(self.font_path, 30)
                elif field_lower in ["ted","tsd","date of birth","dob","qualification"]:
                    self.pil_font = ImageFont.truetype(self.font_path, 28)
                else:
                    self.pil_font = ImageFont.truetype(self.font_path, self.font_size)
                if value:
                    self._draw_text_on_image(draw, value, x, y, w, h)
                self.pil_font = original_font

        output_path = os.path.join(output_folder, f"{srno}_{name.replace(' ','_')}.jpg")
        filled_image.save(output_path, "JPEG", quality=95)

# --- Streamlit App ---
st.set_page_config(page_title="Aiclex Bulk Form Filler", layout="wide")
TEMP_DIR, OUTPUT_DIR = "temp", "output"
FONT_PATH = "assets/DejaVuSans.ttf/DejaVuSans.ttf"
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

st.sidebar.markdown('<h1 style="color:#1E3A8A;">Aiclex Technologies</h1>', unsafe_allow_html=True)
st.sidebar.markdown('<h3>Bulk Form Filler</h3>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["🚀 Overview","🔄 Process Forms"])

# --- Tab 1: Overview ---
with tab1:
    st.header("Welcome!")
    st.write("This app uses a fixed template and mapping JSON. Upload Excel and Photos ZIP to generate JPG forms.")

# --- Tab 2: Process Forms ---
with tab2:
    st.header("🔄 Process Forms")

    # --- Fixed template & mapping ---
    TEMPLATE_PATH = "assets/1A UMESH.jpg"
    MAPPING_PATH = "assets/updated_mapping.json"
    template_image = Image.open(TEMPLATE_PATH)
    st.image(template_image, caption="Fixed Template (Cannot be changed)")

    with open(MAPPING_PATH,"r") as f:
        mapping = json.load(f)

    excel_file = st.file_uploader("Upload Candidate Excel File", type=["xlsx"])
    zip_file = st.file_uploader("Upload Candidate Photos ZIP", type=["zip"])

    if st.button("🚀 Start Processing"):
        if all([excel_file, zip_file]):
            with st.spinner("Processing..."):
                df = get_excel_df(excel_file)
                zip_path = os.path.join(TEMP_DIR, zip_file.name)
                with open(zip_path,"wb") as f:
                    f.write(zip_file.getbuffer())
                photo_dir = unzip_and_organize_files(zip_path, os.path.join(TEMP_DIR,"photos"))

                output_run_dir = os.path.join(OUTPUT_DIR,"run")
                if os.path.exists(output_run_dir):
                    shutil.rmtree(output_run_dir, onerror=lambda func, path, exc: os.chmod(path, stat.S_IWRITE) or func(path))
                os.makedirs(output_run_dir, exist_ok=True)

                filler = ImageFormFiller(template_image, mapping, FONT_PATH)
                progress_bar = st.progress(0)
                total_rows = len(df)

                for i, row in df.iterrows():
                    sr_col = next((c for c in ['SrNo','Sl No.','SNo','Serial'] if c in df.columns), None)
                    if not sr_col:
                        st.error("Excel missing SrNo/Sl No./SNo column")
                        break
                    srno = str(row[sr_col]).split('.')[0]
                    name = row.get('Name', f"Candidate_{srno}")
                    candidate_data = {k.lower().replace(" ","_"): v for k,v in row.to_dict().items()}

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

                final_zip = os.path.join(OUTPUT_DIR,"final_results_jpg.zip")
                create_output_zip(output_run_dir, final_zip)
                with open(final_zip,"rb") as fp:
                    st.download_button("✅ Download Final ZIP of JPGs", fp, "final_results_jpg.zip","application/zip")
                clean_temp_dirs(TEMP_DIR)
        else:
            st.error("Please upload Excel and Photos ZIP to start processing.")
