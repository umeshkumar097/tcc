import streamlit as st
import pandas as pd
import json
import os
import zipfile
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
import shutil
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader

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
        try: shutil.rmtree(directory)
        except Exception as e: print(f"Error cleaning directory: {e}")

def get_excel_df(excel_file_buffer) -> pd.DataFrame:
    try:
        df = pd.read_csv(excel_file_buffer)
    except Exception:
        excel_file_buffer.seek(0)
        df = pd.read_excel(excel_file_buffer)
    return df

# --- PDF Generation Class ---
class ImageFormFiller:
    def __init__(self, template_image: Image.Image, mapping_data: dict, font_path: str, font_size: int = 24):
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
        import textwrap
        avg_char_width = sum(self.pil_font.getlength(char) for char in 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789') / 62
        max_chars_per_line = int(w / avg_char_width) if avg_char_width > 0 else 1
        wrapped_text = textwrap.fill(text, width=max_chars_per_line)
        draw.text((x, y), wrapped_text, font=self.pil_font, fill=(0, 0, 0))

    def fill_and_save_pdf(self, output_folder: str, candidate_data: dict, srno: str, name: str, photo_path: str = None):
        filled_image = self.template_image.copy()
        draw = ImageDraw.Draw(filled_image)
        
        address_parts = []
        if 'Address' in candidate_data and pd.notna(candidate_data['Address']):
            address_parts.append(str(candidate_data['Address']))
        else:
            if 'Address_Line1' in candidate_data and pd.notna(candidate_data['Address_Line1']):
                address_parts.append(str(candidate_data['Address_Line1']))
            if 'Address_Line2' in candidate_data and pd.notna(candidate_data['Address_Line2']):
                address_parts.append(str(candidate_data['Address_Line2']))
        full_address = ", ".join(address_parts)
        candidate_data['full_address_combined'] = full_address

        for field, coords in self.mapping_data["fields"].items():
            x, y, w, h = coords["x"], coords["y"], coords["w"], coords["h"]
            value = None
            if field.lower() == 'address':
                value = candidate_data.get('full_address_combined')
            elif field.lower() == "photo" and photo_path:
                try:
                    with Image.open(photo_path) as photo:
                        photo = photo.resize((w, h), Image.Resampling.LANCZOS)
                        filled_image.paste(photo, (x, y))
                except Exception as e: print(f"Error with photo for {name}: {e}")
            else:
                value = next((str(v) for k, v in candidate_data.items() if k.lower().replace("_", " ") == field.lower()), None)
            
            if value:
                self._draw_text_on_image(draw, value, x, y, w, h)
        
        pdf_path = os.path.join(output_folder, f"{srno}_{name.replace(' ', '_')}.pdf")
        with BytesIO() as img_buffer:
            filled_image.save(img_buffer, format='PNG', dpi=(self.dpi, self.dpi))
            img_buffer.seek(0)
            c = canvas.Canvas(pdf_path, pagesize=(self.page_width_pts, self.page_height_pts))
            c.drawImage(ImageReader(img_buffer), 0, 0, width=self.page_width_pts, height=self.page_height_pts)
            c.save()

# --- Your Hardcoded Mapping Data ---
MAPPING_DATA = {
  "image_size": [
    2480,
    3508
  ],
  "fields": {
    "Photo": {
      "x": 1941,
      "y": 483,
      "w": 381,
      "h": 415
    },
    "Name": {
      "x": 757,
      "y": 1205,
      "w": 1056,
      "h": 50
    },
    "Address_Line1": {
      "x": 754,
      "y": 1279,
      "w": 1060,
      "h": 48
    },
    "Address_Line2": {
      "x": 754,
      "y": 1332,
      "w": 1061,
      "h": 49
    },
    "PIN Code": {
      "x": 2045,
      "y": 1279,
      "w": 161,
      "h": 49
    },
    "City": {
      "x": 757,
      "y": 1391,
      "w": 386,
      "h": 45
    },
    "District": {
      "x": 1222,
      "y": 1393,
      "w": 372,
      "h": 44
    },
    "State": {
      "x": 1699,
      "y": 1394,
      "w": 303,
      "h": 45
    },
    "TSD": {
      "x": 1133,
      "y": 1928,
      "w": 235,
      "h": 41
    },
    "TED": {
      "x": 1496,
      "y": 1926,
      "w": 236,
      "h": 41
    },
    "Date": {
      "x": 427,
      "y": 2085,
      "w": 328,
      "h": 40
    },
    "Qualification": {
      "x": 1018,
      "y": 2380,
      "w": 794,
      "h": 41
    },
    "Date Of Birth": {
      "x": 1018,
      "y": 2445,
      "w": 794,
      "h": 40
    }
  }
}

# --- Streamlit App UI ---
st.set_page_config(page_title="Aiclex Bulk Form Filler", layout="wide")
TEMP_DIR, OUTPUT_DIR = "temp", "output"
FONT_PATH = "assets/DejaVuSans.ttf"
TEMPLATE_PATH = "assets/template.png"
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

st.sidebar.markdown('<h1 style="color:#1E3A8A;">Aiclex Technologies</h1>', unsafe_allow_html=True)
st.sidebar.markdown('<h3>Bulk Form Filler</h3>', unsafe_allow_html=True)

st.title("🚀 Aiclex Bulk Form Filler")
st.info("Your form template and mapping are now fixed. Simply upload your Excel file and Photos ZIP to begin.")

try:
    template_image_process = Image.open(TEMPLATE_PATH)
    
    excel_file = st.file_uploader("1. Upload Candidate Excel File", type=["xlsx", "csv"])
    zip_file = st.file_uploader("2. Upload Candidate Photos ZIP", type=["zip"])

    if st.button("🚀 Start Processing", type="primary"):
        if all([excel_file, zip_file]):
            with st.spinner("Processing... This may take a moment for large files."):
                df = get_excel_df(excel_file)
                
                zip_path = os.path.join(TEMP_DIR, zip_file.name)
                with open(zip_path, "wb") as f: f.write(zip_file.getbuffer())
                
                photo_dir = unzip_and_organize_files(zip_path, os.path.join(TEMP_DIR, "photos"))
                output_run_dir = os.path.join(OUTPUT_DIR, "run")
                if os.path.exists(output_run_dir): shutil.rmtree(output_run_dir)
                os.makedirs(output_run_dir)

                filler = ImageFormFiller(template_image_process, MAPPING_DATA, FONT_PATH)
                
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
            st.error("Please upload both the Excel file and the Photos ZIP.")

except FileNotFoundError:
    st.error(f"FATAL ERROR: Your template file could not be found.")
    st.error(f"Please make sure a file named 'template.png' exists in the 'assets' folder of your GitHub repository.")
