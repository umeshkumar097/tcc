import streamlit as st
import pandas as pd
import json
import os
import zipfile
from PIL import Image, ImageDraw, ImageFont
import shutil
from io import BytesIO
import smtplib
from email.message import EmailMessage

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
                value = ""
                field_lower = field.lower()
                if "ted" in field_lower:
                    ted_value = candidate_data.get("ted", "")
                    value = str(pd.to_datetime(ted_value, dayfirst=True).strftime("%d/%m/%Y")) if pd.notna(ted_value) and ted_value != "" else ""
                elif "tsd" in field_lower:
                    tsd_value = candidate_data.get("tsd", "")
                    value = str(pd.to_datetime(tsd_value, dayfirst=True).strftime("%d/%m/%Y")) if pd.notna(tsd_value) and tsd_value != "" else ""
                elif "date of birth" in field_lower or "dob" in field_lower:
                    dob_value = candidate_data.get("date_of_birth", "")
                    value = str(pd.to_datetime(dob_value, dayfirst=True).strftime("%d/%m/%Y")) if pd.notna(dob_value) and dob_value != "" else ""
                elif "address" in field_lower:
                    parts = []
                    for col in ["address_line1","address_line2","city","district","state"]:
                        val = candidate_data.get(col.lower(), "")
                        if pd.notna(val) and str(val).strip() != "":
                            parts.append(str(val).strip())
                    value = ", ".join(parts) if parts else ""
                else:
                    value = str(candidate_data.get(field_lower, ""))

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

# --- Email Sending ---
def send_email_with_zip(to_email, zip_path):
    sender_email = st.secrets["auth"]["email"]
    sender_pass = st.secrets["auth"]["password"]

    if not os.path.exists(zip_path):
        st.error(f"ZIP not found: {zip_path}")
        return False

    msg = EmailMessage()
    msg["Subject"] = "Your Filled Forms"
    msg["From"] = sender_email
    msg["To"] = to_email
    msg.set_content(f"Hello {to_email.split('@')[0]},\n\nPlease find attached your filled forms.\n\nRegards,\nAiclex")

    try:
        with open(zip_path, "rb") as f:
            msg.add_attachment(f.read(), maintype="application", subtype="zip", filename=os.path.basename(zip_path))

        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as smtp:
            smtp.login(sender_email, sender_pass)
            smtp.send_message(msg)

        print(f"✅ Email sent to {to_email}")
        return True

    except Exception as e:
        print("❌ Error:", e)
        st.error(f"Failed to send email: {e}")
        return False

# --- Streamlit App ---
st.set_page_config(page_title="Aiclex Bulk Form Filler", layout="wide")
TEMP_DIR, OUTPUT_DIR = "temp", "output"
FONT_PATH = "assets/DejaVuSans.ttf/DejaVuSans.ttf"
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- Login ---
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if not st.session_state.logged_in:
    st.title("🔐 TCC Form Generator Login")
    pwd = st.text_input("Enter Access Password:", type="password")
    if st.button("Unlock"):
        if pwd == st.secrets["auth"]["Login_password"]:
            st.session_state.logged_in = True
            st.success("✅ Access granted, now click Unlock button open the app")
        else:
            st.error("❌ Incorrect password")
    st.stop()

# --- Main App ---
st.sidebar.markdown('<h1 style="color:#1E3A8A;">Aiclex Technologies</h1>', unsafe_allow_html=True)
st.sidebar.markdown('<h3>Bulk Form Filler</h3>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["🚀 Overview","🔄 Process Forms"])

with tab1:
    st.header("Welcome!")
    st.write("This app uses a fixed template and mapping JSON. Upload Excel and Photos ZIP to generate JPG forms.")

with tab2:
    TEMPLATE_PATH = "assets/template.png"
    MAPPING_PATH = "assets/updated_mapping(75).json"
    template_image = Image.open(TEMPLATE_PATH)
    st.image(template_image, caption="Fixed Template")
    with open(MAPPING_PATH, "r") as f:
        mapping = json.load(f)

    excel_file = st.file_uploader("Upload Candidate Excel File", type=["xlsx"])
    zip_file = st.file_uploader("Upload Candidate Photos ZIP", type=["zip"])

    if "email_zip_dict" not in st.session_state:
        st.session_state.email_zip_dict = {}

    if st.button("🚀 Start Processing"):
        if not all([excel_file, zip_file]):
            st.error("Please upload Excel and Photos ZIP.")
        else:
            with st.spinner("Processing forms..."):
                df = get_excel_df(excel_file)
                zip_path = os.path.join(TEMP_DIR, zip_file.name)
                with open(zip_file_path := zip_path, "wb") as f:
                    f.write(zip_file.getbuffer())
                photo_dir = unzip_and_organize_files(zip_path, os.path.join(TEMP_DIR, "photos"))

                EMAIL_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "by_email")
                os.makedirs(EMAIL_OUTPUT_DIR, exist_ok=True)

                filler = ImageFormFiller(template_image, mapping, FONT_PATH)
                email_zip_dict = {}

                if "email" in df.columns:
                    email_groups = df.groupby("email")
                    for email, group in email_groups:
                        safe_email = email.replace("@", "_at_").replace(".", "_dot_")
                        email_folder = os.path.join(EMAIL_OUTPUT_DIR, safe_email)
                        os.makedirs(email_folder, exist_ok=True)

                        for i, row in group.iterrows():
                            sr_col = next((c for c in ['SrNo','Sl No.','SNo','Serial'] if c in df.columns), None)
                            srno = str(row[sr_col]).split('.')[0] if sr_col else str(i+1)
                            name = row.get("Name", f"Candidate_{srno}")
                            candidate_data = {k.lower().replace(" ", "_"): v for k, v in row.to_dict().items()}

                            # --- OLD PHOTO LOGIC ---
                            photo_path = None
                            for root, dirs, _ in os.walk(photo_dir):
                                for d in dirs:
                                    if d.strip().startswith(str(srno)) and name.lower() in d.lower():
                                        folder_path = os.path.join(root, d)
                                        for fname in os.listdir(folder_path):
                                            if fname.lower().startswith("photo") and fname.lower().endswith((".jpg", ".jpeg", ".png")):
                                                photo_path = os.path.join(folder_path, fname)
                                                break
                                        break
                                if photo_path:
                                    break

                            if not photo_path:
                                print(f"⚠️ Photo not found for {name} (SrNo={srno})")

                            filler.fill_and_save_jpg(email_folder, candidate_data, srno, name, photo_path)

                        email_zip_path = os.path.join(EMAIL_OUTPUT_DIR, f"{safe_email}.zip")
                        create_output_zip(email_folder, email_zip_path)
                        email_zip_dict[email] = email_zip_path
                        st.success(f"✅ ZIP created for {email}")

                st.session_state.email_zip_dict = email_zip_dict
                clean_temp_dirs(TEMP_DIR)
                st.success("✅ Processing complete!")

    if st.session_state.email_zip_dict:
        st.header("📨 Send or Download Emails")

        # --- Download Buttons ---
        st.subheader("⬇️ Download Each ZIP")
        for email, zip_path in st.session_state.email_zip_dict.items():
            with open(zip_path, "rb") as f:
                st.download_button(
                    label=f"Download ZIP for {email}",
                    data=f,
                    file_name=os.path.basename(zip_path),
                    mime="application/zip"
                )

        # --- Send Emails ---
        if st.button("📧 Send All Emails"):
            with st.spinner("📤 Sending all emails... Please wait, this may take a moment."):
                for email, zip_path in st.session_state.email_zip_dict.items():
                    sent = send_email_with_zip(email, zip_path)
                    if sent:
                        st.success(f"✅ Email sent to {email}")
                    else:
                        st.error(f"❌ Failed to send email to {email}")
            st.success("🎉 All emails processed!")
