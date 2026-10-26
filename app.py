import streamlit as st
import pandas as pd
import json
import os
import zipfile
from PIL import Image
import shutil
from io import BytesIO
import smtplib
from email.message import EmailMessage
import re
import fitz # PyMuPDF

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

def format_date(val):
    if not val or pd.isna(val) or str(val).strip() == "" or str(val).lower() in ["nan", "none"]:
        return ""
    try:
        dt = pd.to_datetime(val, dayfirst=True)
        return f"{dt.day}-{dt.month}-{dt.year}"
    except:
        return str(val).strip()

def normalize_key(k) -> str:
    return re.sub(r'[^a-z0-9]', '', str(k).lower())

def extract_field(row_dict: dict, primary_keys: list, fallback_keywords: list = None, default: str = "") -> str:
    if not isinstance(row_dict, dict):
        return default
    norm_dict = {normalize_key(k): v for k, v in row_dict.items()}
    # 1. Exact normalized match (handles spaces, casing, punctuation, colons, brackets)
    for pk in primary_keys:
        npk = normalize_key(pk)
        if npk in norm_dict and pd.notna(norm_dict[npk]):
            v = str(norm_dict[npk]).strip()
            if v != "" and v.lower() not in ["nan", "none"]:
                return v
    # 2. Fallback keyword match
    if fallback_keywords:
        for orig_k, v in row_dict.items():
            k_lower = str(orig_k).lower()
            for kw in fallback_keywords:
                if kw in k_lower and pd.notna(v):
                    val = str(v).strip()
                    if val != "" and val.lower() not in ["nan", "none"]:
                        return val
    return default

# Alias for backward compatibility
get_val = extract_field

# --- DRACertificateFiller Class ---
class DRACertificateFiller:
    def __init__(self, template_pdf_path: str, font_reg_path: str, font_bold_path: str):
        self.template_pdf_path = template_pdf_path
        self.font_reg_path = font_reg_path
        self.font_bold_path = font_bold_path

    def fill_and_save_jpg(self, output_folder: str, candidate_data: dict, srno: str, name: str, photo_path: str = None):
        doc = fitz.open(self.template_pdf_path)
        page = doc[0]

        # Redact the template placeholder paragraph before inserting candidate-specific text
        page.add_redact_annot(fitz.Rect(52.0, 420.0, 520.0, 485.0), fill=(1, 1, 1))
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        # Register Verdana fonts
        if os.path.exists(self.font_reg_path) and os.path.exists(self.font_bold_path):
            page.insert_font(fontname="Verdana", fontfile=self.font_reg_path)
            page.insert_font(fontname="Verdana-Bold", fontfile=self.font_bold_path)
            f_reg = fitz.Font(fontfile=self.font_reg_path)
            f_bold = fitz.Font(fontfile=self.font_bold_path)
            fn_reg = "Verdana"
            fn_bold = "Verdana-Bold"
        else:
            fn_reg = "helv"
            fn_bold = "hebo"
            f_reg = fitz.Font("helv")
            f_bold = fitz.Font("hebo")

        # 1. Candidate's Name (Bold, 12pt) aligned in same straight vertical line at x = 54.02
        c_name = extract_field(candidate_data, ['name', 'candidate_name', 'candidates_name', 'candidate_name_mr_ms_mrs'], ['name'], default=name)
        disp_name = c_name if c_name else name
        if disp_name:
            page.insert_text(fitz.Point(54.02, 240.0), disp_name, fontname=fn_bold, fontsize=12, color=(0, 0, 0))

        # 2. Date of Birth (Bold, 10pt) aligned with 'Date of Birth: ' label at x = 54.02
        raw_dob = extract_field(candidate_data, ['date_of_birth', 'dob', 'birth_date'], ['birth', 'dob'])
        dob_str = format_date(raw_dob)
        if dob_str:
            dob_label_w = f_bold.text_length('Date of Birth: ', fontsize=10)
            page.insert_text(fitz.Point(54.02 + dob_label_w, 260.5), dob_str, fontname=fn_bold, fontsize=10, color=(0, 0, 0))

        # 3. Training ID (Bold, 10pt) - baseline aligned with heading (279.2)
        trg_id = extract_field(candidate_data, ['training_id', 'trainingid', 'trg_id', 'batch_id'], ['training_id', 'trainingid', 'trg_id', 'batch'])
        if trg_id:
            page.insert_text(fitz.Point(125.0, 279.2), trg_id, fontname=fn_bold, fontsize=10, color=(0, 0, 0))

        # 4. Training Hours (Bold, 10pt) - baseline aligned with heading (299.1)
        hours = extract_field(candidate_data, ['training_hours_(50_/_100)', 'training_hours_(50/100)', 'training_hours', 'hours', 'trg_hours'], ['hour'])
        if hours:
            page.insert_text(fitz.Point(208.5, 299.1), hours, fontname=fn_bold, fontsize=10, color=(0, 0, 0))

        # 5. Training Medium (Bold, 10pt) - baseline aligned with heading (319.1)
        medium = extract_field(candidate_data, ['training_medium_(language):', 'training_medium_(language)', 'training_medium', 'language', 'medium'], ['medium', 'language'])
        if medium:
            page.insert_text(fitz.Point(225.0, 319.1), medium, fontname=fn_bold, fontsize=10, color=(0, 0, 0))

        # 6. Training Mode (Bold, 10pt) - on SAME line as heading, baseline aligned (339.2)
        mode = extract_field(candidate_data, ['training_mode_(online/offline):', 'training_mode_(online/offline)', 'training_mode', 'mode'], ['mode'])
        if mode:
            page.insert_text(fitz.Point(237.0, 339.2), mode, fontname=fn_bold, fontsize=10, color=(0, 0, 0))

        # 7. Training Center Location (Bold, 10pt) - moved up to baseline 359.2 directly after heading (minimized space)
        loc = extract_field(candidate_data, ['training_center_location_(in_case_of_offline_batch):', 'training_center_location_(in_case_of_offline_batch)', 'training_center_location', 'location', 'center_location'], ['location', 'center'], default="NA")
        page.insert_text(fitz.Point(342.0, 359.2), loc, fontname=fn_bold, fontsize=10, color=(0, 0, 0))

        # 8. TSD and TED in acknowledgement paragraph
        raw_tsd = extract_field(candidate_data, ['tsd', 'training_start_date', 'start_date'], ['tsd', 'start'])
        raw_ted = extract_field(candidate_data, ['ted', 'training_end_date', 'end_date'], ['ted', 'end'])
        tsd_str = format_date(raw_tsd)
        ted_str = format_date(raw_ted)

        # Line 1:
        page.insert_text(fitz.Point(54.02, 435.91), "This is to acknowledge that the above-mentioned candidate has completed", fontname=fn_reg, fontsize=11, color=(0, 0, 0))
        # Line 2:
        prefix = "Debt Recovery Agents (DRA) training during the period from "
        page.insert_text(fitz.Point(54.02, 455.95), prefix, fontname=fn_reg, fontsize=11, color=(0, 0, 0))
        p_len = f_reg.text_length(prefix, fontsize=11)
        page.insert_text(fitz.Point(54.02 + p_len, 455.95), tsd_str, fontname=fn_bold, fontsize=11, color=(0, 0, 0))
        t_len = f_bold.text_length(tsd_str, fontsize=11)
        page.insert_text(fitz.Point(54.02 + p_len + t_len, 455.95), " to", fontname=fn_reg, fontsize=11, color=(0, 0, 0))
        # Line 3:
        page.insert_text(fitz.Point(54.02, 475.99), ted_str, fontname=fn_bold, fontsize=11, color=(0, 0, 0))
        ted_len = f_bold.text_length(ted_str, fontsize=11)
        page.insert_text(fitz.Point(54.02 + ted_len, 475.99), " and is eligible to appear for DRA examination.", fontname=fn_reg, fontsize=11, color=(0, 0, 0))

        # 9. Date at bottom left: matches TED (baseline 570.82)
        page.insert_text(fitz.Point(91.5, 570.82), ted_str, fontname=fn_bold, fontsize=11, color=(0, 0, 0))

        # Render page to PIL image
        pix = page.get_pixmap(dpi=300, alpha=False)
        filled_image = Image.open(BytesIO(pix.tobytes("jpeg")))
        doc.close()

        # Photo box coordinates on 300 DPI canvas
        x, y, w, h = 1826, 860, 435, 537

        # Apply photo using PIL (same as before)
        if photo_path:
            try:
                with Image.open(photo_path) as photo:
                    photo = photo.resize((w, h), Image.Resampling.LANCZOS)
                    filled_image.paste(photo, (x, y))
            except Exception as e:
                print(f"Error with photo for {name}: {e}")

        # Save as 300 DPI high resolution JPEG
        clean_name = name.replace(" ", "_").replace("/", "_").replace("\\", "_")
        output_path = os.path.join(output_folder, f"{srno}_{clean_name}.jpg")
        filled_image.save(output_path, "JPEG", quality=95)
        return output_path

# Alias for backward compatibility
ImageFormFiller = DRACertificateFiller

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
st.set_page_config(page_title="DRA Certificate Form Filler", layout="wide")
TEMP_DIR, OUTPUT_DIR = "temp", "output"
TEMPLATE_PDF = "assets/base_template.pdf"
TEMPLATE_PREVIEW = "assets/template.png"
FONT_REG = "assets/verdana.ttf"
FONT_BOLD = "assets/verdanab.ttf"

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- Login ---
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.title("🔐 DRA Certificate Generator Login")
    pwd = st.text_input("Enter Access Password:", type="password")
    if st.button("Unlock"):
        correct_pwd = st.secrets["auth"].get("Login_password", st.secrets["auth"].get("login_password", "admin"))
        if pwd == correct_pwd:
            st.session_state.logged_in = True
            st.success("✅ Access granted!")
            st.rerun()
        else:
            st.error("❌ Incorrect password")
    st.stop()

# --- Main App ---
st.sidebar.markdown('<h1 style="color:#1E3A8A;">Crux / Aiclex</h1>', unsafe_allow_html=True)
st.sidebar.markdown('<h3>DRA Certificate Generator</h3>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["🚀 Overview", "🔄 Process Forms"])

with tab1:
    st.header("DRA Training Completion Certificate Generator")
    st.write("This app generates **DRA Training Completion Certificates** from candidate Excel records and photo ZIP archives.")
    st.markdown("""
    ### Expected Excel Columns:
    - **`Sl No.`** (or `SrNo`)
    - **`Name`**
    - **`Date Of Birth`**
    - **`Training ID`**
    - **`Training Hours (50 / 100)`**
    - **`Training Medium (Language):`**
    - **`Training Mode (Online/Offline):`**
    - **`Training Center Location (in case of offline batch):`**
    - **`TSD`** (Training Start Date)
    - **`TED`** (Training End Date)
    - **`email`** (Used to group candidates into separate ZIP files)
    """)

with tab2:
    if os.path.exists(TEMPLATE_PREVIEW):
        template_image = Image.open(TEMPLATE_PREVIEW)
        st.image(template_image, caption="DRA Training Completion Certificate Template", width=550)

    excel_file = st.file_uploader("Upload Candidate Excel File", type=["xlsx"])
    zip_file = st.file_uploader("Upload Candidate Photos ZIP", type=["zip"])

    if "email_zip_dict" not in st.session_state:
        st.session_state.email_zip_dict = {}

    if st.button("🚀 Start Processing"):
        if not all([excel_file, zip_file]):
            st.error("Please upload both Excel file and Photos ZIP.")
        else:
            with st.spinner("Processing candidate certificates..."):
                df = get_excel_df(excel_file)
                zip_path = os.path.join(TEMP_DIR, zip_file.name)
                with open(zip_file_path := zip_path, "wb") as f:
                    f.write(zip_file.getbuffer())
                photo_dir = unzip_and_organize_files(zip_path, os.path.join(TEMP_DIR, "photos"))

                EMAIL_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "by_email")

                # Clean old output before regenerating
                if os.path.exists(EMAIL_OUTPUT_DIR):
                    shutil.rmtree(EMAIL_OUTPUT_DIR)
                os.makedirs(EMAIL_OUTPUT_DIR, exist_ok=True)

                filler = DRACertificateFiller(TEMPLATE_PDF, FONT_REG, FONT_BOLD)
                email_zip_dict = {}

                # Find email column flexibly
                email_col = next((c for c in df.columns if str(c).strip().lower() in ['email', 'email_id', 'email id', 'email address']), None)

                if email_col:
                    email_groups = df.groupby(email_col)
                    for email_val, group in email_groups:
                        email = str(email_val).strip()
                        safe_email = email.replace("@", "_at_").replace(".", "_dot_")
                        email_folder = os.path.join(EMAIL_OUTPUT_DIR, safe_email)
                        os.makedirs(email_folder, exist_ok=True)

                        for i, row in group.iterrows():
                            candidate_data = row.to_dict()
                            sr_val = extract_field(candidate_data, ['sl_no.', 'sl_no', 'srno', 'sr_no', 'serial'], ['sl', 'sr', 'serial'])
                            srno = str(sr_val).split('.')[0] if sr_val else str(i+1)
                            name = extract_field(candidate_data, ['name', 'candidate_name', 'candidates_name'], ['name'], default=f"Candidate_{srno}")

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
                    st.success("🎉 All certificates processed successfully!")
                else:
                    st.error("❌ 'email' column not found in uploaded Excel file. Please ensure there is an 'email' column.")

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
