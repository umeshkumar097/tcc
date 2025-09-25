# app.py
import streamlit as st
from pathlib import Path
import pandas as pd
import tempfile, shutil, zipfile, io, os, re, json
from PIL import Image, ImageOps, ImageFont, ImageDraw
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
import datetime
import textwrap

st.set_page_config(page_title="Form Filler with Template Mapping", layout="wide")
st.title("Template Mapping + Bulk Form Filler")

st.markdown("""
**Flow:**  
1. Upload blank template image (PNG/JPG).  
2. Click **Mapping** tab → click on image to mark fields (Name, Address, DOB, Photo, etc.) and save mapping JSON.  
3. Switch to **Processing** tab → upload Excel + input ZIP → app will fill template using mapping and create a final ZIP with per-candidate folders (maintains folder names).
""")

# -----------------
# Sidebar settings
# -----------------
st.sidebar.header("Options")
date_format = st.sidebar.text_input("Date format for output", value="%d-%m-%Y")
font_size = st.sidebar.number_input("Font size (pt)", min_value=8, max_value=20, value=11)
address_wrap_width = st.sidebar.number_input("Address wrap width (chars)", min_value=20, max_value=80, value=40)
pdf_dpi = st.sidebar.number_input("PDF / image DPI (affects sizing)", min_value=72, max_value=300, value=150)

# -----------------
# Tabs: Mapping & Processing
# -----------------
tab = st.radio("Mode", ["Mapping", "Processing"])

# Helper funcs
def normalize(s):
    if pd.isna(s):
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()

def detect_sr_column(df, hint_name=""):
    if hint_name and hint_name in df.columns:
        return hint_name
    candidates = [c for c in df.columns if re.search(r"sr|sno|srno|serial|sl ?no", c, re.I)]
    if candidates:
        return candidates[0]
    for c in df.columns:
        if pd.api.types.is_integer_dtype(df[c]) or pd.api.types.is_float_dtype(df[c]):
            return c
    return df.columns[0]

# -----------------
# Mapping Tab
# -----------------
if tab == "Mapping":
    st.header("Template Mapping (Click to place fields)")
    template_file = st.file_uploader("Upload blank template image (PNG/JPG)", type=["png","jpg","jpeg"])
    st.info("Mapping will record coordinates in pixels relative to the uploaded image's width & height.")

    if template_file:
        # load image
        img = Image.open(io.BytesIO(template_file.getvalue())).convert("RGB")
        img_w, img_h = img.size
        st.write(f"Image size: {img_w}px × {img_h}px")
        # show image with streamlit components
        st.markdown("**Click on the image to record a coordinate for a field.**")
        st.markdown("If image doesn't allow click (Streamlit limitation), download mapping helper below and then manually provide JSON.")

        # We'll implement a simple coordinate picker using st.image + user entering coords after visually checking.
        st.image(img, use_column_width=True)

        st.markdown("### Option A (Recommended): Visual + manual entry")
        st.markdown("Zoom the displayed image in another window, click visually where fields are, and enter X,Y (pixels).")

        cols = st.columns(2)
        with cols[0]:
            st.subheader("Enter coords manually (pixels)")
            field_name = st.selectbox("Field name", ["Name","Address","DOB","PIN CODE","Photo","Qualification","TrainingPeriod","Signature"], index=0)
            x_coord = st.number_input("X (pixels from left)", min_value=0, max_value=img_w, value=int(img_w*0.6))
            y_coord = st.number_input("Y (pixels from top)", min_value=0, max_value=img_h, value=int(img_h*0.15))
            w_box = st.number_input("Box width (px) — for photo/text box", min_value=10, max_value=img_w, value=120)
            h_box = st.number_input("Box height (px)", min_value=10, max_value=img_h, value=40)
            add_btn = st.button("Add / Update Mapping")
            if add_btn:
                # load existing mapping from session_state or create
                if "mapping" not in st.session_state:
                    st.session_state.mapping = {"image_size": [img_w, img_h], "fields": {}}
                st.session_state.mapping["fields"][field_name] = {"x": int(x_coord), "y": int(y_coord), "w": int(w_box), "h": int(h_box)}
                st.success(f"Mapping set for {field_name} -> ({x_coord},{y_coord})")
        with cols[1]:
            st.subheader("Current mapping JSON")
            if "mapping" in st.session_state:
                st.json(st.session_state.mapping)
            else:
                st.write("No mapping yet.")

        st.markdown("---")
        st.subheader("Option B: Use pre-built mapping JSON (Upload)")
        upload_map = st.file_uploader("Upload mapping JSON (if you already have one)", type=["json"])
        if upload_map:
            st.session_state.mapping = json.load(upload_map)
            st.success("Mapping loaded into session.")

        # Save mapping to file
        if "mapping" in st.session_state:
            save_name = st.text_input("Save mapping filename", value="mapping_template.json")
            if st.button("Download mapping JSON"):
                b = io.BytesIO()
                b.write(json.dumps(st.session_state.mapping, indent=2).encode("utf-8"))
                b.seek(0)
                st.download_button("Download mapping", data=b, file_name=save_name, mime="application/json")
        st.markdown("**Note:** For Photo coordinate, its `w` and `h` defines photo box size. For text fields, `w` and `h` help if you want to draw a box or limit wrap.")

# -----------------
# Processing Tab
# -----------------
else:
    st.header("Processing — Fill template for each candidate")
    st.markdown("Upload previously created mapping JSON (from Mapping tab), upload template (same image used during mapping), Excel and input ZIP.")

    mapping_upload = st.file_uploader("Upload mapping JSON", type=["json"])
    template_file = st.file_uploader("Upload same template image (PNG/JPG)", type=["png","jpg","jpeg"])
    excel_file = st.file_uploader("Upload Candidate Excel (.xlsx)", type=["xlsx"])
    input_zip = st.file_uploader("Upload Input ZIP (folders with photos)", type=["zip"])

    if st.button("Start Processing"):
        # validations
        if not mapping_upload:
            st.error("Please upload mapping JSON.")
            st.stop()
        if not template_file:
            st.error("Please upload template image.")
            st.stop()
        if not excel_file:
            st.error("Please upload Excel file.")
            st.stop()
        if not input_zip:
            st.error("Please upload input ZIP.")
            st.stop()

        mapping = json.load(mapping_upload)
        template_img = Image.open(io.BytesIO(template_file.getvalue())).convert("RGB")
        img_w, img_h = template_img.size
        # check mapping image size compatibility
        if "image_size" in mapping:
            mx, my = mapping["image_size"]
            if mx != img_w or my != img_h:
                st.warning(f"Mapping was created for image size {mx}x{my}, but uploaded template is {img_w}x{img_h}. Coordinates will still be used but verify positions.")

        # create working dir
        work_dir = Path(tempfile.mkdtemp(prefix="formfill_"))
        st.info(f"Working dir: {work_dir}")

        # save template to working dir
        template_path = work_dir / "template.png"
        with open(template_path, "wb") as f:
            f.write(template_file.getvalue())

        # save excel
        excel_path = work_dir / "data.xlsx"
        with open(excel_path, "wb") as f:
            f.write(excel_file.getvalue())

        # unzip input zip to dir
        unzip_dir = work_dir / "input_unzipped"
        unzip_dir.mkdir()
        with zipfile.ZipFile(io.BytesIO(input_zip.getvalue())) as z:
            z.extractall(path=str(unzip_dir))

        # find all images in unzipped
        images = [p for p in unzip_dir.rglob("*") if p.suffix.lower() in {".jpg",".jpeg",".png"}]
        st.write(f"Found {len(images)} images in input zip (searching recursively).")

        # read excel
        df = pd.read_excel(excel_path)
        st.write("Excel columns:", list(df.columns))
        sr_col = detect_sr_column(df)
        st.write("Using SrNo column:", sr_col)

        # helper to find image in input by sr or name parts
        def find_image_for_row(row):
            sr = normalize(row.get(sr_col,""))
            candidates = set()
            if sr:
                candidates.add(str(sr))
                candidates.add(str(sr).zfill(2))
                candidates.add(str(int(float(sr))) if str(sr).isdigit() else str(sr))
            # name heuristics
            name_cols_try = ["Name","NAME","Name of the Candidate","First_Name","Full Name","Candidate","Candidate Name"]
            name_str = ""
            for c in name_cols_try:
                if c in df.columns and not pd.isna(row.get(c,"")):
                    name_str = normalize(row.get(c,""))
                    break
            if not name_str:
                # first textual column
                for c in df.columns:
                    if df[c].dtype == object and not pd.isna(row.get(c,"")):
                        name_str = normalize(row.get(c,""))
                        break
            if name_str:
                parts = re.split(r"[ ,._\-]+", name_str.lower())
                for p in parts:
                    if p:
                        candidates.add(p)
            # match by filename contains
            for img in images:
                nm = img.name.lower()
                for patt in candidates:
                    if patt and patt.lower() in nm:
                        return img
            return None

        # prepare output dir
        output_dir = work_dir / "output"
        output_dir.mkdir()
        report = []

        # iterate rows
        for idx, row in df.iterrows():
            sr_val = normalize(row.get(sr_col,""))
            # build safe folder name: use Sr + Name (if exists)
            name_val = ""
            for c in ["Name","NAME","Name of the Candidate","First_Name","Full Name","Candidate"]:
                if c in df.columns and not pd.isna(row.get(c,"")):
                    name_val = normalize(row.get(c,""))
                    break
            folder_name = f"{sr_val} {name_val}".strip()
            safe_folder = re.sub(r"[^\w\s\-_.]", "_", folder_name).strip()
            person_dir = output_dir / safe_folder
            person_dir.mkdir(parents=True, exist_ok=True)

            matched_img = find_image_for_row(row)
            if matched_img:
                shutil.copy(matched_img, person_dir / matched_img.name)
                photo_path = person_dir / matched_img.name
            else:
                photo_path = None

            # create filled image based on template and mapping
            filled_img = template_img.copy()
            draw = ImageDraw.Draw(filled_img)
            # optional: load a TTF font if available, else default
            try:
                # common system font fallback; adjust path if needed
                font = ImageFont.truetype("DejaVuSans.ttf", font_size)
            except Exception:
                font = ImageFont.load_default()

            # place text fields
            for field, meta in mapping.get("fields", {}).items():
                x = meta.get("x",0)
                y = meta.get("y",0)
                w = meta.get("w",100)
                h = meta.get("h",30)
                # get matching column value from excel (try flexible matches)
                value = ""
                # exact match
                # try case-insensitive match
                for col in df.columns:
                    if col.strip().lower() == field.strip().lower():
                        value = normalize(row.get(col,""))
                        break
                # if not exact, try contains
                if not value:
                    for col in df.columns:
                        if field.strip().lower() in col.strip().lower():
                            value = normalize(row.get(col,""))
                            break
                # last fallback: try some known mappings
                if not value:
                    if field.lower() in ("name","candidate","full name"):
                        for c in ["Name","NAME","Name of the Candidate","First_Name"]:
                            if c in df.columns and not pd.isna(row.get(c,"")):
                                value = normalize(row.get(c,""))
                                break
                # For photo, handle separately
                if field.lower() == "photo" or field.lower() == "photo_box" or field.lower()=="photo_box":
                    if photo_path and photo_path.exists():
                        # paste photo resized into this box
                        try:
                            ph = Image.open(photo_path).convert("RGB")
                            # fit into w,h preserving aspect ratio
                            ph.thumbnail((w, h))
                            # center inside box
                            paste_x = x + (w - ph.width)//2
                            paste_y = y + (h - ph.height)//2
                            filled_img.paste(ph, (int(paste_x), int(paste_y)))
                        except Exception as e:
                            draw.rectangle([x,y,x+w,y+h], outline="black")
                            draw.text((x+2,y+2), "Photo error", font=font, fill="black")
                    else:
                        draw.rectangle([x,y,x+w,y+h], outline="black")
                        draw.text((x+2,y+2), "Photo missing", font=font, fill="black")
                    continue

                # if field is address or long, wrap text
                if field.lower() in ("address","address of candidate","address_line1","address_line2"):
                    lines = textwrap.wrap(value, width=address_wrap_width)
                    ly = y
                    for ln in lines:
                        draw.text((x, ly), ln, font=font, fill="black")
                        ly += font.getsize(ln)[1] + 2
                else:
                    # single line
                    draw.text((x, y), str(value), font=font, fill="black")

            # save filled image as PDF (one-page)
            out_pdf_path = person_dir / f"{safe_folder}_filled.pdf"
            # Convert to RGB if not
            if filled_img.mode != "RGB":
                filled_img = filled_img.convert("RGB")
            # Save as temporary image then embed into PDF with reportlab to keep A4 sizing
            temp_img_path = person_dir / "temp_filled.png"
            filled_img.save(temp_img_path, dpi=(pdf_dpi, pdf_dpi))
            # Create PDF with reportlab drawing the image full page
            c = canvas.Canvas(str(out_pdf_path), pagesize=A4)
            page_w, page_h = A4
            # Open saved image to get size and scale to fit A4 while preserving aspect
            bg = Image.open(temp_img_path)
            bg_w, bg_h = bg.size
            scale = min(page_w/bg_w, page_h/bg_h)
            draw_w = bg_w * scale
            draw_h = bg_h * scale
            x0 = (page_w - draw_w)/2
            y0 = (page_h - draw_h)/2
            bio = io.BytesIO()
            bg.save(bio, format="PNG")
            bio.seek(0)
            c.drawImage(ImageReader(bio), x0, y0, width=draw_w, height=draw_h)
            c.save()
            # cleanup temp image
            try:
                temp_img_path.unlink()
            except:
                pass

            report.append({
                "SrNo": sr_val,
                "Name": name_val,
                "Folder": str(person_dir.relative_to(work_dir)),
                "PhotoFound": bool(photo_path),
                "PhotoName": photo_path.name if photo_path else ""
            })

        # create final zip preserving per-candidate folders
        final_zip = work_dir / "final_filled_results.zip"
        with zipfile.ZipFile(final_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(output_dir):
                for file in files:
                    full = Path(root) / file
                    arc = str(full.relative_to(output_dir))
                    zf.write(full, arcname=arc)

        st.success("Processing complete.")
        st.write("Report:")
        st.dataframe(pd.DataFrame(report))

        with open(final_zip, "rb") as f:
            st.download_button("Download final ZIP", data=f, file_name="final_filled_results.zip", mime="application/zip")

        st.markdown(f"Temporary working folder: `{work_dir}` (you can inspect if needed).")
        st.info("After confirming results, delete temporary folder. For production, implement scheduled cleanup.")

        st.markdown("**If you want, I can now:**")
        st.markdown("- Provide a small UI to auto-detect field names from Excel and suggest mapping.  \n- Add PDF template support directly (render page to image).  \n- Wrap into Docker for client deployment.")
