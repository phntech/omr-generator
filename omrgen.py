import os
import re
import base64
import pandas as pd
import streamlit as st
from io import BytesIO
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Table, TableStyle
import zipfile
from PIL import Image, ImageDraw

# ===== PATH CONFIG (UNCHANGED) =====
BASE_DIR = Path(__file__).parent
child_omr_template = BASE_DIR / "child_omr.jpg"
master_omr_template = BASE_DIR / "master_omr.jpg"
LOGO_FILE = BASE_DIR / "logo.webp"

# ===== Master OMR Bubble positions (UNCHANGED) =====
MASTER_ROLL_X_CM = [10.1, 11.5, 12.9]
MASTER_BUBBLE_Y_TOP_CM = [22, 22, 22]
MASTER_BUBBLE_SPACING_CM = 0.62
MASTER_BUBBLE_RADIUS_CM = 0.24

# ===== Child OMR Bubble positions (UNCHANGED) =====
SHIFT_X_CM = 6
SHIFT_Y_CM = -0.3
CHILD_ROLL_X_CM = [9.9 + SHIFT_X_CM, 11.3 + SHIFT_X_CM, 12.6 + SHIFT_X_CM]
CHILD_BUBBLE_Y_TOP_CM = [21.5 + SHIFT_Y_CM, 21.5 + SHIFT_Y_CM, 21.5 + SHIFT_Y_CM]
CHILD_BUBBLE_SPACING_CM = 0.61
CHILD_BUBBLE_RADIUS_CM = 0.23

# ----------------------------------------------------------------------
# ----- Utility Functions (UNCHANGED, but kept for completeness) -----
# ----------------------------------------------------------------------
def normalize_col_name(s):
    return re.sub(r'[^a-z0-9]', '', str(s).lower().strip()) if s else ""

def find_column(df_cols_norm, aliases):
    for orig_col, norm in df_cols_norm.items():
        for a in aliases:
            if norm == a:
                return orig_col
    for orig_col, norm in df_cols_norm.items():
        for a in aliases:
            if a in norm:
                return orig_col
    return None

def safe_filename(s):
    s = str(s).strip()
    s = re.sub(r'[\\/*?:"<>|]', '_', s)
    s = re.sub(r'\s+', '_', s)
    return s[:200]

def format_roll_value(v):
    if pd.isna(v) or not str(v).strip():
        return "000"
    try:
        return str(int(float(v))).zfill(3)
    except ValueError:
        # Fallback for non-numeric or complex values
        return str(v).zfill(3)[:3]

def fill_roll_bubbles_master(c, roll_no):
    roll_no = roll_no.zfill(3)
    for i, digit_char in enumerate(roll_no):
        if not digit_char.isdigit():
            continue
        digit = int(digit_char)
        x = (MASTER_ROLL_X_CM[i] * cm) / 2.2
        y = MASTER_BUBBLE_Y_TOP_CM[i] * cm - digit * MASTER_BUBBLE_SPACING_CM * cm - (2.6 * cm) + 0.03 * cm
        c.setFillColor(colors.black)
        c.circle(x, y, MASTER_BUBBLE_RADIUS_CM * cm, stroke=0, fill=1)

def fill_roll_bubbles_child(c, roll_no):
    roll_no = roll_no.zfill(3)
    for i, digit_char in enumerate(roll_no):
        if not digit_char.isdigit():
            continue
        digit = int(digit_char)
        x = (CHILD_ROLL_X_CM[i] * cm) / 2.2
        y = CHILD_BUBBLE_Y_TOP_CM[i] * cm - digit * CHILD_BUBBLE_SPACING_CM * cm - (2.6 * cm) + 0.2 * cm
        c.setFillColor(colors.black)
        c.circle(x, y, CHILD_BUBBLE_RADIUS_CM * cm, stroke=0, fill=1)

def draw_roll_number_text(c, roll_no, template="master"):
    roll_no = roll_no.zfill(3)
    if template == "master":
        text_y = MASTER_BUBBLE_Y_TOP_CM[0] * cm - (2.1 * cm)
        x_positions = MASTER_ROLL_X_CM
    else:
        text_y = CHILD_BUBBLE_Y_TOP_CM[0] * cm - (2.0 * cm)
        x_positions = CHILD_ROLL_X_CM

    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(colors.black)
    for i, digit_char in enumerate(roll_no):
        x = (x_positions[i] * cm) / 2.2
        c.drawCentredString(x, text_y, digit_char)

def parse_class_value(class_val):
    if pd.isna(class_val):
        return None
    s = str(class_val).strip().lower()

    if s.isdigit():
        return int(s)

    match = re.search(r'(\d+)', s)
    if match:
        return int(match.group(1))

    roman_map = {
        "i": 1, "ii": 2, "iii": 3, "iv": 4, "v": 5,
        "vi": 6, "vii": 7, "viii": 8, "ix": 9, "x": 10,
        "xi": 11, "xii": 12
    }
    if s in roman_map:
        return roman_map[s]

    words_map = {
        "first": 1, "second": 2, "third": 3, "fourth": 4, "fifth": 5,
        "sixth": 6, "seventh": 7, "eighth": 8, "ninth": 9, "tenth": 10,
        "eleventh": 11, "twelfth": 12
    }
    if s in words_map:
        return words_map[s]

    return None

# ----------------------------------------------------------------------
# ----- Optimized Image Loading and Caching -----
# ----------------------------------------------------------------------

# Placeholder Images (Only run if file is missing, using Path.exists())
def create_placeholder_image(path, text="OMR Missing"):
    img = Image.new("RGB", (595, 842), color=(255,255,255))
    d = ImageDraw.Draw(img)
    d.text((50,400), text, fill=(0,0,0))
    # Save the placeholder image as a temporary file in memory or local disk if needed.
    # For Streamlit deployment, relying on the actual files is better.
    # The current setup creates placeholder files if they don't exist, which is fine, 
    # but the caching below is the main speed gain.
    try:
        img.save(path)
    except Exception as e:
        # Cannot write in read-only environment (Streamlit Cloud). 
        # For deployment, the actual files MUST be in the repo.
        pass

if not child_omr_template.exists():
    create_placeholder_image(child_omr_template, text="Child OMR Missing")
if not master_omr_template.exists():
    create_placeholder_image(master_omr_template, text="Master OMR Missing")

@st.cache_data
def load_omr_templates():
    """Load OMR images once and cache the ImageReader objects."""
    omr_images = {}
    try:
        omr_images['child'] = ImageReader(child_omr_template)
    except Exception as e:
        st.error(f"Failed to load child OMR template: {e}")
        omr_images['child'] = None
    try:
        omr_images['master'] = ImageReader(master_omr_template)
    except Exception as e:
        st.error(f"Failed to load master OMR template: {e}")
        omr_images['master'] = None
    return omr_images

OMR_TEMPLATES = load_omr_templates()

# ----------------------------------------------------------------------
# ===== Streamlit App =====
# ----------------------------------------------------------------------
st.set_page_config(page_title="OMR Sheet Generator", layout="wide")

# Logo
if Path(LOGO_FILE).exists():
    try:
        with open(LOGO_FILE, "rb") as f:
            data = f.read()
        img_base64 = base64.b64encode(data).decode()
        st.markdown(
            f'<div style="text-align:center;"><img src="data:image/webp;base64,{img_base64}" width="150"></div>',
            unsafe_allow_html=True
        )
    except Exception:
        st.image("https://placehold.co/150x50/3498db/ffffff?text=LOGO+Missing", width=150)
        st.warning(f"Note: Error loading logo file. Using placeholder.")
else:
    st.image("https://placehold.co/150x50/3498db/ffffff?text=LOGO+Missing", width=150)
    st.warning(f"Note: Local logo file '{LOGO_FILE}' not found. Using placeholder.")

st.markdown(
    "<h1 style='text-align: center; color: white;'>OMR Sheet Generator</h1>",
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    # --- Check for loaded OMR templates ---
    if OMR_TEMPLATES['child'] is None and OMR_TEMPLATES['master'] is None:
        st.error("Cannot proceed: Both OMR template images failed to load. Ensure they are in your repository.")
        st.stop()

    with st.spinner("⏳ Please wait, generating PDFs..."):
        xls = pd.ExcelFile(uploaded_file)
        output_zip = BytesIO()

        with zipfile.ZipFile(output_zip, "w") as zipf:
            for sheet_name in xls.sheet_names:
                try:
                    # Reading the sheet is outside the page loop, which is efficient
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, dtype=object)
                except Exception as e:
                    st.warning(f"⚠ Could not read sheet '{sheet_name}': {e}. Skipping sheet.")
                    continue

                # --- Column Mapping (outside row loop) ---
                df_cols_norm = {orig: normalize_col_name(orig) for orig in df.columns}
                aliases = {
                    "school_name": ["schoolname", "scoolname", "school"],
                    "class": ["class", "grade", "standard"],
                    "division": ["division", "section"],
                    "roll_no": ["rollno", "rollnumber", "roll_no"],
                    "student_name": ["nameofthestudent", "name", "studentname"],
                }
                col_map = {canon: find_column(df_cols_norm, al_list) for canon, al_list in aliases.items()}
                
                # Report missing columns once per sheet
                missing_cols = [k for k, v in col_map.items() if v is None]
                if missing_cols:
                    st.info(f"ℹ️ Missing columns in sheet '{sheet_name}': {', '.join(missing_cols)}. Fields will be blank.")

                pdf_buffer = BytesIO()
                c = canvas.Canvas(pdf_buffer, pagesize=A4)
                width, height = A4
                
                # --- Pre-calculate table style (outside row loop) ---
                table_style = TableStyle([
                    ("BOX", (0,0), (-1,-1), 0.8, colors.black),
                    ("INNERGRID", (0,0), (-1,-1), 0.5, colors.black),
                    ("FONTNAME", (0,0), (-1,-1), "Helvetica-Bold"),
                    ("FONTSIZE", (0,0), (-1,-1), 11),
                    ("ALIGN", (0,0), (-1,-1), "LEFT"),
                    ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                    ("LEFTPADDING", (0,0), (-1,-1), 10),
                    ("RIGHTPADDING", (0,0), (-1,-1), 10),
                    ("TOPPADDING", (0,0), (-1,-1), 5),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 5),
                ])
                table_width = width * 0.7
                table_x = (width - table_width)/2
                table_y_offset = height - 4.5*cm 

                # --- Main Row Loop ---
                for _, row in df.iterrows():
                    student_name = row.get(col_map.get("student_name"), "") if col_map.get("student_name") else ""
                    school_name = row.get(col_map.get("school_name"), "") if col_map.get("school_name") else ""
                    class_name = row.get(col_map.get("class"), "") if col_map.get("class") else ""
                    division = row.get(col_map.get("division"), "") if col_map.get("division") else ""
                    roll_no_raw = row.get(col_map.get("roll_no"), "") if col_map.get("roll_no") else ""
                    roll_no = format_roll_value(roll_no_raw)

                    parsed_class = parse_class_value(class_name)
                    is_child = (parsed_class is not None and parsed_class in [1, 2, 3])
                    template_type = "child" if is_child else "master"
                    omr_img = OMR_TEMPLATES.get(template_type) # ***OPTIMIZATION: Get cached ImageReader***

                    if omr_img is None:
                        st.warning(f"Skipping row for Roll No. {roll_no}: Template failed to load.")
                        continue
                    
                    # --- Draw Image ---
                    c.drawImage(omr_img, 0, 0, width=width, height=height, preserveAspectRatio=True)

                    # --- Fill Roll Bubbles ---
                    if is_child:
                        fill_roll_bubbles_child(c, roll_no)
                    else:
                        fill_roll_bubbles_master(c, roll_no)

                    # --- Draw Roll Number Text ---
                    draw_roll_number_text(c, roll_no, template=template_type)

                    # --- Draw Info Table ---
                    data = [
                        [f"Student Name: {student_name or ' '}"],
                        [f"School: {school_name or ' '}"],
                        [f"Class: {class_name or ' '}      Division: {division or ' '}"],
                        ["Question Paper Set: _____________"],
                    ]
                    table = Table(data, colWidths=[table_width])
                    table.setStyle(table_style)
                    w, h = table.wrap(0, 0)
                    table.drawOn(c, table_x, table_y_offset - h) # Draw table

                    c.showPage() # End of page for current student

                # --- Save and Zip PDF ---
                c.save()
                pdf_data = pdf_buffer.getvalue()
                pdf_buffer.close()
                pdf_filename = f"{safe_filename(sheet_name)}.pdf"
                zipf.writestr(pdf_filename, pdf_data)

        st.success("✅ PDFs Generated Successfully! The application is optimized for speed.")
        st.download_button(
            label="⬇ Download All PDFs (ZIP)",
            data=output_zip.getvalue(),
            file_name="Generated_OMRs.zip",
            mime="application/zip"
        )