import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import datetime
import io
import os
import re
import base64
from pdf2image import convert_from_bytes

# --- 1. SETUP & CONNECTION ---
st.set_page_config(page_title="Worksheet Generator", page_icon="📝")
st.title("📝 Worksheet Generator")

# Try to import reportlab and handle font registration
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.enums import TA_CENTER

    font_paths = [
        "Kai.ttf",
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf"
    ]

    CHINESE_FONT = None
    for path in font_paths:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont('ChineseFont', path))
                CHINESE_FONT = 'ChineseFont'
                st.success(f"✅ Font loaded: {path}")
                break
            except Exception:
                continue

    if not CHINESE_FONT:
        st.error("❌ Chinese font not found. Please ensure Kai.ttf is in your GitHub repository.")

except ImportError:
    st.error("❌ reportlab not found. Please add 'reportlab' to your requirements.txt")
    st.stop()

# --- CONNECT TO GOOGLE CLOUD ---
try:
    key_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(
        key_dict,
        scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    )
    client = gspread.authorize(creds)
    SHEET_ID = st.secrets["app_config"]["spreadsheet_id"]
    st.success("✅ Connected to Google Cloud!")
except Exception as e:
    st.error(f"❌ Connection Error: {e}")
    st.stop()

# --- 2. READ DATA ---
@st.cache_data(ttl=60)
def load_data():
    try:
        sh = client.open_by_key(SHEET_ID)
        worksheet = sh.worksheet("standby")
        data = worksheet.get_all_records()
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Error reading sheet: {e}")
        return pd.DataFrame()

if st.button("🔄 Refresh Data"):
    load_data.clear()
    st.rerun()

df = load_data()

if df.empty:
    st.warning("The 'standby' sheet is empty or could not be read.")
    st.stop()

# --- 3. FILTER & SELECT ---
st.subheader("Select Questions")

if "Status" not in df.columns:
    st.error("Column 'Status' not found. Please check your Google Sheet headers.")
    st.stop()

status_norm = (
    df["Status"]
    .astype(str)
    .str.replace("\u00A0", " ", regex=False)
    .str.replace("\u3000", " ", regex=False)
    .str.strip()
)

ready_df = df[status_norm.isin(["Ready", "Waiting"])]

if ready_df.empty:
    st.info("No questions with status 'Ready' or 'Waiting'.")
    st.stop()

edited_df = st.data_editor(
    ready_df,
    column_config={
        "Select": st.column_config.CheckboxColumn("Generate?", default=True)
    },
    disabled=["School", "Word"],
    hide_index=True
)

# --- 4. GENERATE PDF FUNCTION ---
def create_pdf(school_name, questions):
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=letter)
    story = []

    styles = getSampleStyleSheet()
    font_name = CHINESE_FONT if CHINESE_FONT else 'Helvetica'

    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontName=font_name,
        fontSize=20,
        alignment=TA_CENTER,
        spaceAfter=12
    )
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontName=font_name,
        fontSize=14,
        leading=20,
        leftIndent=25,
        firstLineIndent=-25
    )

    story.append(Paragraph(f"<b>{school_name} - 校本填充工作紙</b>", title_style))
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph(f"Date: {datetime.date.today()}", normal_style))
    story.append(Spacer(1, 0.3*inch))

    for i, row in enumerate(questions):
        content = row['Content']
        content = re.sub(r'【】(.+?)【】', r'<u>\1</u>', content)
        content = re.sub(r'【(.+?)】', r'<u>\1</u>', content)
        p = Paragraph(f"{i+1}. {content}", normal_style)
        story.append(p)
        story.append(Spacer(1, 0.15*inch))

    doc.build(story)
    bio.seek(0)
    return bio

# --- Helper: Render PDF pages as images ---
def display_pdf_as_images(pdf_bytes):
    try:
        images = convert_from_bytes(pdf_bytes, dpi=150)
        for i, image in enumerate(images):
            st.image(image, caption=f"Page {i+1}", use_container_width=True)
    except Exception as e:
        st.error(f"Could not render preview: {e}")
        st.info("You can still download the PDF using the button on the left.")

# --- 5. PREVIEW & DOWNLOAD INTERFACE ---
st.divider()
st.subheader("🚀 Finalize Documents")

schools = edited_df['School'].unique() if not edited_df.empty else []

if len(schools) == 0:
    st.info("Select at least one question above to begin.")
else:
    selected_school = st.selectbox("Select School to Preview/Download", schools)
    school_data = edited_df[edited_df['School'] == selected_school]

    col1, col2 = st.columns([1, 2])

    # Generate PDF once — same bytes used for both preview and download
    pdf_buffer = create_pdf(selected_school, school_data.to_dict('records'))
    pdf_bytes = pdf_buffer.getvalue()

    with col1:
        st.write(f"**School:** {selected_school}")
        st.write(f"**Questions:** {len(school_data)}")

        st.download_button(
            label=f"📥 Download {selected_school}.pdf",
            data=pdf_bytes,
            file_name=f"{selected_school}_Review_{datetime.date.today()}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key=f"dl_{selected_school}"
        )

        st.info("💡 Fix typos in Google Sheet, then click 'Refresh Data' above.")

    with col2:
        st.write("🔍 **100% Accurate Preview**")
        display_pdf_as_images(pdf_bytes)
