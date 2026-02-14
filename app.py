import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import datetime
import io
import os
import re

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
    
    # Register a font that supports Chinese if available
    # Common paths for fonts on Linux/Streamlit Cloud
    font_paths = [
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
        "TW-Kai-98_1.ttf", # Upload this to your GitHub repo
        "NotoSansTC-Regular.otf"
    ]
    
    CHINESE_FONT = None
    for path in font_paths:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont('ChineseFont', path))
                CHINESE_FONT = 'ChineseFont'
                break
            except:
                continue
    
    if not CHINESE_FONT:
        st.warning("⚠️ Chinese font not found. Chinese characters may appear as boxes in the PDF.")
        uploaded_font = st.file_uploader("📤 Upload Chinese Font (.ttf or .otf)", type=['ttf', 'otf'])
        if uploaded_font is not None:
            try:
                # Save uploaded font to a temporary file to register it
                with open("temp_font.ttf", "wb") as f:
                    f.write(uploaded_font.getbuffer())
                pdfmetrics.registerFont(TTFont('ChineseFont', "temp_font.ttf"))
                CHINESE_FONT = 'ChineseFont'
                st.success("✅ Font uploaded and registered successfully!")
            except Exception as e:
                st.error(f"❌ Error registering font: {e}")
except ImportError:
    st.error("❌ reportlab not found. Please add 'reportlab' to your requirements.txt")
    st.stop()

# Load Secrets
try:
    # Construct the credentials dictionary from secrets
    key_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(
        key_dict,
        scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    )
    client = gspread.authorize(creds)
    
    # Get Config
    SHEET_ID = st.secrets["app_config"]["spreadsheet_id"]
    
    st.success("✅ Connected to Google Cloud!")
except Exception as e:
    st.error(f"❌ Connection Error: {e}")
    st.stop()

# --- 2. READ DATA ---
@st.cache_data(ttl=60) # Cache for 1 minute so it doesn't reload constantly
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
    st.cache_data.clear()
    st.rerun()

df = load_data()

if df.empty:
    st.warning("The 'standby' sheet is empty or could not be read.")
    st.stop()

# --- 3. FILTER & SELECT ---
st.subheader("Select Questions")

# Filter for 'Ready' status
# We look for columns: 'School', 'Word', 'Type', 'Content', 'Status'
try:
    # Filter rows where Status is 'Ready' or 'Waiting'
    ready_df = df[df['Status'].isin(['Ready', 'Waiting'])]
except KeyError:
    st.error("Column 'Status' not found. Please check your Google Sheet headers.")
    st.write("Available columns:", df.columns.tolist())
    st.stop()

if ready_df.empty:
    st.info("No questions with status 'Ready' or 'Waiting'.")
    st.stop()

# Show data editor
edited_df = st.data_editor(
    ready_df,
    column_config={
        "Select": st.column_config.CheckboxColumn("Generate?", default=True)
    },
    disabled=["School", "Word"],
    hide_index=True
)

# --- 4. GENERATE PDF ---
def create_pdf(school_name, questions):
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=letter)
    story = []
    
    # Styles
    styles = getSampleStyleSheet()
    
    # Use registered Chinese font if found, otherwise fallback
    font_name = CHINESE_FONT if CHINESE_FONT else 'Helvetica'
    
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontName=font_name,
        fontSize=20,
        alignment=TA_CENTER,
        spaceAfter=12
    )
    
    # Hanging indent style: leftIndent moves the whole block, firstLineIndent moves it back
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontName=font_name,
        fontSize=14,
        leading=20,
        leftIndent=25,
        firstLineIndent=-25
    )
    
    # Title
    title = Paragraph(f"<b>{school_name} - Weekly Review</b>", title_style)
    story.append(title)
    story.append(Spacer(1, 0.2*inch))
    
    # Date
    date_text = Paragraph(f"Date: {datetime.date.today()}", normal_style)
    story.append(date_text)
    story.append(Spacer(1, 0.3*inch))
    
    # Questions
    for i, row in enumerate(questions):
        content = row['Content']
        # Convert 【】text【】 to <u>text</u> for underline (專名號)
        content = re.sub(r'【】(.+?)【】', r'<u>\1</u>', content)
        question_text = f"{i+1}. {content}"
        p = Paragraph(question_text, normal_style)
        story.append(p)
        story.append(Spacer(1, 0.15*inch))
    
    # Build PDF
    doc.build(story)
    bio.seek(0)
    return bio

if st.button("🚀 Generate PDF Document"):
    # Group by school
    schools = edited_df['School'].unique()
    
    for school in schools:
        school_data = edited_df[edited_df['School'] == school]
        
        if not school_data.empty:
            pdf_file = create_pdf(school, school_data.to_dict('records'))
            
            st.download_button(
                label=f"📥 Download {school}.pdf",
                data=pdf_file,
                file_name=f"{school}_Review_{datetime.date.today()}.pdf",
                mime="application/pdf"
            )
