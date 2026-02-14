import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import datetime
import io

# --- 1. SETUP & CONNECTION ---
st.set_page_config(page_title="Worksheet Generator", page_icon="📝")
st.title("📝 Worksheet Generator")

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
# Adjust column names if your sheet is different!
# Assuming columns: 'School', 'Word', 'Type', 'Content', 'Status'
try:
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
    disabled=["School", "Word", "Content"],
    hide_index=True
)

# --- 4. GENERATE WORD DOC ---
def create_docx(school_name, questions):
    doc = Document()
    
    # Font Setup
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.element.rPr.rFonts.set(qn('w:eastAsia'), 'TW-Kai') # Try to set Chinese font
    
    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{school_name} - Weekly Review")
    run.font.size = Pt(20)
    run.bold = True
    
    doc.add_paragraph(f"Date: {datetime.date.today()}")
    doc.add_paragraph("-" * 30)
    
    # Questions
    for i, row in enumerate(questions):
        p = doc.add_paragraph()
        run = p.add_run(f"{i+1}. {row['Content']}")
        run.font.size = Pt(14)
        
    # Save to memory
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

if st.button("🚀 Generate Word Document"):
    # Group by school
    schools = edited_df['School'].unique()
    
    for school in schools:
        school_data = edited_df[edited_df['School'] == school]
        
        if not school_data.empty:
            docx_file = create_docx(school, school_data.to_dict('records'))
            
            st.download_button(
                label=f"📥 Download {school}.docx",
                data=docx_file,
                file_name=f"{school}_Review_{datetime.date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
