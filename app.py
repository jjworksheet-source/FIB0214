import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
import datetime
import io
import random

# --- 1. SETUP & CONNECTION ---
st.set_page_config(page_title="Worksheet Generator", page_icon="📝", layout="wide")
st.title("📝 Worksheet Generator (Pro Layout)")

# Load Secrets
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
    st.cache_data.clear()
    st.rerun()

df = load_data()
if df.empty:
    st.warning("The 'standby' sheet is empty.")
    st.stop()

# --- 3. FILTER & SELECT ---
st.subheader("Select Questions")
try:
    ready_df = df[df['Status'].isin(['Ready', 'Waiting'])]
except KeyError:
    st.error("Column 'Status' not found.")
    st.stop()

if ready_df.empty:
    st.info("No questions with status 'Ready'.")
    st.stop()

edited_df = st.data_editor(
    ready_df,
    column_config={"Select": st.column_config.CheckboxColumn("Generate?", default=True)},
    disabled=["School", "Word", "Content"],
    hide_index=True,
    use_container_width=True
)

# --- 4. WORD DOCUMENT ENGINE ---

def set_chinese_font(run, size=14, bold=False):
    """Helper to set Chinese fonts correctly in Word"""
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'BiauKai') # 標楷體
    run.font.size = Pt(size)
    run.bold = bold

def create_header(doc, school_name, grade="P3", lesson_info="中文科詞語填充"):
    """Creates the specific header format from your reference code"""
    # Line 1: School + Grade (Centered)
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = p1.add_run(f"{school_name}   {grade}")
    set_chinese_font(run1, size=18, bold=True)
    
    # Line 2: Lesson Info (Centered)
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run(lesson_info)
    set_chinese_font(run2, size=16, bold=False)
    
    # Line 3: Name (Left) and Date (Right) using a Table
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.allow_autofit = False
    table.width = Cm(16) # Adjust based on page margins
    
    # Left Cell: Name
    cell_left = table.cell(0, 0)
    p_left = cell_left.paragraphs[0]
    run_left = p_left.add_run("學生姓名：__________________")
    set_chinese_font(run_left, size=14)
    
    # Right Cell: Date
    cell_right = table.cell(0, 1)
    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    run_right = p_right.add_run(f"日期：{today_str}")
    set_chinese_font(run_right, size=14)
    
    doc.add_paragraph("_" * 45) # Horizontal Line

def create_student_worksheet(school_name, questions):
    doc = Document()
    
    # 1. Header
    create_header(doc, school_name)
    
    # 2. Questions (Fill in the blank)
    p_title = doc.add_paragraph()
    run_title = p_title.add_run("甲、填充題")
    set_chinese_font(run_title, size=16, bold=True)
    
    for i, row in enumerate(questions):
        word = str(row['Word']).strip()
        sentence = str(row['Content']).strip()
        
        # Logic: Replace the word with underscores
        # If the word exists in the sentence, replace it. 
        # If not, just append blanks (fallback).
        blank = "_______"
        if word in sentence:
            question_text = sentence.replace(word, blank)
        else:
            question_text = sentence + " " + blank
            
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5
        run_num = p.add_run(f"{i+1}. ")
        set_chinese_font(run_num, size=14)
        
        run_q = p.add_run(question_text)
        set_chinese_font(run_q, size=14)
        
    # 3. Word List (Page Break)
    doc.add_page_break()
    p_list = doc.add_paragraph()
    run_list = p_list.add_run("乙、詞語表")
    set_chinese_font(run_list, size=18, bold=True)
    
    # Create a table for words (4 columns)
    words = [q['Word'] for q in questions]
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    
    row_cells = table.rows[0].cells
    idx = 0
    for word in words:
        if idx == 4:
            row_cells = table.add_row().cells
            idx = 0
        p = row_cells[idx].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(word)
        set_chinese_font(run, size=14)
        idx += 1
        
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def create_teacher_key(school_name, questions):
    doc = Document()
    
    # 1. Header
    create_header(doc, school_name, lesson_info="老師專用 - 詞語清單")
    
    # 2. List
    p_title = doc.add_paragraph()
    run_title = p_title.add_run("詞語清單 (依題目順序)")
    set_chinese_font(run_title, size=16, bold=True)
    
    for i, row in enumerate(questions):
        p = doc.add_paragraph()
        run = p.add_run(f"{i+1}. {row['Word']}")
        set_chinese_font(run, size=14)
        
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 5. GENERATION BUTTON ---
if st.button("🚀 Generate Files"):
    schools = edited_df['School'].unique()
    
    for school in schools:
        school_data = edited_df[edited_df['School'] == school].to_dict('records')
        
        if not school_data:
            continue
            
        st.divider()
        st.subheader(f"🏫 {school}")
        
        # 1. Student Worksheet
        # Shuffle questions for students like the reference code
        student_questions = school_data.copy()
        random.shuffle(student_questions)
        
        student_file = create_student_worksheet(school, student_questions)
        st.download_button(
            label=f"📥 下載學生試卷 ({school})",
            data=student_file,
            file_name=f"{school}_Student_Worksheet.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"btn_stu_{school}"
        )
        
        # 2. Teacher Key (Same order as student worksheet)
        teacher_file = create_teacher_key(school, student_questions)
        st.download_button(
            label=f"📥 下載老師清單 ({school})",
            data=teacher_file,
            file_name=f"{school}_Teacher_Key.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"btn_tea_{school}"
        )
