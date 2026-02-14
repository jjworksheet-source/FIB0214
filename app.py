import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import datetime
import io
import random

# --- 1. SETUP & CONNECTION ---
st.set_page_config(page_title="Worksheet Generator", page_icon="📝", layout="wide")
st.title("📝 Worksheet Generator (Smart Headers)")

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

# --- 3. SMART COLUMN MAPPING ---
# This function finds the correct column name regardless of English/Chinese
def get_col_name(df, possible_names):
    for name in possible_names:
        if name in df.columns:
            return name
    return None

# Define possible names for each required field
col_map = {
    "School": get_col_name(df, ["School", "school", "學校"]),
    "Word": get_col_name(df, ["Word", "word", "詞語", "詞"]),
    "Content": get_col_name(df, ["Content", "content", "句子", "Question", "question", "題目"]),
    "Status": get_col_name(df, ["Status", "status", "狀態"])
}

# Check if all columns were found
missing_cols = [k for k, v in col_map.items() if v is None]
if missing_cols:
    st.error(f"❌ Missing columns in your Google Sheet: {', '.join(missing_cols)}")
    st.write("Found columns:", df.columns.tolist())
    st.stop()

# --- 4. FILTER & SELECT ---
st.subheader("Select Questions")
ready_df = df[df[col_map["Status"]].isin(['Ready', 'Waiting', 'Ready', 'Waiting'])]

if ready_df.empty:
    st.info("No questions with status 'Ready'.")
    st.stop()

# Rename columns temporarily for the editor so it looks clean
display_df = ready_df.rename(columns={
    col_map["School"]: "School",
    col_map["Word"]: "Word",
    col_map["Content"]: "Content",
    col_map["Status"]: "Status"
})

edited_df = st.data_editor(
    display_df,
    column_config={"Select": st.column_config.CheckboxColumn("Generate?", default=True)},
    disabled=["School", "Word", "Content"],
    hide_index=True,
    use_container_width=True
)

# --- 5. WORD DOCUMENT ENGINE ---

def set_chinese_font(run, size=14, bold=False):
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'BiauKai')
    run.font.size = Pt(size)
    run.bold = bold

def create_header(doc, school_name):
    # Line 1: School (Centered)
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = p1.add_run(f"{school_name}")
    set_chinese_font(run1, size=18, bold=True)
    
    # Line 2: Title (Centered)
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run("中文科詞語填充")
    set_chinese_font(run2, size=16, bold=False)
    
    # Line 3: Name/Date Table
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.width = Cm(16)
    
    cell_left = table.cell(0, 0)
    p_left = cell_left.paragraphs[0]
    run_left = p_left.add_run("學生姓名：__________________")
    set_chinese_font(run_left, size=14)
    
    cell_right = table.cell(0, 1)
    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    run_right = p_right.add_run(f"日期：{today_str}")
    set_chinese_font(run_right, size=14)
    
    doc.add_paragraph("_" * 45)

def create_student_worksheet(school_name, questions):
    doc = Document()
    create_header(doc, school_name)
    
    p_title = doc.add_paragraph()
    run_title = p_title.add_run("甲、填充題")
    set_chinese_font(run_title, size=16, bold=True)
    
    for i, row in enumerate(questions):
        word = str(row['Word']).strip()
        sentence = str(row['Content']).strip()
        
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
        
    doc.add_page_break()
    p_list = doc.add_paragraph()
    run_list = p_list.add_run("乙、詞語表")
    set_chinese_font(run_list, size=18, bold=True)
    
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
    create_header(doc, f"{school_name} (老師專用)")
    
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

# --- 6. GENERATION BUTTON ---
if st.button("🚀 Generate Files"):
    schools = edited_df['School'].unique()
    
    for school in schools:
        school_data = edited_df[edited_df['School'] == school].to_dict('records')
        
        if not school_data:
            continue
            
        st.divider()
        st.subheader(f"🏫 {school}")
        
        # Shuffle for students
        student_questions = school_data.copy()
        random.shuffle(student_questions)
        
        # Generate Student File
        student_file = create_student_worksheet(school, student_questions)
        st.download_button(
            label=f"📥 下載學生試卷 ({school})",
            data=student_file,
            file_name=f"{school}_Student_Worksheet.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"btn_stu_{school}"
        )
        
        # Generate Teacher File
        teacher_file = create_teacher_key(school, student_questions)
        st.download_button(
            label=f"📥 下載老師清單 ({school})",
            data=teacher_file,
            file_name=f"{school}_Teacher_Key.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"btn_tea_{school}"
        )
