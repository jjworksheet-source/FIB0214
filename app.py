import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import datetime
import io
import os
import re
import base64
import random
import time
from pdf2image import convert_from_bytes
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition, Email
from python_http_client.exceptions import HTTPError

# --- 1. SETUP & CONNECTION ---
st.set_page_config(page_title="Worksheet Generator", page_icon="📝")
st.title("📝 Worksheet Generator")
# --- Initialize session state for shuffled questions ---
if 'shuffled_cache' not in st.session_state:
    st.session_state.shuffled_cache = {}

# Try to import reportlab and handle font registration
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
    from reportlab.lib import colors
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
        st.error(f"Error reading standby sheet: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=60)
def load_students():
    try:
        sh = client.open_by_key(SHEET_ID)
        worksheet = sh.worksheet("學生資料")
        data = worksheet.get_all_records()
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Error reading 學生資料 sheet: {e}")
        return pd.DataFrame()

if st.button("🔄 Refresh Data"):
    load_data.clear()
    load_students.clear()
    st.rerun()

df = load_data()
student_df = load_students()

if df.empty:
    st.warning("The 'standby' sheet is empty or could not be read.")
    st.stop()

# --- 3. FILTER & SELECT ---
st.subheader("Select Questions")
if st.button("🔀 重新打亂題目順序"):
    st.session_state.shuffled_cache = {}
    st.rerun()

if "Status" not in df.columns:
    st.error("Column 'Status' not found. Please check your Google Sheet headers.")
    st.stop()

if "level" not in df.columns and "Level" not in df.columns:
    st.error("Column 'Level' not found. Please check your Google Sheet headers.")
    st.stop()

# Normalize column names
df.columns = [c.strip() for c in df.columns]
level_col = "Level" if "Level" in df.columns else "level"
df = df.rename(columns={level_col: "Level"})

# Clean student_df column names
if not student_df.empty:
    student_df.columns = [c.strip() for c in student_df.columns]
    for col in student_df.columns:
        if student_df[col].dtype == object:
            student_df[col] = student_df[col].astype(str).str.strip()

# Clean standby df
for col in df.columns:
    if df[col].dtype == object:
        df[col] = df[col].astype(str).str.strip()

# --- Sidebar: Level Filter ---
with st.sidebar:
    st.header("🎓 篩選年級")
    available_levels = sorted(df["Level"].astype(str).str.strip().unique().tolist())
    selected_level = st.radio("選擇年級", available_levels, index=0)
    st.divider()
    st.info(f"目前顯示：**{selected_level}** 的題目")

    # --- Sidebar: Mode Toggle ---
    st.divider()
    st.header("📬 發送模式")
    send_mode = st.radio(
        "選擇模式",
        ["📄 按學校預覽下載", "👨‍👩‍👧 按學生寄送 (配對學生資料)"],
        index=0
    )

status_norm = (
    df["Status"]
    .astype(str)
    .str.replace("\u00A0", " ", regex=False)
    .str.replace("\u3000", " ", regex=False)
    .str.strip()
)

level_norm = df["Level"].astype(str).str.strip()
ready_df = df[status_norm.isin(["Ready", "Waiting"]) & (level_norm == selected_level)]

if ready_df.empty:
    st.info(f"No questions with status 'Ready' or 'Waiting' for {selected_level}.")
    st.stop()

edited_df = st.data_editor(
    ready_df,
    column_config={
        "Select": st.column_config.CheckboxColumn("Generate?", default=True)
    },
    disabled=["School", "Level", "Word"],
    hide_index=True
)

# --- HELPER: Shuffle questions once for consistency across all documents ---
def get_shuffled_questions(questions, cache_key):
    """
    Get shuffled questions with caching.
    Same cache_key returns same order within session.
    Different sessions get different random orders.
    """
    # Check if already shuffled in this session
    if cache_key in st.session_state.shuffled_cache:
        return st.session_state.shuffled_cache[cache_key]

    # First time: shuffle and cache
    questions_list = list(questions)
    random.seed(int(time.time() * 1000))
    random.shuffle(questions_list)
    st.session_state.shuffled_cache[cache_key] = questions_list
    return questions_list

# --- 4. GENERATE PDF FUNCTION (Student Version) ---
def create_pdf(school_name, level, questions, student_name=None):
    """
    Create student PDF.
    Questions are displayed in the order provided (no internal shuffling).
    Answers are hidden (replaced with underlines).
    Page 2: Vocabulary table with unique words from the "Word" column.
    """
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=letter)
    story = []

    styles = getSampleStyleSheet()
    font_name = CHINESE_FONT if CHINESE_FONT else 'Helvetica'

    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontName=font_name,
        fontSize=22,
        alignment=TA_CENTER,
        spaceAfter=12
    )
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontName=font_name,
        fontSize=18,
        leading=26,
        leftIndent=30,
        firstLineIndent=-30
    )
    vocab_title_style = ParagraphStyle(
        'VocabTitle',
        parent=styles['Heading2'],
        fontName=font_name,
        fontSize=20,
        alignment=TA_CENTER,
        spaceAfter=20
    )

    if student_name:
        title_text = f"<b>{school_name} ({level}) - {student_name} - 校本填充工作紙</b>"
    else:
        title_text = f"<b>{school_name} ({level}) - 校本填充工作紙</b>"

    story.append(Paragraph(title_text, title_style))
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph(f"日期: {datetime.date.today() + datetime.timedelta(days=1)}", normal_style))
    story.append(Spacer(1, 0.3*inch))

    # Generate questions in the order provided (shuffling done externally)
    for i, row in enumerate(questions):
        content = row['Content']
        # Hide answers: replace 【answer】 with underline
        content = re.sub(r'【】(.+?)【】', r'<u>\1</u>', content)
        content = re.sub(r'【(.+?)】', r'<u>________</u>', content)
        p = Paragraph(f"{i+1}. {content}", normal_style)
        story.append(p)
        story.append(Spacer(1, 0.2*inch))

    # --- PAGE 2: Vocabulary Table ---
    # Extract unique words from the "Word" column
    words = [row.get('Word', '').strip() for row in questions]
    unique_words = list(dict.fromkeys([w for w in words if w]))  # Remove duplicates, preserve order
    
    if unique_words:
        story.append(PageBreak())
        story.append(Paragraph("<b>詞語表</b>", vocab_title_style))
        story.append(Spacer(1, 0.2*inch))
        
        # Organize words into rows (4 columns)
        num_cols = 4
        table_data = []
        for i in range(0, len(unique_words), num_cols):
            row = unique_words[i:i+num_cols]
            # Pad row with empty strings if needed
            while len(row) < num_cols:
                row.append('')
            table_data.append(row)
        
        # Create table with styling
        col_width = 1.5*inch
        vocab_table = Table(table_data, colWidths=[col_width]*num_cols)
        vocab_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), font_name),
            ('FONTSIZE', (0, 0), (-1, -1), 14),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('TOPPADDING', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
            ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ]))
        story.append(vocab_table)

    doc.build(story)
    bio.seek(0)
    return bio

# --- FEATURE 2: Teacher Answer PDF Function ---
def create_answer_pdf(school_name, level, questions, student_name=None):
    """
    Create teacher answer PDF with answers visible.
    Questions are displayed in the order provided (same as student PDF).
    Answers are shown clearly highlighted in RED - using the "Word" column.
    """
    from reportlab.lib.colors import blue, red
    
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=letter)
    story = []

    styles = getSampleStyleSheet()
    font_name = CHINESE_FONT if CHINESE_FONT else 'Helvetica'

    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontName=font_name,
        fontSize=22,
        alignment=TA_CENTER,
        spaceAfter=12
    )
    subtitle_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Heading2'],
        fontName=font_name,
        fontSize=16,
        alignment=TA_CENTER,
        textColor=red,
        spaceAfter=12
    )
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontName=font_name,
        fontSize=18,
        leading=26,
        leftIndent=30,
        firstLineIndent=-30
    )

    # Title with "教師版答案" indicator
    if student_name:
        title_text = f"<b>{school_name} ({level}) - {student_name} - 校本填充工作紙</b>"
    else:
        title_text = f"<b>{school_name} ({level}) - 校本填充工作紙</b>"

    story.append(Paragraph(title_text, title_style))
    story.append(Paragraph("<b>教師版答案 (Answer Key)</b>", subtitle_style))
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph(f"日期: {datetime.date.today() + datetime.timedelta(days=1)}", normal_style))
    story.append(Spacer(1, 0.3*inch))

    # Display questions in the order provided (same order as student PDF)
    for i, row in enumerate(questions):
        content = row['Content']
        
        # Get the answer from the "Word" column
        answer = row.get('Word', '')
        
        # Strategy: The Content field may have blanks in different formats:
        # 1. Underscores: ________ or ＿＿＿＿
        # 2. Brackets with answer: 【answer】
        # 3. Empty brackets: 【】text【】
        
        # First, try to replace underscores/blanks with the answer from Word column
        if answer:
            # Replace various blank patterns with the highlighted answer
            answer_html = f'<font color="red"><b>【{answer}】</b></font>'
            
            # Pattern: Multiple underscores (half-width or full-width)
            content = re.sub(r'_{2,}|＿{2,}', answer_html, content)
            
            # Pattern: 【】text【】 (empty brackets surrounding text - keep original behavior)
            content = re.sub(r'【】(.+?)【】', r'<font color="red"><b>【\1】</b></font>', content)
            
            # Pattern: 【answer】 (answer inside brackets)
            content = re.sub(r'【(.+?)】', r'<font color="red"><b>【\1】</b></font>', content)
        else:
            # No Word answer - try bracket patterns only
            content = re.sub(r'【】(.+?)【】', r'<font color="red"><b>【\1】</b></font>', content)
            content = re.sub(r'【(.+?)】', r'<font color="red"><b>【\1】</b></font>', content)
        
        p = Paragraph(f"{i+1}. {content}", normal_style)
        story.append(p)
        story.append(Spacer(1, 0.2*inch))

    doc.build(story)
    bio.seek(0)
    return bio

# --- SendGrid Email Function (FIXED) ---
def send_email_with_pdf(to_email, student_name, school_name, grade, pdf_bytes, cc_email=None):
    try:
        sg_config = st.secrets["sendgrid"]

        # --- CLEAN & VALIDATE RECIPIENT ---
        recipient = str(to_email).strip()
        if not re.match(r'^[\w\.-]+@[\w\.-]+\.\w+$', recipient):
            return False, f"無效的家長電郵格式: '{recipient}'"

        # --- BUILD MESSAGE (use Email object, not tuple) ---
        from_email_obj = Email(sg_config["from_email"], sg_config.get("from_name", ""))

        # Clean student name for filename (remove non-ASCII)
        safe_name = re.sub(r'[^\w\-]', '_', str(student_name).strip())

        message = Mail(
            from_email=from_email_obj,
            to_emails=recipient,
            subject=f"【工作紙】{school_name} ({grade}) - {student_name} 的校本填充練習",
            html_content=f"""
                <p>親愛的家長您好：</p>
                <p>附件為 <strong>{student_name}</strong> 同學在 <strong>{school_name} ({grade})</strong> 的校本填充工作紙。</p>
                <p>請下載並列印供同學練習。祝 學習愉快！</p>
                <br><p>-- 自動發送系統 --</p>
            """
        )

        # --- CLEAN & VALIDATE CC ---
        if cc_email:
            cc_clean = str(cc_email).strip().lower()
            if cc_clean not in ["n/a", "nan", "", "none"] and "@" in cc_clean and cc_clean != recipient.lower():
                message.add_cc(cc_clean)

        # --- ATTACHMENT ---
        encoded_pdf = base64.b64encode(pdf_bytes).decode()
        attachment = Attachment(
            FileContent(encoded_pdf),
            FileName(f"{safe_name}_Worksheet.pdf"),
            FileType('application/pdf'),
            Disposition('attachment')
        )
        message.add_attachment(attachment)

        # --- SEND ---
        sg = SendGridAPIClient(sg_config["api_key"])
        response = sg.send(message)

        if 200 <= response.status_code < 300:
            return True, "發送成功"
        else:
            return False, f"SendGrid Error: {response.status_code}"

    except HTTPError as e:
        # Shows the REAL detailed error from SendGrid
        try:
            return False, e.body.decode("utf-8")
        except Exception:
            return False, str(e)
    except Exception as e:
        return False, str(e)

def create_docx(school_name, level, questions, student_name=None):
    """
    Create Word document with questions.
    Questions are displayed in the order provided (same as PDFs).
    """
    doc = Document()

    if student_name:
        title_text = f"{school_name} ({level}) - {student_name} - 校本填充工作紙"
    else:
        title_text = f"{school_name} ({level}) - 校本填充工作紙"

    title = doc.add_heading(title_text, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    date_para = doc.add_paragraph(f"日期: {datetime.date.today() + datetime.timedelta(days=1)}")
    date_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    doc.add_paragraph("")

    for i, row in enumerate(questions):
        content = row['Content']
        content_clean = re.sub(r'【|】', '', content)
        p = doc.add_paragraph(style='List Number')
        run = p.add_run(f"{content_clean}")
        run.font.size = Pt(18)

    bio = io.BytesIO()
    doc.save(bio)
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

# ============================================================
# MODE A: 按學校預覽下載
# ============================================================
if send_mode == "📄 按學校預覽下載":
    schools = edited_df['School'].unique() if not edited_df.empty else []

    if len(schools) == 0:
        st.info("Select at least one question above to begin.")
    else:
        selected_school = st.selectbox("Select School to Preview/Download", schools)
        school_data = edited_df[edited_df['School'] == selected_school]

        col1, col2 = st.columns([1, 2])

        # Shuffle questions ONCE for consistency across all documents
        original_questions = school_data.to_dict('records')
        cache_key = f"school_{selected_school}_{selected_level}"
        shuffled_questions = get_shuffled_questions(original_questions, cache_key)
        
        # Generate all documents with the SAME shuffled question order
        pdf_buffer = create_pdf(selected_school, selected_level, shuffled_questions)
        pdf_bytes = pdf_buffer.getvalue()
        
        # Teacher answer PDF uses same order as student PDF
        answer_pdf_buffer = create_answer_pdf(selected_school, selected_level, shuffled_questions)
        answer_pdf_bytes = answer_pdf_buffer.getvalue()
        
        # Word document uses same order
        docx_buffer = create_docx(selected_school, selected_level, shuffled_questions)
        docx_bytes = docx_buffer.getvalue()

        with col1:
            st.write(f"**School:** {selected_school}")
            st.write(f"**Level:** {selected_level}")
            st.write(f"**Questions:** {len(school_data)}")

            st.download_button(
                label=f"📥 Download {selected_school}_{selected_level}.pdf",
                data=pdf_bytes,
                file_name=f"{selected_school}_{selected_level}_Review_{datetime.date.today()}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key=f"dl_{selected_school}_{selected_level}"
            )

            st.download_button(
                label=f"📄 下載 Word 檔（可編輯）",
                data=docx_bytes,
                file_name=f"{selected_school}_{selected_level}_Review_{datetime.date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key=f"dl_docx_{selected_school}_{selected_level}"
            )
            
            # --- FEATURE 2: Teacher Answer PDF Download Button ---
            st.download_button(
                label=f"📥 下載教師版答案 PDF",
                data=answer_pdf_bytes,
                file_name=f"{selected_school}_{selected_level}_教師版答案_{datetime.date.today()}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key=f"dl_answer_{selected_school}_{selected_level}"
            )
            
            st.info("💡 Fix typos in Google Sheet, then click 'Refresh Data' above.")

        with col2:
            st.write("🔍 **100% Accurate Preview**")
            display_pdf_as_images(pdf_bytes)

# ============================================================
# MODE B: 按學生寄送
# ============================================================
else:
    st.subheader("👨‍👩‍👧 學生配對結果")

    if student_df.empty:
        st.error("❌ 無法讀取「學生資料」工作表，請確認工作表名稱正確。")
        st.stop()

    required_cols = ['學校', '年級', '狀態', '學生姓名', '學生編號', '家長 Email']
    missing_cols = [c for c in required_cols if c not in student_df.columns]
    if missing_cols:
        st.error(f"❌ 「學生資料」工作表缺少以下欄位：{missing_cols}")
        st.write("現有欄位：", student_df.columns.tolist())
        st.stop()

    active_students = student_df[student_df['狀態'] == 'Y']

    if active_students.empty:
        st.warning("⚠️ 「學生資料」中沒有「狀態 = Y」的學生。請先將測試學生的狀態改為 Y。")
        st.stop()

    # 只保留有勾選 Select 的題目（避免未勾選都被配對）
    questions_df = edited_df
    if 'Select' in questions_df.columns:
        questions_df = questions_df[questions_df['Select'] == True]

    # 題目去重：優先用 standby 的 ID（最穩陣）；如果冇 ID 就用 Content
    if 'ID' in questions_df.columns:
        questions_df = questions_df.drop_duplicates(subset=['ID'])
    else:
        questions_df = questions_df.drop_duplicates(subset=['School', 'Level', 'Content'])

    merged = active_students.merge(
        questions_df,
        left_on=['學校', '年級'],
        right_on=['School', 'Level'],
        how='inner'
    )

    if merged.empty:
        st.warning("⚠️ 沒有符合條件的配對。請確認：")
        st.write("1. `standby` 表有 Status = Ready/Waiting 的題目")
        st.write("2. `學生資料` 表有 狀態 = Y 的學生")
        st.write("3. 學校名稱和年級在兩張表中**完全一致**（注意空格/全半形）")

        with st.expander("🔍 查看配對資料（協助排查問題）"):
            st.write("**standby 的 School 值：**", edited_df['School'].unique().tolist())
            st.write("**standby 的 Level 值：**", edited_df['Level'].unique().tolist())
            st.write("**學生資料 的 學校 值：**", active_students['學校'].unique().tolist())
            st.write("**學生資料 的 年級 值：**", active_students['年級'].unique().tolist())
        st.stop()

    student_count = merged['學生編號'].nunique()
    st.success(f"✅ 成功配對 {student_count} 位學生（按學生編號），共 {len(merged)} 筆配對資料")

    # ✅ 每位學生一份：按「學生編號」分組
    for student_id, group in merged.groupby('學生編號'):
        # 由 group 取回真正的家長電郵（分組 key 已經唔係 email）
        parent_email = str(group['家長 Email'].iloc[0]).strip()

        student_name  = group['學生姓名'].iloc[0]
        school_name   = group['學校'].iloc[0]
        grade         = group['年級'].iloc[0]
        teacher_email = group['老師 Email'].iloc[0] if '老師 Email' in group.columns else "N/A"

        # 保險：每位學生的題目再去重一次（避免任何上游重覆）
        if 'ID' in group.columns:
            unique_group = group.drop_duplicates(subset=['ID'])
            question_count = unique_group['ID'].nunique()
        else:
            unique_group = group.drop_duplicates(subset=['Content'])
            question_count = unique_group['Content'].nunique()

        st.divider()
        col1, col2 = st.columns([1, 2])

        # Shuffle questions ONCE for consistency across all documents for this student
        original_questions = unique_group.to_dict('records')
        cache_key = f"student_{student_id}_{grade}"
        shuffled_questions = get_shuffled_questions(original_questions, cache_key)
        
        # Generate all documents with the SAME shuffled question order
        pdf_buffer = create_pdf(school_name, grade, shuffled_questions, student_name=student_name)
        pdf_bytes  = pdf_buffer.getvalue()
        
        # Teacher answer PDF uses same order as student PDF
        answer_pdf_buffer = create_answer_pdf(school_name, grade, shuffled_questions, student_name=student_name)
        answer_pdf_bytes = answer_pdf_buffer.getvalue()
        
        # Word document uses same order
        docx_buffer = create_docx(school_name, grade, shuffled_questions, student_name=student_name)
        docx_bytes  = docx_buffer.getvalue()

        with col1:
            st.write(f"**👤 學生：** {student_name}")
            st.write(f"**🆔 學生編號：** {student_id}")
            st.write(f"**🏫 學校：** {school_name} ({grade})")
            st.write(f"**📧 家長：** {parent_email}")
            st.write(f"**👩‍🏫 老師：** {teacher_email}")
            st.write(f"**📝 題目數：** {question_count} 題")

            # ✅ key 用 student_id，避免同一測試 email 撞 key
            st.download_button(
                label=f"📥 下載 {student_name} PDF",
                data=pdf_bytes,
                file_name=f"{student_name}_{grade}_Review_{datetime.date.today()}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key=f"dl_{student_id}"
            )
            st.download_button(
                label=f"📄 下載 {student_name} Word 檔（可編輯）",
                data=docx_bytes,
                file_name=f"{student_name}_{grade}_Review_{datetime.date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key=f"dl_docx_{student_id}"
            )
            
            # --- FEATURE 2: Teacher Answer PDF Download Button ---
            st.download_button(
                label=f"📥 下載教師版答案 PDF",
                data=answer_pdf_bytes,
                file_name=f"{student_name}_{grade}_教師版答案_{datetime.date.today()}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key=f"dl_answer_{student_id}"
            )

            if st.button(
                f"📧 寄送給 {student_name} 家長",
                key=f"send_{student_id}",
                use_container_width=True
            ):
                with st.spinner(f"正在寄送給 {parent_email}..."):
                    success, msg = send_email_with_pdf(
                        parent_email, student_name, school_name, grade, pdf_bytes, cc_email=teacher_email
                    )
                    if success:
                        st.success("✅ 已成功寄送！")
                    else:
                        st.error(f"❌ 發送失敗: {msg}")
                        st.code(msg)

        with col2:
            st.write("🔍 **100% 準確預覽**")
            display_pdf_as_images(pdf_bytes)
