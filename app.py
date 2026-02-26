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
    if cache_key in st.session_state.shuffled_cache:
        return st.session_state.shuffled_cache[cache_key]
    questions_list = list(questions)
    random.seed(int(time.time() * 1000))
    random.shuffle(questions_list)
    st.session_state.shuffled_cache[cache_key] = questions_list
    return questions_list

# --- 4. GENERATE PDF FUNCTION (Student Version) ---
def draw_text_with_underline_wrapped(c, x, y, text, font_name, font_size, max_width, underline_offset=2, line_height=18):
    """
    Draws text with <u>underline</u> tags, wrapping lines automatically.
    Returns the new y position after drawing.
    """
    import re
    from reportlab.pdfbase import pdfmetrics

    # Split text into normal parts and underlined parts
    parts = re.split(r'(<u>.*?</u>)', text)
    tokens = []
    for p in parts:
        if not p:
            continue
        if p.startswith("<u>") and p.endswith("</u>"):
            tokens.append(p)          # keep underlined part as one token
        else:
            # split normal text into individual characters (so we can wrap anywhere)
            tokens.extend(list(p))

    def measure(tok):
        if tok.startswith("<u>") and tok.endswith("</u>"):
            inner = tok[3:-4]
            return pdfmetrics.stringWidth(inner, font_name, font_size)
        else:
            return pdfmetrics.stringWidth(tok, font_name, font_size)

    def draw_line(parts_to_draw, draw_x, draw_y):
        cx = draw_x
        for tp in parts_to_draw:
            if tp.startswith("<u>") and tp.endswith("</u>"):
                inner = tp[3:-4]
                c.setFont(font_name, font_size)
                c.drawString(cx, draw_y, inner)
                w = pdfmetrics.stringWidth(inner, font_name, font_size)
                c.line(cx, draw_y - underline_offset, cx + w, draw_y - underline_offset)
                cx += w
            else:
                c.setFont(font_name, font_size)
                c.drawString(cx, draw_y, tp)
                cx += pdfmetrics.stringWidth(tp, font_name, font_size)

    cur_y = y
    line_buf = []
    line_width = 0
    for tok in tokens:
        tok_w = measure(tok)
        if line_width + tok_w > max_width and line_buf:
            draw_line(line_buf, x, cur_y)
            cur_y -= line_height
            line_buf = [tok]
            line_width = tok_w
        else:
            line_buf.append(tok)
            line_width += tok_w
    if line_buf:
        draw_line(line_buf, x, cur_y)
        cur_y -= line_height

    # add a small gap after the paragraph
    cur_y -= 12
    return cur_y
def create_pdf(school_name, level, questions, student_name=None, original_questions=None):
    """
    Creates a student worksheet PDF using direct canvas drawing.
    Returns a BytesIO object.
    """
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfbase import pdfmetrics
    import io
    import re
    import datetime

    bio = io.BytesIO()
    c = canvas.Canvas(bio, pagesize=letter)
    page_width, page_height = letter

    font_name = CHINESE_FONT if CHINESE_FONT else 'Helvetica'

    left_margin = 60
    top_margin = page_height - 60
    cur_y = top_margin

    # ---- Title ----
    c.setFont(font_name, 22)
    if student_name:
        title = f"{school_name} ({level}) - {student_name} - 校本填充工作紙"
    else:
        title = f"{school_name} ({level}) - 校本填充工作紙"
    c.drawString(left_margin, cur_y, title)
    cur_y -= 30

    # ---- Date ----
    c.setFont(font_name, 18)
    date_str = f"日期: {datetime.date.today() + datetime.timedelta(days=1)}"
    c.drawString(left_margin, cur_y, date_str)
    cur_y -= 30

    # ---- Questions ----
    max_text_width = page_width - left_margin - 40
    line_height = 26
    body_font_size = 18

    for idx, row in enumerate(questions):
        content = row['Content']

        # 1. Proper noun marks: 【】text【】 -> <u>text</u>
        content = re.sub(r'【】(.*?)【】', r'<u>\1</u>', content)

        # 2. Fill-in-the-blanks: 【word】 -> underlined blank of appropriate length (fullwidth underscores)
        def replace_blank(match):
            word = match.group(1)
            blank_length = max(len(word) * 2, 4)          # same logic as original
            blank = '＿' * blank_length                    # fullwidth underscore
            return f'<u>{blank}</u>'
        content = re.sub(r'【([^】]+)】', replace_blank, content)

        # 3. Fix for underlines at start of paragraph (zero‑width space)
        if content.strip().startswith('<u>'):
            content = '\u200B' + content                   # actual zero‑width space, not HTML entity

        # New page if needed
        if cur_y - line_height < 60:
            c.showPage()
            cur_y = page_height - 60

        # Draw question number
        c.setFont(font_name, body_font_size)
        c.drawString(left_margin, cur_y, f"{idx+1}.")

        # Draw question content with underlines and wrapping
        cur_y = draw_text_with_underline_wrapped(
            c,
            left_margin + 30,
            cur_y,
            content,
            font_name,
            body_font_size,
            max_text_width,
            underline_offset=2,
            line_height=line_height
        )

    # ---- Word list page ----
    if original_questions is not None:
        words = [row.get('Word', '').strip() for row in original_questions]
    else:
        words = [row.get('Word', '').strip() for row in questions]
    unique_words = list(dict.fromkeys([w for w in words if w]))

    if unique_words:
        c.showPage()
        cur_y = page_height - 60

        c.setFont(font_name, 20)
        c.drawString(left_margin, cur_y, "詞語表")
        cur_y -= 30

        c.setFont(font_name, 18)
        # Two columns
        col_width = 200
        x1 = left_margin
        x2 = left_margin + col_width + 20
        col_x = x1
        for i, word in enumerate(unique_words):
            if cur_y < 60:
                c.showPage()
                cur_y = page_height - 60
                c.setFont(font_name, 20)
                c.drawString(left_margin, cur_y, "詞語表 (續)")
                cur_y -= 30
                c.setFont(font_name, 18)
            c.drawString(col_x, cur_y, f"{i+1}. {word}")
            # alternate columns
            if (i+1) % 2 == 0:
                cur_y -= 30
                col_x = x1
            else:
                col_x = x2

    c.save()
    bio.seek(0)
    return bio

# --- FEATURE 2: Teacher Answer PDF Function ---
def create_answer_pdf(school_name, level, questions, student_name=None):
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
        leftIndent=0,
        firstLineIndent=0
    )

    if student_name:
        title_text = f"<b>{school_name} ({level}) - {student_name} - 校本填充工作紙</b>"
    else:
        title_text = f"<b>{school_name} ({level}) - 校本填充工作紙</b>"

    story.append(Paragraph(title_text, title_style))
    story.append(Paragraph("<b>教師版答案 (Answer Key)</b>", subtitle_style))
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph(f"日期: {datetime.date.today() + datetime.timedelta(days=1)}", normal_style))
    story.append(Spacer(1, 0.3*inch))

    for i, row in enumerate(questions):
        content = row['Content']
        answer = row.get('Word', '')

        if answer:
            answer_html = f'<font color="red"><b>【{answer}】</b></font>'
            # Replace underscores with answer
            content = re.sub(r'_{2,}|＿{2,}', answer_html, content)
            # Handle 【】text【】 proper noun marks
            content = re.sub(r'【】(.*?)【】', r'<font color="red"><b>【\1】</b></font>', content)
            # Handle 【answer】 blanks — fixed: added capture group ()
            content = re.sub(r'【([^】]+)】', r'<font color="red"><b>【\1】</b></font>', content)
        else:
            # No Word answer - handle bracket patterns only
            content = re.sub(r'【】(.*?)【】', r'<font color="red"><b>【\1】</b></font>', content)
            # Fixed: added capture group ()
            content = re.sub(r'【([^】]+)】', r'<font color="red"><b>【\1】</b></font>', content)

        # 解決開頭紅字標籤失效問題
        if content.strip().startswith('<font'):
            content = '&#8203;' + content

        num_para = Paragraph(f"<b>{i+1}.</b>", normal_style)
        content_para = Paragraph(content, normal_style)

        t = Table([[num_para, content_para]], colWidths=[0.5*inch, 6.7*inch])
        t.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        story.append(t)
        story.append(Spacer(1, 0.15*inch))

    doc.build(story)
    bio.seek(0)
    return bio

# --- SendGrid Email Function ---
def send_email_with_pdf(to_email, student_name, school_name, grade, pdf_bytes, cc_email=None):
    try:
        sg_config = st.secrets["sendgrid"]

        recipient = str(to_email).strip()
        if not re.match(r'^[\w\.-]+@[\w\.-]+\.\w+$', recipient):
            return False, f"無效的家長電郵格式: '{recipient}'"

        from_email_obj = Email(sg_config["from_email"], sg_config.get("from_name", ""))
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

        if cc_email:
            cc_clean = str(cc_email).strip().lower()
            if cc_clean not in ["n/a", "nan", "", "none"] and "@" in cc_clean and cc_clean != recipient.lower():
                message.add_cc(cc_clean)

        encoded_pdf = base64.b64encode(pdf_bytes).decode()
        attachment = Attachment(
            FileContent(encoded_pdf),
            FileName(f"{safe_name}_Worksheet.pdf"),
            FileType('application/pdf'),
            Disposition('attachment')
        )
        message.add_attachment(attachment)

        sg = SendGridAPIClient(sg_config["api_key"])
        response = sg.send(message)

        if 200 <= response.status_code < 300:
            return True, "發送成功"
        else:
            return False, f"SendGrid Error: {response.status_code}"

    except HTTPError as e:
        try:
            return False, e.body.decode("utf-8")
        except Exception:
            return False, str(e)
    except Exception as e:
        return False, str(e)

def create_docx(school_name, level, questions, student_name=None):
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

        original_questions = school_data.to_dict('records')
        cache_key = f"school_{selected_school}_{selected_level}"
        shuffled_questions = get_shuffled_questions(original_questions, cache_key)

        pdf_buffer = create_pdf(selected_school, selected_level, shuffled_questions, original_questions=original_questions)
        pdf_bytes = pdf_buffer.getvalue()

        answer_pdf_buffer = create_answer_pdf(selected_school, selected_level, shuffled_questions)
        answer_pdf_bytes = answer_pdf_buffer.getvalue()

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

    questions_df = edited_df
    if 'Select' in questions_df.columns:
        questions_df = questions_df[questions_df['Select'] == True]

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

    for student_id, group in merged.groupby('學生編號'):
        parent_email = str(group['家長 Email'].iloc[0]).strip()
        student_name  = group['學生姓名'].iloc[0]
        school_name   = group['學校'].iloc[0]
        grade         = group['年級'].iloc[0]
        teacher_email = group['老師 Email'].iloc[0] if '老師 Email' in group.columns else "N/A"

        if 'ID' in group.columns:
            unique_group = group.drop_duplicates(subset=['ID'])
            question_count = unique_group['ID'].nunique()
        else:
            unique_group = group.drop_duplicates(subset=['Content'])
            question_count = unique_group['Content'].nunique()

        st.divider()
        col1, col2 = st.columns([1, 2])

        original_questions = unique_group.to_dict('records')
        cache_key = f"student_{student_id}_{grade}"
        shuffled_questions = get_shuffled_questions(original_questions, cache_key)

        pdf_buffer = create_pdf(school_name, grade, shuffled_questions, student_name=student_name, original_questions=original_questions)
        pdf_bytes  = pdf_buffer.getvalue()

        answer_pdf_buffer = create_answer_pdf(school_name, grade, shuffled_questions, student_name=student_name)
        answer_pdf_bytes = answer_pdf_buffer.getvalue()

        docx_buffer = create_docx(school_name, grade, shuffled_questions, student_name=student_name)
        docx_bytes  = docx_buffer.getvalue()

        with col1:
            st.write(f"**👤 學生：** {student_name}")
            st.write(f"**🆔 學生編號：** {student_id}")
            st.write(f"**🏫 學校：** {school_name} ({grade})")
            st.write(f"**📧 家長：** {parent_email}")
            st.write(f"**👩‍🏫 老師：** {teacher_email}")
            st.write(f"**📝 題目數：** {question_count} 題")

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
