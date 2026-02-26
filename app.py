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

if 'shuffled_cache' not in st.session_state:
    st.session_state.shuffled_cache = {}

# --- ReportLab Import & Font Registration ---
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

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

# --- Connect to Google Cloud ---
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

if not student_df.empty:
    student_df.columns = [c.strip() for c in student_df.columns]
    for col in student_df.columns:
        if student_df[col].dtype == object:
            student_df[col] = student_df[col].astype(str).str.strip()

for col in df.columns:
    if df[col].dtype == object:
        df[col] = df[col].astype(str).str.strip()

# --- Sidebar ---
with st.sidebar:
    st.header("🎓 篩選年級")
    available_levels = sorted(df["Level"].astype(str).str.strip().unique().tolist())
    selected_level = st.radio("選擇年級", available_levels, index=0)
    st.divider()
    st.info(f"目前顯示：**{selected_level}** 的題目")
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

# --- HELPER: Shuffle questions once per session ---
def get_shuffled_questions(questions, cache_key):
    if cache_key in st.session_state.shuffled_cache:
        return st.session_state.shuffled_cache[cache_key]
    questions_list = list(questions)
    random.seed(int(time.time() * 1000))
    random.shuffle(questions_list)
    st.session_state.shuffled_cache[cache_key] = questions_list
    return questions_list

# ============================================================
# --- PDF LAYOUT CONSTANTS (shared by both PDF functions) ---
# ============================================================
PDF_LEFT_NUM    = 60
PDF_TEXT_START  = PDF_LEFT_NUM + 30
PDF_RIGHT_MARGIN = 40
PDF_LINE_HEIGHT  = 26
PDF_FONT_SIZE    = 18

def _get_max_width():
    page_width, _ = letter
    return page_width - PDF_RIGHT_MARGIN - PDF_TEXT_START

# ============================================================
# --- SHARED HELPER: draw text with <u> underline tags ---
# ============================================================
def draw_text_with_underline_wrapped(c, x, y, text, font_name, font_size, max_width,
                                      underline_offset=2, line_height=18):
    """
    Draws text supporting <u>...</u> underline tags with automatic line wrapping.
    Returns new y position.
    """
    parts = re.split(r'(<u>.*?</u>)', text)
    tokens = []
    for p in parts:
        if not p:
            continue
        if p.startswith("<u>") and p.endswith("</u>"):
            tokens.append(p)
        else:
            tokens.extend(list(p))

    def measure(tok):
        inner = tok[3:-4] if tok.startswith("<u>") else tok
        return pdfmetrics.stringWidth(inner, font_name, font_size)

    def draw_line(line_tokens, draw_x, draw_y):
        cx = draw_x
        for tp in line_tokens:
            c.setFont(font_name, font_size)
            if tp.startswith("<u>") and tp.endswith("</u>"):
                inner = tp[3:-4]
                c.drawString(cx, draw_y, inner)
                w = pdfmetrics.stringWidth(inner, font_name, font_size)
                c.line(cx, draw_y - underline_offset, cx + w, draw_y - underline_offset)
                cx += w
            else:
                c.drawString(cx, draw_y, tp)
                cx += pdfmetrics.stringWidth(tp, font_name, font_size)

    cur_y = y
    line_buf, line_width = [], 0
    for tok in tokens:
        tok_w = measure(tok)
        if line_width + tok_w > max_width and line_buf:
            draw_line(line_buf, x, cur_y)
            cur_y -= line_height
            line_buf, line_width = [tok], tok_w
        else:
            line_buf.append(tok)
            line_width += tok_w
    if line_buf:
        draw_line(line_buf, x, cur_y)
        cur_y -= line_height
    cur_y -= 12
    return cur_y

# ============================================================
# --- SHARED HELPER: draw text with <red> colour tags ---
# ============================================================
def _draw_answer_line_wrapped(c, x, y, text, font_name, font_size, max_width,
                               underline_offset=2, line_height=18):
    """
    Draws text supporting <red>...</red> colour tags with automatic line wrapping.
    Returns new y position.
    """
    from reportlab.lib.colors import red as RED

    parts = re.split(r'(<red>.*?</red>)', text)
    tokens = []
    for p in parts:
        if not p:
            continue
        if p.startswith('<red>') and p.endswith('</red>'):
            tokens.append(p)
        else:
            tokens.extend(list(p))

    def measure(tok):
        inner = tok[5:-6] if tok.startswith('<red>') else tok
        return pdfmetrics.stringWidth(inner, font_name, font_size)

    def draw_line(line_tokens, draw_x, draw_y):
        cx = draw_x
        for tp in line_tokens:
            c.setFont(font_name, font_size)
            if tp.startswith('<red>') and tp.endswith('</red>'):
                inner = tp[5:-6]
                c.setFillColor(RED)
                c.drawString(cx, draw_y, inner)
                c.setFillColorRGB(0, 0, 0)
                cx += pdfmetrics.stringWidth(inner, font_name, font_size)
            else:
                c.setFillColorRGB(0, 0, 0)
                c.drawString(cx, draw_y, tp)
                cx += pdfmetrics.stringWidth(tp, font_name, font_size)

    cur_y = y
    line_buf, line_width = [], 0
    for tok in tokens:
        tok_w = measure(tok)
        if line_width + tok_w > max_width and line_buf:
            draw_line(line_buf, x, cur_y)
            cur_y -= line_height
            line_buf, line_width = [tok], tok_w
        else:
            line_buf.append(tok)
            line_width += tok_w
    if line_buf:
        draw_line(line_buf, x, cur_y)
        cur_y -= line_height
    cur_y -= 12
    return cur_y

# ============================================================
# --- SHARED HELPER: draw word list page ---
# ============================================================
def _draw_word_list_page(c, words, font_name, title="詞語表", word_color=None):
    """
    Draws a word list on a new page in two columns.
    word_color: reportlab color object or None (black).
    """
    from reportlab.lib.colors import red as RED
    _, page_height = letter

    unique_words = list(dict.fromkeys([w for w in words if w]))
    if not unique_words:
        return

    c.showPage()
    cur_y = page_height - 60
    col_width = 200
    x1 = PDF_LEFT_NUM
    x2 = PDF_LEFT_NUM + col_width + 20
    col_x = x1

    c.setFont(font_name, 20)
    c.setFillColorRGB(0, 0, 0)
    c.drawString(PDF_LEFT_NUM, cur_y, title)
    cur_y -= 30

    for i, word in enumerate(unique_words):
        if cur_y < 60:
            c.showPage()
            cur_y = page_height - 60
            c.setFont(font_name, 20)
            c.setFillColorRGB(0, 0, 0)
            c.drawString(PDF_LEFT_NUM, cur_y, f"{title} (續)")
            cur_y -= 30

        c.setFont(font_name, PDF_FONT_SIZE)
        if word_color:
            c.setFillColor(word_color)
        else:
            c.setFillColorRGB(0, 0, 0)
        c.drawString(col_x, cur_y, f"{i+1}. {word}")
        c.setFillColorRGB(0, 0, 0)

        if (i + 1) % 2 == 0:
            cur_y -= 30
            col_x = x1
        else:
            col_x = x2

# ============================================================
# --- 4a. STUDENT WORKSHEET PDF ---
# ============================================================
def create_pdf(school_name, level, questions, student_name=None, original_questions=None):
    """
    Student worksheet: blanks shown as underlined spaces.
    Word list appended at the end.
    """
    from reportlab.pdfgen import canvas as rl_canvas

    bio = io.BytesIO()
    c = rl_canvas.Canvas(bio, pagesize=letter)
    _, page_height = letter
    font_name = CHINESE_FONT if CHINESE_FONT else 'Helvetica'
    max_width = _get_max_width()

    cur_y = page_height - 60

    # Title
    c.setFont(font_name, 22)
    title = f"{school_name} ({level}) - {student_name} - 校本填充工作紙" if student_name \
            else f"{school_name} ({level}) - 校本填充工作紙"
    c.drawString(PDF_LEFT_NUM, cur_y, title)
    cur_y -= 30

    # Date
    c.setFont(font_name, PDF_FONT_SIZE)
    c.drawString(PDF_LEFT_NUM, cur_y, f"日期: {datetime.date.today() + datetime.timedelta(days=1)}")
    cur_y -= 30

    # Questions
    def replace_blank(match):
        word = match.group(1)
        blank_spaces = ' ' * max(len(word) * 2, 4)
        return f'<u>{blank_spaces}</u>'

    for idx, row in enumerate(questions):
        content = row['Content']
        content = re.sub(r'【】(.*?)【】', r'<u>\1</u>', content)
        content = re.sub(r'【([^】]+)】', replace_blank, content)

        if cur_y - PDF_LINE_HEIGHT < 60:
            c.showPage()
            cur_y = page_height - 60

        c.setFont(font_name, PDF_FONT_SIZE)
        c.drawString(PDF_LEFT_NUM, cur_y, f"{idx+1}.")
        cur_y = draw_text_with_underline_wrapped(
            c, PDF_TEXT_START, cur_y, content,
            font_name, PDF_FONT_SIZE, max_width,
            underline_offset=2, line_height=PDF_LINE_HEIGHT
        )

    # Word list (use original_questions order if provided)
    source = original_questions if original_questions is not None else questions
    words = [str(row.get('Word', '')).strip() for row in source]
    _draw_word_list_page(c, words, font_name, title="詞語表")

    c.save()
    bio.seek(0)
    return bio

# ============================================================
# --- 4b. TEACHER ANSWER PDF ---
# ============================================================
def create_answer_pdf(school_name, level, questions, student_name=None):
    """
    教師版：
    - 第一頁：題目原文（不做任何替換，保持原樣）
    - 最後一頁：詞語表（題目順序，紅色顯示答案）
    """
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.colors import red as RED

    bio = io.BytesIO()
    c = rl_canvas.Canvas(bio, pagesize=letter)
    page_width, page_height = letter
    font_name = CHINESE_FONT if CHINESE_FONT else 'Helvetica'

    left_m = 60
    text_x = left_m + 30
    max_w  = page_width - 40 - text_x
    line_h = 26
    cur_y  = page_height - 60

    # ── 標題 ──────────────────────────────────────────────
    c.setFont(font_name, 22)
    c.setFillColorRGB(0, 0, 0)
    title = f"{school_name} ({level}) - {student_name} - 教師版答案" if student_name \
            else f"{school_name} ({level}) - 教師版答案"
    c.drawString(left_m, cur_y, title)
    cur_y -= 40

    # ── 題目（原文，不替換任何內容）────────────────────────
    for idx, row in enumerate(questions):
        if cur_y < 80:
            c.showPage()
            cur_y = page_height - 60

        # 清理 Content：移除所有 【...】 標記，只留純文字
        content = str(row.get('Content', ''))
        content_clean = re.sub(r'【】(.*?)【】', r'\1', content)  # 專有名詞標記 → 保留文字
        content_clean = re.sub(r'【[^】]*】', '＿＿', content_clean)  # 答案空格 → 底線

        c.setFont(font_name, 18)
        c.setFillColorRGB(0, 0, 0)
        c.drawString(left_m, cur_y, f"{idx+1}.")

        # 簡單繪製（不換行，如需換行可用 draw_text_with_underline_wrapped）
        c.drawString(text_x, cur_y, content_clean[:60])  # 截斷防止溢出
        cur_y -= line_h

    # ── 詞語表（答案，紅色）────────────────────────────────
    c.showPage()
    cur_y = page_height - 60

    c.setFont(font_name, 20)
    c.setFillColorRGB(0, 0, 0)
    c.drawString(left_m, cur_y, "詞語表（題目順序）")
    cur_y -= 40

    for idx, row in enumerate(questions):
        word = str(row.get('Word', '')).strip()
        if cur_y < 60:
            c.showPage()
            cur_y = page_height - 60
            c.setFont(font_name, 20)
            c.setFillColorRGB(0, 0, 0)
            c.drawString(left_m, cur_y, "詞語表（題目順序）（續）")
            cur_y -= 40

        c.setFont(font_name, 18)
        c.setFillColorRGB(0, 0, 0)
        c.drawString(left_m, cur_y, f"{idx+1}.  ")
        c.setFillColor(RED)
        c.drawString(left_m + 45, cur_y, word)
        c.setFillColorRGB(0, 0, 0)
        cur_y -= 26

    c.save()
    bio.seek(0)
    return bio
    # 3. 詞語表 (答案版)
    c.showPage()
    cur_y = page_height - 60
    c.setFont(font_name, 20)
    c.setFillColorRGB(0, 0, 0)
    c.drawString(left_m, cur_y, "詞語表 (答案)")
    cur_y -= 40

    for idx, row in enumerate(questions):
        word = str(row.get('Word', '')).strip()
        if cur_y < 60:
            c.showPage()
            cur_y = page_height - 60
        
        c.setFont(font_name, 18)
        c.setFillColorRGB(0, 0, 0)
        c.drawString(left_m, cur_y, f"{idx+1}. ")
        c.setFillColor(RED)
        c.drawString(left_m + 40, cur_y, word)
        cur_y -= 25

    c.save()
    bio.seek(0)
    return bio

# ============================================================
# --- SendGrid Email ---
# ============================================================
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

# ============================================================
# --- DOCX Export ---
# ============================================================
def create_docx(school_name, level, questions, student_name=None):
    doc = Document()
    title_text = f"{school_name} ({level}) - {student_name} - 校本填充工作紙" if student_name \
                 else f"{school_name} ({level}) - 校本填充工作紙"

    title = doc.add_heading(title_text, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    date_para = doc.add_paragraph(f"日期: {datetime.date.today() + datetime.timedelta(days=1)}")
    date_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    doc.add_paragraph("")

    for i, row in enumerate(questions):
        content = re.sub(r'【|】', '', row['Content'])
        p = doc.add_paragraph(style='List Number')
        run = p.add_run(content)
        run.font.size = Pt(18)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ============================================================
# --- Helper: Render PDF as images for preview ---
# ============================================================
def display_pdf_as_images(pdf_bytes):
    try:
        images = convert_from_bytes(pdf_bytes, dpi=150)
        for i, image in enumerate(images):
            st.image(image, caption=f"Page {i+1}", use_container_width=True)
    except Exception as e:
        st.error(f"Could not render preview: {e}")
        st.info("You can still download the PDF using the button on the left.")

# ============================================================
# --- 5. PREVIEW & DOWNLOAD INTERFACE ---
# ============================================================
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

        pdf_bytes        = create_pdf(selected_school, selected_level, shuffled_questions, original_questions=original_questions).getvalue()
        answer_pdf_bytes = create_answer_pdf(selected_school, selected_level, shuffled_questions).getvalue()
        docx_bytes       = create_docx(selected_school, selected_level, shuffled_questions).getvalue()

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
                label="📄 下載 Word 檔（可編輯）",
                data=docx_bytes,
                file_name=f"{selected_school}_{selected_level}_Review_{datetime.date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key=f"dl_docx_{selected_school}_{selected_level}"
            )
            st.download_button(
                label="📥 下載教師版答案 PDF",
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
        parent_email  = str(group['家長 Email'].iloc[0]).strip()
        student_name  = group['學生姓名'].iloc[0]
        school_name   = group['學校'].iloc[0]
        grade         = group['年級'].iloc[0]
        teacher_email = group['老師 Email'].iloc[0] if '老師 Email' in group.columns else "N/A"

        if 'ID' in group.columns:
            unique_group   = group.drop_duplicates(subset=['ID'])
            question_count = unique_group['ID'].nunique()
        else:
            unique_group   = group.drop_duplicates(subset=['Content'])
            question_count = unique_group['Content'].nunique()

        st.divider()
        col1, col2 = st.columns([1, 2])

        original_questions = unique_group.to_dict('records')
        cache_key          = f"student_{student_id}_{grade}"
        shuffled_questions = get_shuffled_questions(original_questions, cache_key)

        pdf_bytes        = create_pdf(school_name, grade, shuffled_questions, student_name=student_name, original_questions=original_questions).getvalue()
        answer_pdf_bytes = create_answer_pdf(school_name, grade, shuffled_questions, student_name=student_name).getvalue()
        docx_bytes       = create_docx(school_name, grade, shuffled_questions, student_name=student_name).getvalue()

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
                label="📥 下載教師版答案 PDF",
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
