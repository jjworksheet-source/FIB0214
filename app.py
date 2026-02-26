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

# ============================================================
# --- 1. SETUP & CONNECTION ---
# ============================================================
st.set_page_config(page_title="Worksheet Generator", page_icon="📝", layout="wide")
st.title("📝 校本填充工作紙生成器")

# Session state init
if 'shuffled_cache' not in st.session_state:
    st.session_state.shuffled_cache = {}
# final_pool: { "學校||年級": [ {Word, Content, School, Level, ...}, ... ] }
if 'final_pool' not in st.session_state:
    st.session_state.final_pool = {}
# ai_choices: { "學校||年級||詞語||idx": chosen_sentence_text }
if 'ai_choices' not in st.session_state:
    st.session_state.ai_choices = {}
# confirmed_batches: set of "學校||年級" that have been confirmed
if 'confirmed_batches' not in st.session_state:
    st.session_state.confirmed_batches = set()

# ============================================================
# --- ReportLab Import & Font Registration ---
# ============================================================
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
                break
            except Exception:
                continue

    if not CHINESE_FONT:
        st.error("❌ Chinese font not found. Please ensure Kai.ttf is in your GitHub repository.")

except ImportError:
    st.error("❌ reportlab not found. Please add 'reportlab' to your requirements.txt")
    st.stop()

# ============================================================
# --- Connect to Google Cloud ---
# ============================================================
try:
    key_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(
        key_dict,
        scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    )
    client = gspread.authorize(creds)
    SHEET_ID = st.secrets["app_config"]["spreadsheet_id"]
except Exception as e:
    st.error(f"❌ Connection Error: {e}")
    st.stop()

# ============================================================
# --- 2. DATA LOADING ---
# ============================================================
@st.cache_data(ttl=60)
def load_review():
    """
    讀取 Review 工作表。
    欄位：Timestamp, 學校, 年級, 詞語, 句子, 來源（可選）, 狀態（可選）
    """
    try:
        sh = client.open_by_key(SHEET_ID)
        worksheet = sh.worksheet("Review")
        data = worksheet.get_all_records()
        df_r = pd.DataFrame(data)
        df_r.columns = [c.strip() for c in df_r.columns]
        for col in df_r.columns:
            if df_r[col].dtype == object:
                df_r[col] = df_r[col].astype(str).str.strip()
        return df_r
    except Exception as e:
        st.error(f"Error reading Review sheet: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=60)
def load_students():
    try:
        sh = client.open_by_key(SHEET_ID)
        worksheet = sh.worksheet("學生資料")
        data = worksheet.get_all_records()
        df_s = pd.DataFrame(data)
        df_s.columns = [c.strip() for c in df_s.columns]
        for col in df_s.columns:
            if df_s[col].dtype == object:
                df_s[col] = df_s[col].astype(str).str.strip()
        return df_s
    except Exception as e:
        st.error(f"Error reading 學生資料 sheet: {e}")
        return pd.DataFrame()

def build_review_groups(review_df):
    """
    整理 Review 表成：
    {
      "學校||年級": {
        "詞語A": {
          "original": "句子" or None,
          "ai": ["AI句1", "AI句2", ...],
          "row_keys": ["學校||年級||詞語A||idx", ...]
        }, ...
      }
    }
    只保留有 AI 句子（🟨 開頭）的詞語。
    """
    groups = {}
    if review_df.empty or '句子' not in review_df.columns:
        return groups

    for idx, row in review_df.iterrows():
        school   = str(row.get('學校', '')).strip()
        level    = str(row.get('年級', '')).strip()
        word     = str(row.get('詞語', '')).strip()
        sentence = str(row.get('句子', '')).strip()
        content  = str(row.get('Content', '')).strip()

        if not school or not level or not word or not sentence:
            continue

        batch_key = f"{school}||{level}"
        if batch_key not in groups:
            groups[batch_key] = {}
        if word not in groups[batch_key]:
            groups[batch_key][word] = {
                'original': None,
                'ai': [],
                'row_keys': [],
                'content': content   # store original Content for PDF
            }

        is_ai = sentence.startswith('🟨')
        clean_sentence = sentence.lstrip('🟨').strip()

        if is_ai:
            groups[batch_key][word]['ai'].append(clean_sentence)
            groups[batch_key][word]['row_keys'].append(f"{batch_key}||{word}||{idx}")
        else:
            groups[batch_key][word]['original'] = clean_sentence
            if not groups[batch_key][word]['content']:
                groups[batch_key][word]['content'] = clean_sentence

    # Only keep words that have AI sentences
    filtered = {}
    for batch_key, words in groups.items():
        ai_words = {w: d for w, d in words.items() if d['ai']}
        if ai_words:
            filtered[batch_key] = ai_words
    return filtered

def build_final_pool_from_review(review_df):
    """
    Build final_pool directly from Review table (non-AI rows = original sentences).
    Returns { "學校||年級": [ {Word, Content, School, Level}, ... ] }
    """
    pool = {}
    if review_df.empty:
        return pool

    for idx, row in review_df.iterrows():
        school  = str(row.get('學校', '')).strip()
        level   = str(row.get('年級', '')).strip()
        word    = str(row.get('詞語', '')).strip()
        sentence = str(row.get('句子', '')).strip()
        content  = str(row.get('Content', '')).strip()

        if not school or not level or not word:
            continue

        # Skip AI candidate rows (🟨) — only keep original rows here
        if sentence.startswith('🟨'):
            continue

        batch_key = f"{school}||{level}"
        if batch_key not in pool:
            pool[batch_key] = []

        pool[batch_key].append({
            'Word': word,
            'Content': content if content else sentence,
            'School': school,
            'Level': level,
        })

    return pool

# ============================================================
# --- 3. SIDEBAR ---
# ============================================================
student_df = load_students()
review_df  = load_review()
review_groups = build_review_groups(review_df)

with st.sidebar:
    st.header("⚙️ 控制面板")

    col_r, col_s = st.columns(2)
    with col_r:
        if st.button("🔄 更新資料", use_container_width=True):
            load_review.clear()
            load_students.clear()
            st.session_state.final_pool = {}
            st.session_state.ai_choices = {}
            st.session_state.confirmed_batches = set()
            st.session_state.shuffled_cache = {}
            st.rerun()
    with col_s:
        if st.button("🔀 打亂題目", use_container_width=True):
            st.session_state.shuffled_cache = {}
            st.rerun()

    st.divider()

    # Level selector — derived from Review table
    all_levels = []
    if not review_df.empty and '年級' in review_df.columns:
        all_levels = sorted(review_df['年級'].astype(str).str.strip().unique().tolist())
    if not all_levels:
        all_levels = ["P1"]

    st.subheader("🎓 年級")
    selected_level = st.radio("選擇年級", all_levels, index=0, label_visibility="collapsed")

    st.divider()

    st.subheader("📬 模式")
    send_mode = st.radio(
        "選擇模式",
        ["🤖 AI 句子審核", "📄 按學校預覽下載", "👨‍👩‍👧 按學生寄送"],
        index=0,
        label_visibility="collapsed"
    )

    st.divider()

    # Stats dashboard
    st.subheader("📊 資料概覽")
    level_batches = [k for k in review_groups if k.endswith(f"||{selected_level}")]
    total_words = sum(len(v) for k, v in review_groups.items() if k.endswith(f"||{selected_level}"))
    confirmed_count = len([k for k in st.session_state.confirmed_batches if k.endswith(f"||{selected_level}")])
    pool_count = sum(len(v) for k, v in st.session_state.final_pool.items() if k.endswith(f"||{selected_level}"))

    st.metric(f"{selected_level} 待審核批次", len(level_batches))
    st.metric(f"{selected_level} 待審核詞語", total_words)
    st.metric(f"{selected_level} 已確認批次", confirmed_count)
    st.metric("題庫已鎖定題目", pool_count)

    if not student_df.empty and '狀態' in student_df.columns:
        active_count = (student_df['狀態'] == 'Y').sum()
        st.metric("啟用學生數", int(active_count))

# ============================================================
# --- PDF LAYOUT CONSTANTS ---
# ============================================================
PDF_LEFT_NUM     = 60
PDF_TEXT_START   = PDF_LEFT_NUM + 30
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
# --- HELPER: Shuffle questions once per session ---
# ============================================================
def get_shuffled_questions(questions, cache_key):
    if cache_key in st.session_state.shuffled_cache:
        return st.session_state.shuffled_cache[cache_key]
    questions_list = list(questions)
    random.seed(int(time.time() * 1000))
    random.shuffle(questions_list)
    st.session_state.shuffled_cache[cache_key] = questions_list
    return questions_list

# ============================================================
# --- 4a. STUDENT WORKSHEET PDF ---
# ============================================================
def create_pdf(school_name, level, questions, student_name=None, original_questions=None):
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
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.colors import red as RED

    bio = io.BytesIO()
    c = rl_canvas.Canvas(bio, pagesize=letter)
    _, page_height = letter
    font_name = CHINESE_FONT if CHINESE_FONT else 'Helvetica'
    max_width = _get_max_width()
    cur_y = page_height - 60

    # Title
    c.setFont(font_name, 22)
    c.setFillColorRGB(0, 0, 0)
    title = f"{school_name} ({level}) - {student_name} - 校本填充工作紙" if student_name \
            else f"{school_name} ({level}) - 校本填充工作紙"
    c.drawString(PDF_LEFT_NUM, cur_y, title)
    cur_y -= 30

    # Answer key subtitle
    c.setFont(font_name, 16)
    c.setFillColor(RED)
    c.drawString(PDF_LEFT_NUM, cur_y, "教師版答案 (Answer Key)")
    c.setFillColorRGB(0, 0, 0)
    cur_y -= 30

    # Date
    c.setFont(font_name, PDF_FONT_SIZE)
    c.drawString(PDF_LEFT_NUM, cur_y, f"日期: {datetime.date.today() + datetime.timedelta(days=1)}")
    cur_y -= 30

    for idx, row in enumerate(questions):
        content = row['Content']
        answer  = str(row.get('Word', '')).strip()

        content = re.sub(
            r'【】(.*?)【】',
            lambda m: f'<red>【{m.group(1)}】</red>',
            content
        )
        if answer:
            content = re.sub(
                r'【([^】]+)】',
                f'<red>【{answer}】</red>',
                content
            )
        else:
            content = re.sub(
                r'【([^】]+)】',
                lambda m: f'<red>【{m.group(1)}】</red>',
                content
            )

        if cur_y - PDF_LINE_HEIGHT < 60:
            c.showPage()
            cur_y = page_height - 60

        c.setFont(font_name, PDF_FONT_SIZE)
        c.setFillColorRGB(0, 0, 0)
        c.drawString(PDF_LEFT_NUM, cur_y, f"{idx+1}.")
        cur_y = _draw_answer_line_wrapped(
            c, PDF_TEXT_START, cur_y, content,
            font_name, PDF_FONT_SIZE, max_width,
            underline_offset=2, line_height=PDF_LINE_HEIGHT
        )

    words = [str(row.get('Word', '')).strip() for row in questions]
    _draw_word_list_page(c, words, font_name, title="詞語表（答案）", word_color=RED)

    c.save()
    bio.seek(0)
    return bio

# ============================================================
# --- 4c. DOCX Export ---
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
# --- 4d. SendGrid Email ---
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
# --- 4e. PDF Preview Helper ---
# ============================================================
def display_pdf_as_images(pdf_bytes):
    try:
        images = convert_from_bytes(pdf_bytes, dpi=150)
        for i, image in enumerate(images):
            st.image(image, caption=f"Page {i+1}", use_container_width=True)
    except Exception as e:
        st.error(f"Could not render preview: {e}")
        st.info("You can still download the PDF using the button above.")

# ============================================================
# --- 5. MAIN CONTENT AREA ---
# ============================================================
st.divider()

# ============================================================
# MODE C: AI 句子審核  (shown first — must confirm before PDF)
# ============================================================
if send_mode == "🤖 AI 句子審核":
    st.subheader("🤖 AI 句子審核")
    st.caption("為每個詞語選擇最合適的句子，確認後題目會鎖入題庫，PDF 即可生成。")

    # Filter to selected level
    level_groups = {k: v for k, v in review_groups.items() if k.endswith(f"||{selected_level}")}

    if not level_groups:
        st.success(f"✅ {selected_level} 目前沒有待審核的 AI 句子。")

        # Show what's already in final_pool for this level
        pool_batches = {k: v for k, v in st.session_state.final_pool.items() if k.endswith(f"||{selected_level}")}
        if pool_batches:
            st.subheader("📦 已鎖定題庫")
            for bk, qs in pool_batches.items():
                school_r, level_r = bk.split("||")
                with st.expander(f"🏫 {school_r}  {level_r}  —  {len(qs)} 題", expanded=False):
                    for q in qs:
                        st.markdown(f"- **{q['Word']}**：{q['Content']}")
        st.stop()

    for batch_key, word_dict in level_groups.items():
        school_r, level_r = batch_key.split("||")
        is_confirmed = batch_key in st.session_state.confirmed_batches

        all_chosen = all(
            any(k.startswith(f"{batch_key}||{w}||") for k in st.session_state.ai_choices)
            for w in word_dict
        )
        chosen_count = sum(
            1 for w in word_dict
            if any(k.startswith(f"{batch_key}||{w}||") for k in st.session_state.ai_choices)
        )

        if is_confirmed:
            status_badge = "✅ 已確認"
        elif all_chosen:
            status_badge = "🟢 可確認"
        else:
            status_badge = f"🟡 {len(word_dict) - chosen_count}/{len(word_dict)} 待選"

        with st.expander(f"🏫 {school_r}  {level_r}　　{status_badge}", expanded=not is_confirmed):

            if is_confirmed:
                st.success("此批次已確認並鎖入題庫。如需重新選擇，請按「重置」。")
                col_rst, col_view = st.columns(2)
                with col_rst:
                    if st.button("🔄 重置此批次", key=f"reset_{batch_key}", use_container_width=True):
                        st.session_state.confirmed_batches.discard(batch_key)
                        st.session_state.final_pool.pop(batch_key, None)
                        keys_to_del = [k for k in st.session_state.ai_choices if k.startswith(batch_key)]
                        for k in keys_to_del:
                            del st.session_state.ai_choices[k]
                        st.rerun()
                with col_view:
                    pool_qs = st.session_state.final_pool.get(batch_key, [])
                    if pool_qs:
                        with st.popover(f"📋 查看 {len(pool_qs)} 題"):
                            for q in pool_qs:
                                st.markdown(f"- **{q['Word']}**：{q['Content']}")
                continue

            # --- Per-word selection ---
            for word, data in word_dict.items():
                ai_list  = data['ai']
                original = data['original']
                row_keys = data['row_keys']

                st.markdown(f"---\n**詞語：{word}**")

                options = []
                option_labels = []
                if original:
                    options.append(('original', original))
                    option_labels.append(f"📝 原句：{original}")
                for i, ai_s in enumerate(ai_list):
                    options.append((f'ai_{i}', ai_s))
                    option_labels.append(f"🤖 AI {i+1}：{ai_s}")

                existing_key = next(
                    (k for k in st.session_state.ai_choices if k.startswith(f"{batch_key}||{word}||")),
                    None
                )
                default_idx = 0
                if existing_key:
                    saved = st.session_state.ai_choices[existing_key]
                    for i, (_, txt) in enumerate(options):
                        if txt == saved:
                            default_idx = i
                            break

                chosen_label = st.radio(
                    f"請為「{word}」選擇句子：",
                    option_labels,
                    index=default_idx,
                    key=f"radio_{batch_key}_{word}",
                    label_visibility="collapsed"
                )

                chosen_idx  = option_labels.index(chosen_label)
                chosen_text = options[chosen_idx][1]
                choice_key  = f"{batch_key}||{word}||{row_keys[0] if row_keys else word}"
                st.session_state.ai_choices[choice_key] = chosen_text
                st.info(f"✏️ 已選：{chosen_text}")

            st.divider()

            all_chosen_now = all(
                any(k.startswith(f"{batch_key}||{w}||") for k in st.session_state.ai_choices)
                for w in word_dict
            )

            if all_chosen_now:
                if st.button(
                    f"✅ 確認並鎖入題庫：{school_r} {level_r}",
                    key=f"confirm_{batch_key}",
                    type="primary",
                    use_container_width=True
                ):
                    # Build final question list for this batch
                    final_qs = []
                    for word, data in word_dict.items():
                        ck = next(
                            (k for k in st.session_state.ai_choices if k.startswith(f"{batch_key}||{word}||")),
                            None
                        )
                        chosen_content = st.session_state.ai_choices[ck] if ck else (data['original'] or '')
                        final_qs.append({
                            'Word': word,
                            'Content': chosen_content,
                            'School': school_r,
                            'Level': level_r,
                        })
                    st.session_state.final_pool[batch_key] = final_qs
                    st.session_state.confirmed_batches.add(batch_key)
                    st.success(f"🎉 已確認！{school_r} {level_r} 共 {len(final_qs)} 題鎖入題庫，PDF 現已解鎖。")
                    st.rerun()
            else:
                st.warning("⚠️ 請為所有詞語選擇句子後才能確認。")

    # Summary table
    st.divider()
    st.subheader("📋 已確認選擇一覽")
    if st.session_state.ai_choices:
        rows = []
        for k, v in st.session_state.ai_choices.items():
            parts = k.split("||")
            if len(parts) >= 3:
                rows.append({"學校": parts[0], "年級": parts[1], "詞語": parts[2], "已選句子": v})
        if rows:
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    else:
        st.caption("尚未有任何確認的選擇。")

# ============================================================
# MODE A: 按學校預覽下載
# ============================================================
elif send_mode == "📄 按學校預覽下載":
    st.subheader("🏫 按學校下載")

    # Get confirmed batches for this level
    level_pool = {k: v for k, v in st.session_state.final_pool.items() if k.endswith(f"||{selected_level}")}

    if not level_pool:
        st.warning(
            f"⚠️ {selected_level} 尚未有已確認的題庫。\n\n"
            f"請先切換到「🤖 AI 句子審核」模式，選擇句子並確認後再回來。"
        )
        st.stop()

    available_schools = sorted([k.split("||")[0] for k in level_pool])
    selected_school = st.selectbox("選擇學校", available_schools, label_visibility="collapsed")
    batch_key = f"{selected_school}||{selected_level}"

    original_questions = level_pool.get(batch_key, [])
    if not original_questions:
        st.warning(f"⚠️ {selected_school} {selected_level} 題庫為空。")
        st.stop()

    cache_key = f"school_{selected_school}_{selected_level}"

    with st.spinner("正在生成文件…"):
        shuffled_questions = get_shuffled_questions(original_questions, cache_key)
        pdf_bytes        = create_pdf(selected_school, selected_level, shuffled_questions, original_questions=original_questions).getvalue()
        answer_pdf_bytes = create_answer_pdf(selected_school, selected_level, shuffled_questions).getvalue()
        docx_bytes       = create_docx(selected_school, selected_level, shuffled_questions).getvalue()

    # Info strip
    ic1, ic2, ic3 = st.columns(3)
    ic1.metric("學校", selected_school)
    ic2.metric("年級", selected_level)
    ic3.metric("題目數", len(original_questions))

    # Download buttons
    dl1, dl2, dl3 = st.columns(3)
    with dl1:
        st.download_button(
            label="📥 學生版 PDF",
            data=pdf_bytes,
            file_name=f"{selected_school}_{selected_level}_{datetime.date.today()}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key=f"dl_{selected_school}_{selected_level}"
        )
    with dl2:
        st.download_button(
            label="📄 Word 檔（可編輯）",
            data=docx_bytes,
            file_name=f"{selected_school}_{selected_level}_{datetime.date.today()}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key=f"dl_docx_{selected_school}_{selected_level}"
        )
    with dl3:
        st.download_button(
            label="🔑 教師版答案 PDF",
            data=answer_pdf_bytes,
            file_name=f"{selected_school}_{selected_level}_教師版_{datetime.date.today()}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key=f"dl_answer_{selected_school}_{selected_level}"
        )

    st.caption("💡 如需修改題目，請在 Google Sheet 更正後點擊側欄「更新資料」。")
    st.divider()
    st.subheader("🔍 學生版預覽")
    display_pdf_as_images(pdf_bytes)

# ============================================================
# MODE B: 按學生寄送
# ============================================================
elif send_mode == "👨‍👩‍👧 按學生寄送":
    st.subheader("👨‍👩‍👧 按學生寄送")

    if student_df.empty:
        st.error("❌ 無法讀取「學生資料」工作表，請確認工作表名稱正確。")
        st.stop()

    required_cols = ['學校', '年級', '狀態', '學生姓名', '學生編號', '家長 Email']
    missing_cols = [c for c in required_cols if c not in student_df.columns]
    if missing_cols:
        st.error(f"❌ 「學生資料」工作表缺少以下欄位：{missing_cols}")
        st.write("現有欄位：", student_df.columns.tolist())
        st.stop()

    # Get confirmed pool for this level
    level_pool_b = {k: v for k, v in st.session_state.final_pool.items() if k.endswith(f"||{selected_level}")}

    if not level_pool_b:
        st.warning(
            f"⚠️ {selected_level} 尚未有已確認的題庫。\n\n"
            f"請先切換到「🤖 AI 句子審核」模式，選擇句子並確認後再回來。"
        )
        st.stop()

    active_students = student_df[
        (student_df['狀態'] == 'Y') &
        (student_df['年級'] == selected_level)
    ]

    if active_students.empty:
        st.warning(f"⚠️ 沒有 {selected_level} 且狀態 = Y 的學生。")
        st.stop()

    # Build questions_df from final_pool
    all_pool_rows = []
    for bk, qs in level_pool_b.items():
        all_pool_rows.extend(qs)
    questions_df = pd.DataFrame(all_pool_rows)

    # Merge students with their school's questions
    merged = active_students.merge(
        questions_df,
        left_on=['學校', '年級'],
        right_on=['School', 'Level'],
        how='inner'
    )

    if merged.empty:
        st.warning("⚠️ 沒有符合條件的配對。請確認學校名稱和年級在兩張表中完全一致。")
        with st.expander("🔍 查看配對資料（協助排查問題）"):
            st.write("**題庫的 School 值：**", questions_df['School'].unique().tolist())
            st.write("**題庫的 Level 值：**", questions_df['Level'].unique().tolist())
            st.write("**學生資料 的 學校 值：**", active_students['學校'].unique().tolist())
            st.write("**學生資料 的 年級 值：**", active_students['年級'].unique().tolist())
        st.stop()

    # Session state for sent/generated tracking
    if 'sent_status' not in st.session_state:
        st.session_state.sent_status = {}
    if 'pdf_generated' not in st.session_state:
        st.session_state.pdf_generated = {}

    # School filter
    all_schools_b = sorted(merged['學校'].unique().tolist())
    selected_school_b = st.selectbox("🏫 選擇學校", all_schools_b)
    school_merged = merged[merged['學校'] == selected_school_b]

    # Build per-student summary
    student_rows = []
    for sid, grp in school_merged.groupby('學生編號'):
        sname  = grp['學生姓名'].iloc[0]
        sgrade = grp['年級'].iloc[0]
        pdf_done  = "📄 已生成" if sid in st.session_state.pdf_generated else "—"
        sent_done = "✅ 已發送" if sid in st.session_state.sent_status else "☐ 未發送"
        student_rows.append({
            '_id': sid, '姓名': sname, '年級': sgrade,
            'PDF': pdf_done, '發送狀態': sent_done,
        })

    st.caption(f"共 {len(student_rows)} 位學生")

    # Two-column layout
    col_list, col_detail = st.columns([1, 2], gap="medium")

    with col_list:
        st.markdown(f"### 👥 學生列表")
        h1, h2, h3, h4 = st.columns([3, 2, 2, 3])
        h1.markdown("**姓名**"); h2.markdown("**年級**")
        h3.markdown("**PDF**");  h4.markdown("**發送狀態**")
        st.divider()

        student_names = [r['姓名'] for r in student_rows]
        if 'selected_student_name_b' not in st.session_state:
            st.session_state.selected_student_name_b = student_names[0] if student_names else None

        for r in student_rows:
            rc1, rc2, rc3, rc4 = st.columns([3, 2, 2, 3])
            is_selected = (st.session_state.selected_student_name_b == r['姓名'])
            label = f"**→ {r['姓名']}**" if is_selected else r['姓名']
            if rc1.button(label, key=f"btn_{r['_id']}", use_container_width=True):
                st.session_state.selected_student_name_b = r['姓名']
                st.rerun()
            rc2.markdown(f"<small>{r['年級']}</small>", unsafe_allow_html=True)
            rc3.markdown(f"<small>{r['PDF']}</small>", unsafe_allow_html=True)
            rc4.markdown(f"<small>{r['發送狀態']}</small>", unsafe_allow_html=True)

    with col_detail:
        sel_row = next(
            (r for r in student_rows if r['姓名'] == st.session_state.get('selected_student_name_b')),
            None
        )
        if sel_row is None:
            st.info("👈 請從左側列表選擇一位學生。")
        else:
            student_id   = sel_row['_id']
            student_name = sel_row['姓名']
            grade        = sel_row['年級']
            group        = school_merged[school_merged['學生編號'] == student_id]
            parent_email  = str(group['家長 Email'].iloc[0]).strip()
            teacher_email = group['老師 Email'].iloc[0] if '老師 Email' in group.columns else "N/A"

            unique_group   = group.drop_duplicates(subset=['Content'])
            question_count = len(unique_group)

            with st.container(border=True):
                ic1, ic2, ic3, ic4 = st.columns(4)
                ic1.markdown(f"**👤 {student_name}**")
                ic2.markdown(f"**🏫** {selected_school_b}")
                ic3.markdown(f"**🎓** {grade}")
                ic4.markdown(f"**📝** {question_count} 題")
                st.markdown(f"📧 家長電郵：`{parent_email}`")

            original_questions = unique_group.to_dict('records')
            cache_key = f"student_{student_id}_{grade}"

            with st.spinner(f"正在生成 {student_name} 的文件…"):
                shuffled_questions = get_shuffled_questions(original_questions, cache_key)
                pdf_bytes        = create_pdf(selected_school_b, grade, shuffled_questions, student_name=student_name, original_questions=original_questions).getvalue()
                answer_pdf_bytes = create_answer_pdf(selected_school_b, grade, shuffled_questions, student_name=student_name).getvalue()
                docx_bytes       = create_docx(selected_school_b, grade, shuffled_questions, student_name=student_name).getvalue()
                st.session_state.pdf_generated[student_id] = True

            tab_gen, tab_preview = st.tabs(["📄 生成與發送", "🔍 預覽工作紙"])

            with tab_gen:
                dl1, dl2, dl3 = st.columns(3)
                with dl1:
                    st.download_button(
                        label="📥 學生版 PDF",
                        data=pdf_bytes,
                        file_name=f"{student_name}_{grade}_{datetime.date.today()}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key=f"dl_{student_id}"
                    )
                with dl2:
                    st.download_button(
                        label="📄 Word 檔",
                        data=docx_bytes,
                        file_name=f"{student_name}_{grade}_{datetime.date.today()}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        key=f"dl_docx_{student_id}"
                    )
                with dl3:
                    st.download_button(
                        label="🔑 教師版答案",
                        data=answer_pdf_bytes,
                        file_name=f"{student_name}_{grade}_教師版_{datetime.date.today()}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key=f"dl_answer_{student_id}"
                    )

                st.divider()
                if st.button(
                    "📧 立即寄送給家長",
                    key=f"send_{student_id}",
                    use_container_width=True,
                    type="primary"
                ):
                    with st.spinner(f"正在寄送給 {parent_email}…"):
                        success, msg = send_email_with_pdf(
                            parent_email, student_name, selected_school_b, grade,
                            pdf_bytes, cc_email=teacher_email
                        )
                        if success:
                            st.session_state.sent_status[student_id] = True
                            st.success(f"✅ 已成功寄送給 {parent_email}！")
                            st.rerun()
                        else:
                            st.error(f"❌ 發送失敗: {msg}")
                            st.code(msg)

            with tab_preview:
                display_pdf_as_images(pdf_bytes)
