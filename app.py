import pandas as pd
import streamlit as st
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
# --- Streamlit Setup ---
# ============================================================

st.set_page_config(page_title="Worksheet Generator", page_icon="📝", layout="wide")
st.title("📝 校本填充工作紙生成器")

# Session state
st.session_state.setdefault("shuffled_cache", {})
st.session_state.setdefault("final_pool", {})
st.session_state.setdefault("ai_choices", {})
st.session_state.setdefault("confirmed_batches", set())
st.session_state.setdefault("last_selected_level", None)

# ============================================================
# --- ReportLab Font Setup ---
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
                pdfmetrics.registerFont(TTFont("ChineseFont", path))
                CHINESE_FONT = "ChineseFont"
                break
            except Exception:
                continue

    if not CHINESE_FONT:
        st.error("❌ Chinese font not found. Please ensure Kai.ttf is in your GitHub repository.")

except ImportError:
    st.error("❌ reportlab not found. Please add 'reportlab' to your requirements.txt")
    st.stop()

# ============================================================
# --- Google Sheet Connection ---
# ============================================================

try:
    key_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(
        key_dict,
        scopes=[
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
    )
    client = gspread.authorize(creds)
    SHEET_ID = st.secrets["app_config"]["spreadsheet_id"]

except Exception as e:
    st.error(f"❌ Google Sheet Connection Error: {e}")
    st.stop()

# ============================================================
# --- Google Sheet Loader (Refactored) ---
# ============================================================

def load_sheet(sheet_name: str) -> pd.DataFrame:
    """讀取 Google Sheet 並清洗欄位。"""
    try:
        sh = client.open_by_key(SHEET_ID)
        ws = sh.worksheet(sheet_name)
        df = pd.DataFrame(ws.get_all_records())

        df.columns = [c.strip() for c in df.columns]
        for col in df.columns:
            df[col] = df[col].astype(str).str.strip()


        return df

    except Exception as e:
        st.error(f"❌ 無法讀取工作表「{sheet_name}」: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=60)
def load_review():
    return load_sheet("Review")


@st.cache_data(ttl=60)
def load_students():
    return load_sheet("學生資料")

# ============================================================
# --- Review Parser (Refactored) ---
# ============================================================

def parse_review_table(df: pd.DataFrame):
    groups = {}

    for idx, row in df.iterrows():
        school = row.get("學校", "").strip()
        level = row.get("年級", "").strip()
        word = row.get("詞語", "").strip()
        sentence = row.get("句子", "").strip()

        if not (school and level and word and sentence):
            continue

        batch_key = f"{school}||{level}"
        groups.setdefault(batch_key, {})
        groups[batch_key].setdefault(word, {
            "original": None,
            "ai": [],
            "needs_review": False,
            "row_indices": []
        })

        is_ai = sentence.startswith("🟨")
        clean_sentence = sentence.lstrip("🟨").strip()

        if is_ai:
            groups[batch_key][word]["ai"].append(clean_sentence)
            groups[batch_key][word]["needs_review"] = True
        else:
            groups[batch_key][word]["original"] = clean_sentence

        groups[batch_key][word]["row_indices"].append(idx)

    return groups

# ============================================================
# --- Batch Readiness Checker ---
# ============================================================

def compute_batch_readiness(batch_key: str, word_dict: dict):
    ready_words = []
    pending_words = []
    for word, data in word_dict.items():
        if data["needs_review"]:
            # 統一使用新的 key 格式
            chosen = st.session_state.ai_choices.get(f"{batch_key}||{word}||0", None)
            if chosen:
                ready_words.append((word, chosen))
            else:
                pending_words.append(word)
        else:
            if data["original"]:
                ready_words.append((word, data["original"]))
    is_ready = len(pending_words) == 0
    return ready_words, pending_words, is_ready

# ============================================================
# --- Final Pool Builder ---
# ============================================================

def build_final_pool_for_batch(batch_key: str, word_dict: dict):
    school, level = batch_key.split("||")
    questions = []
    for word, data in word_dict.items():
        if data["needs_review"]:
            content = st.session_state.ai_choices.get(f"{batch_key}||{word}||0", "")
        else:
            content = data["original"] or ""
        if content:
            questions.append({
                "Word": word,
                "Content": content,
                "School": school,
                "Level": level,
            })
    return questions

# ============================================================
# --- PDF Text Rendering Helpers ---
# ============================================================

def draw_text_with_underline_wrapped(c, x, y, text, font_name, font_size, max_width,
                                     underline_offset=2, line_height=18):
    """
    支援 <u>底線</u> 的自動換行文字繪製。
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
# --- Student Worksheet PDF Generator (MODIFIED on mar 2 16:00) ---
# ============================================================

def create_pdf(school_name, level, questions, student_name=None):
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER

    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=letter)
    story = []

    styles = getSampleStyleSheet()
    font_name = CHINESE_FONT if CHINESE_FONT else 'Helvetica'

    # --- TITLE STYLE ---
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontName=font_name,
        fontSize=22,
        alignment=TA_CENTER,
        spaceAfter=12
    )

    # --- BODY TEXT STYLE WITH LINE SPACING ---
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontName=font_name,
        fontSize=18,
        leading=26,              # 👈 Controls line spacing
        leftIndent=0,
        firstLineIndent=0
    )

    # --- VOCABULARY TABLE TITLE STYLE ---
    vocab_title_style = ParagraphStyle(
        'VocabTitle',
        parent=styles['Heading2'],
        fontName=font_name,
        fontSize=20,
        alignment=TA_CENTER,
        spaceAfter=20
    )

    # --- TITLE ---
    title_text = f"<b>{school_name} ({level}) - {student_name} - 校本填充工作紙</b>" if student_name \
                 else f"<b>{school_name} ({level}) - 校本填充工作紙</b>"
    story.append(Paragraph(title_text, title_style))
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph(f"日期: {datetime.date.today() + datetime.timedelta(days=1)}", normal_style))
    story.append(Spacer(1, 0.3*inch))

    # --- QUESTIONS ---
    for i, row in enumerate(questions):
        content = row['Content']
        content = re.sub(r'【】(.*?)【】', r'<u>\1</u>', content)
        content = re.sub(r'【(.+?)】', r'<u>________</u>', content)

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

    # --- VOCABULARY TABLE (Second Page) ---
    words = [row.get('Word', '').strip() for row in questions]
    unique_words = list(dict.fromkeys([w for w in words if w]))

    if unique_words:
        story.append(PageBreak())
        story.append(Paragraph("<b>詞語表</b>", vocab_title_style))
        story.append(Spacer(1, 0.2*inch))

        # Organize into 4-column table
        num_cols = 4
        table_data = []
        for i in range(0, len(unique_words), num_cols):
            row = unique_words[i:i+num_cols]
            while len(row) < num_cols:
                row.append('')
            table_data.append(row)

        col_width = 1.8*inch
        vocab_table = Table(table_data, colWidths=[col_width]*num_cols)
        vocab_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), font_name),
            ('FONTSIZE', (0, 0), (-1, -1), 22),         # 👈 Larger font for readability
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('TOPPADDING', (0, 0), (-1, -1), 16),       # 👈 Vertical padding
            ('BOTTOMPADDING', (0, 0), (-1, -1), 16),
            ('LEFTPADDING', (0, 0), (-1, -1), 12),      # 👈 Horizontal padding
            ('RIGHTPADDING', (0, 0), (-1, -1), 12),
        ]))
        story.append(vocab_table)

    doc.build(story)
    bio.seek(0)
    return bio

# ============================================================
# --- Teacher Answer PDF Generator ---
# ============================================================

def create_answer_pdf(school_name, level, questions):
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.colors import red as RED

    bio = io.BytesIO()
    c = rl_canvas.Canvas(bio, pagesize=letter)
    page_width, page_height = letter
    font_name = CHINESE_FONT or "Helvetica"

    cur_y = page_height - 80
    left_m = 60

    c.setFont(font_name, 22)
    c.drawString(left_m, cur_y, "詞語清單（題目順序）")
    cur_y -= 40

    c.setFont(font_name, 18)

    for idx, row in enumerate(questions, start=1):
        word = row["Word"]

        if cur_y < 60:
            c.showPage()
            cur_y = page_height - 80
            c.setFont(font_name, 22)
            c.drawString(left_m, cur_y, "詞語清單（續）")
            cur_y -= 40
            c.setFont(font_name, 18)

        c.drawString(left_m, cur_y, f"{idx}. ")
        c.setFillColor(RED)
        c.drawString(left_m + 40, cur_y, word)
        c.setFillColorRGB(0, 0, 0)
        cur_y -= 26

    c.save()
    bio.seek(0)
    return bio

# ============================================================
# --- DOCX Worksheet Generator ---
# ============================================================

def create_docx(school_name, level, questions, student_name=None):
    doc = Document()

    # 標題
    title_text = f"{school_name} ({level}) - {student_name} - 校本填充工作紙" if student_name \
                 else f"{school_name} ({level}) - 校本填充工作紙"
    title = doc.add_heading(title_text, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 日期
    date_para = doc.add_paragraph(f"日期: {datetime.date.today() + datetime.timedelta(days=1)}")
    date_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    doc.add_paragraph("")

    # 題目
    for i, row in enumerate(questions):
        content = re.sub(r'【|】', '', row["Content"])
        p = doc.add_paragraph(style="List Number")
        run = p.add_run(content)
        run.font.size = Pt(18)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ============================================================
# --- SendGrid Email Sender ---
# ============================================================

def send_email_with_pdf(to_email, student_name, school_name, grade, pdf_bytes, cc_email=None):
    try:
        sg_config = st.secrets["sendgrid"]
        recipient = str(to_email).strip()

        # 基本 email 格式檢查
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

        # CC email
        if cc_email:
            cc_clean = str(cc_email).strip().lower()
            if cc_clean not in ["n/a", "nan", "", "none"] and "@" in cc_clean and cc_clean != recipient.lower():
                message.add_cc(cc_clean)

        # 附件
        encoded_pdf = base64.b64encode(pdf_bytes).decode()
        attachment = Attachment(
            FileContent(encoded_pdf),
            FileName(f"{safe_name}_Worksheet.pdf"),
            FileType("application/pdf"),
            Disposition("attachment")
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
# --- PDF Preview Helper ---
# ============================================================

def display_pdf_as_images(pdf_bytes):
    try:
        images = convert_from_bytes(pdf_bytes, dpi=150)
        for i, image in enumerate(images):
            st.image(image, caption=f"Page {i+1}", use_container_width=True)
    except Exception as e:
        st.error(f"無法顯示 PDF 預覽: {e}")
        st.info("你仍然可以使用下載按鈕下載 PDF。")

# ============================================================
# --- Sidebar Controls ---
# ============================================================

student_df = load_students()
review_df = load_review()
review_groups = parse_review_table(review_df)

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

    # 年級選擇
    all_levels = sorted(review_df["年級"].astype(str).unique().tolist()) if not review_df.empty else ["P1"]
    st.subheader("🎓 年級")
    selected_level = st.radio("選擇年級", all_levels, index=0, label_visibility="collapsed")

    if st.session_state.last_selected_level != selected_level:
        st.session_state.last_selected_level = selected_level
        st.session_state.selected_student_name_b = None

    st.divider()

    # 模式選擇
    st.subheader("📬 模式")
    send_mode = st.radio(
        "選擇模式",
        ["🤖 AI 句子審核", "📄 按學校預覽下載", "👨‍👩‍👧 按學生寄送"],
        index=0,
        label_visibility="collapsed"
    )

    st.divider()

    # 統計資訊
    st.subheader("📊 資料概覽")

    level_batches = [k for k in review_groups if k.endswith(f"||{selected_level}")]
    total_words = sum(len(v) for k, v in review_groups.items() if k.endswith(f"||{selected_level}"))
    ai_words = sum(
        1 for k, v in review_groups.items() if k.endswith(f"||{selected_level}")
        for w, d in v.items() if d["needs_review"]
    )
    ready_words_cnt = sum(
        1 for k, v in review_groups.items() if k.endswith(f"||{selected_level}")
        for w, d in v.items() if not d["needs_review"]
    )
    confirmed_count = len([k for k in st.session_state.confirmed_batches if k.endswith(f"||{selected_level}")])
    pool_count = sum(len(v) for k, v in st.session_state.final_pool.items() if k.endswith(f"||{selected_level}"))

    st.metric(f"{selected_level} 批次數", len(level_batches))
    st.metric("總詞語數", total_words)
    st.metric("🟨 待選 AI 句", ai_words)
    st.metric("✅ 已就緒（原句）", ready_words_cnt)
    st.metric("已確認批次", confirmed_count)
    st.metric("題庫已鎖定題目", pool_count)

    if not student_df.empty and "狀態" in student_df.columns:
        active_count = (student_df["狀態"] == "Y").sum()
        st.metric("啟用學生數", int(active_count))

# ============================================================
# --- Shuffle Helper ---
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
# --- PDF Layout Constants ---
# ============================================================

PDF_LEFT_NUM = 60
PDF_TEXT_START = PDF_LEFT_NUM + 30
PDF_RIGHT_MARGIN = 40
PDF_LINE_HEIGHT = 26
PDF_FONT_SIZE = 18

def _get_max_width():
    page_width, _ = letter
    return page_width - PDF_RIGHT_MARGIN - PDF_TEXT_START


# ============================================================
# --- Mode B: 按學校預覽下載 ---
# ============================================================

if send_mode == "📄 按學校預覽下載":
    st.subheader("📄 按學校預覽下載")

    # 只顯示選定年級的批次
    level_batches = {k: v for k, v in st.session_state.final_pool.items() if k.endswith(f"||{selected_level}")}

    if not level_batches:
        st.info("⚠️ 尚未有任何批次完成 AI 審核並鎖定題庫。")
        st.stop()

    for batch_key, questions in level_batches.items():
        school, level = batch_key.split("||")
        st.markdown(f"### 🏫 {school}（{level}）")

        # 生成 PDF
        pdf_bytes = create_pdf(school, level, questions)
        answer_pdf_bytes = create_answer_pdf(school, level, questions)

        col1, col2 = st.columns(2)

        with col1:
            st.download_button(
                label="⬇️ 下載學生版 PDF",
                data=pdf_bytes,
                file_name=f"{school}_{level}_worksheet.pdf",
                mime="application/pdf"
            )

        with col2:
            st.download_button(
                label="⬇️ 下載教師版 PDF（答案）",
                data=answer_pdf_bytes,
                file_name=f"{school}_{level}_answers.pdf",
                mime="application/pdf"
            )

        # 預覽 PDF
        with st.expander("📘 預覽學生版 PDF"):
            display_pdf_as_images(pdf_bytes)

        st.divider()

# ============================================================
# --- Mode C: 按學生寄送 ---
# ============================================================

if send_mode == "👨‍👩‍👧 按學生寄送":
    st.subheader("👨‍👩‍👧 按學生寄送")

    if student_df.empty:
        st.error("❌ 學生資料表為空，無法寄送。")
        st.stop()

    # 過濾選定年級
    df_level = student_df[student_df["年級"].astype(str) == selected_level]

    if df_level.empty:
        st.info(f"⚠️ {selected_level} 沒有學生資料。")
        st.stop()

    # 學生選擇（使用「學生姓名」欄）
    student_names = df_level["學生姓名"].tolist()
    selected_student = st.selectbox("選擇學生", [""] + student_names)

    if not selected_student:
        st.stop()

    # 取得學生資料
    row = df_level[df_level["學生姓名"] == selected_student].iloc[0]
    school = row["學校"]
    grade = row["年級"]

    # Email 欄位名稱修正
    parent_email = row.get("家長 Email", "")
    cc_email = row.get("老師 Email", "")

    batch_key = f"{school}||{grade}"

    if batch_key not in st.session_state.final_pool:
        st.error("⚠️ 此學生所屬批次尚未完成 AI 審核並鎖定題庫。")
        st.stop()

    questions = st.session_state.final_pool[batch_key]

    # 生成 PDF
    pdf_bytes = create_pdf(school, grade, questions, student_name=selected_student)

    st.download_button(
        label="⬇️ 下載學生版 PDF",
        data=pdf_bytes,
        file_name=f"{selected_student}_worksheet.pdf",
        mime="application/pdf"
    )

    st.divider()

    # 寄送 email
    st.markdown("### ✉️ 寄送工作紙至家長電郵")

    if st.button("📨 寄出工作紙"):
        ok, msg = send_email_with_pdf(
            parent_email,
            selected_student,
            school,
            grade,
            pdf_bytes,
            cc_email=cc_email
        )

        if ok:
            st.success("🎉 已成功寄出！")
        else:
            st.error(f"❌ 寄送失敗：{msg}")


# ============================================================
# --- End of App ---
# ============================================================

st.write("")
st.write("© 2026 校本填充工作紙生成器 — 自動化教學工具")

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
