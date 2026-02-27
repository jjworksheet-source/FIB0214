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

st.session_state.setdefault("selected_student_name_b", None)  # ← 新增

# 防止 final_pool 被污染
if not isinstance(st.session_state.final_pool, dict):
    st.session_state.final_pool = {}
	
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
# --- Student Worksheet PDF Generator ---
# ============================================================

def create_pdf(school_name, level, questions, student_name=None):
    from reportlab.pdfgen import canvas as rl_canvas

    bio = io.BytesIO()
    c = rl_canvas.Canvas(bio, pagesize=letter)
    _, page_height = letter
    font_name = CHINESE_FONT or "Helvetica"
    max_width = 500
    cur_y = page_height - 60

    # 標題
    c.setFont(font_name, 22)
    title = f"{school_name} ({level}) - {student_name} - 校本填充工作紙" if student_name \
            else f"{school_name} ({level}) - 校本填充工作紙"
    c.drawString(60, cur_y, title)
    cur_y -= 30

    # 日期
    c.setFont(font_name, 18)
    c.drawString(60, cur_y, f"日期: {datetime.date.today() + datetime.timedelta(days=1)}")
    cur_y -= 30

    # 題目
    for idx, row in enumerate(questions):
        content = row["Content"]

        # 處理底線格式
        content = re.sub(r'【】(.*?)【】', r'<u>\1</u>', content)

        if cur_y < 80:
            c.showPage()
            cur_y = page_height - 60

        c.setFont(font_name, 18)
        c.drawString(60, cur_y, f"{idx+1}.")
        cur_y = draw_text_with_underline_wrapped(
            c, 100, cur_y, content, font_name, 18, max_width
        )

    c.save()
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
    pool_count = sum(len(v) for k, v in st.session_state.final_pool.items() if k.endswith(f"||{selected_level}") and isinstance(v, list))

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
# --- Mode A: AI 句子審核 ---
# ============================================================

st.divider()

if send_mode == "🤖 AI 句子審核":
    st.subheader("🤖 AI 句子審核")
    
    level_groups = {k: v for k, v in review_groups.items() if k.endswith(f"||{selected_level}")}

    if not level_groups:
        st.success(f"✅ {selected_level} 目前沒有任何題目。")
        st.stop()

    for batch_key, word_dict in level_groups.items():
        school, level = batch_key.split("||")
        st.markdown(f"### 🏫 {school}（{level}）")

        has_any_ai_review = any(d["needs_review"] for d in word_dict.values())

        if not has_any_ai_review:
            # --- 情況 1：全部都是原句，不需要審核 ---
            st.info(f"💡 這次 **{school}** 學校句子沒有需要審核，請直接到「按學校預覽下載」或「按學生寄送」使用工作紙。")
            
            if batch_key not in st.session_state.confirmed_batches:
                final_qs = build_final_pool_for_batch(batch_key, word_dict)
                st.session_state.final_pool[batch_key] = final_qs
                st.session_state.confirmed_batches.add(batch_key)

        else:
            # --- 情況 2：有 AI 句需要審核 ---
            ready_words, pending_words, is_ready = compute_batch_readiness(batch_key, word_dict)

            for word, data in word_dict.items():
                if data["needs_review"]:
                    st.markdown(f"#### 詞語：{word}")
                    ai_list = data["ai"]
                    key_radio = f"{batch_key}||{word}||choice"
                    key_custom = f"{batch_key}||{word}||custom"

                    # 選項：AI 句子 + 自行輸入（移除「不選」）
                    options = ai_list + ["✏️ 自行輸入句子"]

                    # 決定預設選哪一個
                    current = st.session_state.ai_choices.get(f"{batch_key}||{word}||0", None)
                    if current in ai_list:
                        default_index = ai_list.index(current)
                    elif current and current not in ai_list:
                        default_index = len(options) - 1
                    else:
                        default_index = 0  # 預設選第一句

                    selected = st.radio(
                        f"請為「{word}」選擇最合適的句子：",
                        options,
                        index=default_index,
                        key=key_radio
                    )

                    if selected == "✏️ 自行輸入句子":
                        prev_custom = st.session_state.get(key_custom, "")
                        custom_input = st.text_input(
                            f"請輸入「{word}」的自定義句子（請用【】詞語【】標示）：",
                            value=prev_custom,
                            key=key_custom,
                            placeholder="例如：小明【定期】到牙科診所檢查牙齒。"
                        )
                        if custom_input.strip():
                            st.session_state.ai_choices[f"{batch_key}||{word}||0"] = custom_input.strip()
                        else:
                            st.session_state.ai_choices.pop(f"{batch_key}||{word}||0", None)
                    else:
                        st.session_state.ai_choices[f"{batch_key}||{word}||0"] = selected
                        if key_custom in st.session_state:
                            del st.session_state[key_custom]

            # 顯示待確認詞語提示
            if pending_words:
                st.warning(f"⚠️ 以下詞語尚未選擇句子：{', '.join(pending_words)}")

            # 確認鎖定按鈕
            if is_ready and batch_key not in st.session_state.confirmed_batches:
                if st.button(f"🔒 確認並鎖定題庫：{school}", key=f"confirm_{batch_key}"):
                    final_qs = build_final_pool_for_batch(batch_key, word_dict)
                    st.session_state.final_pool[batch_key] = final_qs
                    st.session_state.confirmed_batches.add(batch_key)
                    st.success("✅ 已鎖定題庫！")
                    st.rerun()
            elif batch_key in st.session_state.confirmed_batches:
                st.success("✅ 此批次已完成審核並鎖定。")

        st.divider()
	

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
