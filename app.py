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
st.session_state.setdefault("selected_student_name_b", None)

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
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.file"
        ]
    )
    client = gspread.authorize(creds)
    SHEET_ID = st.secrets["app_config"]["spreadsheet_id"]

except Exception as e:
    st.error(f"❌ Google Sheet Connection Error: {e}")
    st.stop()

# ============================================================
# --- Google Sheet Loader ---
# ============================================================

def load_sheet(sheet_name: str) -> pd.DataFrame:
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


@st.cache_data(ttl=60)
def load_used_sentences():
    """載入已使用的句子工作表"""
    try:
        df = load_sheet("已使用")
        return df
    except Exception:
        # 如果工作表不存在，返回空的 DataFrame
        return pd.DataFrame()


def write_used_sentences(sentences_data):
    """將已使用的句子寫入「已使用」工作表"""
    try:
        sh = client.open_by_key(SHEET_ID)

        # 嘗試打開已使用工作表，如果不存在則創建
        sheet_exists = True
        try:
            ws = sh.worksheet("已使用")
        except Exception:
            sheet_exists = False
            # 創建新工作表
            ws = sh.add_worksheet("已使用", rows=1000, cols=5)
            # 設定標題行
            ws.update('A1:E1', [['學校', '年級', '詞語', '句子', '使用日期']])

        # 準備要寫入的資料
        today = datetime.date.today().strftime("%Y-%m-%d")
        rows_to_add = []

        for item in sentences_data:
            row = [
                item.get("school", ""),
                item.get("level", ""),
                item.get("word", ""),
                item.get("sentence", ""),
                today
            ]
            rows_to_add.append(row)

        # 讀取現有所有資料找出正確的下一行
        all_values = ws.get_all_values()
        next_row = len(all_values) + 1  # 自動計算下一行

        # 寫入資料
        if rows_to_add:
            cell_range = f'A{next_row}:E{next_row + len(rows_to_add) - 1}'
            ws.update(cell_range, rows_to_add)

        return True, f"成功寫入 {len(rows_to_add)} 筆記錄"

    except Exception as e:
        return False, str(e)

# ============================================================
# --- Review Parser ---
# ============================================================

def parse_review_table(df: pd.DataFrame, used_df: pd.DataFrame = None):
    """
    解析審核表格
    - df: Review 工作表的資料
    - used_df: 已使用句子的工作表資料（用於過濾）
    """
    groups = {}

    # 建立已使用句子的集合，用於快速查詢
    used_sentences = set()
    if used_df is not None and not used_df.empty:
        for _, row in used_df.iterrows():
            # 用 (學校+年級+詞語+句子) 作為唯一識別
            key = f"{row.get('學校', '').strip()}||{row.get('年級', '').strip()}||{row.get('詞語', '').strip()}||{row.get('句子', '').strip()}"
            used_sentences.add(key)

    for idx, row in df.iterrows():
        school = row.get("學校", "").strip()
        level = row.get("年級", "").strip()
        word = row.get("詞語", "").strip()
        sentence = row.get("句子", "").strip()

        if not (school and level and word and sentence):
            continue

        # 檢查這個句子是否已經被使用過
        sentence_key = f"{school}||{level}||{word}||{sentence}"
        if sentence_key in used_sentences:
            continue  # 跳過已使用的句子

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
    max_width = 450
    cur_y = page_height - 60

    c.setFont(font_name, 22)
    title = f"{school_name} ({level}) - {student_name} - 校本填充工作紙" if student_name \
            else f"{school_name} ({level}) - 校本填充工作紙"
    c.drawString(60, cur_y, title)
    cur_y -= 30

    c.setFont(font_name, 18)
    c.drawString(60, cur_y, f"日期: {datetime.date.today() + datetime.timedelta(days=1)}")
    cur_y -= 30

    for idx, row in enumerate(questions):
        content = row["Content"]
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

    title_text = f"{school_name} ({level}) - {student_name} - 校本填充工作紙" if student_name \
                 else f"{school_name} ({level}) - 校本填充工作紙"
    title = doc.add_heading(title_text, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    date_para = doc.add_paragraph(f"日期: {datetime.date.today() + datetime.timedelta(days=1)}")
    date_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    doc.add_paragraph("")

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

# 預先載入資料（加入載入狀態）
with st.spinner("正在載入資料，請稍候..."):
    student_df = load_students()
    used_df = load_used_sentences()  # 載入已使用的句子

# 在 spinner 外面定義 review_df，確保後續程式碼可以存取
review_df = load_review()
review_groups = parse_review_table(review_df, used_df)

with st.sidebar:
    st.header("⚙️ 控制面板")

    # === 控制區塊 ===
    with st.container(border=True):
        col_r, col_s = st.columns(2)

        with col_r:
            if st.button("🔄 更新資料", use_container_width=True, help="點擊重新載入 Google Sheets 資料"):
                with st.spinner("正在同步最新資料..."):
                    load_review.clear()
                    load_students.clear()
                    load_used_sentences.clear()  # 清除已使用句子的快取
                    st.session_state.final_pool = {}
                    st.session_state.ai_choices = {}
                    st.session_state.confirmed_batches = set()
                    st.session_state.shuffled_cache = {}
                    st.rerun()

        with col_s:
            if st.button("🔀 打亂題目", use_container_width=True, help="重新隨機排序題目順序"):
                st.session_state.shuffled_cache = {}
                st.rerun()

    st.divider()

    # === 篩選區塊 ===
    with st.container(border=True):
        all_levels = sorted(review_df["年級"].astype(str).unique().tolist()) if not review_df.empty else ["P1"]
        st.subheader("🎓 選擇年級")
        selected_level = st.selectbox(
            "年級",
            all_levels,
            index=0,
            label_visibility="collapsed",
            help="選擇要處理的工作表年級"
        )

        if st.session_state.last_selected_level != selected_level:
            st.session_state.last_selected_level = selected_level
            st.session_state.selected_student_name_b = None

    st.divider()

    # === 狀態儀表板 ===
    with st.container(border=True):
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
        pool_count = sum(
            len(v) for k, v in st.session_state.final_pool.items()
            if k.endswith(f"||{selected_level}") and isinstance(v, list)
        )

        # 使用更視覺化的指標顯示
        col_stat1, col_stat2 = st.columns(2)
        with col_stat1:
            st.metric("批次數", len(level_batches))
            st.metric("待選 AI 句", ai_words, delta="⚠️ 待處理" if ai_words > 0 else None)
        with col_stat2:
            st.metric("總詞語", total_words)
            st.metric("已就緒", ready_words_cnt, delta="✅ 完成" if ready_words_cnt > 0 else None)

        st.metric("已鎖定題庫", pool_count)

        # 顯示已使用句子的數量
        used_count = len(used_df) if used_df is not None and not used_df.empty else 0
        st.metric("已使用句子", used_count, help="已記錄在「已使用」工作表中的句子總數")

        if not student_df.empty and "狀態" in student_df.columns:
            active_count = (student_df["狀態"] == "Y").sum()
            st.metric("啟用學生", int(active_count))

    st.divider()

    # === 說明區塊 ===
    with st.expander("📖 使用說明", expanded=False):
        st.markdown("""
        **操作流程：**

        1. **AI 審核**：選擇 AI 生成的句子或輸入自定義句子
        2. **鎖定題庫**：確認審核完成後鎖定題目
        3. **預覽下載**：生成並下載工作紙 PDF
        4. **寄送郵件**：將工作紙寄送給學生家長

        **小提示：**
        - 使用【詞語】標記需要填寫的部分
        - 寄送前請確認學生資料正確
        """)

    # === 系統狀態 ===
    with st.container(border=True):
        st.caption("🔗 系統狀態")
        if not student_df.empty:
            st.success("✅ Google Sheets 已連接")
        else:
            st.warning("⚠️ 請檢查資料連接")

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

# ============================================================
# --- 頂部標籤頁導航 ---
# ============================================================

st.divider()

# 建立三個標籤頁
tab_review, tab_preview, tab_email = st.tabs([
    "🤖 AI 句子審核",
    "📄 預覽下載",
    "✉️ 寄送郵件"
])

# ============================================================
# --- 標籤頁 1: AI 句子審核 ---
# ============================================================

with tab_review:
    st.subheader("🤖 AI 句子審核")

    level_groups = {k: v for k, v in review_groups.items() if k.endswith(f"||{selected_level}")}

    if not level_groups:
        with st.container(border=True):
            st.success(f"✅ {selected_level} 目前沒有任何題目。")
            st.info("請確認 Google Sheets 中的資料是否正確，或嘗試點擊側邊欄的「更新資料」按鈕。")
        st.stop()

    for batch_key, word_dict in level_groups.items():
        with st.container(border=True):
            school, level = batch_key.split("||")
            st.markdown(f"### 🏫 {school}（{level}）")

            has_any_ai_review = any(d["needs_review"] for d in word_dict.values())

            if not has_any_ai_review:
                st.info(f"💡 這次 **{school}** 學校句子沒有需要審核，請切換到「預覽下載」或「寄送郵件」使用工作紙。")
                if batch_key not in st.session_state.confirmed_batches:
                    final_qs = build_final_pool_for_batch(batch_key, word_dict)
                    st.session_state.final_pool[batch_key] = final_qs
                    st.session_state.confirmed_batches.add(batch_key)

            else:
                ready_words, pending_words, is_ready = compute_batch_readiness(batch_key, word_dict)

                for word, data in word_dict.items():
                    if data["needs_review"]:
                        with st.expander(f"📝 詞語：{word}", expanded=True):
                            ai_list = data["ai"]
                            key_radio = f"{batch_key}||{word}||choice"
                            key_custom = f"{batch_key}||{word}||custom"

                            options = ai_list + ["✏️ 自行輸入句子"]

                            current = st.session_state.ai_choices.get(f"{batch_key}||{word}||0", None)
                            if current in ai_list:
                                default_index = ai_list.index(current)
                            elif current and current not in ai_list:
                                default_index = len(options) - 1
                            else:
                                default_index = 0

                            selected = st.radio(
                                "請選擇最合適的句子：",
                                options,
                                index=default_index,
                                key=key_radio,
                                label_visibility="collapsed"
                            )

                            if selected == "✏️ 自行輸入句子":
                                prev_custom = st.session_state.get(key_custom, "")
                                custom_input = st.text_input(
                                    "請輸入自定義句子（使用【】詞語【】標示）：",
                                    value=prev_custom,
                                    key=key_custom,
                                    placeholder="例如：小明【定期】到牙科診所檢查牙齒。",
                                    help="請用【】符號標示需要填寫的詞語"
                                )
                                if custom_input.strip():
                                    st.session_state.ai_choices[f"{batch_key}||{word}||0"] = custom_input.strip()
                                else:
                                    st.session_state.ai_choices.pop(f"{batch_key}||{word}||0", None)
                            else:
                                st.session_state.ai_choices[f"{batch_key}||{word}||0"] = selected
                                if key_custom in st.session_state:
                                    del st.session_state[key_custom]

                if pending_words:
                    st.warning(f"⚠️ 以下詞語尚未選擇句子：{', '.join(pending_words)}")

                # 確認鎖定區塊
                if is_ready and batch_key not in st.session_state.confirmed_batches:
                    with st.container(border=True):
                        st.markdown("### 🔒 確認並鎖定題庫")
                        st.info("請確認所有詞語都已選擇句子後，再鎖定題庫。鎖定後將寫入使用記錄。")

                        # 二次確認機制
                        confirm_checkbox = st.checkbox(
                            "我確認已完成所有詞語的審核，同意鎖定題庫並寫入使用記錄",
                            key=f"confirm_check_{batch_key}"
                        )

                        if confirm_checkbox:
                            if st.button(f"✅ 確認並鎖定題庫：{school}", key=f"confirm_{batch_key}", type="primary"):
                                with st.spinner("正在鎖定題庫並寫入使用記錄..."):
                                    # 構建最終題庫
                                    final_qs = build_final_pool_for_batch(batch_key, word_dict)
                                    st.session_state.final_pool[batch_key] = final_qs
                                    st.session_state.confirmed_batches.add(batch_key)

                                    # 寫入已使用句子到 Google Sheets
                                    sentences_to_save = []
                                    for q in final_qs:
                                        # 找出原始句子（包含 🟨 符號）
                                        original_sentence = None
                                        for word_data in word_dict.values():
                                            if word_data.get("original"):
                                                if word_data["original"] == q["Content"]:
                                                    original_sentence = word_data["original"]
                                                    break
                                            if q["Content"] in word_data.get("ai", []):
                                                # 如果是 AI 句子，需要找到帶 🟨 的原始版本
                                                for original_idx in word_data.get("row_indices", []):
                                                    if original_idx < len(review_df):
                                                        original_row = review_df.iloc[original_idx]
                                                        orig_sent = original_row.get("句子", "").strip()
                                                        if q["Content"] in orig_sent:
                                                            original_sentence = orig_sent
                                                            break
                                                if original_sentence:
                                                    break

                                        sentences_to_save.append({
                                            "school": q["School"],
                                            "level": q["Level"],
                                            "word": q["Word"],
                                            "sentence": original_sentence if original_sentence else q["Content"]
                                        })

                                    # 寫入到「已使用」工作表
                                    if sentences_to_save:
                                        write_ok, write_msg = write_used_sentences(sentences_to_save)
                                        if write_ok:
                                            st.success(f"✅ 已記錄 {len(sentences_to_save)} 個句子到「已使用」工作表")
                                        else:
                                            st.error(f"❌ 寫入失敗：{write_msg}")
                                            st.info("💡 請確保 Google Service Account 有試算表的編輯權限")

                                st.success("✅ 已成功鎖定題庫並記錄使用！")
                                st.rerun()
                        else:
                            st.caption("請勾選上方確認方塊以啟用鎖定按鈕")

                elif batch_key in st.session_state.confirmed_batches:
                    st.success("✅ 此批次已完成審核並鎖定。")

# ============================================================
# --- 標籤頁 2: 預覽下載 ---
# ============================================================

with tab_preview:
    st.subheader("📄 預覽下載")

    level_batches = {k: v for k, v in st.session_state.final_pool.items() if k.endswith(f"||{selected_level}")}

    if not level_batches:
        with st.container(border=True):
            st.warning("⚠️ 尚未有任何批次完成 AI 審核並鎖定題庫。")
            st.info("請先到「AI 句子審核」標籤頁完成審核並鎖定題庫後，再回到此處下載工作紙。")
        st.stop()

    for batch_key, questions in level_batches.items():
        with st.container(border=True):
            school, level = batch_key.split("||")
            st.markdown(f"### 🏫 {school}（{level}）")
            st.caption(f"共 {len(questions)} 題")

            # 生成 PDF（加入載入狀態）
            with st.spinner("正在生成 PDF..."):
                pdf_bytes = create_pdf(school, level, questions)
                answer_pdf_bytes = create_answer_pdf(school, level, questions)

            # 下載按鈕區塊
            col1, col2 = st.columns(2)

            with col1:
                st.download_button(
                    label="⬇️ 下載學生版 PDF",
                    data=pdf_bytes,
                    file_name=f"{school}_{level}_worksheet.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    help="下載學生版本的工作紙 PDF"
                )

            with col2:
                st.download_button(
                    label="⬇️ 下載教師版 PDF（答案）",
                    data=answer_pdf_bytes,
                    file_name=f"{school}_{level}_answers.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    help="下載包含答案的教師版 PDF"
                )

            # 預覽區塊
            with st.expander("📘 預覽學生版 PDF", expanded=False):
                display_pdf_as_images(pdf_bytes)

# ============================================================
# --- 標籤頁 3: 寄送郵件 ---
# ============================================================

with tab_email:
    st.subheader("✉️ 寄送郵件")

    if student_df.empty:
        with st.container(border=True):
            st.error("❌ 學生資料表為空，無法寄送。")
            st.info("請檢查 Google Sheets 中的「學生資料」工作表是否正確設定。")
        st.stop()

    df_level = student_df[student_df["年級"].astype(str) == selected_level]

    if df_level.empty:
        with st.container(border=True):
            st.warning(f"⚠️ {selected_level} 沒有學生資料。")
            st.info("請確認該年級的學生資料是否存在於「學生資料」工作表中。")
        st.stop()

    # 學生選擇區塊
    with st.container(border=True):
        st.markdown("### 👤 選擇學生")

        student_names = df_level["學生姓名"].tolist()
        selected_student = st.selectbox(
            "選擇學生",
            [""] + student_names,
            help="選擇要寄送工作紙的學生"
        )

    if not selected_student:
        st.info("👆 請從上方選擇一位學生")
        st.stop()

    row = df_level[df_level["學生姓名"] == selected_student].iloc[0]
    school = row["學校"]
    grade = row["年級"]

    parent_email = row.get("家長 Email", "")
    cc_email = row.get("老師 Email", "")

    batch_key = f"{school}||{grade}"

    if batch_key not in st.session_state.final_pool:
        with st.container(border=True):
            st.error("⚠️ 此學生所屬批次尚未完成 AI 審核並鎖定題庫。")
            st.info("請先到「AI 句子審核」標籤頁完成審核並鎖定題庫。")
        st.stop()

    questions = st.session_state.final_pool[batch_key]

    # PDF 生成區塊
    with st.container(border=True):
        st.markdown("### 📄 工作紙預覽")

        with st.spinner("正在生成 PDF..."):
            pdf_bytes = create_pdf(school, grade, questions, student_name=selected_student)

        st.download_button(
            label="⬇️ 下載學生版 PDF",
            data=pdf_bytes,
            file_name=f"{selected_student}_worksheet.pdf",
            mime="application/pdf",
            use_container_width=True
        )

    st.divider()

    # 郵件寄送區塊
    with st.container(border=True):
        st.markdown("### ✉️ 寄送工作紙")

        # 顯示寄送資訊摘要
        with st.expander("📋 寄送資訊摘要", expanded=True):
            st.markdown(f"""
            - **學生姓名**：{selected_student}
            - **學校**：{school}
            - **年級**：{grade}
            - **家長電郵**：{parent_email if parent_email else '（未提供）'}
            - **老師電郵**：{cc_email if cc_email else '（未提供）'}
            """)

        # 二次確認機制
        st.markdown("#### ⚠️ 確認寄送")

        if not parent_email or parent_email.lower() in ["n/a", "nan", "", "none"]:
            st.error("❌ 該學生的家長電郵地址為空，無法寄送。")
            st.stop()

        confirm_email = st.checkbox(
            f"我確認要將工作紙寄送至以下電郵：{parent_email}",
            key="email_confirm_checkbox"
        )

        if not confirm_email:
            st.caption("請勾選上方確認方塊以啟用寄送按鈕")
            st.stop()

        # 寄送按鈕
        if st.button("📨 寄出工作紙", type="primary", use_container_width=True):
            with st.spinner("正在發送郵件，請稍候..."):
                ok, msg = send_email_with_pdf(
                    parent_email,
                    selected_student,
                    school,
                    grade,
                    pdf_bytes,
                    cc_email=cc_email
                )

            if ok:
                st.success("🎉 已成功寄出工作紙！")
                st.balloons()
                st.toast(f"工作紙已成功寄送給 {selected_student} 的家長！", icon="✅")
            else:
                st.error(f"❌ 寄送失敗：{msg}")
                st.info("請檢查網路連線或稍後再試。")

# ============================================================
# --- End of App ---
# ============================================================

st.write("")
st.write("© 2026 校本填充工作紙生成器 — 自動化教學工具")
