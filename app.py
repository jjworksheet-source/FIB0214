# main.py — Full Production Code
# Architecture: Google Form → GAS → Review Sheet → Streamlit (One-Stop) → PDF/Email
# Review Sheet Status Flow: Ready/Pending → Loaded → Sent

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import datetime
import io
import os
import re
import base64
from pdf2image import convert_from_bytes
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import (
    Mail, Attachment, FileContent, FileName, FileType, Disposition, Email
)
from python_http_client.exceptions import HTTPError

# ============================================================
# 1. PAGE CONFIG
# ============================================================
st.set_page_config(page_title="Worksheet Admin", page_icon="🎯", layout="wide")

st.markdown("""
<style>
[data-testid="stSidebar"] { background-color: #f0f4f8; }
.stTabs [data-baseweb="tab"] { font-size:16px; font-weight:600; padding:10px 20px; }
.word-card {
    background:#fff; border:1px solid #dee2e6; border-radius:12px;
    padding:16px 20px; margin-bottom:12px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
.badge-db      { background:#d4edda; color:#155724; padding:3px 10px; border-radius:20px; font-size:13px; font-weight:600; }
.badge-ai      { background:#fff3cd; color:#856404; padding:3px 10px; border-radius:20px; font-size:13px; font-weight:600; }
.badge-pending { background:#f8d7da; color:#721c24; padding:3px 10px; border-radius:20px; font-size:13px; font-weight:600; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 2. FONT SETUP
# ============================================================
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.enums import TA_CENTER

    CHINESE_FONT = None
    for path in ["Kai.ttf",
                 "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
                 "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf"]:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("ChineseFont", path))
                CHINESE_FONT = "ChineseFont"
                break
            except Exception:
                continue
except ImportError:
    st.error("❌ reportlab not found. Add 'reportlab' to requirements.txt")
    st.stop()

# ============================================================
# 3. GOOGLE SHEETS CONNECTION
# ============================================================
@st.cache_resource
def get_gspread_client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/drive"]
    )
    return gspread.authorize(creds)

try:
    gc       = get_gspread_client()
    SHEET_ID = st.secrets["app_config"]["spreadsheet_id"]
except Exception as e:
    st.error(f"❌ Connection Error: {e}")
    st.stop()

# ============================================================
# 4. DATA LOADERS
# ============================================================
@st.cache_data(ttl=30)
def load_review() -> pd.DataFrame:
    """Load rows from Review sheet that are Ready or Pending (not Loaded/Sent)."""
    try:
        sh  = gc.open_by_key(SHEET_ID)
        ws  = sh.worksheet("Review")
        df  = pd.DataFrame(ws.get_all_records())
        if df.empty:
            return df
        df.columns = [c.strip() for c in df.columns]
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Error loading Review sheet: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=30)
def load_students() -> pd.DataFrame:
    try:
        sh  = gc.open_by_key(SHEET_ID)
        ws  = sh.worksheet("學生資料")
        df  = pd.DataFrame(ws.get_all_records())
        if df.empty:
            return df
        df.columns = [c.strip() for c in df.columns]
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Error loading 學生資料: {e}")
        return pd.DataFrame()

def clear_cache():
    load_review.clear()
    load_students.clear()

# ============================================================
# 5. GOOGLE SHEETS WRITE-BACK
# ============================================================
def mark_rows_in_review(timestamps: list[str], new_status: str,
                         sentence_updates: dict = None):
    """
    Update 狀態 column (col G = 7) for rows matching given Timestamps.
    Optionally update 句子 column (col E = 5) via sentence_updates = {timestamp: sentence}.
    """
    try:
        sh  = gc.open_by_key(SHEET_ID)
        ws  = sh.worksheet("Review")
        all_vals = ws.get_all_values()   # list of lists, row 0 = header

        # Build col index map from header
        header = [h.strip() for h in all_vals[0]]
        ts_col     = header.index("Timestamp") + 1   # 1-based
        status_col = header.index("狀態")      + 1
        sentence_col = header.index("句子")    + 1

        updates = []
        for i, row in enumerate(all_vals[1:], start=2):   # row 2 onward
            ts = str(row[ts_col - 1]).strip()
            if ts in timestamps:
                updates.append({"range": f"{chr(64+status_col)}{i}",
                                 "values": [[new_status]]})
                if sentence_updates and ts in sentence_updates:
                    updates.append({"range": f"{chr(64+sentence_col)}{i}",
                                     "values": [[sentence_updates[ts]]]})

        if updates:
            ws.batch_update(updates)
        return True
    except Exception as e:
        st.error(f"Google Sheets update error: {e}")
        return False

# ============================================================
# 6. PDF BUILDER
# ============================================================
def create_pdf(school: str, level: str, questions: list,
               student_name: str = None) -> bytes:
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=letter)
    styles = getSampleStyleSheet()
    fn = CHINESE_FONT or "Helvetica"

    title_style = ParagraphStyle("T", parent=styles["Heading1"], fontName=fn,
                                 fontSize=20, alignment=TA_CENTER, spaceAfter=12)
    body_style  = ParagraphStyle("B", parent=styles["Normal"], fontName=fn,
                                 fontSize=14, leading=20,
                                 leftIndent=25, firstLineIndent=-25)

    title_text = (f"<b>{school} ({level}) - {student_name} - 校本填充工作紙</b>"
                  if student_name
                  else f"<b>{school} ({level}) - 校本填充工作紙</b>")

    story = [
        Paragraph(title_text, title_style),
        Spacer(1, 0.2 * inch),
        Paragraph(f"日期: {datetime.date.today() + datetime.timedelta(days=1)}", body_style),
        Spacer(1, 0.3 * inch),
    ]
    for i, row in enumerate(questions):
        content = str(row.get("句子", row.get("Content", "")))
        content = re.sub(r'【】(.+?)【】', r'<u>\1</u>', content)
        content = re.sub(r'【(.+?)】',    r'<u>\1</u>', content)
        story.append(Paragraph(f"{i+1}. {content}", body_style))
        story.append(Spacer(1, 0.15 * inch))

    doc.build(story)
    bio.seek(0)
    return bio.getvalue()

# ============================================================
# 7. EMAIL SENDER
# ============================================================
def send_email_with_pdf(to_email, student_name, school, grade,
                         pdf_bytes, cc_email=None):
    try:
        cfg       = st.secrets["sendgrid"]
        recipient = str(to_email).strip()
        if not re.match(r'^[\w\.-]+@[\w\.-]+\.\w+$', recipient):
            return False, f"無效電郵格式: '{recipient}'"

        safe_name = re.sub(r'[^\w\-]', '_', str(student_name).strip())
        msg = Mail(
            from_email=Email(cfg["from_email"], cfg.get("from_name", "")),
            to_emails=recipient,
            subject=f"【工作紙】{school} ({grade}) - {student_name} 的校本填充練習",
            html_content=f"""
                <p>親愛的家長您好：</p>
                <p>附件為 <strong>{student_name}</strong> 同學在
                <strong>{school} ({grade})</strong> 的校本填充工作紙。</p>
                <p>請下載並列印供同學練習。祝 學習愉快！</p>
                <br><p>-- 自動發送系統 --</p>
            """
        )
        if cc_email:
            cc = str(cc_email).strip().lower()
            if cc not in ["n/a", "nan", "", "none"] and "@" in cc and cc != recipient.lower():
                msg.add_cc(cc)

        encoded = base64.b64encode(pdf_bytes).decode()
        msg.add_attachment(Attachment(
            FileContent(encoded),
            FileName(f"{safe_name}_Worksheet.pdf"),
            FileType("application/pdf"),
            Disposition("attachment")
        ))
        resp = SendGridAPIClient(cfg["api_key"]).send(msg)
        return (True, "發送成功") if 200 <= resp.status_code < 300 \
               else (False, f"HTTP {resp.status_code}")
    except HTTPError as e:
        try:    return False, e.body.decode("utf-8")
        except: return False, str(e)
    except Exception as e:
        return False, str(e)

# ============================================================
# 8. PDF PREVIEW
# ============================================================
def show_pdf_preview(pdf_bytes: bytes):
    try:
        images = convert_from_bytes(pdf_bytes, dpi=150)
        for i, img in enumerate(images):
            st.image(img, caption=f"Page {i+1}", use_container_width=True)
    except Exception as e:
        st.warning(f"Preview unavailable: {e}")

# ============================================================
# 9. SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("## 🎯 Worksheet Admin")
    st.divider()

    if not CHINESE_FONT:
        st.error("⚠️ Chinese font not found.\nAdd Kai.ttf to your repo root.")
    else:
        st.success("✅ Font OK")

    st.divider()
    if st.button("🔄 Refresh Data", use_container_width=True):
        clear_cache()
        st.rerun()
    st.caption("Data auto-refreshes every 30 seconds.")

    st.divider()
    st.markdown("### 📊 Status Legend")
    st.markdown("""
- 🟢 **Ready** — DB 句子，可直接使用
- 🟡 **Pending** — AI 句子，需要審批
- 🔵 **Loaded** — 已被 App 取走處理中
- ✅ **Sent** — 已發送，不再顯示
""")

# ============================================================
# 10. LOAD DATA
# ============================================================
st.title("🎯 Worksheet Admin")

raw_review  = load_review()
student_df  = load_students()

# ============================================================
# 11. VALIDATE REVIEW SHEET
# ============================================================
REQUIRED_COLS = ["Timestamp", "學校", "年級", "詞語", "句子", "來源", "狀態"]

if raw_review.empty:
    st.info("📭 Review 表目前沒有資料。等待老師填寫 Google Form。")
    st.stop()

missing = [c for c in REQUIRED_COLS if c not in raw_review.columns]
if missing:
    st.error(f"❌ Review 表缺少欄位：{missing}")
    st.write("現有欄位：", raw_review.columns.tolist())
    st.stop()

# Filter: only show Ready + Pending (not Loaded / Sent)
active_df = raw_review[raw_review["狀態"].isin(["Ready", "Pending"])].copy()

if active_df.empty:
    st.success("🎉 目前沒有待處理的詞語。所有資料已發送或正在處理中。")
    st.stop()

# ============================================================
# 12. LEVEL & SCHOOL SELECTOR (Sidebar-style inside main)
# ============================================================
col_ctrl, col_main = st.columns([1, 3])

with col_ctrl:
    st.markdown("### ⚙️ 篩選")
    levels  = sorted(active_df["年級"].astype(str).unique().tolist())
    sel_lvl = st.selectbox("年級", levels, key="sel_level")

    lvl_df  = active_df[active_df["年級"] == sel_lvl]
    schools = sorted(lvl_df["學校"].astype(str).unique().tolist())
    sel_sch = st.selectbox("學校", schools, key="sel_school")

    lot_df  = lvl_df[lvl_df["學校"] == sel_sch].copy()

    # Stats
    n_ready   = len(lot_df[lot_df["狀態"] == "Ready"])
    n_pending = len(lot_df[lot_df["狀態"] == "Pending"])
    st.metric("🟢 Ready (DB)", n_ready)
    st.metric("🟡 Pending (AI)", n_pending)

    st.divider()
    send_mode = st.radio("發送模式", ["📄 預覽 & 下載", "📧 按學生寄送"], key="send_mode")

# ============================================================
# 13. MAIN PANEL — WORD CARDS
# ============================================================
with col_main:
    st.markdown(f"### 📋 {sel_sch} · {sel_lvl} 詞語清單")

    if lot_df.empty:
        st.info("此學校/年級沒有待處理的詞語。")
        st.stop()

    # --- Session state: store final chosen sentences ---
    # Key: timestamp, Value: final sentence string
    if "chosen" not in st.session_state:
        st.session_state["chosen"] = {}

    # Reset chosen if school/level changed
    state_key = f"{sel_sch}_{sel_lvl}"
    if st.session_state.get("last_lot") != state_key:
        st.session_state["chosen"] = {}
        st.session_state["last_lot"] = state_key

    words = lot_df["詞語"].unique().tolist()
    all_ready = True   # track if all AI words have been approved

    for word in words:
        word_rows = lot_df[lot_df["詞語"] == word]
        source    = str(word_rows.iloc[0]["來源"]).strip()
        status    = str(word_rows.iloc[0]["狀態"]).strip()
        ts        = str(word_rows.iloc[0]["Timestamp"]).strip()
        # Use DataFrame index as unique key suffix to avoid duplicate key errors
        row_idx   = word_rows.index[0]

        if source == "DB":
            badge = '<span class="badge-db">📗 資料庫</span>'
        elif status == "Pending":
            badge = '<span class="badge-pending">⏳ AI 待審批</span>'
            all_ready = False
        else:
            badge = '<span class="badge-ai">🤖 AI 已審批</span>'

        st.markdown(f"""
        <div class="word-card">
            <b style="font-size:17px">{word}</b>&nbsp;&nbsp;{badge}
        </div>
        """, unsafe_allow_html=True)

        if source == "DB":
            # DB: single sentence, editable, auto-approved
            sentence = str(word_rows.iloc[0]["句子"]).strip()
            final = st.text_area(
                f"句子（可修改）", value=sentence,
                key=f"db_{row_idx}", height=75, label_visibility="collapsed"
            )
            st.session_state["chosen"][ts] = final

        else:
            # AI: radio select among options + optional manual override
            options = word_rows["句子"].astype(str).tolist()
            chosen_opt = st.radio(
                "選擇 AI 句子", options,
                key=f"rad_{row_idx}", horizontal=False
            )
            override = st.text_input(
                "✏️ 手動輸入（留空則使用上方選擇）",
                value="", placeholder=chosen_opt,
                key=f"ovr_{row_idx}"
            )
            final = override.strip() if override.strip() else chosen_opt
            st.session_state["chosen"][ts] = final

            if status == "Pending":
                all_ready = False   # still needs explicit approval

        st.write("")  # spacing

    # ============================================================
    # 14. MARK AS LOADED BUTTON
    # ============================================================
    st.divider()

    if not all_ready:
        st.warning("⚠️ 仍有 AI 句子未選定。請在上方為每個 AI 詞語選擇句子後再繼續。")

    # Build final questions list from session state
    def build_questions() -> list:
        rows = []
        for word in words:
            word_rows = lot_df[lot_df["詞語"] == word]
            ts = str(word_rows.iloc[0]["Timestamp"]).strip()
            sentence = st.session_state["chosen"].get(ts, str(word_rows.iloc[0]["句子"]))
            rows.append({
                "詞語": word,
                "句子": sentence,
                "Timestamp": ts,
                "學校": sel_sch,
                "年級": sel_lvl,
            })
        return rows

    def mark_lot_loaded():
        """Mark all words in this lot as Loaded in Review sheet."""
        timestamps = [str(lot_df.iloc[i]["Timestamp"]).strip()
                      for i in range(len(lot_df))]
        sentence_updates = {ts: st.session_state["chosen"].get(ts, "")
                            for ts in timestamps}
        return mark_rows_in_review(timestamps, "Loaded",
                                   sentence_updates=sentence_updates)

    def mark_lot_sent():
        """Mark all words in this lot as Sent in Review sheet."""
        timestamps = [str(lot_df.iloc[i]["Timestamp"]).strip()
                      for i in range(len(lot_df))]
        return mark_rows_in_review(timestamps, "Sent")

    # ============================================================
    # 15A. MODE: PREVIEW & DOWNLOAD
    # ============================================================
    if send_mode == "📄 預覽 & 下載":
        if st.button("📄 生成 PDF 預覽", use_container_width=True,
                     disabled=not all_ready, type="primary"):
            questions = build_questions()
            pdf_bytes = create_pdf(sel_sch, sel_lvl, questions)

            # Mark as Loaded immediately
            with st.spinner("更新 Review 表狀態為 Loaded..."):
                mark_lot_loaded()
                clear_cache()

            st.download_button(
                label=f"📥 下載 {sel_sch}_{sel_lvl}.pdf",
                data=pdf_bytes,
                file_name=f"{sel_sch}_{sel_lvl}_{datetime.date.today()}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
            st.markdown("#### 🔍 PDF 預覽")
            show_pdf_preview(pdf_bytes)

            if st.button("✅ 確認完成，標記為 Sent", use_container_width=True):
                with st.spinner("更新狀態為 Sent..."):
                    mark_lot_sent()
                    clear_cache()
                st.success("✅ 已標記為 Sent，下次不再顯示。")
                st.rerun()

    # ============================================================
    # 15B. MODE: SEND BY STUDENT
    # ============================================================
    else:
        st.markdown("#### 👨‍👩‍👧 按學生寄送")

        if student_df.empty:
            st.error("❌ 無法讀取「學生資料」工作表。")
            st.stop()

        req_cols = ["學校", "年級", "狀態", "學生姓名", "家長 Email"]
        miss     = [c for c in req_cols if c not in student_df.columns]
        if miss:
            st.error(f"「學生資料」缺少欄位：{miss}")
            st.stop()

        active_students = student_df[student_df["狀態"] == "Y"]
        matched = active_students[
            (active_students["學校"] == sel_sch) &
            (active_students["年級"] == sel_lvl)
        ]

        if matched.empty:
            st.warning("⚠️ 沒有符合此學校/年級的學生（狀態 = Y）。")
            with st.expander("🔍 排查資料"):
                st.write("Review 學校:", sel_sch, "| 年級:", sel_lvl)
                st.write("學生資料 學校:", active_students["學校"].unique().tolist())
                st.write("學生資料 年級:", active_students["年級"].unique().tolist())
        else:
            st.success(f"✅ 找到 {len(matched)} 位學生")

            questions = build_questions()
            sent_all  = []

            for _, student in matched.iterrows():
                student_name  = student["學生姓名"]
                parent_email  = student["家長 Email"]
                teacher_email = student.get("老師 Email", None)

                pdf_bytes = create_pdf(sel_sch, sel_lvl, questions,
                                       student_name=student_name)

                with st.container(border=True):
                    c1, c2 = st.columns([1, 2])
                    with c1:
                        st.markdown(f"**👤 {student_name}**")
                        st.caption(f"📧 {parent_email}")
                        if teacher_email:
                            st.caption(f"👩‍🏫 CC: {teacher_email}")

                        st.download_button(
                            "📥 下載 PDF", data=pdf_bytes,
                            file_name=f"{student_name}_{sel_lvl}_{datetime.date.today()}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                            key=f"dl_{parent_email}",
                            disabled=not all_ready
                        )

                        if st.button(f"📧 寄送給家長",
                                     key=f"send_{parent_email}",
                                     use_container_width=True,
                                     disabled=not all_ready):
                            with st.spinner(f"寄送給 {parent_email}..."):
                                # Mark Loaded before sending
                                mark_lot_loaded()
                                ok, msg = send_email_with_pdf(
                                    parent_email, student_name,
                                    sel_sch, sel_lvl, pdf_bytes,
                                    cc_email=teacher_email
                                )
                            if ok:
                                st.success(f"✅ 已寄出！")
                                sent_all.append(parent_email)
                            else:
                                st.error(f"❌ {msg}")
                    with c2:
                        if all_ready:
                            show_pdf_preview(pdf_bytes)
                        else:
                            st.info("請先完成所有 AI 句子審批才能預覽。")

            # After all sent, mark Sent
            if sent_all and len(sent_all) == len(matched):
                st.divider()
                if st.button("✅ 全部已寄出，標記為 Sent", use_container_width=True, type="primary"):
                    with st.spinner("更新 Review 表..."):
                        mark_lot_sent()
                        clear_cache()
                    st.success("✅ 已標記為 Sent，下次不再顯示。")
                    st.rerun()
