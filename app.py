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
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition, Email
from python_http_client.exceptions import HTTPError

# ============================================================
# 1. PAGE CONFIG & CUSTOM CSS
# ============================================================
st.set_page_config(page_title="Worksheet Admin", page_icon="🎯", layout="wide")

st.markdown("""
<style>
[data-testid="stSidebar"] { background-color: #f0f4f8; }
.stTabs [data-baseweb="tab"] {
    font-size: 16px; font-weight: 600; padding: 10px 20px;
}
.word-card {
    background: #ffffff;
    border: 1px solid #dee2e6;
    border-radius: 12px;
    padding: 18px 22px;
    margin-bottom: 14px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
.badge-db   { background:#d4edda; color:#155724; padding:3px 10px; border-radius:20px; font-size:13px; font-weight:600; }
.badge-ai   { background:#fff3cd; color:#856404; padding:3px 10px; border-radius:20px; font-size:13px; font-weight:600; }
.badge-done { background:#cce5ff; color:#004085; padding:3px 10px; border-radius:20px; font-size:13px; font-weight:600; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 2. FONT & PDF SETUP
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
    for path in ["Kai.ttf", "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
                 "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf"]:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont('ChineseFont', path))
                CHINESE_FONT = 'ChineseFont'
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
def get_client():
    key_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(
        key_dict,
        scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    )
    return gspread.authorize(creds)

try:
    gc = get_client()
    SHEET_ID = st.secrets["app_config"]["spreadsheet_id"]
except Exception as e:
    st.error(f"❌ Connection Error: {e}")
    st.stop()

# ============================================================
# 4. DATA LOADERS
# ============================================================
@st.cache_data(ttl=60)
def load_review():
    sh = gc.open_by_key(SHEET_ID)
    df = pd.DataFrame(sh.worksheet("Review").get_all_records())
    if df.empty:
        return df
    df.columns = [c.strip() for c in df.columns]
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
    return df

@st.cache_data(ttl=60)
def load_standby():
    sh = gc.open_by_key(SHEET_ID)
    df = pd.DataFrame(sh.worksheet("standby").get_all_records())
    if df.empty:
        return df
    df.columns = [c.strip() for c in df.columns]
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
    return df

@st.cache_data(ttl=60)
def load_students():
    sh = gc.open_by_key(SHEET_ID)
    df = pd.DataFrame(sh.worksheet("學生資料").get_all_records())
    if df.empty:
        return df
    df.columns = [c.strip() for c in df.columns]
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
    return df

def clear_all_cache():
    load_review.clear()
    load_standby.clear()
    load_students.clear()

# ============================================================
# 5. WRITE-BACK HELPERS
# ============================================================
def move_word_to_standby(review_row: dict, final_sentence: str) -> tuple[bool, str]:
    """Write one approved word to standby and mark Review row as Transferred."""
    try:
        sh = gc.open_by_key(SHEET_ID)
        standby_ws = sh.worksheet("standby")
        review_ws  = sh.worksheet("Review")

        now_str    = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        unique_id  = f"ID_{datetime.datetime.now().strftime('%m%d%H%M%S%f')[-12:]}"

        # Append to standby: ID, School, Grade, Word, Type, Content, Answer, Status, Date
        standby_ws.append_row([
            unique_id,
            review_row.get("學校", ""),
            review_row.get("年級", ""),
            review_row.get("詞語", ""),
            "填空題",
            final_sentence,
            review_row.get("詞語", ""),   # Answer = the word itself
            "Ready",
            now_str
        ])

        # Mark Review row as Transferred using Timestamp as key
        ts = str(review_row.get("Timestamp", "")).strip()
        if ts:
            cell = review_ws.find(ts)
            if cell:
                review_ws.update_cell(cell.row, 7, "Transferred")  # Column G = 狀態

        return True, "OK"
    except Exception as e:
        return False, str(e)

# ============================================================
# 6. PDF BUILDER
# ============================================================
def create_pdf(school_name: str, level: str, questions: list, student_name: str = None) -> bytes:
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=letter)
    styles = getSampleStyleSheet()
    fn = CHINESE_FONT or "Helvetica"

    title_style = ParagraphStyle("T", parent=styles["Heading1"], fontName=fn,
                                 fontSize=20, alignment=TA_CENTER, spaceAfter=12)
    body_style  = ParagraphStyle("B", parent=styles["Normal"], fontName=fn,
                                 fontSize=14, leading=20, leftIndent=25, firstLineIndent=-25)

    title_text = (f"<b>{school_name} ({level}) - {student_name} - 校本填充工作紙</b>"
                  if student_name else f"<b>{school_name} ({level}) - 校本填充工作紙</b>")

    story = [
        Paragraph(title_text, title_style),
        Spacer(1, 0.2 * inch),
        Paragraph(f"日期: {datetime.date.today() + datetime.timedelta(days=1)}", body_style),
        Spacer(1, 0.3 * inch),
    ]

    for i, row in enumerate(questions):
        content = str(row.get("Content", ""))
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
def send_email_with_pdf(to_email, student_name, school_name, grade, pdf_bytes, cc_email=None):
    try:
        sg_cfg    = st.secrets["sendgrid"]
        recipient = str(to_email).strip()
        if not re.match(r'^[\w\.-]+@[\w\.-]+\.\w+$', recipient):
            return False, f"無效電郵格式: '{recipient}'"

        safe_name = re.sub(r'[^\w\-]', '_', str(student_name).strip())
        msg = Mail(
            from_email=Email(sg_cfg["from_email"], sg_cfg.get("from_name", "")),
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
            cc = str(cc_email).strip().lower()
            if cc not in ["n/a", "nan", "", "none"] and "@" in cc and cc != recipient.lower():
                msg.add_cc(cc)

        encoded = base64.b64encode(pdf_bytes).decode()
        msg.add_attachment(Attachment(
            FileContent(encoded), FileName(f"{safe_name}_Worksheet.pdf"),
            FileType("application/pdf"), Disposition("attachment")
        ))

        resp = SendGridAPIClient(sg_cfg["api_key"]).send(msg)
        return (True, "發送成功") if 200 <= resp.status_code < 300 else (False, f"HTTP {resp.status_code}")

    except HTTPError as e:
        try:    return False, e.body.decode("utf-8")
        except: return False, str(e)
    except Exception as e:
        return False, str(e)

# ============================================================
# 8. PDF PREVIEW HELPER
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
    st.image("https://placehold.co/260x60/4A90D9/white?text=Worksheet+Admin", use_container_width=True)
    st.divider()

    if not CHINESE_FONT:
        st.error("⚠️ Chinese font not found.\nPlease add Kai.ttf to your repo.")
    else:
        st.success("✅ Font OK")

    st.divider()
    if st.button("🔄 Refresh All Data", use_container_width=True):
        clear_all_cache()
        st.rerun()

    st.caption("Data auto-refreshes every 60 seconds.")

# ============================================================
# 10. MAIN TABS
# ============================================================
st.title("🎯 Worksheet Admin")
tab_review, tab_generate = st.tabs(["📥  Step 1 — 審批新詞", "📄  Step 2 — 生成工作紙"])

# ============================================================
# TAB 1 — REVIEW & APPROVAL
# ============================================================
with tab_review:
    st.subheader("審批新詞語 · 移交至題庫")
    st.caption("從 Google Form 自動進入 Review 表的詞語，在這裡選句、修改，然後移交至 Standby 題庫。")

    review_df = load_review()

    if review_df.empty:
        st.info("📭 Review 表目前沒有資料。")
    else:
        # Only show non-transferred rows
        pending_df = review_df[review_df.get("狀態", review_df.get("Status", pd.Series(dtype=str))).astype(str).str.strip() != "Transferred"].copy()

        if pending_df.empty:
            st.success("🎉 所有詞語已審批完成！")
        else:
            # Level selector
            levels = sorted(pending_df["年級"].astype(str).unique().tolist())
            sel_level = st.selectbox("📚 選擇年級", levels, key="review_level")

            level_data = pending_df[pending_df["年級"].astype(str) == sel_level]
            words = level_data["詞語"].unique().tolist()

            # Stats row
            c1, c2, c3 = st.columns(3)
            c1.metric("待審批詞語", len(words))
            c2.metric("DB 句子", len(level_data[level_data["來源"] == "DB"]["詞語"].unique()))
            c3.metric("AI 句子", len(level_data[level_data["來源"] == "AI"]["詞語"].unique()))

            st.divider()

            # --- Word Cards ---
            for word in words:
                word_rows = level_data[level_data["詞語"] == word]
                source    = str(word_rows.iloc[0].get("來源", "AI")).strip()
                school    = str(word_rows.iloc[0].get("學校", "")).strip()
                ts        = str(word_rows.iloc[0].get("Timestamp", "")).strip()

                badge = (f'<span class="badge-db">📗 資料庫</span>' if source == "DB"
                         else f'<span class="badge-ai">🤖 AI 生成</span>')

                st.markdown(f"""
                <div class="word-card">
                    <b style="font-size:18px">{word}</b>&nbsp;&nbsp;{badge}
                    &nbsp;&nbsp;<span style="color:#888;font-size:13px">🏫 {school} · {sel_level}</span>
                </div>
                """, unsafe_allow_html=True)

                with st.container():
                    if source == "DB":
                        # Single sentence — just confirm
                        content = str(word_rows.iloc[0].get("句子", "")).strip()
                        final   = st.text_area("✏️ 確認句子（可修改）", value=content, key=f"ta_{word}_{ts}", height=80)
                        if st.button(f"✅ 移交「{word}」", key=f"btn_{word}_{ts}", type="primary"):
                            with st.spinner("移交中..."):
                                ok, msg = move_word_to_standby(word_rows.iloc[0].to_dict(), final)
                            if ok:
                                st.toast(f"✅ 「{word}」已移交！", icon="🎉")
                                clear_all_cache()
                                st.rerun()
                            else:
                                st.error(f"移交失敗：{msg}")

                    else:
                        # Multiple AI options — radio select
                        options = word_rows["句子"].astype(str).tolist()
                        chosen  = st.radio("選擇最合適的 AI 句子", options, key=f"rad_{word}_{ts}")
                        final   = st.text_area("✏️ 手動微調（選填，留空則使用上方選擇）",
                                               value="", placeholder=chosen,
                                               key=f"ta_{word}_{ts}", height=80)
                        use_sentence = final.strip() if final.strip() else chosen

                        if st.button(f"🚀 批准並移交「{word}」", key=f"btn_{word}_{ts}", type="primary"):
                            with st.spinner("移交中..."):
                                ok, msg = move_word_to_standby(word_rows.iloc[0].to_dict(), use_sentence)
                            if ok:
                                st.toast(f"✅ 「{word}」已移交！", icon="🎉")
                                clear_all_cache()
                                st.rerun()
                            else:
                                st.error(f"移交失敗：{msg}")

                st.write("")  # spacing

# ============================================================
# TAB 2 — GENERATE WORKSHEETS
# ============================================================
with tab_generate:
    st.subheader("生成工作紙 · 下載或寄送")
    st.caption("從 Standby 題庫讀取已審批的題目，生成 PDF 並寄送給家長。")

    standby_df = load_standby()
    student_df = load_students()

    if standby_df.empty:
        st.warning("⚠️ Standby 題庫是空的。請先在 Step 1 完成審批移交。")
        st.stop()

    # Normalize column names
    col_map = {"Grade": "Level", "grade": "Level", "level": "Level",
               "school": "School", "word": "Word", "content": "Content", "status": "Status"}
    standby_df = standby_df.rename(columns={k: v for k, v in col_map.items() if k in standby_df.columns})

    required = ["School", "Level", "Word", "Content", "Status"]
    missing  = [c for c in required if c not in standby_df.columns]
    if missing:
        st.error(f"Standby 表缺少欄位：{missing}。現有欄位：{standby_df.columns.tolist()}")
        st.stop()

    # Normalize status
    standby_df["_status_clean"] = (standby_df["Status"].astype(str)
                                   .str.replace("\u00A0", " ").str.replace("\u3000", " ").str.strip())
    ready_df = standby_df[standby_df["_status_clean"].isin(["Ready", "Waiting"])]

    if ready_df.empty:
        st.info("Standby 中沒有 Ready/Waiting 的題目。")
        st.stop()

    # --- Sidebar-style controls inside tab ---
    ctrl_col, main_col = st.columns([1, 2])

    with ctrl_col:
        st.markdown("#### ⚙️ 設定")
        levels_sb = sorted(ready_df["Level"].astype(str).unique().tolist())
        sel_level = st.selectbox("年級", levels_sb, key="gen_level")

        level_ready = ready_df[ready_df["Level"].astype(str) == sel_level]
        schools_sb  = sorted(level_ready["School"].unique().tolist())
        sel_school  = st.selectbox("學校", schools_sb, key="gen_school")

        mode = st.radio("發送模式", ["📄 預覽 & 下載", "📧 按學生寄送"], key="gen_mode")

        school_data = level_ready[level_ready["School"] == sel_school]
        st.metric("題目數", len(school_data))

    with main_col:
        if school_data.empty:
            st.info("請在左側選擇學校。")
        else:
            # Show question list
            with st.expander("📋 查看題目列表", expanded=False):
                st.dataframe(school_data[["Word", "Content"]].reset_index(drop=True),
                             use_container_width=True, hide_index=True)

            # ---- MODE A: Preview & Download ----
            if mode == "📄 預覽 & 下載":
                pdf_bytes = create_pdf(sel_school, sel_level, school_data.to_dict("records"))

                dl_col, _ = st.columns([1, 1])
                with dl_col:
                    st.download_button(
                        label=f"📥 下載 {sel_school}_{sel_level}.pdf",
                        data=pdf_bytes,
                        file_name=f"{sel_school}_{sel_level}_{datetime.date.today()}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )

                st.markdown("#### 🔍 PDF 預覽")
                show_pdf_preview(pdf_bytes)

            # ---- MODE B: Send by Student ----
            else:
                if student_df.empty:
                    st.error("❌ 無法讀取「學生資料」工作表。")
                    st.stop()

                req_cols = ["學校", "年級", "狀態", "學生姓名", "家長 Email"]
                miss     = [c for c in req_cols if c not in student_df.columns]
                if miss:
                    st.error(f"「學生資料」缺少欄位：{miss}")
                    st.stop()

                active = student_df[student_df["狀態"] == "Y"]
                merged = active.merge(school_data, left_on=["學校", "年級"],
                                      right_on=["School", "Level"], how="inner")

                if merged.empty:
                    st.warning("⚠️ 沒有符合條件的學生配對。")
                    with st.expander("🔍 排查資料"):
                        st.write("Standby School:", school_data["School"].unique().tolist())
                        st.write("Standby Level:", school_data["Level"].unique().tolist())
                        st.write("學生資料 學校:", active["學校"].unique().tolist())
                        st.write("學生資料 年級:", active["年級"].unique().tolist())
                else:
                    unique_students = merged["家長 Email"].nunique()
                    st.success(f"✅ 配對到 {unique_students} 位學生")

                    for parent_email, grp in merged.groupby("家長 Email"):
                        student_name  = grp["學生姓名"].iloc[0]
                        school_name   = grp["學校"].iloc[0]
                        grade         = grp["年級"].iloc[0]
                        teacher_email = grp["老師 Email"].iloc[0] if "老師 Email" in grp.columns else None

                        pdf_bytes = create_pdf(school_name, grade, grp.to_dict("records"), student_name=student_name)

                        with st.container(border=True):
                            s1, s2 = st.columns([1, 2])
                            with s1:
                                st.markdown(f"**👤 {student_name}**")
                                st.caption(f"🏫 {school_name} ({grade})")
                                st.caption(f"📧 {parent_email}")
                                if teacher_email:
                                    st.caption(f"👩‍🏫 CC: {teacher_email}")

                                st.download_button(
                                    f"📥 下載 PDF",
                                    data=pdf_bytes,
                                    file_name=f"{student_name}_{grade}_{datetime.date.today()}.pdf",
                                    mime="application/pdf",
                                    use_container_width=True,
                                    key=f"dl_{parent_email}"
                                )
                                if st.button(f"📧 寄送給家長", key=f"send_{parent_email}", use_container_width=True):
                                    with st.spinner("寄送中..."):
                                        ok, msg = send_email_with_pdf(
                                            parent_email, student_name, school_name, grade,
                                            pdf_bytes, cc_email=teacher_email
                                        )
                                    if ok:
                                        st.success("✅ 已寄出！")
                                    else:
                                        st.error(f"❌ {msg}")

                            with s2:
                                show_pdf_preview(pdf_bytes)
