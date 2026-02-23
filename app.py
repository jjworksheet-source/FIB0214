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

# --- 2. DATA LOADING ---

@st.cache_data(ttl=60)
def load_review_data():
    """讀取 Review 工作表（草稿/審批區）"""
    try:
        sh = client.open_by_key(SHEET_ID)
        worksheet = sh.worksheet("Review")
        data = worksheet.get_all_records()
        df = pd.DataFrame(data)
        if df.empty:
            return df
        df.columns = [c.strip() for c in df.columns]
        rename_map = {
            "學校": "School",
            "年級": "Level",
            "詞語": "Word",
            "句子": "Content",
            "狀態": "Status",
            "來源": "Source",
            "Timestamp": "Timestamp",
        }
        df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Error reading Review sheet: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=60)
def load_standby_data():
    """讀取 standby 工作表（正式題庫，用於生成 PDF）"""
    try:
        sh = client.open_by_key(SHEET_ID)
        worksheet = sh.worksheet("standby")
        data = worksheet.get_all_records()
        df = pd.DataFrame(data)
        if df.empty:
            return df
        df.columns = [c.strip() for c in df.columns]
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Error reading standby sheet: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=60)
def load_students():
    try:
        sh = client.open_by_key(SHEET_ID)
        worksheet = sh.worksheet("學生資料")
        data = worksheet.get_all_records()
        df = pd.DataFrame(data)
        if not df.empty:
            df.columns = [c.strip() for c in df.columns]
            for col in df.columns:
                if df[col].dtype == object:
                    df[col] = df[col].astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Error reading 學生資料 sheet: {e}")
        return pd.DataFrame()

if st.button("🔄 Refresh Data"):
    load_review_data.clear()
    load_standby_data.clear()
    load_students.clear()
    st.rerun()

# --- 3. HELPER: WRITE TO GOOGLE SHEETS ---

def transfer_to_standby(selected_rows_df):
    """將審批通過的句子從 Review 移交至 standby"""
    try:
        sh = client.open_by_key(SHEET_ID)
        standby_ws = sh.worksheet("standby")
        review_ws = sh.worksheet("Review")

        new_standby_data = []
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for _, row in selected_rows_df.iterrows():
            unique_id = f"{str(row['School'])[:4]}_{datetime.datetime.now().strftime('%S%f')[-6:]}_{row['Word']}_f"
            # standby 欄位: ID, School, Grade, Word, Type, Content, Answer, Status, Date
            new_row = [
                unique_id,
                row['School'],
                row['Level'],
                row['Word'],
                "填空題",
                row['EditedContent'],  # 使用 Admin 修改後的句子
                row['Word'],           # Answer = 詞語本身
                "Ready",
                now_str
            ]
            new_standby_data.append(new_row)

        if new_standby_data:
            standby_ws.append_rows(new_standby_data)

            # 在 Review 表中把已移交的行標記為 Transferred
            # 用 (Timestamp, School, Level, Word, Content) 找回列號
            all_values = review_ws.get_all_values()
            key_to_row = {}
            for i, r in enumerate(all_values[1:], start=2):
                r = r + [""] * (7 - len(r))
                k = (r[0].strip(), r[1].strip(), r[2].strip(), r[3].strip(), r[4].strip())
                if k not in key_to_row:
                    key_to_row[k] = i

            for _, row in selected_rows_df.iterrows():
                k = (
                    str(row['Timestamp']).strip(),
                    str(row['School']).strip(),
                    str(row['Level']).strip(),
                    str(row['Word']).strip(),
                    str(row['Content']).strip()
                )
                row_index = key_to_row.get(k)
                if row_index:
                    review_ws.update_cell(row_index, 7, "Transferred")  # G欄: 狀態

            return True, len(new_standby_data)
    except Exception as e:
        return False, str(e)

# --- 4. PDF GENERATION ---

def create_pdf(school_name, level, questions, student_name=None):
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=letter)
    story = []

    styles = getSampleStyleSheet()
    font_name = CHINESE_FONT if CHINESE_FONT else 'Helvetica'

    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontName=font_name,
        fontSize=20,
        alignment=TA_CENTER,
        spaceAfter=12
    )
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontName=font_name,
        fontSize=14,
        leading=20,
        leftIndent=25,
        firstLineIndent=-25
    )

    if student_name:
        title_text = f"<b>{school_name} ({level}) - {student_name} - 校本填充工作紙</b>"
    else:
        title_text = f"<b>{school_name} ({level}) - 校本填充工作紙</b>"

    story.append(Paragraph(title_text, title_style))
    story.append(Spacer(1, 0.2 * inch))
    story.append(Paragraph(f"日期: {datetime.date.today() + datetime.timedelta(days=1)}", normal_style))
    story.append(Spacer(1, 0.3 * inch))

    for i, row in enumerate(questions):
        content = row['Content']
        content = re.sub(r'【】(.+?)【】', r'<u>\1</u>', content)
        content = re.sub(r'【(.+?)】', r'<u>\1</u>', content)
        p = Paragraph(f"{i + 1}. {content}", normal_style)
        story.append(p)
        story.append(Spacer(1, 0.15 * inch))

    doc.build(story)
    bio.seek(0)
    return bio

# --- 5. EMAIL ---

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

# --- 6. PDF PREVIEW ---

def display_pdf_as_images(pdf_bytes):
    try:
        images = convert_from_bytes(pdf_bytes, dpi=150)
        for i, image in enumerate(images):
            st.image(image, caption=f"Page {i + 1}", use_container_width=True)
    except Exception as e:
        st.error(f"Could not render preview: {e}")
        st.info("You can still download the PDF using the button on the left.")

# ============================================================
# SECTION A: 審批區 (Review → Standby)
# ============================================================
st.markdown("---")
st.header("📥 Step 1：審批新詞語")
st.caption("從 Review 表讀取新詞，選擇/修改句子後移交至 Standby 題庫")

review_df = load_review_data()

if review_df.empty:
    st.info("✅ Review 表目前沒有新資料。")
else:
    # --- Sidebar ---
    with st.sidebar:
        st.header("🎓 篩選年級")
        available_levels_review = sorted(review_df["Level"].astype(str).str.strip().unique().tolist())
        selected_level_review = st.radio("審批年級", available_levels_review, index=0, key="review_level")
        st.divider()

    # 篩選選定年級，且尚未移交的資料
    review_filtered = review_df[
        (review_df["Level"].astype(str).str.strip() == selected_level_review) &
        (~review_df["Status"].astype(str).str.strip().isin(["Transferred"]))
    ].copy()

    if review_filtered.empty:
        st.info(f"✅ {selected_level_review} 的所有詞語已審批完畢。")
    else:
        # 分開顯示 DB (Ready) 和 AI (Pending)
        db_df = review_filtered[review_filtered["Source"].astype(str).str.strip() == "DB"].copy()
        ai_df = review_filtered[review_filtered["Source"].astype(str).str.strip() == "AI"].copy()

        # --- DB 句子（直接確認）---
        if not db_df.empty:
            st.subheader("🟢 資料庫句子 (DB) — 建議直接移交")
            db_df.insert(0, "移交?", True)
            db_df["EditedContent"] = db_df["Content"]
            edited_db = st.data_editor(
                db_df[["移交?", "Timestamp", "School", "Level", "Word", "Content", "EditedContent", "Source", "Status"]],
                key="db_editor",
                hide_index=True,
                column_config={
                    "移交?": st.column_config.CheckboxColumn("移交?", default=True),
                    "EditedContent": st.column_config.TextColumn("句子（可修改）"),
                },
                disabled=["Timestamp", "School", "Level", "Word", "Content", "Source", "Status"]
            )
        else:
            edited_db = pd.DataFrame()

        # --- AI 句子（需審批）---
        if not ai_df.empty:
            st.subheader("🟨 AI 生成句子 (Pending) — 請選擇或修改")
            st.caption("每個詞語可能有 3 個 AI 句子選項，請勾選你想要的（可只選 1 個），並可在「句子（可修改）」欄直接改寫。")
            ai_df.insert(0, "移交?", False)
            ai_df["EditedContent"] = ai_df["Content"]
            edited_ai = st.data_editor(
                ai_df[["移交?", "Timestamp", "School", "Level", "Word", "Content", "EditedContent", "Source", "Status"]],
                key="ai_editor",
                hide_index=True,
                column_config={
                    "移交?": st.column_config.CheckboxColumn("移交?", default=False),
                    "EditedContent": st.column_config.TextColumn("句子（可修改/手動輸入）"),
                },
                disabled=["Timestamp", "School", "Level", "Word", "Content", "Source", "Status"]
            )
        else:
            edited_ai = pd.DataFrame()

        # --- 移交按鈕 ---
        st.divider()
        if st.button("🚀 確認移交至 Standby 題庫", use_container_width=True, type="primary"):
            # 合併 DB 和 AI 中勾選的行
            to_transfer_list = []

            if not edited_db.empty:
                to_transfer_list.append(edited_db[edited_db["移交?"] == True])
            if not edited_ai.empty:
                to_transfer_list.append(edited_ai[edited_ai["移交?"] == True])

            if to_transfer_list:
                to_transfer = pd.concat(to_transfer_list, ignore_index=True)
            else:
                to_transfer = pd.DataFrame()

            if to_transfer.empty:
                st.warning("⚠️ 請先勾選至少一行再移交。")
            else:
                with st.spinner("正在移交資料至 Standby..."):
                    success, result = transfer_to_standby(to_transfer)
                    if success:
                        st.success(f"✅ 成功移交 {result} 筆資料至 Standby！請按「🔄 Refresh Data」重新載入。")
                        load_review_data.clear()
                        load_standby_data.clear()
                    else:
                        st.error(f"❌ 移交失敗：{result}")

# ============================================================
# SECTION B: 生成 PDF（從 Standby 讀取）
# ============================================================
st.markdown("---")
st.header("📄 Step 2：生成工作紙")
st.caption("從 Standby 題庫讀取 Ready 的題目，生成 PDF 並寄送")

standby_df = load_standby_data()
student_df = load_students()

if standby_df.empty:
    st.info("Standby 題庫目前是空的，請先完成 Step 1 的審批移交。")
    st.stop()

if "Status" not in standby_df.columns:
    st.error("Standby 表缺少 'Status' 欄位。")
    st.stop()

if "Level" not in standby_df.columns and "level" not in standby_df.columns:
    st.error("Standby 表缺少 'Level' 欄位。")
    st.stop()

level_col = "Level" if "Level" in standby_df.columns else "level"
standby_df = standby_df.rename(columns={level_col: "Level"})

# --- Sidebar: Level + Mode ---
with st.sidebar:
    st.divider()
    st.header("📄 生成設定")
    available_levels_standby = sorted(standby_df["Level"].astype(str).str.strip().unique().tolist())
    selected_level = st.radio("生成年級", available_levels_standby, index=0, key="standby_level")
    st.divider()
    st.header("📬 發送模式")
    send_mode = st.radio(
        "選擇模式",
        ["📄 按學校預覽下載", "👨‍👩‍👧 按學生寄送"],
        index=0
    )

status_norm = (
    standby_df["Status"]
    .astype(str)
    .str.replace("\u00A0", " ", regex=False)
    .str.replace("\u3000", " ", regex=False)
    .str.strip()
)
level_norm = standby_df["Level"].astype(str).str.strip()
ready_df = standby_df[status_norm.isin(["Ready", "Waiting"]) & (level_norm == selected_level)]

if ready_df.empty:
    st.info(f"Standby 中沒有 {selected_level} 的 Ready/Waiting 題目。")
    st.stop()

st.subheader("📋 選擇題目")
edited_df = st.data_editor(
    ready_df,
    column_config={
        "Select": st.column_config.CheckboxColumn("Generate?", default=True)
    },
    disabled=["School", "Level", "Word"],
    hide_index=True
)

st.divider()
st.subheader("🚀 Finalize Documents")

# ============================================================
# MODE A: 按學校預覽下載
# ============================================================
if send_mode == "📄 按學校預覽下載":
    schools = edited_df['School'].unique() if not edited_df.empty else []

    if len(schools) == 0:
        st.info("請先在上方選擇題目。")
    else:
        selected_school = st.selectbox("選擇學校預覽/下載", schools)
        school_data = edited_df[edited_df['School'] == selected_school]

        col1, col2 = st.columns([1, 2])
        pdf_buffer = create_pdf(selected_school, selected_level, school_data.to_dict('records'))
        pdf_bytes = pdf_buffer.getvalue()

        with col1:
            st.write(f"**學校：** {selected_school}")
            st.write(f"**年級：** {selected_level}")
            st.write(f"**題目數：** {len(school_data)}")
            st.download_button(
                label=f"📥 下載 {selected_school}_{selected_level}.pdf",
                data=pdf_bytes,
                file_name=f"{selected_school}_{selected_level}_Review_{datetime.date.today()}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key=f"dl_{selected_school}_{selected_level}"
            )
            st.info("💡 如需修改句子，請在 Google Sheet 更改後按「🔄 Refresh Data」。")

        with col2:
            st.write("🔍 **PDF 預覽**")
            display_pdf_as_images(pdf_bytes)

# ============================================================
# MODE B: 按學生寄送
# ============================================================
else:
    st.subheader("👨‍👩‍👧 學生配對結果")

    if student_df.empty:
        st.error("❌ 無法讀取「學生資料」工作表。")
        st.stop()

    required_cols = ['學校', '年級', '狀態', '學生姓名', '家長 Email']
    missing_cols = [c for c in required_cols if c not in student_df.columns]
    if missing_cols:
        st.error(f"❌ 「學生資料」工作表缺少以下欄位：{missing_cols}")
        st.write("現有欄位：", student_df.columns.tolist())
        st.stop()

    active_students = student_df[student_df['狀態'] == 'Y']

    if active_students.empty:
        st.warning("⚠️ 「學生資料」中沒有「狀態 = Y」的學生。")
        st.stop()

    merged = active_students.merge(
        edited_df,
        left_on=['學校', '年級'],
        right_on=['School', 'Level'],
        how='inner'
    )

    if merged.empty:
        st.warning("⚠️ 沒有符合條件的配對。")
        with st.expander("🔍 查看配對資料（協助排查問題）"):
            st.write("**Standby 的 School 值：**", edited_df['School'].unique().tolist())
            st.write("**Standby 的 Level 值：**", edited_df['Level'].unique().tolist())
            st.write("**學生資料 的 學校 值：**", active_students['學校'].unique().tolist())
            st.write("**學生資料 的 年級 值：**", active_students['年級'].unique().tolist())
        st.stop()

    st.success(f"✅ 成功配對 {merged['家長 Email'].nunique()} 位學生，共 {len(merged)} 題")

    for parent_email, group in merged.groupby('家長 Email'):
        student_name  = group['學生姓名'].iloc[0]
        school_name   = group['學校'].iloc[0]
        grade         = group['年級'].iloc[0]
        teacher_email = group['老師 Email'].iloc[0] if '老師 Email' in group.columns else "N/A"

        st.divider()
        col1, col2 = st.columns([1, 2])

        pdf_buffer = create_pdf(school_name, grade, group.to_dict('records'), student_name=student_name)
        pdf_bytes  = pdf_buffer.getvalue()

        with col1:
            st.write(f"**👤 學生：** {student_name}")
            st.write(f"**🏫 學校：** {school_name} ({grade})")
            st.write(f"**📧 家長：** {parent_email}")
            st.write(f"**👩‍🏫 老師：** {teacher_email}")
            st.write(f"**📝 題目數：** {len(group)} 題")

            st.download_button(
                label=f"📥 下載 {student_name} PDF",
                data=pdf_bytes,
                file_name=f"{student_name}_{grade}_Review_{datetime.date.today()}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key=f"dl_{parent_email}"
            )

            if st.button(f"📧 寄送給 {student_name} 家長", key=f"send_{parent_email}", use_container_width=True):
                with st.spinner(f"正在寄送給 {parent_email}..."):
                    success, msg = send_email_with_pdf(
                        parent_email, student_name, school_name, grade, pdf_bytes, cc_email=teacher_email
                    )
                    if success:
                        st.success(f"✅ 已成功寄送！")
                    else:
                        st.error(f"❌ 發送失敗: {msg}")
                        st.code(msg)

        with col2:
            st.write("🔍 **100% 準確預覽**")
            display_pdf_as_images(pdf_bytes)
