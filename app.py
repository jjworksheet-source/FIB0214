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

# --- 4. GENERATE PDF FUNCTION ---
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
    story.append(Spacer(1, 0.2*inch))
    story.append(Paragraph(f"日期: {datetime.date.today() + datetime.timedelta(days=1)}", normal_style))
    story.append(Spacer(1, 0.3*inch))

    for i, row in enumerate(questions):
        content = row['Content']
        content = re.sub(r'【】(.+?)【】', r'<u>\1</u>', content)
        content = re.sub(r'【(.+?)】', r'<u>\1</u>', content)
        p = Paragraph(f"{i+1}. {content}", normal_style)
        story.append(p)
        story.append(Spacer(1, 0.15*inch))

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

        pdf_buffer = create_pdf(selected_school, selected_level, school_data.to_dict('records'))
        pdf_bytes = pdf_buffer.getvalue()

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

    required_cols = ['學校', '年級', '狀態', '學生姓名', '家長 Email']
    missing_cols = [c for c in required_cols if c not in student_df.columns]
    if missing_cols:
        st.error(f"❌ 「學生資料」工作表缺少以下欄位：{missing_cols}")
        st.write("現有欄位：", student_df.columns.tolist())
        st.stop()

    active_students = student_df[student_df['狀態'] == 'Y']

    if active_students.empty:
        st.warning("⚠️ 「學生資料」中沒有「狀態 = Y」的學生。請先將測試學生的狀態改為 Y。")
        st.stop()

    merged = active_students.merge(
        edited_df,
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
