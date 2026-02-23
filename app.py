import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import datetime
import io
import os
import re

# --- 1. SETUP & CONNECTION ---
st.set_page_config(page_title="Worksheet Admin", page_icon="🎯", layout="wide")

# Sidebar
with st.sidebar:
    st.title("🎯 Worksheet Admin")
    st.divider()
    if st.button("🔄 Refresh Data"):
        st.cache_data.clear()
        st.rerun()
    st.caption("Data auto-refreshes every 30 seconds.")
    st.divider()
    st.markdown("### 📊 Status Legend")
    st.markdown("- 🟢 **Ready** — DB 句子，可直接使用")
    st.markdown("- 🟡 **Pending** — AI 句子，需要審批")
    st.markdown("- 🔵 **Loaded** — 已被 App 取走處理")

# Try to import reportlab
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch, cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.colors import HexColor

    font_paths = [
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
        "TW-Kai-98_1.ttf",
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

    if CHINESE_FONT:
        with st.sidebar:
            st.success("✅ Font OK")
    else:
        with st.sidebar:
            st.warning("⚠️ No Chinese font found")
            uploaded_font = st.file_uploader("📤 Upload Chinese Font (.ttf/.otf)", type=['ttf', 'otf'])
            if uploaded_font:
                with open("temp_font.ttf", "wb") as f:
                    f.write(uploaded_font.getbuffer())
                pdfmetrics.registerFont(TTFont('ChineseFont', "temp_font.ttf"))
                CHINESE_FONT = 'ChineseFont'
                st.success("✅ Font registered!")

except ImportError:
    st.error("❌ reportlab not found. Add 'reportlab' to requirements.txt")
    st.stop()

# --- Google Cloud Connection ---
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

# --- 2. LOAD DATA ---
@st.cache_data(ttl=30)
def load_standby_data():
    try:
        sh = client.open_by_key(SHEET_ID)
        ws = sh.worksheet("standby")
        data = ws.get_all_records()
        df = pd.DataFrame(data)
        return df
    except Exception as e:
        st.error(f"Error reading standby sheet: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=30)
def load_review_data():
    try:
        sh = client.open_by_key(SHEET_ID)
        ws = sh.worksheet("Review")
        data = ws.get_all_records()
        df = pd.DataFrame(data)
        # Rename columns to English for internal use
        col_map = {
            'Timestamp': 'Timestamp',
            '學校': 'School',
            '年級': 'Level',
            '詞語': 'Word',
            '句子': 'Content',
            '來源': 'Source',
            '狀態': 'Status'
        }
        df.rename(columns=col_map, inplace=True)
        return df
    except Exception as e:
        st.error(f"Error reading Review sheet: {e}")
        return pd.DataFrame()

def get_review_worksheet():
    sh = client.open_by_key(SHEET_ID)
    return sh.worksheet("Review")

def get_standby_worksheet():
    sh = client.open_by_key(SHEET_ID)
    return sh.worksheet("standby")

def transfer_to_standby(rows_to_transfer):
    """Transfer approved rows from Review to standby sheet."""
    try:
        standby_ws = get_standby_worksheet()
        existing = standby_ws.get_all_records()
        existing_df = pd.DataFrame(existing)

        rows_added = 0
        for _, row in rows_to_transfer.iterrows():
            # Check for duplicate (same School + Level + Word)
            if not existing_df.empty:
                dup = existing_df[
                    (existing_df.get('School', pd.Series(dtype=str)) == row['School']) &
                    (existing_df.get('Level', pd.Series(dtype=str)) == row['Level']) &
                    (existing_df.get('Word', pd.Series(dtype=str)) == row['Word'])
                ]
                if not dup.empty:
                    continue  # Skip duplicate

            new_row = [
                row.get('School', ''),
                row.get('Level', ''),
                row.get('Word', ''),
                row.get('Content', ''),
                'Ready',
                row.get('Source', 'DB'),
                str(datetime.datetime.now())
            ]
            standby_ws.append_row(new_row)
            rows_added += 1

        return rows_added
    except Exception as e:
        st.error(f"Transfer error: {e}")
        return 0

def update_review_status(word, school, level, new_status):
    """Update status of a row in Review sheet."""
    try:
        ws = get_review_worksheet()
        all_data = ws.get_all_values()
        headers = all_data[0]

        # Find column indices
        try:
            word_col = headers.index('詞語') + 1
            school_col = headers.index('學校') + 1
            level_col = headers.index('年級') + 1
            status_col = headers.index('狀態') + 1
        except ValueError:
            return False

        for i, row in enumerate(all_data[1:], start=2):
            if (len(row) >= max(word_col, school_col, level_col) and
                row[word_col-1] == word and
                row[school_col-1] == school and
                row[level_col-1] == level):
                ws.update_cell(i, status_col, new_status)
                return True
        return False
    except Exception as e:
        st.error(f"Update error: {e}")
        return False

# --- 3. CREATE PDF ---
def make_blank_sentence(content, word):
    """Replace the target word in the sentence with a blank line."""
    blank = "____________"
    if word and word in content:
        return content.replace(word, blank, 1)
    # If word not found, append blank at end (before 。if present)
    if content.endswith('。'):
        return content[:-1] + blank + '。'
    return content + blank

def create_pdf(school_name, level, questions):
    """
    questions: list of dicts with keys: Word, Content
    Each question gets a unique sentence with the word replaced by blank.
    """
    bio = io.BytesIO()
    doc = SimpleDocTemplate(
        bio,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )
    story = []

    font_name = CHINESE_FONT if CHINESE_FONT else 'Helvetica'

    title_style = ParagraphStyle(
        'Title',
        fontName=font_name,
        fontSize=18,
        alignment=TA_CENTER,
        spaceAfter=6,
        textColor=HexColor('#1a1a2e')
    )
    subtitle_style = ParagraphStyle(
        'Subtitle',
        fontName=font_name,
        fontSize=12,
        alignment=TA_CENTER,
        spaceAfter=4,
        textColor=HexColor('#555555')
    )
    question_style = ParagraphStyle(
        'Question',
        fontName=font_name,
        fontSize=13,
        leading=22,
        leftIndent=20,
        firstLineIndent=-20,
        spaceAfter=8,
        textColor=HexColor('#1a1a2e')
    )

    # Title
    story.append(Paragraph(f"<b>{school_name} ({level}) - 校本填充工作紙</b>", title_style))
    story.append(Paragraph(f"日期: {datetime.date.today()}", subtitle_style))
    story.append(Spacer(1, 0.1*inch))
    story.append(HRFlowable(width="100%", thickness=1, color=HexColor('#cccccc')))
    story.append(Spacer(1, 0.2*inch))

    # Questions — each row is a DIFFERENT word/sentence
    for i, row in enumerate(questions):
        word = str(row.get('Word', ''))
        content = str(row.get('Content', ''))

        # ✅ KEY FIX: replace the word with blank to make fill-in-the-blank
        blank_sentence = make_blank_sentence(content, word)

        # Handle special markup
        blank_sentence = re.sub(r'【】(.+?)【】', r'<u>\1</u>', blank_sentence)

        question_text = f"{i+1}. {blank_sentence}"
        story.append(Paragraph(question_text, question_style))

    doc.build(story)
    bio.seek(0)
    return bio

# ============================================================
# MAIN UI
# ============================================================

tab1, tab2 = st.tabs(["📋 Step 1: 審批 & 移交", "📄 Step 2: 生成工作紙"])

# ============================================================
# TAB 1: REVIEW & APPROVE
# ============================================================
with tab1:
    st.subheader("📋 審批 AI 句子 & 移交至 Standby")

    review_df = load_review_data()

    if review_df.empty:
        st.info("Review 表格為空或無法讀取。")
    else:
        # Show available levels
        available_levels = sorted(review_df['Level'].dropna().unique().tolist())
        selected_level = st.selectbox("選擇年級", available_levels, key="review_level")

        level_df = review_df[review_df['Level'] == selected_level].copy()

        # Split by status
        pending_df = level_df[level_df['Status'] == 'Pending'].copy()
        ready_df_review = level_df[level_df['Status'] == 'Ready'].copy()

        col1, col2, col3 = st.columns(3)
        col1.metric("總詞語數", len(level_df))
        col2.metric("🟡 待審批 (Pending)", len(pending_df))
        col3.metric("🟢 已就緒 (Ready)", len(ready_df_review))

        st.divider()

        # --- Pending AI sentences ---
        if not pending_df.empty:
            st.markdown("### 🟡 待審批 AI 句子")
            st.caption("可直接編輯句子，然後點擊「✅ 批准並移交」")

            edited_pending = st.data_editor(
                pending_df[['School', 'Level', 'Word', 'Content', 'Source', 'Status']].reset_index(drop=True),
                column_config={
                    "School": st.column_config.TextColumn("學校", disabled=True),
                    "Level": st.column_config.TextColumn("年級", disabled=True),
                    "Word": st.column_config.TextColumn("詞語", disabled=True),
                    "Content": st.column_config.TextColumn("句子 (可編輯)", width="large"),
                    "Source": st.column_config.TextColumn("來源", disabled=True),
                    "Status": st.column_config.TextColumn("狀態", disabled=True),
                },
                hide_index=True,
                key="pending_editor"
            )

            if st.button("✅ 批准並移交至 Standby", type="primary"):
                transferred = transfer_to_standby(edited_pending)
                # Update Review status to Ready
                for _, row in edited_pending.iterrows():
                    update_review_status(row['Word'], row['School'], row['Level'], 'Ready')
                st.cache_data.clear()
                st.success(f"✅ 成功移交 {transferred} 條句子至 Standby！")
                st.rerun()
        else:
            st.success("✅ 沒有待審批的 AI 句子！")

        st.divider()

        # --- Ready sentences (from DB, auto-approved) ---
        if not ready_df_review.empty:
            st.markdown("### 🟢 已就緒句子 (DB 來源，可直接移交)")
            st.dataframe(
                ready_df_review[['School', 'Level', 'Word', 'Content', 'Source']].reset_index(drop=True),
                hide_index=True,
                use_container_width=True
            )

            if st.button("📤 將所有 Ready 句子移交至 Standby"):
                transferred = transfer_to_standby(ready_df_review)
                st.cache_data.clear()
                st.success(f"✅ 成功移交 {transferred} 條句子！")
                st.rerun()

# ============================================================
# TAB 2: GENERATE PDF
# ============================================================
with tab2:
    st.subheader("📄 生成填充工作紙")

    standby_df = load_standby_data()

    if standby_df.empty:
        st.warning("Standby 表格為空。請先在 Step 1 移交句子。")
        st.stop()

    # Ensure correct column names
    # Expected: School, Level, Word, Content, Status, Source, Timestamp
    required_cols = ['School', 'Level', 'Word', 'Content', 'Status']
    missing = [c for c in required_cols if c not in standby_df.columns]
    if missing:
        st.error(f"Standby 表格缺少欄位: {missing}")
        st.write("現有欄位:", standby_df.columns.tolist())
        st.stop()

    # Filter Ready
    ready_df = standby_df[standby_df['Status'].isin(['Ready', 'Waiting'])].copy()

    if ready_df.empty:
        st.info("Standby 中沒有 Ready/Waiting 的句子。")
        st.stop()

    # Select Level
    levels = sorted(ready_df['Level'].dropna().unique().tolist())
    selected_level_pdf = st.selectbox("選擇年級", levels, key="pdf_level")

    level_ready = ready_df[ready_df['Level'] == selected_level_pdf].copy()

    # Select Schools
    schools = sorted(level_ready['School'].dropna().unique().tolist())
    selected_schools = st.multiselect("選擇學校", schools, default=schools)

    filtered_df = level_ready[level_ready['School'].isin(selected_schools)]

    if filtered_df.empty:
        st.info("沒有符合條件的句子。")
        st.stop()

    st.markdown("### 📝 預覽句子（可編輯）")
    st.caption("句子中的詞語將自動替換為填充空格 ____________")

    # Show preview with blank substitution
    preview_df = filtered_df[['School', 'Level', 'Word', 'Content']].copy()
    preview_df['填充句子預覽'] = preview_df.apply(
        lambda r: make_blank_sentence(str(r['Content']), str(r['Word'])), axis=1
    )

    edited_df = st.data_editor(
        preview_df.reset_index(drop=True),
        column_config={
            "School": st.column_config.TextColumn("學校", disabled=True),
            "Level": st.column_config.TextColumn("年級", disabled=True),
            "Word": st.column_config.TextColumn("詞語", disabled=True),
            "Content": st.column_config.TextColumn("原句", disabled=True),
            "填充句子預覽": st.column_config.TextColumn("填充句子（可修改）", width="large"),
        },
        hide_index=True,
        key="pdf_editor"
    )

    st.divider()

    if st.button("🚀 生成工作紙 PDF", type="primary"):
        pdf_schools = edited_df['School'].unique()
        generated = 0

        for school in pdf_schools:
            school_data = edited_df[edited_df['School'] == school].copy()
            if school_data.empty:
                continue

            # Use the edited blank sentence as Content for PDF
            school_data['Content'] = school_data['填充句子預覽']
            # Pass Word as empty so make_blank_sentence won't double-replace
            school_data['Word'] = ''

            pdf_file = create_pdf(school, selected_level_pdf, school_data.to_dict('records'))
            generated += 1

            st.download_button(
                label=f"📥 下載 {school} ({selected_level_pdf}) 工作紙",
                data=pdf_file,
                file_name=f"{school}_{selected_level_pdf}_worksheet_{datetime.date.today()}.pdf",
                mime="application/pdf",
                key=f"dl_{school}_{selected_level_pdf}"
            )

        if generated:
            st.success(f"✅ 已生成 {generated} 份工作紙！")
