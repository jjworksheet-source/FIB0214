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
def map_headers(df):
    """Map Chinese headers to English for internal logic."""
    col_map = {
        '學校': 'School',
        '年級': 'Level',
        '詞語': 'Word',
        '句子': 'Content',
        '來源': 'Source',
        '狀態': 'Status',
        'Timestamp': 'Timestamp'
    }
    # Rename only if the Chinese column exists
    df.rename(columns={k: v for k, v in col_map.items() if k in df.columns}, inplace=True)
    
    # Ensure all required columns exist to avoid KeyError
    required = ['School', 'Level', 'Word', 'Content', 'Source', 'Status']
    for col in required:
        if col not in df.columns:
            df[col] = ""
    return df

@st.cache_data(ttl=30)
def load_standby_data():
    try:
        sh = client.open_by_key(SHEET_ID)
        ws = sh.worksheet("standby")
        data = ws.get_all_records()
        df = pd.DataFrame(data)
        return map_headers(df)
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
        return map_headers(df)
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
        existing_df = map_headers(existing_df)

        rows_added = 0
        for _, row in rows_to_transfer.iterrows():
            # Check for duplicate
            if not existing_df.empty:
                dup = existing_df[
                    (existing_df['School'] == row['School']) &
                    (existing_df['Level'] == row['Level']) &
                    (existing_df['Word'] == row['Word'])
                ]
                if not dup.empty:
                    continue

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

        # Find column indices (1-based for gspread)
        try:
            word_idx = headers.index('詞語') if '詞語' in headers else headers.index('Word')
            school_idx = headers.index('學校') if '學校' in headers else headers.index('School')
            level_idx = headers.index('年級') if '年級' in headers else headers.index('Level')
            status_idx = headers.index('狀態') if '狀態' in headers else headers.index('Status')
        except ValueError:
            return False

        for i, row in enumerate(all_data[1:], start=2):
            if (len(row) > max(word_idx, school_idx, level_idx) and
                row[word_idx] == word and
                row[school_idx] == school and
                row[level_idx] == level):
                ws.update_cell(i, status_idx + 1, new_status)
                return True
        return False
    except Exception as e:
        st.error(f"Update error: {e}")
        return False

# --- 3. CREATE PDF ---
def make_blank_sentence(content, word):
    """Replace the target word in the sentence with a blank line."""
    blank = "____________"
    content = str(content)
    word = str(word)
    if word and word in content:
        return content.replace(word, blank, 1)
    if content.endswith('。'):
        return content[:-1] + blank + '。'
    return content + blank

def create_pdf(school_name, level, questions):
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=A4, rightMargin=2*cm, leftMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm)
    story = []
    font_name = CHINESE_FONT if CHINESE_FONT else 'Helvetica'

    title_style = ParagraphStyle('Title', fontName=font_name, fontSize=18, alignment=TA_CENTER, spaceAfter=6)
    subtitle_style = ParagraphStyle('Subtitle', fontName=font_name, fontSize=12, alignment=TA_CENTER, spaceAfter=4)
    question_style = ParagraphStyle('Question', fontName=font_name, fontSize=13, leading=22, leftIndent=20, firstLineIndent=-20, spaceAfter=8)

    story.append(Paragraph(f"<b>{school_name} ({level}) - 校本填充工作紙</b>", title_style))
    story.append(Paragraph(f"日期: {datetime.date.today()}", subtitle_style))
    story.append(Spacer(1, 0.1*inch))
    story.append(HRFlowable(width="100%", thickness=1, color=HexColor('#cccccc')))
    story.append(Spacer(1, 0.2*inch))

    for i, row in enumerate(questions):
        word = str(row.get('Word', ''))
        content = str(row.get('Content', ''))
        blank_sentence = make_blank_sentence(content, word)
        blank_sentence = re.sub(r'【】(.+?)【】', r'<u>\1</u>', blank_sentence)
        story.append(Paragraph(f"{i+1}. {blank_sentence}", question_style))

    doc.build(story)
    bio.seek(0)
    return bio

# ============================================================
# MAIN UI
# ============================================================
tab1, tab2 = st.tabs(["📋 Step 1: 審批 & 移交", "📄 Step 2: 生成工作紙"])

with tab1:
    st.subheader("📋 審批 AI 句子 & 移交至 Standby")
    review_df = load_review_data()

    if review_df.empty:
        st.info("Review 表格為空。")
    else:
        levels = sorted(review_df['Level'].dropna().unique().tolist())
        selected_level = st.selectbox("選擇年級", levels, key="review_level")
        level_df = review_df[review_df['Level'] == selected_level].copy()
        
        pending_df = level_df[level_df['Status'] == 'Pending'].copy()
        ready_df_review = level_df[level_df['Status'] == 'Ready'].copy()

        if not pending_df.empty:
            st.markdown("### 🟡 待審批 AI 句子")
            edited_pending = st.data_editor(
                pending_df[['School', 'Level', 'Word', 'Content', 'Source', 'Status']].reset_index(drop=True),
                column_config={"Content": st.column_config.TextColumn("句子 (可編輯)", width="large")},
                hide_index=True, key="pending_editor"
            )
            if st.button("✅ 批准並移交至 Standby", type="primary"):
                transferred = transfer_to_standby(edited_pending)
                for _, row in edited_pending.iterrows():
                    update_review_status(row['Word'], row['School'], row['Level'], 'Ready')
                st.cache_data.clear()
                st.success(f"✅ 成功移交 {transferred} 條句子！")
                st.rerun()
        else:
            st.success("✅ 沒有待審批的 AI 句子！")

        if not ready_df_review.empty:
            st.divider()
            st.markdown("### 🟢 已就緒句子 (可直接移交)")
            if st.button("📤 將所有 Ready 句子移交至 Standby"):
                transferred = transfer_to_standby(ready_df_review)
                st.cache_data.clear()
                st.success(f"✅ 成功移交 {transferred} 條句子！")
                st.rerun()

with tab2:
    st.subheader("📄 生成填充工作紙")
    standby_df = load_standby_data()

    if standby_df.empty:
        st.warning("Standby 表格為空。請先在 Step 1 移交句子。")
    else:
        ready_df = standby_df[standby_df['Status'].isin(['Ready', 'Waiting'])].copy()
        if ready_df.empty:
            st.info("沒有 Ready 的句子。")
        else:
            levels_pdf = sorted(ready_df['Level'].dropna().unique().tolist())
            sel_level_pdf = st.selectbox("選擇年級", levels_pdf, key="pdf_level")
            level_ready = ready_df[ready_df['Level'] == sel_level_pdf].copy()
            
            schools = sorted(level_ready['School'].dropna().unique().tolist())
            sel_schools = st.multiselect("選擇學校", schools, default=schools)
            filtered_df = level_ready[level_ready['School'].isin(sel_schools)]

            if not filtered_df.empty:
                preview_df = filtered_df[['School', 'Level', 'Word', 'Content']].copy()
                preview_df['填充句子預覽'] = preview_df.apply(lambda r: make_blank_sentence(r['Content'], r['Word']), axis=1)
                
                edited_df = st.data_editor(
                    preview_df.reset_index(drop=True),
                    column_config={"填充句子預覽": st.column_config.TextColumn("填充句子 (可修改)", width="large")},
                    hide_index=True, key="pdf_editor"
                )

                if st.button("🚀 生成工作紙 PDF", type="primary"):
                    for school in edited_df['School'].unique():
                        school_data = edited_df[edited_df['School'] == school].copy()
                        school_data['Content'] = school_data['填充句子預覽']
                        school_data['Word'] = "" # Prevent double blanking
                        pdf = create_pdf(school, sel_level_pdf, school_data.to_dict('records'))
                        st.download_button(label=f"📥 下載 {school} 工作紙", data=pdf, file_name=f"{school}_worksheet.pdf", mime="application/pdf")
