import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import datetime
import io
import re
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import tempfile
import os

# --- 1. SETUP & CONNECTION ---
st.set_page_config(page_title="Worksheet Generator", page_icon="📝")
st.title("📝 Worksheet Generator")

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

# --- 2. FONT REGISTRATION FOR PDF ---
@st.cache_resource
def register_chinese_font():
    """
    註冊中文字型（支援多種字型）
    Register Chinese fonts for PDF generation
    """
    font_paths = [
        (r"C:\Windows\Fonts\kaiu.ttf", "KaiU"),
        (r"C:\Windows\Fonts\mingliu.ttc", "MingLiU"),
        (r"C:\Windows\Fonts\msjh.ttc", "MSJH"),
        (r"C:\Windows\Fonts\STKAITI.TTF", "Kaiti"),
        # Add Linux/Mac paths if needed
        ("/usr/share/fonts/truetype/arphic/ukai.ttc", "UKai"),
        ("/System/Library/Fonts/PingFang.ttc", "PingFang"),
    ]
    for path, name in font_paths:
        try:
            if os.path.exists(path):
                pdfmetrics.registerFont(TTFont(name, path))
                st.info(f"✅ Registered font: {name}")
                return name
        except Exception as e:
            st.warning(f"⚠️ Failed to register {name}: {e}")
    st.warning("⚠️ No Chinese font found, using default")
    return "Helvetica"

# --- 3. 專名號 HANDLING FUNCTIONS ---
def draw_text_with_underline_wrapped(c, x, y, text, font_name, font_size, max_width, underline_offset=2, line_height=18):
    """
    處理專名號標記 【】 並轉換為底線顯示
    Handles proper name marks 【】 and converts them to underlines
    Supports automatic line wrapping
    """
    # Split text into underlined and normal parts
    parts = re.split(r'(<u>.*?</u>)', text)
    tokens = []
    for p in parts:
        if not p:
            continue
        if p.startswith("<​u>") and p.endswith("<​/u>"):
            tokens.append(p)
        else:
            tokens.extend(list(p))
    
    def measure(tok):
        if tok.startswith("<​u>") and tok.endswith("<​/u>"):
            inner = tok[3:-4]
            return pdfmetrics.stringWidth(inner, font_name, font_size)
        else:
            return pdfmetrics.stringWidth(tok, font_name, font_size)
    
    def draw_line(parts_to_draw, draw_x, draw_y):
        cx = draw_x
        for tp in parts_to_draw:
            if tp.startswith("<​u>") and tp.endswith("<​/u>"):
                inner = tp[3:-4]
                c.setFont(font_name, font_size)
                c.drawString(cx, draw_y, inner)
                w = pdfmetrics.stringWidth(inner, font_name, font_size)
                # Draw underline for proper names
                c.line(cx, draw_y - underline_offset, cx + w, draw_y - underline_offset)
                cx += w
            else:
                c.setFont(font_name, font_size)
                c.drawString(cx, draw_y, tp)
                cx += pdfmetrics.stringWidth(tp, font_name, font_size)
    
    # Auto line wrapping
    cur_y = y
    line_buf = []
    line_width = 0
    for tok in tokens:
        tok_w = measure(tok)
        if line_width + tok_w > max_width and line_buf:
            draw_line(line_buf, x, cur_y)
            cur_y -= line_height
            line_buf = [tok]
            line_width = tok_w
        else:
            line_buf.append(tok)
            line_width += tok_w
    if line_buf:
        draw_line(line_buf, x, cur_y)
        cur_y -= line_height
    
    cur_y -= 12  # Paragraph spacing
    return cur_y

def draw_pdf_header(c, margin_left, margin_right, page_w, cur_y, style, info, use_font):
    """
    繪製 PDF 標題區域
    Draw PDF header with school info, date, and student name
    """
    title_size = style["title_size"]
    body_size = style["body_size"]
    line_height = style["line_height"]
    
    # Line 1: School + Grade (centered)
    header1 = " ".join([str(info.get("學校", "")), str(info.get("年級", ""))]).strip()
    c.setFont(use_font, title_size)
    if header1:
        text_width = pdfmetrics.stringWidth(header1, use_font, title_size)
        x_center = (page_w - text_width) / 2
        c.drawString(x_center, cur_y, header1)
    cur_y -= line_height
    
    # Line 2: Textbook + Lesson + Title (centered)
    header2 = "      ".join(filter(None, [
        str(info.get("教科書名稱", "")),
        f"第{info['第幾課']}課" if info.get("第幾課") else "",
        str(info.get("課文名稱", "")),
        "童學童樂詞語填充"
    ]))
    c.setFont(use_font, body_size)
    if header2:
        text_width = pdfmetrics.stringWidth(header2, use_font, body_size)
        x_center = (page_w - text_width) / 2
        c.drawString(x_center, cur_y, header2)
    cur_y -= line_height
    
    # Line 3: Student Name (left) + Date (right)
    today_str = datetime.date.today().isoformat()
    left_str = "學生姓名：________________________"
    right_str = f"日期：{today_str}"
    
    c.drawString(margin_left, cur_y, left_str)
    right_text_width = pdfmetrics.stringWidth(right_str, use_font, body_size)
    right_x = page_w - margin_right - right_text_width
    c.drawString(right_x, cur_y, right_str)
    
    cur_y -= line_height * 2
    return cur_y

def get_level_style(school_level):
    """
    根據年級返回字體大小設定
    Return font size settings based on grade level
    """
    if school_level == "P3" or "三年級" in str(school_level):
        return {"title_size": 18, "body_size": 16, "line_height": 30}
    else:
        return {"title_size": 16, "body_size": 14, "line_height": 30}

# --- 4. PDF GENERATION FUNCTION ---
def create_pdf_worksheet(school_name, questions, info=None):
    """
    生成學生練習 PDF（支援專名號處理）
    Generate student worksheet PDF with proper name mark handling
    """
    # Create temporary file
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_filename = temp_file.name
    temp_file.close()
    
    # Get font
    font_name = register_chinese_font()
    
    # Create PDF
    c = canvas.Canvas(temp_filename, pagesize=A4)
    page_w, page_h = A4
    
    # Default info if not provided
    if info is None:
        info = {
            "學校": school_name,
            "年級": "",
            "教科書名稱": "",
            "第幾課": "",
            "課文名稱": "Weekly Review"
        }
    
    style = get_level_style(info.get("年級", "P3"))
    body_size = style["body_size"]
    line_height = style["line_height"]
    max_text_width = page_w - 120 - 40
    
    cur_y = page_h - 80
    use_font = font_name if font_name else "Helvetica"
    
    # Draw header
    cur_y = draw_pdf_header(c, 60, 60, page_w, cur_y, style, info, use_font)
    
    extra_gap = 6
    
    # Draw questions
    for idx, item in enumerate(questions, start=1):
        word = item.get('Word', '')
        content = item.get('Content', '')
        
        # Convert 【】content【】 to <u>content</u> for underline
        processed = re.sub(r'【】(.*?)【】', r'<u>\1</u>', content).strip()
        
        # Replace word with blanks
        blank = '＿' * max(len(str(word)) * 2, 4) if word else '＿' * 4
        if word and word in processed:
            processed = processed.replace(word, blank, 1)
        else:
            if '＿' not in processed and '___' not in processed:
                processed = (processed + " " + blank).strip()
        
        # Check if new page needed
        if cur_y - line_height - extra_gap < 60:
            c.showPage()
            cur_y = page_h - 80
            cur_y = draw_pdf_header(c, 60, 60, page_w, cur_y, style, info, use_font)
        
        # Draw question number
        c.setFont(use_font, body_size)
        c.drawString(60, cur_y, f"{idx}. ")
        
        # Draw question content with underline support
        cur_y = draw_text_with_underline_wrapped(
            c, 100, cur_y, processed, use_font, body_size, max_text_width, 
            underline_offset=2, line_height=line_height
        )
        cur_y -= extra_gap
    
    # Add word list on new page
    c.showPage()
    cur_y = page_h - 80
    c.setFont(use_font, style["title_size"])
    c.drawString(60, cur_y, "詞語清單 Word List")
    cur_y -= line_height
    
    c.setFont(use_font, body_size)
    for idx, item in enumerate(questions, start=1):
        word = item.get('Word', '')
        if cur_y - line_height < 60:
            c.showPage()
            cur_y = page_h - 80
            c.setFont(use_font, style["title_size"])
            c.drawString(60, cur_y, "詞語清單 Word List")
            cur_y -= line_height
            c.setFont(use_font, body_size)
        c.drawString(60, cur_y, f"{idx}. {word}")
        cur_y -= line_height
    
    c.save()
    
    # Read file and return
    with open(temp_filename, 'rb') as f:
        pdf_data = f.read()
    
    # Clean up
    try:
        os.unlink(temp_filename)
    except:
        pass
    
    return io.BytesIO(pdf_data)

# --- 5. READ DATA ---
@st.cache_data(ttl=60)
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

# --- 6. FILTER & SELECT ---
st.subheader("Select Questions")

try:
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
    disabled=["School", "Word", "Content"],
    hide_index=True
)

# --- 7. GENERATE DOCUMENTS ---
st.subheader("Generate Documents")

# Choose format
doc_format = st.radio("Select Format:", ["Word (.docx)", "PDF (.pdf)", "Both"])

def create_docx(school_name, questions):
    doc = Document()
    
    # Font Setup
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.element.rPr.rFonts.set(qn('w:eastAsia'), 'TW-Kai') 
    
    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{school_name} - Weekly Review")
    run.font.size = Pt(20)
    run.bold = True
    
    doc.add_paragraph(f"Date: {datetime.date.today()}")
    doc.add_paragraph("-" * 30)
    
    # Questions
    for i, row in enumerate(questions):
        p = doc.add_paragraph()
        run = p.add_run(f"{i+1}. {row['Content']}")
        run.font.size = Pt(14)
        
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

if st.button("🚀 Generate Documents"):
    schools = edited_df['School'].unique()
    
    for school in schools:
        school_data = edited_df[edited_df['School'] == school]
        
        if not school_data.empty:
            questions = school_data.to_dict('records')
            
            # Generate Word
            if doc_format in ["Word (.docx)", "Both"]:
                docx_file = create_docx(school, questions)
                st.download_button(
                    label=f"📥 Download {school}.docx",
                    data=docx_file,
                    file_name=f"{school}_Review_{datetime.date.today()}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"docx_{school}"
                )
            
            # Generate PDF
            if doc_format in ["PDF (.pdf)", "Both"]:
                pdf_file = create_pdf_worksheet(school, questions)
                st.download_button(
                    label=f"📥 Download {school}.pdf",
                    data=pdf_file,
                    file_name=f"{school}_Review_{datetime.date.today()}.pdf",
                    mime="application/pdf",
                    key=f"pdf_{school}"
                )

st.markdown("---")
st.caption("💡 Tip: Use 【】text【】 in your questions to add underlines for proper names in PDF")
