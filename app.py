import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import datetime
import io
import re
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

# --- PDF Libraries ---
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- 1. SETUP & CONNECTION ---
st.set_page_config(page_title="Worksheet Generator", page_icon="📝")
st.title("📝 Worksheet Generator (PDF & Email)")

# Load Secrets
try:
    # Google Sheets Secrets
    key_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(
        key_dict,
        scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    )
    client = gspread.authorize(creds)
    SHEET_ID = st.secrets["app_config"]["spreadsheet_id"]
    
    # Email Secrets (You need to add these to secrets.toml)
    # [email]
    # smtp_host = "mail.jj.edu.hk"
    # smtp_port = 587
    # smtp_user = "your_email@jj.edu.hk"
    # smtp_pass = "your_password"
    SMTP_HOST = st.secrets["email"]["smtp_host"]
    SMTP_PORT = st.secrets["email"]["smtp_port"]
    SMTP_USER = st.secrets["email"]["smtp_user"]
    SMTP_PASS = st.secrets["email"]["smtp_pass"]
    
    st.success("✅ Connected to Google Cloud & Email Config Loaded!")
except Exception as e:
    st.error(f"❌ Configuration Error: {e}")
    st.stop()

# --- 2. HELPER FUNCTIONS ---

def register_chinese_font():
    """
    Registers a Chinese font. 
    IMPORTANT: On Streamlit Cloud, you must upload a font file (e.g., kaiu.ttf) 
    to your repository and reference it here. Windows paths will not work.
    """
    # Example: If you upload 'kaiu.ttf' to the root of your repo
    font_path = "kaiu.ttf" 
    font_name = "KaiU"
    
    if os.path.exists(font_path):
        try:
            pdfmetrics.registerFont(TTFont(font_name, font_path))
            return font_name
        except Exception as e:
            st.warning(f"Font loading failed: {e}")
            return "Helvetica" # Fallback
    else:
        st.warning(f"Chinese font file '{font_path}' not found. Using default (Chinese may not display).")
        return "Helvetica"

def draw_text_with_underline_wrapped(c, x, y, text, font_name, font_size, max_width, underline_offset=2, line_height=18):
    """
    Handles the drawing of text with 專名號 (Proper Noun Mark).
    Splits text by <u> tags and draws lines under specific parts.
    """
    parts = re.split(r'(<u>.*?</u>)', text)
    tokens = []
    for p in parts:
        if not p: continue
        if p.startswith("<​u>") and p.endswith("<​/u>"):
            tokens.append(p)
        else:
            tokens.extend(list(p)) # Split normal text into chars

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
                # DRAW THE LINE (專名號)
                c.line(cx, draw_y - underline_offset, cx + w, draw_y - underline_offset)
                cx += w
            else:
                c.setFont(font_name, font_size)
                c.drawString(cx, draw_y, tp)
                cx += pdfmetrics.stringWidth(tp, font_name, font_size)

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
    
    cur_y -= 12 # Extra paragraph spacing
    return cur_y

def PDF Generator(school_name, questions, font_name):
    """Generates the PDF in memory (BytesIO)"""
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    page_w, page_h = A4
    
    # Style Settings
    title_size = 18
    body_size = 14
    line_height = 25
    max_text_width = page_w - 120
    
    cur_y = page_h - 80
    
    # --- Header ---
    c.setFont(font_name, title_size)
    header_text = f"{school_name} - Worksheet"
    text_width = pdfmetrics.stringWidth(header_text, font_name, title_size)
    c.drawString((page_w - text_width) / 2, cur_y, header_text)
    cur_y -= line_height * 2
    
    c.setFont(font_name, body_size)
    c.drawString(60, cur_y, f"Date: {datetime.date.today()}")
    cur_y -= line_height * 1.5
    
    # --- Questions ---
    for i, row in enumerate(questions):
        # 1. Pre-process: Convert 【】 to <u> tags for proper noun handling
        raw_content = row.get('Content', '')
        word = row.get('Word', '')
        
        # Regex replace 【】 with <u>
        processed = re.sub(r'【】(.*?)【】', r'<u>\1</u>', raw_content).strip()
        
        # Handle Blanks (if word exists in sentence)
        blank = '＿' * max(len(str(word)) * 2, 4) if word else '＿' * 4
        if word and word in processed:
            processed = processed.replace(word, blank, 1)
        
        # Check page break
        if cur_y < 100:
            c.showPage()
            cur_y = page_h - 80
            
        # Draw Question Number
        c.setFont(font_name, body_size)
        c.drawString(60, cur_y, f"{i+1}. ")
        
        # Draw Text with Proper Noun Handling
        cur_y = draw_text_with_underline_wrapped(
            c, 90, cur_y, processed, font_name, body_size, max_text_width, 
            underline_offset=2, line_height=line_height
        )
        
    c.save()
    buffer.seek(0)
    return buffer

def send_email_with_attachment(to_email, subject, body, pdf_buffer, filename):
    try:
        msg = MIMEMultipart()
        msg['From'] = SMTP_USER
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html', 'utf-8'))
        
        # Attach PDF from memory buffer
        part = MIMEApplication(pdf_buffer.getvalue(), Name=filename)
        part['Content-Disposition'] = f'attachment; filename="{filename}"'
        msg.attach(part)
        
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(SMTP_USER, [to_email], msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)

# --- 3. MAIN APP LOGIC ---

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

# --- 4. FILTER & SELECT ---
st.subheader("Select Questions")

try:
    ready_df = df[df['Status'].isin(['Ready', 'Waiting'])]
except KeyError:
    st.error("Column 'Status' not found.")
    st.stop()

if ready_df.empty:
    st.info("No questions with status 'Ready' or 'Waiting'.")
    st.stop()

edited_df = st.data_editor(
    ready_df,
    column_config={
        "Select": st.column_config.CheckboxColumn("Generate?", default=True)
    },
    disabled=["School", "Word", "Content"],
    hide_index=True
)

# --- 5. GENERATE PDF & EMAIL ---
st.divider()
st.subheader("🚀 Actions")

# Register Font
font_name = register_chinese_font()

col1, col2 = st.columns(2)

with col1:
    if st.button("📄 Generate PDF Only"):
        schools = edited_df['School'].unique()
        for school in schools:
            school_data = edited_df[edited_df['School'] == school]
            if not school_data.empty:
                # Corrected function call
                pdf_buffer = PDF Generator(school, school_data.to_dict('records'), font_name)
                
                st.download_button(
                    label=f"📥 Download {school}.pdf",
                    data=pdf_buffer,
                    file_name=f"{school}_Worksheet.pdf",
                    mime="application/pdf"
                )

with col2:
    with st.form("email_form"):
        st.write("📧 **Send PDF via Email**")
        recipient_email = st.text_input("Recipient Email")
        email_subject = st.text_input("Subject", value="[Worksheet] Latest Practice")
        email_body = st.text_area("Message", value="<p>Please find the worksheet attached.</p>")
        submit_email = st.form_submit_button("Send Email")
        
        if submit_email and recipient_email:
            schools = edited_df['School'].unique()
            success_count = 0
            
            for school in schools:
                school_data = edited_df[edited_df['School'] == school]
                if not school_data.empty:
                    # Corrected function call
                    pdf_buffer = PDF Generator(school, school_data.to_dict('records'), font_name)
                    filename = f"{school}_Worksheet.pdf"
                    
                    # Send
                    success, error = send_email_with_attachment(
                        recipient_email, email_subject, email_body, pdf_buffer, filename
                    )
                    
                    if success:
                        st.success(f"✅ Sent to {recipient_email} for {school}")
                        success_count += 1
                    else:
                        st.error(f"❌ Failed to send {school}: {error}")
