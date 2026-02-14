import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

st.title("🕵️‍♂️ Connection Doctor")

# 1. Check Secrets File
st.write("### 1. Checking Secrets...")
if "gcp_service_account" not in st.secrets:
    st.error("❌ Missing [gcp_service_account] section.")
    st.stop()
else:
    st.success("✅ [gcp_service_account] found.")

if "app_config" not in st.secrets:
    st.error("❌ Missing [app_config] section.")
    st.stop()
else:
    st.success("✅ [app_config] found.")

# 2. Check Private Key Format
st.write("### 2. Checking Key Format...")
key = st.secrets["gcp_service_account"]["private_key"]
if "-----BEGIN PRIVATE KEY-----" in key and "-----END PRIVATE KEY-----" in key:
    st.success("✅ Private Key looks correct (has BEGIN/END tags).")
else:
    st.error("❌ Private Key is malformed. Missing BEGIN or END tags.")
    st.stop()

# 3. Test Connection
st.write("### 3. Testing Google Connection...")
try:
    key_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(
        key_dict,
        scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    )
    client = gspread.authorize(creds)
    st.success("✅ Login Successful!")
except Exception as e:
    st.error(f"❌ Login Failed: {e}")
    st.stop()

# 4. Test Spreadsheet Access
st.write("### 4. Finding Spreadsheet...")
sheet_id = st.secrets["app_config"]["spreadsheet_id"]
try:
    sh = client.open_by_key(sheet_id)
    st.success(f"✅ Found Spreadsheet: **{sh.title}**")
    st.balloons()
except Exception as e:
    st.error(f"❌ Could not open Spreadsheet. Check ID or Sharing.\nError: {e}")
