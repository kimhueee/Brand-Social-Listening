#!/usr/bin/env python
# coding: utf-8

# In[1]:


pip install streamlit gspread oauth2client pandas


# In[3]:


pip install streamlit


# In[5]:


pip install --upgrade numpy pandas


# In[7]:


pip install --upgrade matplotlib scipy


# In[8]:


import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- Cấu hình ---
SPREADSHEET_ID = "1qsGHOEEQRlFAitOsQ3i-P9E5_sLpyi1nERMyySi5l7U"
SHEET_NAME = "Data"
CREDENTIALS_FILE = "famous-analyzer-458803-n6-aa0f7555cf4f.json"

# --- Kết nối Google Sheet ---
@st.cache_data(ttl=60)
def load_data():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    data = sheet.get_all_records()
    df = pd.DataFrame(data)
    return df

# --- Giao diện Streamlit ---
st.title("📊 Dashboard từ Google Sheet riêng tư")
try:
    df = load_data()
    st.success("✅ Dữ liệu đã tải thành công.")
    st.dataframe(df)

    if "Doanh thu" in df.columns and "Ngày" in df.columns:
        df["Ngày"] = pd.to_datetime(df["Ngày"])
        st.line_chart(df.set_index("Ngày")["Doanh thu"])
    else:
        st.warning("Không có cột 'Ngày' hoặc 'Doanh thu' để hiển thị biểu đồ.")

except Exception as e:
    st.error(f"❌ Lỗi khi tải dữ liệu: {e}")


# In[ ]:




