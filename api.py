import gspread
import pandas as pd
from google.oauth2 import service_account
import streamlit as st

# Autentikasi
google_cloud_secrets = st.secrets["google_cloud"]
creds = service_account.Credentials.from_service_account_info(
    google_cloud_secrets,
    scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
)
client = gspread.authorize(creds)

def get_data_gsheet(spreadsheet_id, sheetname,range):
    # Mengambil semua data dari worksheet
    all_data = client.open_by_key(spreadsheet_id).worksheet(sheetname).get(range)
    
    # Memisahkan header (baris pertama) dan data
    headers = all_data[0]  # Baris pertama sebagai header
    rows = all_data[1:]    # Baris berikutnya sebagai data
    
    # Mengubah list of lists menjadi list of dictionaries dengan header sebagai key
    data = [dict(zip(headers, row)) for row in rows]
    
    return data