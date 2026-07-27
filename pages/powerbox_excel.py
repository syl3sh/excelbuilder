import streamlit as st
import numpy as np
import pandas as pd
from io import BytesIO
from datetime import datetime
import pytz
import streamlit_authenticator as stauth
from streamlit_gsheets import GSheetsConnection
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload,MediaIoBaseUpload

sgt = pytz.timezone("Asia/Singapore")
df = pd.read_excel('data/ATE_Tracking_Record_10726.xlsx')

creds_dict = dict(st.secrets["gdrive"])
history_folder_id = creds_dict.pop("history_folder_id")
crendentials = service_account.Credentials.from_service_account_info(
    creds_dict, scopes=["https://www.googleapis.com/auth/drive"]
)
service=build("drive", "v3", credentials=credentials)



config = {
    "credentials": {
        "usernames": {
            "admin": {
                "email": st.secrets["credentials"]["usernames"]["admin"]["email"],
                "first_name": st.secrets["credentials"]["usernames"]["admin"]["first_name"],
                "last_name": st.secrets["credentials"]["usernames"]["admin"]["last_name"],
                "username": st.secrets["credentials"]["usernames"]["admin"]["username"],
                "password": st.secrets["credentials"]["usernames"]["admin"]["password"],
                "logged_in": False,
                "failed_login_attempts": 0
            }
        }
    },
    "cookie": {
        "name": st.secrets["cookie"]["name"],
        "key": st.secrets["cookie"]["key"],
        "expiry_days": st.secrets["cookie"]["expiry_days"]
    }
}
authenticator = stauth.Authenticate(
    config['credentials'],
    config['cookie']['name'],
    config['cookie']['key'],
    config['cookie']['expiry_days']
)

authenticator.login(location='unrendered')

if not st.session_state.get('authentication_status'):
    st.error("You must log in first.")
    if st.button("Go to Login"):
        st.switch_page("loginexcel.py")
    st.stop()

authenticator.logout(location='sidebar')

def list_versions():
           results = service.files().list(
               q="'{history_folder_id}' in parents",
               orderBy="createdTime desc",
               fields="files(id, name, createdTime)"
           ).execute()
           return results.get("files", [])

versions = list_versions()
if not versions:
    current_df = pd.read_excel("data/ATE_Tracking_Record_10726.xlsx")
    save_version(current_df)
    versions = list_versions()

latest_version = versions[0]
current_df = download_version(latest_version["id"])

edited_df=st.data_editor(current_df, num_rows="dynamic")



if st.button("Save Changes on Dashboard"):
    save_version(edited_df)
    st.success("Save a new version!")
    st.rerun()
    
version_labels=[v["createdTime"] for v in versions]
selected_label=st.selectbox("View a previous version", version_labels)
selected_file_id=next(v["id"] for v in versions if v["createdTime"] == selected_label)

st.write("Viewing selected version:")
st.dataframe(download_version(selected_file_id))

def list_versions():
           results = service.files().list(
               q="'{history_folder_id}' in parents",
               orderBy="createdTime desc",
               fields="files(id, name, createdTime)"
           ).execute()
           return results.get("files", [])




def convert_df_to_excel(df):
  output = BytesIO()
  with pd.ExcelWriter(output, engine = "openpyxl") as writer:
    df.to_excel(writer, index = False, sheet_name = "sheet1")
    return output.getvalue()

excel_data = convert_df_to_excel(edited_df)





def download_version(file_id):
    request = service.files().get_media(fileId=file_id)
    buffer = BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buffer.seek(0)
    return pd.read_excel(buffer)

def save_version(df):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    buffer.seek(0)
    file_metadata = {
        "name": f"ATE_Tracking_Record_{timestamp}.xlsx",
        "parents": [history_folder_id]
    }
    media = MediaIoBaseUpload(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    service.files().create(body=file_metadata, media_body=media).execute()
    
st.download_button(
    label="Save Copy as Excel",
    data=excel_data,
    file_name="ATE Tracking Record 10726_edited.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
  
  
  
