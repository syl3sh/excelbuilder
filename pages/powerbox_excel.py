import streamlit as st
import numpy as np
import pandas as pd
from io import BytesIO
from datetime import datetime
import pytz
import streamlit_authenticator as stauth
from streamlit_gsheets import GSheetsConnection
from github import Github

sgt = pytz.timezone("Asia/Singapore")
df = pd.read_excel('data/ATE_Tracking_Record_10726.xlsx')




g = Github(st.secrets["github"]["token"])
repo = g.get_repo(st.secrets["github"]["repo"])
branch = st.secrets["github"]["branch"]
history_path = "data/history"



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
    try:
        content = repo.get_contents(history_path, ref=branch)
        return sorted(contents, key=lambda f:f.name, reverse = True)
    except Exception:
        return[]
def save_version(df):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"{history_path}/ATE_Tracking_Record_{timestamp}.xlsx"
    
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    buffer.seek(0)
    repo.create_file(
        path=history_path,
        message=f"Save version {timestamp}",
        content=buffer.getvalue(),
        branch= branch
    )
def download_version(file_content):
    buffer=BytesIO(file_content.decoded_content)
    return pd.read_excel(buffer)
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
    
version_labels = [v.name for v in versions]
selected_label = st.selectbox("View a previous version", version_labels)
selected_file = next(v for v in versions if v.name == selected_label)
st.dataframe(download_version(selected_file))

def convert_df_to_excel(df):
  output = BytesIO()
  with pd.ExcelWriter(output, engine = "openpyxl") as writer:
    df.to_excel(writer, index = False, sheet_name = "sheet1")
    return output.getvalue()

excel_data = convert_df_to_excel(edited_df)
    
st.download_button(
    label="Save Copy as Excel",
    data=excel_data,
    file_name="ATE Tracking Record 10726_edited.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
  
  
  
