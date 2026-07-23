import streamlit as st
import numpy as np
import pandas as pd
from io import BytesIO
import streamlit_authenticator as stauth
df = pd.read_excel('data/ATE_Tracking_Record_10726.xlsx')

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

edited_df = st.data_editor(df, num_rows = "dynamic")

if "version_history" not in st.session_state:
    st.session_state.version_history = [pd.read_excel("data/ATE_Tracking_Record_10726.xlsx")]
    current_df=st.session_state.version_history[-1]
    edited_df = st.data_editor(current_df, num rows = "dynamic")



if st.button("Save Changes on Dashboard"):
  st.session_state.version_history.append(edited_df.copy())
  st.success("Saved")

version_num = st.selectbox(
    "View a previous version",
    range(len(st.session_state.version_history)),
    format_func=lambda i: f"Version {i+1}"
)
st.dataframe(st.session_state.version_history[version_num])



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
  
  
  
