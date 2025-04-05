import streamlit as st
import pandas as pd
from view_role import display_view_report
from msme_role import display_msme_report
from creditcontrol_role import display_creditcontrol_report
from centralOps_role import display_centralOps_report
from admin_role import admin

st.set_page_config(layout="wide")
st.logo(r'data/logo.jpg',size="large")
@st.cache_data
def load_all_sheets():
    xls = pd.ExcelFile(r"data/Users.xlsx")
    
    sheet_data = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
    return sheet_data

def get_user_role(email, sheets):

    agusers = sheets.get("Agusers")
    if agusers is None:
        return None, "Agusers sheet not found."

    # Normalize emails (strip, lowercase) for reliable comparison
    email = email.strip().lower()

    agusers['email'] = agusers['email'].astype(str).str.strip().str.lower()

    if email not in agusers['email'].values:
        return None, "Access Denied. Email not found in Agusers."

    role_found = []
    for sheet_name, df in sheets.items():

        if sheet_name == "Agusers":
            continue
        
        if 'email' in df.columns:
            
            df['email'] = df['email'].astype(str).str.strip().str.lower()
            if email in df['email'].values:
                
                role_found.append(sheet_name)

    if not role_found:
        return "view", None
    else:
        return role_found[0], None  # Return the first matching role


def main():
    st.markdown("""
    <h1 style='text-align: center; color: #1f77b4;'>🚛 MSME Shipment Tracker</h1>
    <p style='text-align: center; color: grey;'>Virya Logistics Technologies Pvt Ltd</p>
""", unsafe_allow_html=True)

    sheets = load_all_sheets()

    # Initialize session state if not set
    if 'logged_in' not in st.session_state:
        st.session_state.logged_in = False
    if 'role' not in st.session_state:
        st.session_state.role = None
    if 'email' not in st.session_state:
        st.session_state.email = None

    if not st.session_state.logged_in:
        email_input = st.text_input("Enter your Email ID", key="email_input")
        if email_input:
            role, error = get_user_role(email_input.strip(), sheets)
            if error:
                st.error(error)
            else:
                st.session_state.logged_in = True
                st.session_state.role = role
                st.session_state.email = email_input
                st.rerun()  # Triggers page "refresh" to show next view
    else:
        show_role_page(st.session_state.email,st.session_state.role)

def show_role_page(email,role):
    st.subheader(f"Welcome, {email.upper()}!")

    if role == "Admin":
        admin()
    elif role == "MSME":
        display_msme_report()
    elif role == "Central Ops":
        display_centralOps_report()
    elif role == "Credit Control":
        display_creditcontrol_report()
    elif role == "view":
        display_view_report()
    else:
        st.warning("Unrecognized role.")

if __name__ == "__main__":
    main()
