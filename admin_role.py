import streamlit as st
from streamlit_option_menu import option_menu

from msme_role import display_msme_report
from creditcontrol_role import display_creditcontrol_report
from centralOps_role import display_centralOps_report

def admin():
    # --- SIDEBAR NAVIGATION ---
    with st.sidebar:
        selected = option_menu(
            menu_title="👤 Admin Panel",
            options=["MSME Team", "Credit Control Team", "Central Ops Team"],
            icons=["people", "cash-stack", "gear"],
            default_index=0,
            menu_icon="cast"
        )

    # --- MAIN CONTENT ---
    if selected == "MSME Team":
        display_msme_report()

    elif selected == "Credit Control Team":
        display_creditcontrol_report()

    elif selected == "Central Ops Team":
        display_centralOps_report()
