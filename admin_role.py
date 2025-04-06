import streamlit as st
import pandas as pd
from streamlit_option_menu import option_menu
from msme_role import display_msme_report
from creditcontrol_role import display_creditcontrol_report
from centralOps_role import display_centralOps_report



def load_team_names():
    xls = pd.ExcelFile(r"data/Users.xlsx")
    return [sheet for sheet in xls.sheet_names if sheet != "Agusers"]

def load_team_data(sheet_name):
    return pd.read_excel(r"data/Users.xlsx", sheet_name=sheet_name)

def save_team_data(sheet_name, df):
    with pd.ExcelWriter(r"data/Users.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

def admin():
    # --- SIDEBAR NAVIGATION ---
    with st.sidebar:
        selected = option_menu(
            menu_title="Admin Panel",
            options=["MSME Team", "Credit Control Team", "Central Ops Team", "UAM", "Logs Download"],
            icons=["file-earmark-check-fill", "cash-stack", "tools", "people-fill", "cloud-download-fill"],
            default_index=3,
            menu_icon="cast"
        )

    # --- MAIN CONTENT ---
    if selected == "MSME Team":
        display_msme_report()

    elif selected == "Credit Control Team":
        display_creditcontrol_report()

    elif selected == "Central Ops Team":
        display_centralOps_report()

    elif selected == "UAM":
        st.title("👥 User Access Management")

        try:
            team_names = load_team_names()
            selected_team = st.selectbox("Select Team", team_names)

            if selected_team:
                df = load_team_data(selected_team)
                st.write(f"### 📄 Users in {selected_team}")
                edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="edit_uam")

                if st.button("💾 Save Changes", key="save_uam"):
                    # Save updated team sheet
                    save_team_data(selected_team, edited_df)

                    # ---- Also update Agusers sheet ----
                    agusers_df = pd.read_excel(r"data/Users.xlsx", sheet_name="Agusers")
                    agusers_emails = agusers_df["email"].astype(str).str.lower().str.strip().tolist()
                    new_emails = edited_df["email"].astype(str).str.lower().str.strip().tolist()

                    # Identify new emails to add
                    emails_to_add = [email for email in new_emails if email not in agusers_emails]

                    if emails_to_add:
                        new_entries = pd.DataFrame({"email": emails_to_add})
                        updated_agusers = pd.concat([agusers_df, new_entries], ignore_index=True)

                        with pd.ExcelWriter(r"data/Users.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            updated_agusers.to_excel(writer, sheet_name="Agusers", index=False)

                        st.success(f"✅ Added {len(emails_to_add)} new user(s) to Agusers.")
                    else:
                        st.info("No new users to add to Agusers.")

                    st.cache_data.clear()
                    st.success("✅ Changes saved and cache cleared. New roles will be reflected on next login.")
                    st.success("✅ Team data updated!")

        except Exception as e:
            st.error(f"Error loading user data: {e}")

