import streamlit as st
import pandas as pd
from io import BytesIO

def convert_df_to_excel(df):
    """Convert DataFrame to an Excel file and return as bytes for downloading."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name='Report')
    processed_data = output.getvalue()
    return processed_data

def display_creditcontrol_report():
    try:
        df = pd.read_excel("data/report.xlsx")
        df["DO Release Approved?"] = df["DO Release Approved?"].astype(str).str.strip().fillna('')

        st.write("### 💳 Credit Control Editable Report")

        # --- FILTER SECTION ---
        booking_options = sorted(df["Agraga Booking #"].dropna().unique())
        customer_options = sorted(df["Customer Name"].dropna().unique())

        col1, col2 = st.columns(2)
        with col1:
            selected_booking = st.selectbox("Filter by Agraga Booking #", options=["All"] + booking_options)
        with col2:
            selected_customer = st.selectbox("Filter by Customer Name", options=["All"] + customer_options)

        # Apply filters
        filtered_df = df.copy()
        if selected_booking != "All":
            filtered_df = filtered_df[filtered_df["Agraga Booking #"] == selected_booking]
        if selected_customer != "All":
            filtered_df = filtered_df[filtered_df["Customer Name"] == selected_customer]

        # Apply same cleaning to filtered data
        filtered_df["DO Release Approved?"] = filtered_df["DO Release Approved?"].astype(str).str.strip().fillna('')
        filtered_df["DO Release Approved?"] = filtered_df["DO Release Approved?"].map({"Yes": True})


        # --- EDITABLE TABLE ---
        edited_df = st.data_editor(
            filtered_df,
            use_container_width=True,
            hide_index = True,
            column_config={
                "Agraga Booking #": st.column_config.Column(pinned=True),
                "Customer Name": st.column_config.Column(pinned=True),
                "DO Release Approved?": st.column_config.CheckboxColumn("DO Release Approved?",pinned=True)
            },
            disabled=[
                col for col in df.columns if col not in ["DO Release Approved?"]
            ],
            key="creditcontrol_editor"
        )
        edited_df["DO Release Approved?"] = edited_df["DO Release Approved?"].map({True: "Yes", False: ""})
        edited_df["DO Release Approved?"] = edited_df["DO Release Approved?"].astype(str).str.strip()

        # Replace 'nan' strings with empty string
        edited_df.replace("nan", "", inplace=True)
        edited_df.replace("None", "", inplace=True)  # If some values are 'None' as strings
        edited_df = edited_df.fillna("")  # If actual NaN exists
        # --- SAVE BUTTON ---
        if st.button("💾 Save Changes"):
            try:
                # Read the original full report
                original_df = pd.read_excel("data/report.xlsx")

                # Ensure indices match for correct merging
                edited_df.index = filtered_df.index  # Maintain correct row alignment

                # Update only the edited columns
                columns_to_update = ["DO Release Approved?"]
                for col in columns_to_update:
                    # Ensure both DataFrames treat the column as string type
                    original_df[col] = original_df[col].astype(str)
                    edited_df[col] = edited_df[col].astype(str)

                    # Apply updates safely
                    original_df.loc[filtered_df.index, col] = edited_df[col].fillna(original_df[col])
                    # Replace 'nan' strings with empty string
                    original_df.replace("nan", "", inplace=True)
                    original_df.replace("None", "", inplace=True)  # If some values are 'None' as strings
                    original_df = original_df.fillna("")  # If actual NaN exists

                # Save the updated DataFrame back to Excel
                original_df.to_excel("data/report.xlsx", index=False)

                st.success("✅ Changes saved successfully!")
            except Exception as e:
                st.error(f"❌ Error saving file: {e}")

        # --- DOWNLOAD BUTTON ---
        excel_data = convert_df_to_excel(edited_df)
        st.download_button(
            label="📥 Download Report",
            data=excel_data,
            file_name="MSME Tracker Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error loading Credit Control report: {e}")

