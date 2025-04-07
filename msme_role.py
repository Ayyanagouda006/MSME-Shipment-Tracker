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

def display_msme_report():
    try:
        df = pd.read_excel("data/report.xlsx")
        df["Freight Broker"] = df["Freight Broker"].astype(str).str.strip().fillna('')
        df["Transporter"] = df["Transporter"].astype(str).str.strip().fillna('')

        st.write("### 📝 MSME Editable Report")

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

        # --- DROPDOWN OPTIONS ---
        freight_brokers = ['Nolan Transportation Group','HeyPrimo','Ex-Freight','YouParcel']
        transporters = ["A Duie Pyle", "AAA Cooper", "ABF Freight System", "Amazon Freight", "Averitt Express", "California Sierra", "Central Transport", "Daylight Transport", "Estes Express", "Exclusive Transportation", "FedEx", "Forward Air", "Frontline Freight", "GoTo Logistics", "JTS Express", "Old Dominion", "Pitt-Ohio", "R+L Cariers", "Rist Transport", "Road Runner Transportation", "SAIA Motor", "South-Eastern Freight Lines", "Sunset Pacific Transportation", "T Central Transport", "TForce Freight", "Unis Transportation", "Ward Trucking", "WARP", "XPO Freight"]
        # Apply same cleaning to filtered data
        filtered_df["Freight Broker"] = filtered_df["Freight Broker"].astype(str).str.strip().fillna('')
        filtered_df["Transporter"] = filtered_df["Transporter"].astype(str).str.strip().fillna('')
        filtered_df["Delivery Quote"] = pd.to_numeric(filtered_df["Delivery Quote"], errors='coerce').fillna(0.0)

        # --- EDITABLE TABLE ---
        edited_df = st.data_editor(
            filtered_df,
            column_order=[
                "Customer Name", "MBL#", "HBL#", "Agraga Booking #", "Booking Status", "FBA?", "ISF Filing", "Stuffing Date",
                "Container #", "ETD", "ETA", "SOB", "ATA", "Carrier", "Consolidator", "FPOD", "CFS", "Delivery Address",
                "FBA Code", "Freight Broker", "Transporter", "Delivery Quote", "Packages", "Pallets", "Clearance Date",
                "Duty Invoice", "Actual # of Pallets", "Ready for Pick-up Date", "LFD", "DO Release Approved?",
                "HBL Released Date", "DO Released Date", "Pick-up Date", "Pick up number", "Delivery Appointment Date",
                "Delivery Date", "Vendor Delivery Invoice", "Updated Status Remarks", "PRO Number", "Storage Incurred (Days)"
            ],
            use_container_width=True,
            hide_index = True,
            column_config={
                "Customer Name": st.column_config.Column(pinned=True),
                "MBL#": st.column_config.Column(pinned=True),
                "HBL#": st.column_config.Column(pinned=True),
                "Agraga Booking #": st.column_config.Column(pinned=True),
                "Booking Status": st.column_config.Column(pinned=True),
                "Freight Broker": st.column_config.SelectboxColumn(
                    "Freight Broker",
                    options=freight_brokers,
                    required=False
                ),
                "Transporter": st.column_config.SelectboxColumn(
                    "Transporter",
                    options=transporters,
                    required=False
                ),
                "Delivery Quote": st.column_config.NumberColumn(
                    "Delivery Quote",
                    step=0.01,
                    format="$%.2f",
                    help="in USD"
                )
            },
            disabled=[
                col for col in df.columns if col not in ["Freight Broker", "Transporter", "Delivery Quote"]
            ],
            key="msme_editor"
        )
        edited_df["Freight Broker"] = edited_df["Freight Broker"].astype(str).str.strip()
        edited_df["Transporter"] = edited_df["Transporter"].astype(str).str.strip()

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
                columns_to_update = ["Freight Broker", "Transporter", "Delivery Quote"]
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
        st.error(f"Error loading MSME report: {e}")

