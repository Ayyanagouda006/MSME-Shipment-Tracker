import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date

def convert_df_to_excel(df):
    """Convert DataFrame to an Excel file and return as bytes for downloading."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name='Report')
    processed_data = output.getvalue()
    return processed_data

def is_filled(val):
    return pd.notna(val) and str(val).strip() != ''

def determine_status(row):
    all_required_columns = [
        'ISF Filing', 'CFS', 'Freight Broker', 'Transporter', 'Delivery Quote',
        'Actual # of Pallets', 'Ready for Pick-up Date', 'DO Release Approved?',
        'HBL Released Date', 'Pick up number', 'Delivery Appointment Date',
        'Vendor Delivery Invoice', 'PRO Number', 'Storage Incurred (Days)', 'Remarks'
    ]
    basic_6_fields = ['CFS', 'Actual # of Pallets', 'Ready for Pick-up Date', 'Freight Broker', 'Transporter', 'Delivery Quote']

    # Condition 3: All 6 fields updated
    if all(is_filled(row[col]) for col in basic_6_fields):
        # Check which of the 15 total required fields are still not filled
        pending_fields = [col for col in all_required_columns if not is_filled(row[col])]
        if pending_fields:
            row['status'] = ', '.join(pending_fields) + ' pending'
    # Condition 1: Basic transport info
    elif is_filled(row['CFS']) and is_filled(row['Actual # of Pallets']) and is_filled(row['Ready for Pick-up Date']) and int(row['Actual # of Pallets']) != 0:
        row['status'] = 'Transport Assignment Pending'
    # Condition 2: Broker + Transporter + Quote
    elif is_filled(row['Freight Broker']) and is_filled(row['Transporter']) and is_filled(row['Delivery Quote']) and float(row['Delivery Quote']) != 0.0:
        row['status'] = 'Delivery Order Release Approval Pending'

    return row

def display_centralOps_report():
    # try:
    df = pd.read_excel("data/report.xlsx")
    df = df[df['Booking Status']=='INPROGRESS']
    df["ISF Filing"] = df["ISF Filing"].astype(str).str.strip().fillna('')
    df["CFS"] = df["CFS"].astype(str).str.strip().fillna('')
    df["Pick up number"] = df["Pick up number"].astype(str).str.strip().fillna('')
    df["Vendor Delivery Invoice"] = df["Vendor Delivery Invoice"].astype(str).str.strip().fillna('')
    df["PRO Number"] = df["PRO Number"].astype(str).str.strip().fillna('')
    df["Ready for Pick-up Date"] = pd.to_datetime(df["Ready for Pick-up Date"], errors='coerce').dt.date
    df["HBL Released Date"] = pd.to_datetime(df["HBL Released Date"], errors='coerce').dt.date
    df["Delivery Appointment Date"] = pd.to_datetime(df["Delivery Appointment Date"], errors='coerce').dt.date
    df["Remarks"] = df["Remarks"].astype(str).str.strip().fillna('')

    st.write("### 🛠️ Central Ops Editable Report")

    # --- FILTER SECTION ---

    # Prepare filter options
    booking_options = sorted(df["Agraga Booking #"].dropna().unique())
    customer_options = sorted(df["Customer Name"].dropna().unique())
    fba_options = sorted(df["FBA Code"].dropna().unique())
    pickup_options = sorted(df["Pick up number"].dropna().unique())
    cfs_options = sorted(df["CFS"].dropna().unique())
    eta_options = sorted(df["ETA"].dropna().astype(str).unique())  # convert to string if datetime

    # First row: Booking # and Customer Name
    col1, col2 = st.columns(2)
    with col1:
        selected_booking = st.selectbox("Filter by Agraga Booking #", options=["All"] + booking_options)
    with col2:
        selected_customer = st.selectbox("Filter by Customer Name", options=["All"] + customer_options)

    # Second row: FBA Code, Pick up number, CFS, ETA
    col3, col4, col5, col6 = st.columns(4)
    with col3:
        selected_fba = st.selectbox("Filter by FBA Code", options=["All"] + fba_options)
    with col4:
        selected_pickup = st.selectbox("Filter by Pick up number", options=["All"] + pickup_options)
    with col5:
        selected_cfs = st.selectbox("Filter by CFS", options=["All"] + cfs_options)
    with col6:
        selected_eta = st.selectbox("Filter by ETA", options=["All"] + eta_options)

    # Apply filters
    filtered_df = df.copy()
    if selected_booking != "All":
        filtered_df = filtered_df[filtered_df["Agraga Booking #"] == selected_booking]
    if selected_customer != "All":
        filtered_df = filtered_df[filtered_df["Customer Name"] == selected_customer]
    if selected_fba != "All":
        filtered_df = filtered_df[filtered_df["FBA Code"] == selected_fba]
    if selected_pickup != "All":
        filtered_df = filtered_df[filtered_df["Pick up number"] == selected_pickup]
    if selected_cfs != "All":
        filtered_df = filtered_df[filtered_df["CFS"] == selected_cfs]
    if selected_eta != "All":
        filtered_df = filtered_df[filtered_df["ETA"].astype(str) == selected_eta]
        

    # --- DROPDOWN OPTIONS ---
    cfs = ["New Jersey (ICT - 07201)", "New Jersey (St. George - 07047)", "Charleston (St. George - 29492)", 
            "Los Angeles (St. George - 90220)", "Charleston (Guardian Logistics Solutions - 29483)", 
            "Houston (St. George - 77507)"]
    # Apply same cleaning to filtered data
    filtered_df["ISF Filing"] = filtered_df["ISF Filing"].astype(str).str.strip().fillna('')
    filtered_df["ISF Filing"] = filtered_df["ISF Filing"].map({"Yes": True})
    filtered_df["CFS"] = filtered_df["CFS"].astype(str).str.strip().fillna('')
    filtered_df["Actual # of Pallets"] = pd.to_numeric(filtered_df["Actual # of Pallets"], errors='coerce').fillna(0).astype('Int64')
    # Convert to datetime safely and format as dd-mm-yyyy
    filtered_df["Ready for Pick-up Date"] = pd.to_datetime(filtered_df["Ready for Pick-up Date"], errors='coerce').dt.date
    filtered_df["HBL Released Date"] = pd.to_datetime(filtered_df["HBL Released Date"], errors='coerce').dt.date
    filtered_df["Pick up number"] = filtered_df["Pick up number"].astype(str).str.strip().fillna('')
    filtered_df["Delivery Appointment Date"] = pd.to_datetime(filtered_df["Delivery Appointment Date"], errors='coerce').dt.date
    filtered_df["Vendor Delivery Invoice"] = filtered_df["Vendor Delivery Invoice"].astype(str).str.strip().fillna('')
    filtered_df["Vendor Delivery Invoice"] = filtered_df["Vendor Delivery Invoice"].map({"Yes": True})
    filtered_df["PRO Number"] = filtered_df["PRO Number"].astype(str).str.strip().fillna('')
    filtered_df["Storage Incurred (Days)"] = pd.to_numeric(filtered_df["Storage Incurred (Days)"], errors='coerce').fillna(0).astype('Int64')
    filtered_df["Remarks"] = filtered_df["Remarks"].astype(str).str.strip().fillna('')
    # --- EDITABLE TABLE ---
    edited_df = st.data_editor(
        filtered_df,
        column_order=[
            "status","pickup type","Customer Name", "MBL#", "HBL#", "Agraga Booking #", "Booking Status", "FBA?", "ISF Filing", "Stuffing Date",
            "Container #", "ETD", "ETA", "SOB", "ATA", "Carrier", "Consolidator", "FPOD", "CFS", "Delivery Address",
            "FBA Code", "Freight Broker", "Transporter", "Delivery Quote", "Packages", "Pallets", "Clearance Date",
            "Duty Invoice", "Actual # of Pallets", "Ready for Pick-up Date", "LFD", "DO Release Approved?",
            "HBL Released Date", "DO Released Date", "Pick-up Date", "Pick up number", "Delivery Appointment Date",
            "Delivery Date", "Vendor Delivery Invoice", "Updated Status Remarks", "PRO Number", "Storage Incurred (Days)", "Remarks"
        ],
        use_container_width=True,
        hide_index = True,
        column_config={
            "status": st.column_config.Column(pinned=True),
            "pickup type": st.column_config.Column(pinned=True),
            "Customer Name": st.column_config.Column(pinned=True),
            "MBL#": st.column_config.Column(pinned=True),
            "HBL#": st.column_config.Column(pinned=True),
            "Agraga Booking #": st.column_config.Column(pinned=True),
            "Booking Status": st.column_config.Column(pinned=True),
            "ISF Filing": st.column_config.CheckboxColumn("ISF Filing"),
            "CFS": st.column_config.SelectboxColumn(
                "CFS",
                options=cfs,
                required=False
            ),
            "Actual # of Pallets": st.column_config.NumberColumn(
                "Actual # of Pallets",
                step=1,
                default="int"
            ),
            "Ready for Pick-up Date": st.column_config.DateColumn(
                "Ready for Pick-up Date",
                format="iso8601",
                min_value=date(2025, 1, 1),
                required=False
            ),
            "HBL Released Date": st.column_config.DateColumn(
                "HBL Released Date",
                format="iso8601",
                min_value=date(2025, 1, 1),
                required=False
            ),
            "Pick up number": st.column_config.TextColumn(
                "Pick up number",
                required=False
            ),
            "Delivery Appointment Date": st.column_config.DateColumn(
                "Delivery Appointment Date",
                format="iso8601",
                min_value=date(2025, 1, 1),
                required=False
            ),
            "Vendor Delivery Invoice": st.column_config.CheckboxColumn("Vendor Delivery Invoice"),
            "PRO Number": st.column_config.TextColumn(
                "PRO Number",
                required=False
            ),
            "Storage Incurred (Days)": st.column_config.NumberColumn(
                "Storage Incurred (Days)",
                step=1,
                default="int"
            ),
            "Remarks": st.column_config.TextColumn(
                "Remarks",
                required=False
            )
        },
        disabled=[
            col for col in df.columns if col not in ["ISF Filing", "CFS", "Actual # of Pallets", "Ready for Pick-up Date",
                                                        "HBL Released Date", "Pick up number", "Delivery Appointment Date",
                                                        "Vendor Delivery Invoice", "PRO Number", "Storage Incurred (Days)", "Remarks"]
        ],
        key="centralops_editor"
    )
    edited_df["ISF Filing"] = edited_df["ISF Filing"].map({True:"Yes"})
    edited_df["ISF Filing"] = edited_df["ISF Filing"].astype(str).str.strip()
    edited_df["CFS"] = edited_df["CFS"].astype(str).str.strip()
    edited_df["Actual # of Pallets"] = pd.to_numeric(edited_df["Actual # of Pallets"], errors='coerce').astype('Int64')
    # Convert to datetime safely and format as dd-mm-yyyy
    edited_df["Ready for Pick-up Date"] = pd.to_datetime(edited_df["Ready for Pick-up Date"], errors='coerce').dt.date
    edited_df["HBL Released Date"] = pd.to_datetime(edited_df["HBL Released Date"], errors='coerce').dt.date
    edited_df["Pick up number"] = edited_df["Pick up number"].astype(str).str.strip()
    edited_df["Delivery Appointment Date"] = pd.to_datetime(edited_df["Delivery Appointment Date"], errors='coerce').dt.date
    edited_df["Vendor Delivery Invoice"] = edited_df["Vendor Delivery Invoice"].map({True:"Yes"})
    edited_df["Vendor Delivery Invoice"] = edited_df["Vendor Delivery Invoice"].astype(str).str.strip()
    edited_df["PRO Number"] = edited_df["PRO Number"].astype(str).str.strip()
    edited_df["Storage Incurred (Days)"] = pd.to_numeric(edited_df["Storage Incurred (Days)"], errors='coerce').astype('Int64')
    edited_df["Remarks"] = edited_df["Remarks"].astype(str).str.strip()

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
            columns_to_update = ["ISF Filing", "CFS", "Actual # of Pallets", "Ready for Pick-up Date",
                                "HBL Released Date", "Pick up number", "Delivery Appointment Date",
                                "Vendor Delivery Invoice", "PRO Number", "Storage Incurred (Days)", "Remarks"]
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

            # Apply the function
            original_df = original_df.apply(determine_status, axis=1)
            # Ensure clean values
            original_df['Pick up number'] = original_df['Pick up number'].fillna('').astype(str).str.strip()
            original_df['FBA Code'] = original_df['FBA Code'].fillna('').astype(str).str.strip()

            # Filter rows with valid 'Pick up number'
            valid_pickup_df = original_df[
                (original_df['Pick up number'] != '') &
                (original_df['Pick up number'].str.lower() != 'nan')
            ]

            # Group only valid rows
            pickup_counts = valid_pickup_df.groupby(['Pick up number', 'FBA Code']).size().reset_index(name='count')

            # Merge back the counts to original_df
            original_df = original_df.merge(pickup_counts, on=['Pick up number', 'FBA Code'], how='left')

            # Assign pickup type based on count, default to empty or 'Single Pick-Up'
            original_df['pickup type'] = original_df['count'].apply(
                lambda x: 'Combined Pick-Up' if pd.notnull(x) and x > 1 else 'Single Pick-Up'
            )

            # Optional: Set pickup type to empty where Pick up number is invalid
            original_df.loc[
                (original_df['Pick up number'] == '') | (original_df['Pick up number'].str.lower() == 'nan'),
                'pickup type'
            ] = ''

            # Drop helper column
            original_df.drop(columns='count', inplace=True)

            # Save the updated DataFrame back to Excel
            original_df.to_excel("data/report.xlsx", index=False)

            st.success("✅ Changes saved successfully!")
            st.rerun()
        except Exception as e:
            st.error(f"❌ Error saving file: {e}")

    # --- DOWNLOAD BUTTON ---
    excel_data = convert_df_to_excel(pd.read_excel("data/report.xlsx"))
    st.download_button(
        label="📥 Download Report",
        data=excel_data,
        file_name="MSME Tracker Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # except Exception as e:
    #     st.error(f"Error loading MSME report: {e}")

