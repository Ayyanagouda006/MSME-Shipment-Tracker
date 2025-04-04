import io
import streamlit as st
import pandas as pd

def display_view_report():
    try:
        report_df = pd.read_excel("data/report.xlsx")

        st.write("### 📊 View Report")
        st.dataframe(report_df, use_container_width=True)

        # Download logic
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            report_df.to_excel(writer, index=False, sheet_name='Report')
        buffer.seek(0)

        st.download_button(
            label="📥 Download Report",
            data=buffer,
            file_name="MSME Tracker Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error loading report: {e}")