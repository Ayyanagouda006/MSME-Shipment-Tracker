import streamlit as st
import pandas as pd
import requests

BASE_URL = "http://localhost:8000"

st.title("MongoDB Data Editor")

# Fetch Data
response = requests.get(f"{BASE_URL}/data")
if response.status_code == 200:
    data = response.json()
    df = pd.DataFrame(data)
else:
    st.error("Failed to fetch data from API")
    st.stop()

# Editable Table
st.write("Edit the data below:")
edited_df = st.data_editor(df, use_container_width=True)

# Save Changes
if st.button("Save Changes"):
    for index, row in edited_df.iterrows():
        response = requests.put(f"{BASE_URL}/update", json={"unique_key": row["unique_key"], "data": row.to_dict()})
        if response.status_code == 200:
            st.success(f"Row {index+1} updated successfully")
        else:
            st.error(f"Failed to update Row {index+1}: {response.json()['detail']}")
