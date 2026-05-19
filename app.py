import streamlit as st
import requests
import pandas as pd
from io import BytesIO

# ==========================
# PAGE TITLE
# ==========================

st.title("TfL Route Fleet Extractor")

# ==========================
# ROUTE INPUT
# ==========================

route = st.text_input("Enter Route Number", "133")

# ==========================
# BUTTON
# ==========================

if st.button("Generate Excel"):

    url = f"https://headway.plumby.io/api/proxy?reg={route}"

    response = requests.get(url)

    if response.status_code != 200:
        st.error("Failed to fetch data")
    else:

        data = response.json()

        rows = []

        for vehicle_reg, vehicle_data in data.items():

            rows.append({
                "Running Number": vehicle_data.get("rnum"),
                "Fleet Number": vehicle_data.get("fnum"),
                "Vehicle Registration": vehicle_data.get("reg")
            })

        df = pd.DataFrame(rows)

        df = df.drop_duplicates()

        df = df.sort_values(by="Running Number")

        st.dataframe(df)

        # ==========================
        # CREATE EXCEL IN MEMORY
        # ==========================

        output = BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)

        output.seek(0)

        # ==========================
        # DOWNLOAD BUTTON
        # ==========================

        st.download_button(
            label="Download Excel",
            data=output,
            file_name=f"Route_{route}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

