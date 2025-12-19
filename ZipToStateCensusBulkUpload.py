import streamlit as st
import pandas as pd
from pathlib import Path

st.set_page_config(page_title="Bulk Census Enrollment Summary", layout="wide")

st.title("Bulk Census Upload – Enrollment Summary")

st.markdown(
    """
Upload **multiple census CSV files** and a **Zip → State mapping file** to generate
an enrollment summary by prospect.
"""
)

# -------------------------
# File Uploads
# -------------------------
census_files = st.file_uploader(
    "Upload Census CSV files",
    type=["csv"],
    accept_multiple_files=True
)

zip_mapping_file = st.file_uploader(
    "Upload Zip Code → State mapping file (Excel)",
    type=["xlsx"]
)

# -------------------------
# Helper Function
# -------------------------
def normalize_zip(zip_val):
    if pd.isna(zip_val):
        return None
    zip_str = str(zip_val).split(".")[0]
    return zip_str.zfill(5)[:5]

# -------------------------
# Main Logic
# -------------------------
if census_files and zip_mapping_file:
    try:
        # Load Zip → State mapping
        zip_df = pd.read_excel(zip_mapping_file)

        ZIP_COL = "Zip Code"
        STATE_COL = "Name"

        zip_df[ZIP_COL] = zip_df[ZIP_COL].apply(normalize_zip)
        zip_to_state = dict(zip(zip_df[ZIP_COL], zip_df[STATE_COL]))

        summary_rows = []

        for file in census_files:
            census_name = Path(file.name).stem
            df = pd.read_csv(file)

            # Clean column names
            df.columns = df.columns.str.strip()

            required_cols = {"Zip Code", "Health Election", "Relationship"}
            if not required_cols.issubset(df.columns):
                st.warning(f"Skipping {file.name}: Missing required columns.")
                continue

            # Apply filters:
            # 1. Relationship == Employee
            # 2. Health Election == enroll
            filtered_df = df[
                (df["Relationship"].astype(str).str.lower() == "employee") &
                (df["Health Election"].astype(str).str.lower() == "enroll")
            ].copy()

            if filtered_df.empty:
                summary_rows.append({
                    "Prospect Name": census_name,
                    "Prospect Location": "",
                    "Projected Enrolled EEs": 0
                })
                continue

            # Normalize zip codes
            filtered_df["Zip Code"] = filtered_df["Zip Code"].apply(normalize_zip)

            # Map zip → state
            filtered_df["State"] = filtered_df["Zip Code"].map(zip_to_state)

            states = (
                filtered_df["State"]
                .dropna()
                .unique()
                .tolist()
            )

            summary_rows.append({
                "Prospect Name": census_name,
                "Prospect Location": ", ".join(sorted(states)),
                "Projected Enrolled EEs": len(filtered_df)
            })

        summary_df = pd.DataFrame(summary_rows)

        st.subheader("Enrollment Summary")
        st.dataframe(summary_df, use_container_width=True)

        # Optional download
        csv = summary_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download Summary as CSV",
            csv,
            "enrollment_summary.csv",
            "text/csv"
        )

    except Exception as e:
        st.error(f"Error processing files: {e}")

else:
    st.info("Please upload at least one census CSV and the Zip → State mapping file.")
