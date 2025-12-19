import streamlit as st
import pandas as pd
from pathlib import Path
import requests
from io import BytesIO

# -------------------------
# Page Config
# -------------------------
st.set_page_config(
    page_title="Bulk Census Enrollment Summary",
    layout="wide"
)

st.title("Bulk Census Upload – Enrollment Summary")

st.markdown(
    """
Upload **multiple census CSV files** to generate an enrollment summary.
Employee location is derived automatically from ZIP codes.
"""
)

# -------------------------
# Constants
# -------------------------
ZIP_MAPPING_URL = (
    "https://raw.githubusercontent.com/kc-zh/CSA-Tools/"
    "46c67cdbf71645d7fe5b08b2799e960c66694352/"
    "State_County_Zip%20Mapping_StreamLit.xlsx"
)

ZIP_COL = "Zip Code"
STATE_COL = "Name"

REQUIRED_CENSUS_COLS = {
    "Zip Code",
    "Health Election",
    "Relationship"
}

# -------------------------
# File Upload
# -------------------------
census_files = st.file_uploader(
    "Upload Census CSV files",
    type=["csv"],
    accept_multiple_files=True
)

# -------------------------
# Helper Functions
# -------------------------
def normalize_zip(zip_val):
    """Normalize ZIP codes to 5-digit strings."""
    if pd.isna(zip_val):
        return None
    zip_str = str(zip_val).split(".")[0]
    return zip_str.zfill(5)[:5]

@st.cache_data
def load_zip_mapping():
    """Load ZIP → State mapping from GitHub safely."""
    response = requests.get(ZIP_MAPPING_URL)
    response.raise_for_status()

    zip_df = pd.read_excel(BytesIO(response.content))
    zip_df[ZIP_COL] = zip_df[ZIP_COL].apply(normalize_zip)

    return dict(zip(zip_df[ZIP_COL], zip_df[STATE_COL]))

# -------------------------
# Main Logic
# -------------------------
if census_files:
    try:
        zip_to_state = load_zip_mapping()
        summary_rows = []

        for file in census_files:
            census_name = Path(file.name).stem
            df = pd.read_csv(file)

            # Clean column names
            df.columns = df.columns.str.strip()

            if not REQUIRED_CENSUS_COLS.issubset(df.columns):
                st.warning(
                    f"Skipping {file.name}: Missing one or more required columns."
                )
                continue

            # Filter: Employee + Enrolled
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

            # Normalize ZIPs
            filtered_df["Zip Code"] = filtered_df["Zip Code"].apply(normalize_zip)

            # Map ZIP → State
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

        # -------------------------
        # Output
        # -------------------------
        st.subheader("Enrollment Summary")
        st.dataframe(summary_df, use_container_width=True)

        st.download_button(
            label="Download Summary as CSV",
            data=summary_df.to_csv(index=False).encode("utf-8"),
            file_name="enrollment_summary.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"Error processing files: {e}")

else:
    st.info("Please upload at least one census CSV to begin.")
