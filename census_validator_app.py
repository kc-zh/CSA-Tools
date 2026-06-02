"""
Census Validator & Reformatter  —  Streamlit App
=================================================
Run:
    pip install streamlit pandas zipcodes openpyxl
    streamlit run census_validator_app.py

Zip-to-State logic mirrors the VBA InsertStateColumn macro:
  • Looks up state abbreviation/name and county from the zip code in
    columns L (Zip Code) and N (Primary Worksite Zip Code).
  • Zero-pads short numeric zip codes (e.g. "601" → "00601"), matching
    the VBA Format(CLng(zipCode), "00000") behaviour.
  • Auto-fills County if the field is blank and a match is found.
  • Flags unrecognised zips as "Not Found" (red in VBA, warning here).
"""

from __future__ import annotations

import io
import re

import pandas as pd
import streamlit as st

# ── Optional zip-lookup ───────────────────────────────────────────────────────
try:
    import zipcodes as _zc
    ZIPCODES_AVAILABLE = True
except ImportError:
    ZIPCODES_AVAILABLE = False

# =============================================================================
# FIELD SCHEMA
# =============================================================================

FIELDS: list[dict] = [
    {"col": "A",  "name": "First Name",                     "required": True,  "type": "alpha"},
    {"col": "B",  "name": "Last Name",                      "required": True,  "type": "alpha"},
    {"col": "C",  "name": "Employee ID",                    "required": False, "type": "alpha"},
    {"col": "D",  "name": "Relationship",                   "required": True,  "type": "enum",
     "values": ["Employee", "Spouse", "Child"]},
    {"col": "E",  "name": "DOB",                            "required": True,  "type": "date"},
    {"col": "F",  "name": "Gender",                         "required": False, "type": "enum",
     "values": ["1", "2", ""]},
    {"col": "G",  "name": "Email",                          "required": False, "type": "alpha"},
    {"col": "H",  "name": "Address Line 1",                 "required": False, "type": "alpha"},
    {"col": "I",  "name": "Address Line 2",                 "required": False, "type": "alpha"},
    {"col": "J",  "name": "City",                           "required": False, "type": "alpha"},
    {"col": "K",  "name": "State",                          "required": False, "type": "alpha"},
    {"col": "L",  "name": "Zip Code",                       "required": True,  "type": "zipcode"},
    {"col": "M",  "name": "County",                         "required": False, "type": "alpha"},
    {"col": "N",  "name": "Primary Worksite Zip Code",      "required": True,  "type": "zipcode"},
    {"col": "O",  "name": "Primary Worksite County",        "required": False, "type": "alpha"},
    {"col": "P",  "name": "ICHRA Class",                    "required": False, "type": "alpha"},
    {"col": "Q",  "name": "Health Election",                "required": True,  "type": "enum",
     "values": ["Enroll", "E", "Waive", "W"]},
    {"col": "R",  "name": "Current Health Vendor",          "required": False, "type": "alpha"},
    {"col": "S",  "name": "Current Health Plan",            "required": False, "type": "alpha"},
    {"col": "T",  "name": "Current Health Plan Tier",       "required": False, "type": "enum",
     "values": ["Employee Only", "Employee + Spouse", "Employee + Children", "Family", ""]},
    {"col": "U",  "name": "Health Plan OOP (single)",       "required": False, "type": "numeric"},
    {"col": "V",  "name": "Health Plan OOP (family)",       "required": False, "type": "numeric"},
    {"col": "W",  "name": "Health Plan Deductible (single)", "required": False, "type": "numeric"},
    {"col": "X",  "name": "Health Plan Deductible (family)", "required": False, "type": "numeric"},
    {"col": "Y",  "name": "Current Health Plan ER Cost",    "required": False, "type": "numeric",
     "note": "Required for Cost Comparison"},
    {"col": "Z",  "name": "Current Health Plan EE Cost",    "required": False, "type": "numeric",
     "note": "Required for Cost Comparison"},
    {"col": "AA", "name": "Annual Salary",                  "required": False, "type": "numeric",
     "note": "Required for Affordability Testing"},
    {"col": "AB", "name": "Hourly Rate",                    "required": False, "type": "numeric",
     "note": "Required for Affordability Testing"},
    {"col": "AC", "name": "Hours Per Week",                 "required": False, "type": "numeric",
     "note": "Required for Affordability Testing"},
    {"col": "AD", "name": "Notes",                          "required": False, "type": "alpha"},
]

# Field indices for quick reference
_FIELD_IDX_BY_NAME: dict[str, int] = {f["name"]: i for i, f in enumerate(FIELDS)}

# =============================================================================
# ZIP-CODE LOOKUP  (mirrors VBA ZipStateMapping dictionary)
# =============================================================================

@st.cache_data(show_spinner=False)
def build_zip_cache() -> dict[str, dict]:
    """
    Build {zip_str: {"abbr": "WI", "county": "Milwaukee"}} once per session.
    Equivalent to loading the PERSONAL.XLSB ZipStateMapping sheet into a
    Scripting.Dictionary in the VBA macro.
    """
    if not ZIPCODES_AVAILABLE:
        return {}
    cache: dict[str, dict] = {}
    for entry in _zc.list_all():
        z = entry.get("zip_code", "")
        if z and z not in cache:
            county_raw = (entry.get("county") or "")
            county = county_raw.replace(" County", "").strip()
            cache[z] = {
                "abbr":   entry.get("state", ""),
                "county": county,
            }
    return cache


def _pad_zip(z: str) -> str:
    """
    Mirrors VBA: Format(CLng(zipCode), "00000")
    Zero-pads numeric zips shorter than 5 digits.
    """
    if z.isdigit() and 0 < len(z) < 5:
        return z.zfill(5)
    return z


def lookup_zip(zip_raw: str, cache: dict) -> dict | None:
    z = _pad_zip(str(zip_raw).strip())
    return cache.get(z)

# =============================================================================
# NORMALIZATION
# =============================================================================

def normalize_date(v: str) -> str:
    v = v.strip()
    if re.match(r"^\d{2}/\d{2}/\d{4}$", v):
        return v
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", v)
    if m:
        return f"{m.group(1).zfill(2)}/{m.group(2).zfill(2)}/{m.group(3)}"
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", v)
    if m:
        return f"{m.group(2)}/{m.group(3)}/{m.group(1)}"
    return v


def normalize_enum(field: dict, v: str) -> str:
    for fv in field.get("values", []):
        if fv.lower() == v.lower():
            return fv
    return v

# =============================================================================
# CELL-LEVEL VALIDATOR
# =============================================================================

def validate_cell(field: dict, raw: str) -> dict:
    """
    Returns:
        ok        – bool
        msg       – error string or None
        fixed_val – cleaned/normalised value
        fix_note  – human-readable description of the auto-fix, or None
    """
    v = str(raw).strip() if raw is not None else ""
    fixed = v
    fix_note = None

    if field["required"] and v == "":
        return {"ok": False, "msg": f"{field['name']} is required",
                "fixed_val": v, "fix_note": None}
    if v == "":
        return {"ok": True, "msg": None, "fixed_val": v, "fix_note": None}

    ftype = field["type"]

    # ── Date ──────────────────────────────────────────────────────────────────
    if ftype == "date":
        norm = normalize_date(v)
        if norm != v:
            fix_note = f"Reformatted date: '{v}' → '{norm}'"
            fixed = norm
        v = fixed
        m = re.match(r"^(\d{2})/(\d{2})/(\d{4})$", v)
        if not m:
            return {"ok": False, "msg": f"Invalid date format '{v}' — expected mm/dd/yyyy",
                    "fixed_val": fixed, "fix_note": fix_note}
        mo, d, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if not (1 <= mo <= 12):
            return {"ok": False, "msg": f"Month out of range in '{v}'",
                    "fixed_val": fixed, "fix_note": fix_note}
        if not (1 <= d <= 31):
            return {"ok": False, "msg": f"Day out of range in '{v}'",
                    "fixed_val": fixed, "fix_note": fix_note}
        if not (1900 <= yr <= 2100):
            return {"ok": False, "msg": f"Year out of range in '{v}'",
                    "fixed_val": fixed, "fix_note": fix_note}
        return {"ok": True, "msg": None, "fixed_val": fixed, "fix_note": fix_note}

    # ── Numeric ───────────────────────────────────────────────────────────────
    if ftype == "numeric":
        clean = v.replace(",", "").replace("$", "").strip()
        if not re.match(r"^-?\d+(\.\d+)?$", clean):
            return {"ok": False, "msg": f"{field['name']} must be numeric, got '{v}'",
                    "fixed_val": v, "fix_note": None}
        if clean != v:
            fix_note = f"Stripped non-numeric characters: '{v}' → '{clean}'"
        return {"ok": True, "msg": None, "fixed_val": clean, "fix_note": fix_note}

    # ── Zip code ──────────────────────────────────────────────────────────────
    if ftype == "zipcode":
        padded = _pad_zip(v)
        if padded != v:
            fix_note = f"Zero-padded zip: '{v}' → '{padded}'"
            fixed = padded
        if not re.match(r"^\d{5}(-\d{4})?$", fixed):
            return {"ok": False, "msg": f"{field['name']} must be a 5-digit zip, got '{v}'",
                    "fixed_val": fixed, "fix_note": fix_note}
        return {"ok": True, "msg": None, "fixed_val": fixed, "fix_note": fix_note}

    # ── Enum ──────────────────────────────────────────────────────────────────
    if ftype == "enum":
        canonical = normalize_enum(field, v)
        if canonical not in field["values"]:
            valid_vals = ", ".join(x for x in field["values"] if x)
            return {"ok": False,
                    "msg": f"Invalid value '{v}' for {field['name']} — expected: {valid_vals}",
                    "fixed_val": v, "fix_note": None}
        if canonical != v:
            fix_note = f"Casing corrected: '{v}' → '{canonical}'"
        return {"ok": True, "msg": None, "fixed_val": canonical, "fix_note": fix_note}

    # ── Alpha (no content restriction) ───────────────────────────────────────
    return {"ok": True, "msg": None, "fixed_val": v, "fix_note": None}

# =============================================================================
# FILE INGESTION
# =============================================================================

def read_uploaded_file(uploaded) -> pd.DataFrame:
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded, dtype=str, keep_default_na=False)
    elif name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded, dtype=str, keep_default_na=False)
    raise ValueError(f"Unsupported file type: {uploaded.name}")


def map_columns(df: pd.DataFrame) -> dict[int, int]:
    """
    Map FIELDS index → DataFrame column index.
    Accepts column-letter headers (A, B … AA, AB …) OR full field names.
    """
    headers     = [str(c).strip() for c in df.columns]
    headers_up  = [h.upper() for h in headers]
    headers_low = [h.lower() for h in headers]
    mapping: dict[int, int] = {}
    for fi, field in enumerate(FIELDS):
        letter = field["col"].upper()
        name_l = field["name"].lower()
        for ci, (hu, hl) in enumerate(zip(headers_up, headers_low)):
            if hu == letter or hl == name_l:
                mapping[fi] = ci
                break
    return mapping

# =============================================================================
# CORE VALIDATION + REFORMAT ENGINE
# =============================================================================

def run_validation(
    df: pd.DataFrame,
    insert_state: bool,
    state_output: str,   # "ABBR" or "FULL"
    zip_cache: dict,
) -> dict:
    """
    Returns a results dict with:
        errors, warnings, fixes  – lists of issue dicts
        reformatted_df           – cleaned DataFrame
        valid_positions          – list[int] of reformatted_df row indices with no errors
        error_positions          – list[int] of reformatted_df row indices with errors
        total_rows               – int
    """
    col_map = map_columns(df)
    errors:   list[dict] = []
    warnings: list[dict] = []
    fixes:    list[dict] = []

    out_rows:        list[dict] = []
    valid_positions: list[int]  = []
    error_positions: list[int]  = []
    output_pos = 0

    # ── Per-row loop ──────────────────────────────────────────────────────────
    for df_ri, row in df.iterrows():
        # Skip entirely blank rows
        if all(str(v).strip() == "" for v in row):
            continue

        display_row = int(df_ri) + 2   # 1-based row number with header offset
        new_row: dict[str, str] = {}
        row_has_error = False

        # Validate every field
        for fi, field in enumerate(FIELDS):
            ci  = col_map.get(fi)
            raw = str(row.iloc[ci]).strip() if ci is not None else ""
            res = validate_cell(field, raw)

            if not res["ok"]:
                errors.append({
                    "Row": display_row, "Col": field["col"],
                    "Field": field["name"], "Value": raw,
                    "Issue": res["msg"], "Kind": "Error",
                })
                row_has_error = True

            if res["fix_note"]:
                fixes.append({
                    "Row": display_row, "Col": field["col"],
                    "Field": field["name"], "Value": raw,
                    "Issue": res["fix_note"], "Kind": "Auto-fix",
                })

            new_row[field["name"]] = res["fixed_val"]

        # ── Zip → State injection (VBA InsertStateColumn parity) ─────────────
        if insert_state and ZIPCODES_AVAILABLE:
            # Pairs: (census zip field, injected state col name, county field name)
            zip_pairs = [
                ("Zip Code",                 "State (from Zip)",                  "County"),
                ("Primary Worksite Zip Code","Primary Worksite State (from Zip)", "Primary Worksite County"),
            ]
            for zip_fn, state_cn, county_fn in zip_pairs:
                zv = new_row.get(zip_fn, "").strip()
                if not zv:
                    continue
                info = lookup_zip(zv, zip_cache)
                if info:
                    state_val = (
                        info["abbr"] if state_output == "ABBR"
                        else info.get("state_name", info["abbr"])
                    )
                    new_row[state_cn] = state_val
                    # Auto-fill county if blank — mirrors VBA "Not Found" flow
                    existing_county = new_row.get(county_fn, "").strip()
                    if not existing_county and info.get("county"):
                        new_row[county_fn] = info["county"]
                        fixes.append({
                            "Row": display_row, "Col": "—",
                            "Field": county_fn, "Value": "",
                            "Issue": f"County auto-filled from zip lookup: '{info['county']}'",
                            "Kind": "Auto-fix",
                        })
                else:
                    # VBA equivalent: .Value = "Not Found" + red font
                    new_row[state_cn] = "Not Found"
                    warnings.append({
                        "Row": display_row,
                        "Col": "L" if zip_fn == "Zip Code" else "N",
                        "Field": zip_fn, "Value": zv,
                        "Issue": f"Zip code '{zv}' not found in lookup table",
                        "Kind": "Warning",
                    })

        # Track row position
        if row_has_error:
            error_positions.append(output_pos)
        else:
            valid_positions.append(output_pos)

        out_rows.append(new_row)
        output_pos += 1

    # ── Cross-row warnings ────────────────────────────────────────────────────
    # Cost Comparison: Y and Z should both be present or both absent
    ci_y = col_map.get(_FIELD_IDX_BY_NAME["Current Health Plan ER Cost"])
    ci_z = col_map.get(_FIELD_IDX_BY_NAME["Current Health Plan EE Cost"])
    for df_ri, row in df.iterrows():
        if all(str(v).strip() == "" for v in row):
            continue
        display_row = int(df_ri) + 2
        y_val = str(row.iloc[ci_y]).strip() if ci_y is not None else ""
        z_val = str(row.iloc[ci_z]).strip() if ci_z is not None else ""
        if bool(y_val) ^ bool(z_val):
            warnings.append({
                "Row": display_row, "Col": "Y / Z",
                "Field": "Cost Comparison", "Value": "",
                "Issue": "Both ER Cost (Y) and EE Cost (Z) should be present for Cost Comparison",
                "Kind": "Warning",
            })

    # ── Build reformatted DataFrame ───────────────────────────────────────────
    if insert_state and ZIPCODES_AVAILABLE:
        col_order: list[str] = []
        for f in FIELDS:
            col_order.append(f["name"])
            if f["name"] == "Zip Code":
                col_order.append("State (from Zip)")
            if f["name"] == "Primary Worksite Zip Code":
                col_order.append("Primary Worksite State (from Zip)")
    else:
        col_order = [f["name"] for f in FIELDS]

    reformatted_df = pd.DataFrame(out_rows)
    for c in col_order:
        if c not in reformatted_df.columns:
            reformatted_df[c] = ""
    reformatted_df = reformatted_df[col_order].reset_index(drop=True)

    return {
        "errors":           errors,
        "warnings":         warnings,
        "fixes":            fixes,
        "reformatted_df":   reformatted_df,
        "valid_positions":  valid_positions,
        "error_positions":  error_positions,
        "total_rows":       len(out_rows),
    }

# =============================================================================
# DOWNLOAD HELPERS
# =============================================================================

def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")   # utf-8-sig for Excel compatibility


def issues_to_csv_bytes(errors, warnings, fixes) -> bytes:
    combined = errors + warnings + fixes
    combined.sort(key=lambda x: x["Row"])
    out = pd.DataFrame(combined, columns=["Row", "Col", "Field", "Value", "Issue", "Kind"])
    return out.to_csv(index=False).encode("utf-8-sig")

# =============================================================================
# STREAMLIT APP
# =============================================================================

SAMPLE_CSV = """\
A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD
John,Smith,EMP001,Employee,01/15/1980,1,john@example.com,123 Main St,,Milwaukee,WI,53201,,53201,,,Enroll,Anthem,Gold PPO,Employee Only,2000,6000,1000,3000,500,200,60000,,,
Jane,Smith,,spouse,03/22/1982,2,,,,,53201,,53201,,,enroll,,,,,,,,,,,,
Tim,Smith,,Child,6/10/2010,1,,,,,,53201,,53201,,,E,,,,,,,,,,,,
,Johnson,EMP002,Employee,13/45/1975,3,,,,,ABCD,,99999,,,Waive,,,,,,,,,,,,
Mike,,EMP003,Partner,1975-05-10,1,,,,,,53210,,53210,,,EnrolL,,,,,,500,,
"""


def main() -> None:
    st.set_page_config(
        page_title="Census Validator",
        page_icon="📋",
        layout="wide",
    )

    st.markdown("""
    <style>
        .block-container { padding-top: 1.75rem; max-width: 1160px; }
        [data-testid="metric-container"] { background: #f7f7f5; border-radius: 8px; padding: 10px 14px; }
        .stTabs [data-baseweb="tab"] { font-size: 13px; }
        .dl-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 12px; margin-bottom: 1rem; }
        .dl-tile {
            border: 1px solid #e0e0e0; border-radius: 10px;
            padding: 14px 16px; background: #fff;
        }
        .dl-tile h5 { margin: 0 0 3px; font-size: 14px; }
        .dl-tile p  { margin: 0 0 10px; font-size: 12px; color: #666; }
        .banner-ok {
            background: #EAF3DE; border: 1px solid #97C459; border-radius: 8px;
            padding: 12px 16px; color: #3B6D11; font-size: 14px; margin-bottom: 1rem;
        }
    </style>
    """, unsafe_allow_html=True)

    # ── Header ────────────────────────────────────────────────────────────────
    st.title("📋 Census Validator & Reformatter")
    st.caption(
        "Upload a `.csv` or `.xlsx` census file to validate all 30 fields, "
        "auto-correct common issues, look up State/County from zip codes, "
        "and download a clean reformatted file."
    )

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️ Options")

        insert_state = st.toggle(
            "Inject State from Zip Code",
            value=ZIPCODES_AVAILABLE,
            disabled=not ZIPCODES_AVAILABLE,
            help=(
                "Mirrors the VBA InsertStateColumn macro. "
                "Looks up state & county from zip columns L and N, "
                "zero-pads short zips, and auto-fills blank County fields."
            ),
        )
        if not ZIPCODES_AVAILABLE:
            st.caption("⚠️ Install `zipcodes` to enable: `pip install zipcodes`")

        state_output = st.radio(
            "State format",
            ["ABBR", "FULL"],
            index=0,
            format_func=lambda x: "Abbreviation  (e.g. WI)" if x == "ABBR" else "Full name  (e.g. Wisconsin)",
            disabled=not insert_state,
        )

        st.divider()
        st.markdown("**Supported input formats**")
        st.caption("`.csv`  ·  `.xlsx`  ·  `.xls`")
        st.caption("Headers: column letters (A–AD) **or** field names")
        st.divider()
        st.markdown("**Validation rules**")
        st.caption("• Required-field checks  \n• Date format mm/dd/yyyy  \n• Numeric-only fields  \n• 5-digit zip codes  \n• Enum value lists  \n• Cost Comparison pair check")

    # ── Input area ────────────────────────────────────────────────────────────
    col_up, col_paste = st.columns([1, 1], gap="large")

    with col_up:
        uploaded = st.file_uploader(
            "Upload census file  (.csv / .xlsx)",
            type=["csv", "xlsx", "xls"],
        )

    with col_paste:
        # Preserve sample text across reruns
        default_paste = st.session_state.get("paste_text", "")
        pasted = st.text_area(
            "— or paste CSV data —",
            value=default_paste,
            height=150,
            placeholder="Paste CSV including header row…",
        )

    btn_col1, btn_col2, btn_col3 = st.columns([1, 1, 5])
    with btn_col1:
        run_btn = st.button("✅  Validate", type="primary", use_container_width=True)
    with btn_col2:
        if st.button("📄  Load sample", use_container_width=True):
            st.session_state["paste_text"] = SAMPLE_CSV
            st.rerun()

    # ── Execute validation ────────────────────────────────────────────────────
    if run_btn:
        df_input: pd.DataFrame | None = None
        try:
            if uploaded:
                df_input = read_uploaded_file(uploaded)
            elif pasted.strip():
                df_input = pd.read_csv(io.StringIO(pasted), dtype=str, keep_default_na=False)
            else:
                st.warning("Please upload a file or paste CSV data before clicking Validate.")
        except Exception as exc:
            st.error(f"Could not read input: {exc}")

        if df_input is not None:
            with st.spinner("Validating and reformatting…"):
                zc = build_zip_cache() if (insert_state and ZIPCODES_AVAILABLE) else {}
                results = run_validation(df_input, insert_state, state_output, zc)

            # Derive valid / error subset DataFrames now, while we have results
            ref_df = results["reformatted_df"]
            valid_df = ref_df.iloc[results["valid_positions"]].reset_index(drop=True)
            error_df = ref_df.iloc[results["error_positions"]].reset_index(drop=True)
            results["valid_df"] = valid_df
            results["error_df"] = error_df

            st.session_state["results"] = results

    # ── Render results ────────────────────────────────────────────────────────
    if "results" not in st.session_state:
        return

    res       = st.session_state["results"]
    errors    = res["errors"]
    warnings  = res["warnings"]
    fixes     = res["fixes"]
    ref_df    = res["reformatted_df"]
    valid_df  = res["valid_df"]
    error_df  = res["error_df"]
    total     = res["total_rows"]
    err_rows  = len(res["error_positions"])
    valid_rows = total - err_rows

    st.divider()

    # ── Metrics ───────────────────────────────────────────────────────────────
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total rows",       total)
    c2.metric("Valid rows",        valid_rows)
    c3.metric("Rows with errors",  err_rows)
    c4.metric("Errors",            len(errors))
    c5.metric("Warnings",          len(warnings))
    c6.metric("Auto-fixes",        len(fixes))

    st.divider()

    # ── Download panel ────────────────────────────────────────────────────────
    st.subheader("📥 Download census")

    zip_note = (
        "State & County injected from zip lookup · " if insert_state and ZIPCODES_AVAILABLE else ""
    )
    st.caption(
        f"{zip_note}Dates normalised to mm/dd/yyyy · "
        "Enum casing corrected · Zip codes zero-padded · Blank rows removed"
    )

    dc1, dc2, dc3, dc4 = st.columns(4)

    with dc1:
        st.markdown("##### Full reformatted census")
        st.caption(f"All {total} rows, cleaned and standardised")
        st.download_button(
            "⬇ Download",
            data=df_to_csv_bytes(ref_df),
            file_name="census_reformatted.csv",
            mime="text/csv",
            use_container_width=True,
            key="dl_full",
        )

    with dc2:
        st.markdown("##### Valid rows only")
        st.caption(f"{valid_rows} rows that passed all validations")
        st.download_button(
            "⬇ Download",
            data=df_to_csv_bytes(valid_df),
            file_name="census_valid_rows.csv",
            mime="text/csv",
            use_container_width=True,
            key="dl_valid",
        )

    with dc3:
        st.markdown("##### Error rows only")
        st.caption(f"{err_rows} rows needing manual review")
        st.download_button(
            "⬇ Download",
            data=df_to_csv_bytes(error_df),
            file_name="census_error_rows.csv",
            mime="text/csv",
            use_container_width=True,
            key="dl_errors",
        )

    with dc4:
        st.markdown("##### Validation report")
        st.caption("All errors, warnings & auto-fix notes")
        st.download_button(
            "⬇ Download",
            data=issues_to_csv_bytes(errors, warnings, fixes),
            file_name="census_validation_report.csv",
            mime="text/csv",
            use_container_width=True,
            key="dl_report",
        )

    st.divider()

    # ── Issues table ──────────────────────────────────────────────────────────
    all_issues = errors + warnings + fixes

    if not all_issues:
        st.markdown(
            f'<div class="banner-ok">✅  All {total} rows passed validation with no issues.</div>',
            unsafe_allow_html=True,
        )
    else:
        tab_all, tab_err, tab_warn, tab_fix = st.tabs([
            f"All  ({len(all_issues)})",
            f"Errors  ({len(errors)})",
            f"Warnings  ({len(warnings)})",
            f"Auto-fixes  ({len(fixes)})",
        ])

        def _styled(df_issues: list[dict]) -> pd.io.formats.style.Styler:
            df = pd.DataFrame(
                df_issues, columns=["Row", "Col", "Field", "Value", "Issue", "Kind"]
            ).sort_values("Row").reset_index(drop=True)

            def _color(val):
                if val == "Error":    return "color:#993C1D;font-weight:600"
                if val == "Warning":  return "color:#854F0B;font-weight:600"
                if val == "Auto-fix": return "color:#3B6D11;font-weight:600"
                return ""

            # applymap was renamed to map in pandas 2.1
            styler = df.style
            if hasattr(styler, "map"):
                return styler.map(_color, subset=["Kind"])
            return styler.applymap(_color, subset=["Kind"])

        with tab_all:
            if all_issues:
                st.dataframe(_styled(all_issues), use_container_width=True, hide_index=True)
        with tab_err:
            if errors:
                st.dataframe(_styled(errors), use_container_width=True, hide_index=True)
            else:
                st.success("No errors found.")
        with tab_warn:
            if warnings:
                st.dataframe(_styled(warnings), use_container_width=True, hide_index=True)
            else:
                st.success("No warnings found.")
        with tab_fix:
            if fixes:
                st.dataframe(_styled(fixes), use_container_width=True, hide_index=True)
            else:
                st.info("No auto-fixes applied.")

    # ── Data preview ──────────────────────────────────────────────────────────
    st.divider()
    with st.expander("🔍 Preview reformatted data", expanded=False):
        st.dataframe(ref_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
