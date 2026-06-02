"""
Census Validator & Reformatter  —  Streamlit App
=================================================
Install:  pip install streamlit pandas zipcodes openpyxl rapidfuzz
Run:      streamlit run census_validator_app.py

Key behaviours
──────────────
• Fuzzy + exact header matching: maps ANY reasonable column naming to canonical fields.
• Fixed output column order (30 fields) regardless of input order.
• Cleared fields: Employee ID, Gender, Email, Address Line 1, Address Line 2, City, State
  are always blanked in output regardless of input values.
• Relationship normalisation: Partner / Domestic Partner / Life Partner / etc. → Spouse.
• Health Election normalisation: Enroll / E → Enroll  |  Waive / W → Waive.
• Zip-to-County lookup (VBA InsertStateColumn parity).
• Validation report: errors, warnings, auto-fixes, cleared fields, unrecognised columns.
• Pandas 2.1+ compatible (uses Styler.map not applymap).
"""

from __future__ import annotations

import io
import re
from typing import Optional

import pandas as pd
import streamlit as st

# ── Optional dependencies ─────────────────────────────────────────────────────
try:
    import zipcodes as _zc
    ZIPCODES_AVAILABLE = True
except ImportError:
    ZIPCODES_AVAILABLE = False

try:
    from rapidfuzz import fuzz, process as rf_process
    FUZZY_AVAILABLE = True
except ImportError:
    FUZZY_AVAILABLE = False

# =============================================================================
# 1. CANONICAL FIELD SCHEMA
# =============================================================================

# Fields that are ALWAYS blanked in the output
CLEARED_FIELDS = {
    "Employee ID", "Gender", "Email",
    "Address Line 1", "Address Line 2", "City", "State",
}

FIELDS: list[dict] = [
    # col  canonical name                  required  type       enum values / notes
    {"col": "A",  "name": "First Name",                      "required": True,  "type": "alpha"},
    {"col": "B",  "name": "Last Name",                       "required": True,  "type": "alpha"},
    {"col": "C",  "name": "Employee ID",                     "required": False, "type": "alpha"},   # CLEARED
    {"col": "D",  "name": "Relationship",                    "required": True,  "type": "relationship"},
    {"col": "E",  "name": "DOB",                             "required": True,  "type": "date"},
    {"col": "F",  "name": "Gender",                          "required": False, "type": "alpha"},   # CLEARED
    {"col": "G",  "name": "Email",                           "required": False, "type": "alpha"},   # CLEARED
    {"col": "H",  "name": "Address Line 1",                  "required": False, "type": "alpha"},   # CLEARED
    {"col": "I",  "name": "Address Line 2",                  "required": False, "type": "alpha"},   # CLEARED
    {"col": "J",  "name": "City",                            "required": False, "type": "alpha"},   # CLEARED
    {"col": "K",  "name": "State",                           "required": False, "type": "alpha"},   # CLEARED
    {"col": "L",  "name": "Zip Code",                        "required": True,  "type": "zipcode"},
    {"col": "M",  "name": "County",                          "required": False, "type": "alpha"},
    {"col": "N",  "name": "Primary Worksite Zip Code",       "required": True,  "type": "zipcode"},
    {"col": "O",  "name": "Primary Worksite County",         "required": False, "type": "alpha"},
    {"col": "P",  "name": "ICHRA Class",                     "required": False, "type": "alpha"},
    {"col": "Q",  "name": "Health Election",                 "required": True,  "type": "election"},
    {"col": "R",  "name": "Current Health Plan Vendor",      "required": False, "type": "alpha"},
    {"col": "S",  "name": "Current Health Plan",             "required": False, "type": "alpha"},
    {"col": "T",  "name": "Current Health Plan Tier",        "required": False, "type": "enum",
     "values": ["Employee Only", "Employee + Spouse", "Employee + Children", "Family", ""]},
    {"col": "U",  "name": "Current Health Plan OOP (single)","required": False, "type": "numeric"},
    {"col": "V",  "name": "Current Health Plan OOP (family)","required": False, "type": "numeric"},
    {"col": "W",  "name": "Current Health Plan Deductible (single)", "required": False, "type": "numeric"},
    {"col": "X",  "name": "Current Health Plan Deductible (family)", "required": False, "type": "numeric"},
    {"col": "Y",  "name": "Current Health Plan ER Cost",     "required": False, "type": "numeric",
     "note": "Required for Cost Comparison"},
    {"col": "Z",  "name": "Current Health Plan EE Cost",     "required": False, "type": "numeric",
     "note": "Required for Cost Comparison"},
    {"col": "AA", "name": "Annual Salary",                   "required": False, "type": "numeric",
     "note": "Required for Affordability Testing"},
    {"col": "AB", "name": "Hourly Rate",                     "required": False, "type": "numeric",
     "note": "Required for Affordability Testing"},
    {"col": "AC", "name": "Hours Per Week",                  "required": False, "type": "numeric",
     "note": "Required for Affordability Testing"},
    {"col": "AD", "name": "Notes",                           "required": False, "type": "alpha"},
]

OUTPUT_COLUMNS = [f["name"] for f in FIELDS]   # canonical order for all outputs
_FIELD_BY_NAME = {f["name"]: f for f in FIELDS}
_FIELD_IDX_BY_NAME = {f["name"]: i for i, f in enumerate(FIELDS)}

# =============================================================================
# 2. HEADER ALIASES  (fuzzy matching seed list)
# =============================================================================

# Maps lower-case alias → canonical field name
HEADER_ALIASES: dict[str, str] = {
    # First / Last name
    "first name": "First Name", "firstname": "First Name", "first": "First Name",
    "given name": "First Name", "employee first name": "First Name",
    "last name": "Last Name", "lastname": "Last Name", "last": "Last Name",
    "surname": "Last Name", "family name": "Last Name", "employee last name": "Last Name",
    # Employee ID
    "employee id": "Employee ID", "emp id": "Employee ID", "empid": "Employee ID",
    "employee number": "Employee ID", "emp number": "Employee ID", "id": "Employee ID",
    "employee #": "Employee ID",
    # Relationship
    "relationship": "Relationship", "relation": "Relationship",
    "dependent type": "Relationship", "member type": "Relationship",
    # DOB
    "dob": "DOB", "date of birth": "DOB", "birth date": "DOB", "birthdate": "DOB",
    "date of birth (mm/dd/yyyy)": "DOB",
    # Gender
    "gender": "Gender", "sex": "Gender",
    # Email
    "email": "Email", "email address": "Email", "e-mail": "Email",
    "work email": "Email", "personal email": "Email",
    # Address
    "address line 1": "Address Line 1", "address1": "Address Line 1",
    "address line1": "Address Line 1", "street address": "Address Line 1",
    "address": "Address Line 1",
    "address line 2": "Address Line 2", "address2": "Address Line 2",
    "address line2": "Address Line 2", "apt": "Address Line 2", "suite": "Address Line 2",
    # City / State / Zip / County
    "city": "City", "town": "City",
    "state": "State", "st": "State", "state code": "State",
    "zip code": "Zip Code", "zip": "Zip Code", "zipcode": "Zip Code",
    "postal code": "Zip Code", "zip/postal": "Zip Code",
    "county": "County",
    # Worksite
    "primary worksite zip code": "Primary Worksite Zip Code",
    "worksite zip": "Primary Worksite Zip Code", "work zip": "Primary Worksite Zip Code",
    "worksite zip code": "Primary Worksite Zip Code", "site zip": "Primary Worksite Zip Code",
    "primary worksite county": "Primary Worksite County",
    "worksite county": "Primary Worksite County", "work county": "Primary Worksite County",
    # ICHRA
    "ichra class": "ICHRA Class", "ichra": "ICHRA Class", "class": "ICHRA Class",
    "benefit class": "ICHRA Class",
    # Health Election
    "health election": "Health Election", "election": "Health Election",
    "coverage election": "Health Election", "medical election": "Health Election",
    "benefit election": "Health Election",
    # Health plan fields
    "current health vendor": "Current Health Plan Vendor",
    "current health plan vendor": "Current Health Plan Vendor",
    "health vendor": "Current Health Plan Vendor", "carrier": "Current Health Plan Vendor",
    "insurance carrier": "Current Health Plan Vendor", "health carrier": "Current Health Plan Vendor",
    "current health plan": "Current Health Plan", "health plan": "Current Health Plan",
    "plan name": "Current Health Plan", "medical plan": "Current Health Plan",
    "current health plan tier": "Current Health Plan Tier",
    "health plan tier": "Current Health Plan Tier", "coverage tier": "Current Health Plan Tier",
    "plan tier": "Current Health Plan Tier", "tier": "Current Health Plan Tier",
    "current health plan oop (single)": "Current Health Plan OOP (single)",
    "health plan oop (single)": "Current Health Plan OOP (single)",
    "oop single": "Current Health Plan OOP (single)", "oop (single)": "Current Health Plan OOP (single)",
    "out of pocket single": "Current Health Plan OOP (single)",
    "current health plan oop (family)": "Current Health Plan OOP (family)",
    "health plan oop (family)": "Current Health Plan OOP (family)",
    "oop family": "Current Health Plan OOP (family)", "oop (family)": "Current Health Plan OOP (family)",
    "out of pocket family": "Current Health Plan OOP (family)",
    "current health plan deductible (single)": "Current Health Plan Deductible (single)",
    "health plan deductible (single)": "Current Health Plan Deductible (single)",
    "deductible single": "Current Health Plan Deductible (single)",
    "deductible (single)": "Current Health Plan Deductible (single)",
    "current health plan deductible (family)": "Current Health Plan Deductible (family)",
    "health plan deductible (family)": "Current Health Plan Deductible (family)",
    "deductible family": "Current Health Plan Deductible (family)",
    "deductible (family)": "Current Health Plan Deductible (family)",
    "current health plan er cost": "Current Health Plan ER Cost",
    "health plan er cost": "Current Health Plan ER Cost",
    "er cost": "Current Health Plan ER Cost", "employer cost": "Current Health Plan ER Cost",
    "employer contribution": "Current Health Plan ER Cost",
    "current health plan ee cost": "Current Health Plan EE Cost",
    "health plan ee cost": "Current Health Plan EE Cost",
    "ee cost": "Current Health Plan EE Cost", "employee cost": "Current Health Plan EE Cost",
    "employee contribution": "Current Health Plan EE Cost",
    # Compensation
    "annual salary": "Annual Salary", "salary": "Annual Salary",
    "yearly salary": "Annual Salary", "base salary": "Annual Salary",
    "hourly rate": "Hourly Rate", "hourly": "Hourly Rate", "rate": "Hourly Rate",
    "wage": "Hourly Rate", "hourly wage": "Hourly Rate",
    "hours per week": "Hours Per Week", "hours/week": "Hours Per Week",
    "weekly hours": "Hours Per Week", "hrs per week": "Hours Per Week",
    # Notes
    "notes": "Notes", "note": "Notes", "comments": "Notes", "comment": "Notes",
    "remarks": "Notes",
    # Column-letter pass-through (A→AD)
    **{
        col: f["name"]
        for col, f in zip(
            ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O",
             "P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD"],
            FIELDS
        )
    },
    # Old field names from previous spec
    "current health vendor": "Current Health Plan Vendor",
    "health plan oop (single)": "Current Health Plan OOP (single)",
    "health plan oop (family)": "Current Health Plan OOP (family)",
    "health plan deductible (single)": "Current Health Plan Deductible (single)",
    "health plan deductible (family)": "Current Health Plan Deductible (family)",
}

# =============================================================================
# 3. RELATIONSHIP NORMALISATION TABLE
# =============================================================================

_RELATIONSHIP_MAP: dict[str, str] = {}

def _add_rel(canonical: str, *aliases: str) -> None:
    for a in aliases:
        _RELATIONSHIP_MAP[a.lower()] = canonical

_add_rel("Employee",
    "employee", "ee", "emp", "subscriber", "self", "primary",
    "insured", "member",
)
_add_rel("Spouse",
    "spouse", "sp", "husband", "wife", "partner", "domestic partner",
    "life partner", "registered domestic partner", "common law spouse",
    "common-law spouse", "common law partner", "significant other",
    "spse", "sps",
)
_add_rel("Child",
    "child", "ch", "dependent", "dep", "daughter", "son",
    "stepchild", "step child", "adopted child", "foster child",
    "stepson", "stepdaughter",
)


def normalize_relationship(v: str) -> tuple[str, Optional[str]]:
    """
    Returns (canonical_value, fix_note_or_None).
    If unrecognised returns (original_value, None) so the validator can flag it.
    """
    key = v.strip().lower()
    if key in _RELATIONSHIP_MAP:
        canonical = _RELATIONSHIP_MAP[key]
        note = f"Relationship normalised: '{v}' → '{canonical}'" if canonical.lower() != key else None
        return canonical, note

    # Fuzzy fallback
    if FUZZY_AVAILABLE:
        best, score, _ = rf_process.extractOne(
            key, _RELATIONSHIP_MAP.keys(), scorer=fuzz.ratio
        )
        if score >= 80:
            canonical = _RELATIONSHIP_MAP[best]
            return canonical, f"Relationship fuzzy-matched: '{v}' → '{canonical}' (score {score})"

    return v, None  # unrecognised — will fail enum check below

# =============================================================================
# 4. HEALTH ELECTION NORMALISATION
# =============================================================================

_ELECTION_ENROLL = {"enroll", "e", "yes", "y", "enrolled", "participating", "active"}
_ELECTION_WAIVE  = {"waive", "w", "no", "n", "waived", "decline", "declined", "opt out",
                    "opt-out", "not participating", "waiving"}


def normalize_election(v: str) -> tuple[str, Optional[str]]:
    """Returns (canonical, fix_note_or_None). Canonical is 'Enroll' or 'Waive'."""
    key = v.strip().lower()
    if key in _ELECTION_ENROLL:
        canonical = "Enroll"
    elif key in _ELECTION_WAIVE:
        canonical = "Waive"
    else:
        return v, None   # unrecognised

    note = f"Health Election normalised: '{v}' → '{canonical}'" if v != canonical else None
    return canonical, note

# =============================================================================
# 5. ZIP CODE LOOKUP
# =============================================================================

@st.cache_data(show_spinner=False)
def build_zip_cache() -> dict[str, dict]:
    if not ZIPCODES_AVAILABLE:
        return {}
    cache: dict[str, dict] = {}
    for entry in _zc.list_all():
        z = entry.get("zip_code", "")
        if z and z not in cache:
            county_raw = (entry.get("county") or "")
            county = county_raw.replace(" County", "").strip()
            cache[z] = {"abbr": entry.get("state", ""), "county": county}
    return cache


def _pad_zip(z: str) -> str:
    """Zero-pad numeric zips shorter than 5 digits (mirrors VBA behaviour)."""
    if z.isdigit() and 0 < len(z) < 5:
        return z.zfill(5)
    return z


def lookup_zip(zip_raw: str, cache: dict) -> Optional[dict]:
    z = _pad_zip(str(zip_raw).strip())
    return cache.get(z)

# =============================================================================
# 6. COLUMN MAPPING  (exact → alias → fuzzy)
# =============================================================================

FUZZY_THRESHOLD = 82   # minimum score to auto-accept a fuzzy match

def map_columns(df: pd.DataFrame) -> tuple[dict[int, int], list[str], list[str]]:
    """
    Returns:
        mapping          – {field_index: df_column_index}
        unrecognised     – list of df column headers that could not be mapped
        ambiguous_notes  – list of human-readable notes about fuzzy matches
    """
    raw_headers = [str(c) for c in df.columns]
    headers_norm = [h.strip().lower() for h in raw_headers]

    mapping: dict[int, int] = {}
    used_df_cols: set[int] = set()
    ambiguous_notes: list[str] = []

    # ── Pass 1: exact alias match ─────────────────────────────────────────────
    for fi, field in enumerate(FIELDS):
        for ci, hn in enumerate(headers_norm):
            if ci in used_df_cols:
                continue
            canonical = HEADER_ALIASES.get(hn)
            if canonical == field["name"]:
                mapping[fi] = ci
                used_df_cols.add(ci)
                break

    # ── Pass 2: fuzzy match for remaining unmapped fields ─────────────────────
    if FUZZY_AVAILABLE:
        unmapped_fields = [fi for fi in range(len(FIELDS)) if fi not in mapping]
        unused_cols = [ci for ci in range(len(raw_headers)) if ci not in used_df_cols]

        for fi in unmapped_fields:
            field_name = FIELDS[fi]["name"]
            candidates = [(ci, headers_norm[ci]) for ci in unused_cols]
            if not candidates:
                break
            best_ci, best_score = None, 0
            for ci, hn in candidates:
                score = fuzz.ratio(field_name.lower(), hn)
                # also try token_sort for multi-word headers
                score2 = fuzz.token_sort_ratio(field_name.lower(), hn)
                s = max(score, score2)
                if s > best_score:
                    best_score, best_ci = s, ci
            if best_ci is not None and best_score >= FUZZY_THRESHOLD:
                mapping[fi] = best_ci
                used_df_cols.add(best_ci)
                ambiguous_notes.append(
                    f"Fuzzy mapped '{raw_headers[best_ci]}' → '{field_name}' "
                    f"(score {best_score})"
                )

    # ── Identify unrecognised columns ─────────────────────────────────────────
    unrecognised = [
        raw_headers[ci] for ci in range(len(raw_headers))
        if ci not in used_df_cols
    ]

    return mapping, unrecognised, ambiguous_notes

# =============================================================================
# 7. CELL-LEVEL VALIDATION
# =============================================================================

def validate_date(v: str) -> dict:
    """Attempt several date format normalisations then validate."""
    original = v
    # yyyy-mm-dd → mm/dd/yyyy
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", v)
    if m:
        v = f"{m.group(2)}/{m.group(3)}/{m.group(1)}"
    # m/d/yyyy or m/dd/yyyy
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", v)
    if m:
        v = f"{m.group(1).zfill(2)}/{m.group(2).zfill(2)}/{m.group(3)}"

    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})$", v)
    if not m:
        return {"ok": False, "msg": f"Invalid date format '{original}' — expected mm/dd/yyyy",
                "fixed_val": original, "fix_note": None}
    mo, d, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if not (1 <= mo <= 12):
        return {"ok": False, "msg": f"Month out of range ({mo}) in date '{original}'",
                "fixed_val": v, "fix_note": None}
    if not (1 <= d <= 31):
        return {"ok": False, "msg": f"Day out of range ({d}) in date '{original}'",
                "fixed_val": v, "fix_note": None}
    if not (1900 <= yr <= 2100):
        return {"ok": False, "msg": f"Year out of range ({yr}) in date '{original}'",
                "fixed_val": v, "fix_note": None}
    fix_note = f"Date reformatted: '{original}' → '{v}'" if v != original else None
    return {"ok": True, "msg": None, "fixed_val": v, "fix_note": fix_note}


def validate_cell(field: dict, raw: str) -> dict:
    """
    Returns dict with keys: ok, msg, fixed_val, fix_note.
    Cleared fields are handled upstream before this is called.
    """
    v = str(raw).strip() if raw is not None else ""

    # Required check
    if field["required"] and v == "":
        return {"ok": False, "msg": f"{field['name']} is required",
                "fixed_val": "", "fix_note": None}
    if v == "":
        return {"ok": True, "msg": None, "fixed_val": "", "fix_note": None}

    ftype = field["type"]

    # ── Date ──────────────────────────────────────────────────────────────────
    if ftype == "date":
        return validate_date(v)

    # ── Relationship (special enum with alias normalisation) ──────────────────
    if ftype == "relationship":
        canonical, fix_note = normalize_relationship(v)
        if canonical not in ("Employee", "Spouse", "Child"):
            return {"ok": False,
                    "msg": f"Invalid Relationship '{v}' — expected Employee, Spouse, or Child",
                    "fixed_val": canonical, "fix_note": fix_note}
        return {"ok": True, "msg": None, "fixed_val": canonical, "fix_note": fix_note}

    # ── Health Election ───────────────────────────────────────────────────────
    if ftype == "election":
        canonical, fix_note = normalize_election(v)
        if canonical not in ("Enroll", "Waive"):
            return {"ok": False,
                    "msg": f"Invalid Health Election '{v}' — expected Enroll or Waive",
                    "fixed_val": v, "fix_note": None}
        return {"ok": True, "msg": None, "fixed_val": canonical, "fix_note": fix_note}

    # ── Zip code ──────────────────────────────────────────────────────────────
    if ftype == "zipcode":
        # Guard: reject state abbreviations in zip fields
        us_states = {
            "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN",
            "IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV",
            "NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN",
            "TX","UT","VT","VA","WA","WV","WI","WY","DC","PR","VI","GU","AS","MP",
        }
        if v.upper() in us_states:
            return {"ok": False,
                    "msg": f"{field['name']}: looks like a state abbreviation ('{v}'), not a zip code",
                    "fixed_val": v, "fix_note": None}
        padded = _pad_zip(v)
        fix_note = f"Zero-padded zip: '{v}' → '{padded}'" if padded != v else None
        if not re.match(r"^\d{5}(-\d{4})?$", padded):
            return {"ok": False,
                    "msg": f"{field['name']} must be a 5-digit numeric zip, got '{v}'",
                    "fixed_val": padded, "fix_note": fix_note}
        return {"ok": True, "msg": None, "fixed_val": padded, "fix_note": fix_note}

    # ── Numeric ───────────────────────────────────────────────────────────────
    if ftype == "numeric":
        clean = v.replace(",", "").replace("$", "").strip()
        if not re.match(r"^-?\d+(\.\d+)?$", clean):
            return {"ok": False,
                    "msg": f"{field['name']} must be numeric, got '{v}'",
                    "fixed_val": v, "fix_note": None}
        fix_note = f"Stripped formatting: '{v}' → '{clean}'" if clean != v else None
        return {"ok": True, "msg": None, "fixed_val": clean, "fix_note": fix_note}

    # ── Generic enum ─────────────────────────────────────────────────────────
    if ftype == "enum":
        valid_vals = field.get("values", [])
        # case-insensitive match
        for fv in valid_vals:
            if fv.lower() == v.lower():
                fix_note = f"Casing corrected: '{v}' → '{fv}'" if fv != v else None
                return {"ok": True, "msg": None, "fixed_val": fv, "fix_note": fix_note}
        return {"ok": False,
                "msg": f"Invalid value '{v}' for {field['name']} — expected: "
                       f"{', '.join(x for x in valid_vals if x)}",
                "fixed_val": v, "fix_note": None}

    # ── Alpha ─────────────────────────────────────────────────────────────────
    return {"ok": True, "msg": None, "fixed_val": v, "fix_note": None}

# =============================================================================
# 8. FILE INGESTION
# =============================================================================

def read_uploaded_file(uploaded) -> pd.DataFrame:
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded, dtype=str, keep_default_na=False)
    elif name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(uploaded, dtype=str, keep_default_na=False)
    else:
        raise ValueError(f"Unsupported file type: {uploaded.name}")
    # Ensure all cells are strings
    return df.fillna("").astype(str)

# =============================================================================
# 9. CORE VALIDATION + REFORMAT ENGINE
# =============================================================================

def run_validation(df: pd.DataFrame, zip_cache: dict) -> dict:
    """
    Full pipeline: column mapping → per-row validation → reformat → output.

    Returns dict with:
        errors, warnings, fixes, clears     – lists of issue dicts
        reformatted_df                      – canonical-order DataFrame
        valid_positions, error_positions    – row index lists
        total_rows                          – int
        mapping_notes                       – list of mapping diagnostic strings
        unrecognised_cols                   – list of unmapped input header names
    """
    errors:   list[dict] = []
    warnings: list[dict] = []
    fixes:    list[dict] = []
    clears:   list[dict] = []

    # ── Column mapping ────────────────────────────────────────────────────────
    col_map, unrecognised_cols, mapping_notes = map_columns(df)

    # Warn about unrecognised columns
    for uc in unrecognised_cols:
        warnings.append({
            "Row": "—", "Col": "—", "Field": uc,
            "Value": "", "Kind": "Warning",
            "Issue": f"Unrecognised column '{uc}' — not mapped to any canonical field",
        })

    out_rows:        list[dict] = []
    valid_positions: list[int]  = []
    error_positions: list[int]  = []
    output_pos = 0

    # ── Per-row processing ────────────────────────────────────────────────────
    for df_ri, row in df.iterrows():
        if all(str(v).strip() == "" for v in row):
            continue   # skip blank rows

        display_row = int(df_ri) + 2   # 1-based with header offset
        new_row: dict[str, str] = {col: "" for col in OUTPUT_COLUMNS}
        row_has_error = False

        for fi, field in enumerate(FIELDS):
            fname = field["name"]
            ci    = col_map.get(fi)
            raw   = str(row.iloc[ci]).strip() if ci is not None else ""

            # ── Cleared fields: always blank, log if value was present ────────
            if fname in CLEARED_FIELDS:
                if raw:
                    clears.append({
                        "Row": display_row, "Col": field["col"],
                        "Field": fname, "Value": raw,
                        "Kind": "Cleared",
                        "Issue": f"Field '{fname}' automatically cleared (value was '{raw}')",
                    })
                new_row[fname] = ""
                continue

            # ── Normal validation ─────────────────────────────────────────────
            res = validate_cell(field, raw)

            if not res["ok"]:
                errors.append({
                    "Row": display_row, "Col": field["col"],
                    "Field": fname, "Value": raw,
                    "Kind": "Error", "Issue": res["msg"],
                })
                row_has_error = True

            if res["fix_note"]:
                fixes.append({
                    "Row": display_row, "Col": field["col"],
                    "Field": fname, "Value": raw,
                    "Kind": "Auto-fix", "Issue": res["fix_note"],
                })

            new_row[fname] = res["fixed_val"] if res["ok"] else raw

        # ── Zip → County lookup ───────────────────────────────────────────────
        if ZIPCODES_AVAILABLE and zip_cache:
            for zip_fname, county_fname in [
                ("Zip Code",                "County"),
                ("Primary Worksite Zip Code","Primary Worksite County"),
            ]:
                zv = new_row.get(zip_fname, "").strip()
                if zv and zv != "":
                    info = lookup_zip(zv, zip_cache)
                    if info:
                        existing_county = new_row.get(county_fname, "").strip()
                        if not existing_county:
                            new_row[county_fname] = info.get("county", "")
                            if info.get("county"):
                                fixes.append({
                                    "Row": display_row, "Col": "—",
                                    "Field": county_fname, "Value": "",
                                    "Kind": "Auto-fix",
                                    "Issue": f"County auto-filled from zip '{zv}': '{info['county']}'",
                                })
                    else:
                        warnings.append({
                            "Row": display_row,
                            "Col": "L" if zip_fname == "Zip Code" else "N",
                            "Field": zip_fname, "Value": zv,
                            "Kind": "Warning",
                            "Issue": f"Zip code '{zv}' not found in lookup table",
                        })

        # Track positions
        if row_has_error:
            error_positions.append(output_pos)
        else:
            valid_positions.append(output_pos)

        out_rows.append(new_row)
        output_pos += 1

    # ── Cross-row: Cost Comparison pair warning ───────────────────────────────
    er_fi = _FIELD_IDX_BY_NAME["Current Health Plan ER Cost"]
    ee_fi = _FIELD_IDX_BY_NAME["Current Health Plan EE Cost"]
    ci_er = col_map.get(er_fi)
    ci_ee = col_map.get(ee_fi)
    for df_ri, row in df.iterrows():
        if all(str(v).strip() == "" for v in row):
            continue
        display_row = int(df_ri) + 2
        er_v = str(row.iloc[ci_er]).strip() if ci_er is not None else ""
        ee_v = str(row.iloc[ci_ee]).strip() if ci_ee is not None else ""
        if bool(er_v) ^ bool(ee_v):
            warnings.append({
                "Row": display_row, "Col": "Y/Z",
                "Field": "Cost Comparison", "Value": "",
                "Kind": "Warning",
                "Issue": "Both ER Cost and EE Cost should be present for Cost Comparison",
            })

    # ── Build canonical output DataFrame ─────────────────────────────────────
    reformatted_df = pd.DataFrame(out_rows, columns=OUTPUT_COLUMNS)
    reformatted_df = reformatted_df.reset_index(drop=True)

    return {
        "errors":            errors,
        "warnings":          warnings,
        "fixes":             fixes,
        "clears":            clears,
        "reformatted_df":    reformatted_df,
        "valid_positions":   valid_positions,
        "error_positions":   error_positions,
        "total_rows":        len(out_rows),
        "mapping_notes":     mapping_notes,
        "unrecognised_cols": unrecognised_cols,
    }

# =============================================================================
# 10. DOWNLOAD HELPERS
# =============================================================================

def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def issues_to_csv_bytes(errors, warnings, fixes, clears) -> bytes:
    cols = ["Row", "Col", "Field", "Value", "Kind", "Issue"]
    combined = errors + warnings + fixes + clears
    combined.sort(key=lambda x: (str(x.get("Row", "0")).zfill(6), x.get("Kind", "")))
    df = pd.DataFrame(combined, columns=cols)
    return df.to_csv(index=False).encode("utf-8-sig")

# =============================================================================
# 11. SAMPLE DATA
# =============================================================================

SAMPLE_CSV = """\
First Name,Last Name,Employee ID,Relationship,DOB,Gender,Email,Address Line 1,Address Line 2,City,State,Zip Code,County,Primary Worksite Zip Code,Primary Worksite County,ICHRA Class,Health Election,Current Health Plan Vendor,Current Health Plan,Current Health Plan Tier,Current Health Plan OOP (single),Current Health Plan OOP (family),Current Health Plan Deductible (single),Current Health Plan Deductible (family),Current Health Plan ER Cost,Current Health Plan EE Cost,Annual Salary,Hourly Rate,Hours Per Week,Notes
John,Smith,EMP001,Employee,01/15/1980,1,john@acme.com,123 Main St,Apt 4,Milwaukee,WI,53201,Milwaukee,53201,Milwaukee,Group A,Enroll,Anthem,Gold PPO,Employee Only,2000,6000,1000,3000,450,180,75000,,,
Jane,Smith,EMP001S,spouse,03/22/1982,2,jane@acme.com,123 Main St,,Milwaukee,WI,53201,,53201,,Group A,enroll,,,,,,,,,,,,
Tim,Smith,EMP001C,Child,6/10/2010,1,,,,,,53201,,53201,,,E,,,,,,,,,,,,
,Johnson,EMP002,Employee,13/45/1975,1,hr@co.com,456 Oak Ave,,Chicago,IL,ABCD,,99999,,,Waive,,,,,,,,,,,,
Mike,,EMP003,Partner,1975-05-10,1,,,,,,53210,,53210,,,enrolled,,,,,,500,,
Sarah,Connor,EMP004,Domestic Partner,08/30/1985,2,,,,,,53202,,53202,,,W,,,,,,,,,,,,
Bob,Jones,EMP005,Employee,05/14/1978,,,,,,,53203,,53203,,,Decline,,,,,,,,,,,,
"""

# =============================================================================
# 12. STREAMLIT UI
# =============================================================================

def _style_issues(df_issues: list[dict]) -> "pd.io.formats.style.Styler":
    cols = ["Row", "Col", "Field", "Value", "Kind", "Issue"]
    df = pd.DataFrame(df_issues, columns=cols).sort_values(
        ["Row", "Kind"], key=lambda s: s.map(
            lambda x: str(x).zfill(6) if str(x).isdigit() else str(x)
        )
    ).reset_index(drop=True)

    def _color(val):
        if val == "Error":    return "color:#993C1D;font-weight:600"
        if val == "Warning":  return "color:#854F0B;font-weight:600"
        if val == "Auto-fix": return "color:#3B6D11;font-weight:600"
        if val == "Cleared":  return "color:#555599;font-weight:600"
        return ""

    styler = df.style
    # pandas ≥2.1 renamed applymap → map
    if hasattr(styler, "map"):
        return styler.map(_color, subset=["Kind"])
    return styler.applymap(_color, subset=["Kind"])   # type: ignore[attr-defined]


def main() -> None:
    st.set_page_config(
        page_title="Census Validator",
        page_icon="📋",
        layout="wide",
    )

    st.markdown("""
    <style>
        .block-container { padding-top: 1.5rem; max-width: 1200px; }
        [data-testid="metric-container"] {
            background:#f7f7f5; border-radius:8px; padding:10px 14px;
        }
        .stTabs [data-baseweb="tab"] { font-size:13px; }
        .banner-ok {
            background:#EAF3DE; border:1px solid #97C459; border-radius:8px;
            padding:12px 16px; color:#3B6D11; font-size:14px; margin-bottom:1rem;
        }
        .info-box {
            background:#EEF2FB; border:1px solid #ADC0EF; border-radius:8px;
            padding:12px 16px; font-size:13px; color:#1a3a7a; margin-bottom:1rem;
        }
        .warn-box {
            background:#FFF8E7; border:1px solid #F0C040; border-radius:8px;
            padding:10px 14px; font-size:13px; color:#7a5a00; margin-bottom:.75rem;
        }
    </style>
    """, unsafe_allow_html=True)

    # ── Header ────────────────────────────────────────────────────────────────
    st.title("📋 Census Validator & Reformatter")
    st.caption(
        "Accepts any column order or naming convention · "
        "Validates all 30 fields · "
        "Auto-clears 7 PII fields · "
        "Exports a system-ready CSV"
    )

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️ Options")

        use_zip_lookup = st.toggle(
            "Auto-fill County from Zip",
            value=ZIPCODES_AVAILABLE,
            disabled=not ZIPCODES_AVAILABLE,
            help="Looks up County from Zip Code and Primary Worksite Zip Code.",
        )
        if not ZIPCODES_AVAILABLE:
            st.caption("Install `zipcodes` to enable:  \n`pip install zipcodes`")

        st.divider()
        st.markdown("**Output column order** (30 fields)")
        for i, f in enumerate(FIELDS, 1):
            cleared = " 🚫" if f["name"] in CLEARED_FIELDS else ""
            req     = " ✳" if f["required"] else ""
            st.caption(f"{i}. {f['name']}{req}{cleared}")
        st.caption("✳ = Required   🚫 = Always cleared")

        st.divider()
        st.markdown("**Column mapping**")
        st.caption(
            f"Exact match → alias table ({len(HEADER_ALIASES)} aliases) "
            f"→ fuzzy match ({'on' if FUZZY_AVAILABLE else 'off — install rapidfuzz'})"
        )

    # ── Input area ────────────────────────────────────────────────────────────
    col_up, col_paste = st.columns([1, 1], gap="large")

    with col_up:
        uploaded = st.file_uploader(
            "Upload census file  (.csv / .xlsx / .xls)",
            type=["csv", "xlsx", "xls"],
        )

    with col_paste:
        default_paste = st.session_state.get("paste_text", "")
        pasted = st.text_area(
            "— or paste CSV data —",
            value=default_paste,
            height=150,
            placeholder="Paste CSV including header row…",
        )

    btn1, btn2, btn3 = st.columns([1, 1, 5])
    with btn1:
        run_btn = st.button("✅  Validate", type="primary", use_container_width=True)
    with btn2:
        if st.button("📄  Load sample", use_container_width=True):
            st.session_state["paste_text"] = SAMPLE_CSV
            st.rerun()

    # ── Run ───────────────────────────────────────────────────────────────────
    if run_btn:
        df_input = None
        try:
            if uploaded:
                df_input = read_uploaded_file(uploaded)
            elif pasted.strip():
                df_input = pd.read_csv(
                    io.StringIO(pasted), dtype=str, keep_default_na=False
                ).fillna("").astype(str)
            else:
                st.warning("Please upload a file or paste CSV data first.")
        except Exception as exc:
            st.error(f"Could not read input: {exc}")

        if df_input is not None:
            with st.spinner("Validating and reformatting…"):
                zc = build_zip_cache() if (use_zip_lookup and ZIPCODES_AVAILABLE) else {}
                results = run_validation(df_input, zc)

            ref_df = results["reformatted_df"]
            results["valid_df"] = ref_df.iloc[results["valid_positions"]].reset_index(drop=True)
            results["error_df"] = ref_df.iloc[results["error_positions"]].reset_index(drop=True)
            st.session_state["results"] = results

    # ── Display ───────────────────────────────────────────────────────────────
    if "results" not in st.session_state:
        return

    res      = st.session_state["results"]
    errors   = res["errors"]
    warnings = res["warnings"]
    fixes    = res["fixes"]
    clears   = res["clears"]
    ref_df   = res["reformatted_df"]
    valid_df = res["valid_df"]
    error_df = res["error_df"]
    total    = res["total_rows"]
    err_rows = len(res["error_positions"])
    valid_rows = total - err_rows
    mapping_notes     = res.get("mapping_notes", [])
    unrecognised_cols = res.get("unrecognised_cols", [])

    st.divider()

    # ── Column mapping diagnostics ────────────────────────────────────────────
    if mapping_notes or unrecognised_cols:
        with st.expander(
            f"🗺️ Column mapping notes  "
            f"({'fuzzy matches: ' + str(len(mapping_notes)) if mapping_notes else ''}"
            f"{'  |  ' if mapping_notes and unrecognised_cols else ''}"
            f"{'unrecognised: ' + str(len(unrecognised_cols)) if unrecognised_cols else ''})",
            expanded=bool(unrecognised_cols),
        ):
            if mapping_notes:
                st.markdown("**Fuzzy-matched headers**")
                for note in mapping_notes:
                    st.caption(f"• {note}")
            if unrecognised_cols:
                st.markdown("**Unrecognised headers** (not mapped to any canonical field)")
                for uc in unrecognised_cols:
                    st.caption(f"• `{uc}`")

    # ── Summary metrics ───────────────────────────────────────────────────────
    m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
    m1.metric("Total rows",      total)
    m2.metric("✅ Valid rows",    valid_rows)
    m3.metric("❌ Error rows",   err_rows)
    m4.metric("Errors",           len(errors))
    m5.metric("Warnings",         len(warnings))
    m6.metric("Auto-fixes",       len(fixes))
    m7.metric("Cleared fields",   len(clears))

    st.divider()

    # ── Download panel ────────────────────────────────────────────────────────
    st.subheader("📥 Download")

    cleared_note = (
        "Employee ID · Gender · Email · Address Lines · City · State "
        "are blanked in all outputs"
    )
    st.caption(
        f"Columns in canonical order · Dates mm/dd/yyyy · "
        f"Relationship & Election normalised · Zips zero-padded · {cleared_note}"
    )

    dc1, dc2, dc3, dc4 = st.columns(4)
    with dc1:
        st.markdown("##### Full reformatted census")
        st.caption(f"All {total} rows · ready for import")
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
        st.caption(f"{valid_rows} rows · no validation errors")
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
        st.caption(f"{err_rows} rows · needs review")
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
        st.caption("Errors · Warnings · Fixes · Clears")
        st.download_button(
            "⬇ Download",
            data=issues_to_csv_bytes(errors, warnings, fixes, clears),
            file_name="census_validation_report.csv",
            mime="text/csv",
            use_container_width=True,
            key="dl_report",
        )

    st.divider()

    # ── Issues table ──────────────────────────────────────────────────────────
    all_issues = errors + warnings + fixes + clears

    if not all_issues:
        st.markdown(
            f'<div class="banner-ok">✅  All {total} rows passed validation with no issues.</div>',
            unsafe_allow_html=True,
        )
    else:
        tab_all, tab_err, tab_warn, tab_fix, tab_clr = st.tabs([
            f"All  ({len(all_issues)})",
            f"❌ Errors  ({len(errors)})",
            f"⚠️ Warnings  ({len(warnings)})",
            f"🔧 Auto-fixes  ({len(fixes)})",
            f"🚫 Cleared  ({len(clears)})",
        ])
        with tab_all:
            st.dataframe(_style_issues(all_issues), use_container_width=True, hide_index=True)
        with tab_err:
            if errors:
                st.dataframe(_style_issues(errors), use_container_width=True, hide_index=True)
            else:
                st.success("No errors found.")
        with tab_warn:
            if warnings:
                st.dataframe(_style_issues(warnings), use_container_width=True, hide_index=True)
            else:
                st.success("No warnings found.")
        with tab_fix:
            if fixes:
                st.dataframe(_style_issues(fixes), use_container_width=True, hide_index=True)
            else:
                st.info("No auto-fixes applied.")
        with tab_clr:
            if clears:
                st.dataframe(_style_issues(clears), use_container_width=True, hide_index=True)
            else:
                st.info("No fields were cleared.")

    # ── Data preview ──────────────────────────────────────────────────────────
    st.divider()
    with st.expander("🔍 Preview reformatted output", expanded=False):
        st.caption("Showing canonical 30-column output. Cleared fields will appear blank.")
        st.dataframe(ref_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
