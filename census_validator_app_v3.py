"""
Census Validator & Reformatter — Streamlit App
==============================================
Install:  pip install streamlit pandas zipcodes openpyxl rapidfuzz
Run:      streamlit run census_validator_app_horizontal.py

Key behaviours
──────────────
• Accepts standard vertical CSA census files and horizontal dependent census files.
• Horizontal mode creates one Employee row, then Spouse row(s), then Child row(s).
• Dependent Health Election is derived from the employee's health/medical coverage tier:
  Employee Only → dependents Waive
  Employee + Spouse → spouse Enroll, children Waive
  Employee + Children → children Enroll, spouse Waive
  Family → spouse and children Enroll
• If no health/medical tier exists but a health/medical plan exists, the app can infer listed
  dependents as enrolled. This is controlled in the sidebar.
• Only health/medical fields are used for horizontal plan/tier/election detection. Dental,
  vision, life, disability, and other ancillary fields are ignored.
• Fixed output column order matching the CSA import template.
• Fuzzy + exact header matching for vertical and horizontal inputs.
• Auto-clears selected PII fields in the final output.
• Zip-to-County lookup when the optional zipcodes package is installed.
• Robust file encoding detection — handles UTF-8, Windows-1252/CP1252 (smart quotes,
  curly apostrophes), Latin-1, and UTF-8-with-BOM exports from Excel/Windows tools.
• ZIP normalization converts ZIP+4 / 9-digit ZIPs to 5-digit ZIPs.
• Invalid, unsupported, and non-quotable ZIPs are flagged as errors.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from typing import Any, Optional

import pandas as pd
import streamlit as st

# ── Optional dependencies ─────────────────────────────────────────────────────
try:
    import zipcodes as _zc
    ZIPCODES_AVAILABLE = True
except ImportError:  # pragma: no cover - optional runtime dependency
    ZIPCODES_AVAILABLE = False

try:
    from rapidfuzz import fuzz, process as rf_process
    FUZZY_AVAILABLE = True
except ImportError:  # pragma: no cover - optional runtime dependency
    FUZZY_AVAILABLE = False


SUPPORTED_ZIP_STATES = {
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA",
    "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
    "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT",
    "VA", "WA", "WV", "WI", "WY", "DC",
}

# ZIPs below are either unsupported by the quoting tool, discontinued, territory-based,
# or otherwise not accepted by the import even when they are 5 digits.
# Keep this small and explicit so it is easy to update if the quoting tool flags more.
KNOWN_INVALID_OR_UNSUPPORTED_ZIPS = {
    "00820",  # U.S. Virgin Islands; not supported for CSA quoting
    "03333",
    "33210",
    "57507",
    "85403",
}


# =============================================================================
# 1. CANONICAL FIELD SCHEMA
# =============================================================================

CLEARED_FIELDS = {
    "Employee ID", "Gender", "Email",
    "Address Line 1", "Address Line 2", "City", "State",
}

FIELDS: list[dict[str, Any]] = [
    {"col": "A",  "name": "First Name",                      "required": True,  "type": "alpha"},
    {"col": "B",  "name": "Last Name",                       "required": True,  "type": "alpha"},
    {"col": "C",  "name": "Employee ID",                     "required": False, "type": "alpha"},
    {"col": "D",  "name": "Relationship",                    "required": True,  "type": "relationship"},
    {"col": "E",  "name": "DOB",                             "required": True,  "type": "date"},
    {"col": "F",  "name": "Gender",                          "required": False, "type": "alpha"},
    {"col": "G",  "name": "Email",                           "required": False, "type": "alpha"},
    {"col": "H",  "name": "Address Line 1",                  "required": False, "type": "alpha"},
    {"col": "I",  "name": "Address Line 2",                  "required": False, "type": "alpha"},
    {"col": "J",  "name": "City",                            "required": False, "type": "alpha"},
    {"col": "K",  "name": "State",                           "required": False, "type": "alpha"},
    {"col": "L",  "name": "Zip Code",                        "required": True,  "type": "zipcode"},
    {"col": "M",  "name": "County",                          "required": False, "type": "alpha"},
    {"col": "N",  "name": "Primary Worksite Zip Code",       "required": True,  "type": "zipcode"},
    {"col": "O",  "name": "Primary Worksite County",         "required": False, "type": "alpha"},
    {"col": "P",  "name": "ICHRA Class",                     "required": False, "type": "alpha"},
    {"col": "Q",  "name": "Health Election",                 "required": True,  "type": "election"},
    {"col": "R",  "name": "Current Health Plan Vendor",      "required": False, "type": "alpha"},
    {"col": "S",  "name": "Current Health Plan",             "required": False, "type": "alpha"},
    {"col": "T",  "name": "Current Health Plan Tier",        "required": False, "type": "tier"},
    {"col": "U",  "name": "Current Health Plan OOP (single)","required": False, "type": "numeric"},
    {"col": "V",  "name": "Current Health Plan OOP (family)","required": False, "type": "numeric"},
    {"col": "W",  "name": "Current Health Plan Deductible (single)", "required": False, "type": "numeric"},
    {"col": "X",  "name": "Current Health Plan Deductible (family)", "required": False, "type": "numeric"},
    {"col": "Y",  "name": "Current Health Plan ER Cost",     "required": False, "type": "numeric"},
    {"col": "Z",  "name": "Current Health Plan EE Cost",     "required": False, "type": "numeric"},
    {"col": "AA", "name": "Annual Salary",                   "required": False, "type": "numeric"},
    {"col": "AB", "name": "Hourly Rate",                     "required": False, "type": "numeric"},
    {"col": "AC", "name": "Hours Per Week",                  "required": False, "type": "numeric"},
    {"col": "AD", "name": "Notes",                           "required": False, "type": "alpha"},
]

OUTPUT_COLUMNS = [f["name"] for f in FIELDS]
_FIELD_IDX_BY_NAME = {f["name"]: i for i, f in enumerate(FIELDS)}
_FIELD_BY_NAME = {f["name"]: f for f in FIELDS}

# =============================================================================
# 2. HEADER NORMALISATION + ALIASES
# =============================================================================

def _norm(s: object) -> str:
    s = str(s or "").strip().lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _compact(s: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


ANCILLARY_TERMS = {
    "dental", "vision", "life", "disability", "std", "ltd", "accident",
    "critical", "illness", "hospital", "indemnity", "cancer", "legal",
    "pet", "commuter", "parking", "transit", "supplemental", "voluntary",
}

HEALTH_TERMS = {
    "health", "medical", "med", "current health", "plan name", "product name", "carrier",
    "coverage tier", "coverage level", "benefit tier", "election tier",
    "employee cost", "employer cost", "ee cost", "er cost",
}

HEADER_ALIASES: dict[str, str] = {
    # First / Last name
    "first name": "First Name", "firstname": "First Name", "first": "First Name",
    "given name": "First Name", "employee first name": "First Name", "ee first name": "First Name",
    "member first name": "First Name", "subscriber first name": "First Name",
    "last name": "Last Name", "lastname": "Last Name", "last": "Last Name",
    "surname": "Last Name", "family name": "Last Name", "employee last name": "Last Name",
    "ee last name": "Last Name", "member last name": "Last Name", "subscriber last name": "Last Name",
    # Employee ID
    "employee id": "Employee ID", "emp id": "Employee ID", "empid": "Employee ID",
    "employee number": "Employee ID", "emp number": "Employee ID", "id": "Employee ID",
    "employee #": "Employee ID", "worker id": "Employee ID", "person number": "Employee ID",
    # Relationship
    "relationship": "Relationship", "relation": "Relationship",
    "dependent type": "Relationship", "member type": "Relationship", "covered person type": "Relationship",
    # DOB
    "dob": "DOB", "date of birth": "DOB", "birth date": "DOB", "birthdate": "DOB",
    "date of birth mm dd yyyy": "DOB", "employee dob": "DOB", "employee date of birth": "DOB",
    # Gender
    "gender": "Gender", "sex": "Gender",
    # Email
    "email": "Email", "email address": "Email", "e mail": "Email",
    "work email": "Email", "personal email": "Email",
    # Address
    "address line 1": "Address Line 1", "address1": "Address Line 1", "street address": "Address Line 1",
    "primary address line 1": "Address Line 1", "home address line 1": "Address Line 1", "address": "Address Line 1",
    "address line 2": "Address Line 2", "address2": "Address Line 2", "apt": "Address Line 2",
    "suite": "Address Line 2", "primary address line 2": "Address Line 2",
    # City / State / Zip / County
    "city": "City", "town": "City", "primary address city": "City", "home city": "City",
    "state": "State", "st": "State", "state code": "State", "primary address state territory": "State",
    "primary address state": "State", "home state": "State",
    "zip code": "Zip Code", "zip": "Zip Code", "zipcode": "Zip Code", "postal code": "Zip Code",
    "primary address zip code": "Zip Code", "primary address postal code": "Zip Code", "home zip": "Zip Code",
    "county": "County", "home county": "County", "primary address county": "County",
    # Worksite
    "primary worksite zip code": "Primary Worksite Zip Code", "worksite zip": "Primary Worksite Zip Code",
    "work zip": "Primary Worksite Zip Code", "worksite zip code": "Primary Worksite Zip Code",
    "site zip": "Primary Worksite Zip Code", "work location zip": "Primary Worksite Zip Code",
    "primary worksite county": "Primary Worksite County", "worksite county": "Primary Worksite County",
    "work county": "Primary Worksite County", "work location county": "Primary Worksite County",
    # Class
    "ichra class": "ICHRA Class", "ichra": "ICHRA Class", "class": "ICHRA Class",
    "benefit class": "ICHRA Class", "worker category": "ICHRA Class", "employee class": "ICHRA Class",
    # Health Election
    "health election": "Health Election", "election": "Health Election", "coverage election": "Health Election",
    "medical election": "Health Election", "benefit election": "Health Election", "medical waive reason": "Health Election",
    # Health plan fields
    "current health vendor": "Current Health Plan Vendor", "current health plan vendor": "Current Health Plan Vendor",
    "health vendor": "Current Health Plan Vendor", "medical vendor": "Current Health Plan Vendor",
    "carrier": "Current Health Plan Vendor", "insurance carrier": "Current Health Plan Vendor",
    "health carrier": "Current Health Plan Vendor", "medical carrier": "Current Health Plan Vendor",
    "current health plan": "Current Health Plan", "health plan": "Current Health Plan",
    "plan name": "Current Health Plan", "medical plan": "Current Health Plan", "medical plan name": "Current Health Plan",
    "health plan name": "Current Health Plan", "current plan": "Current Health Plan",
    "current health plan tier": "Current Health Plan Tier", "health plan tier": "Current Health Plan Tier",
    "medical plan tier": "Current Health Plan Tier", "coverage tier": "Current Health Plan Tier",
    "coverage level": "Current Health Plan Tier", "plan tier": "Current Health Plan Tier", "tier": "Current Health Plan Tier",
    "medical coverage tier": "Current Health Plan Tier", "medical coverage level": "Current Health Plan Tier",
    "current health plan oop single": "Current Health Plan OOP (single)", "health plan oop single": "Current Health Plan OOP (single)",
    "oop single": "Current Health Plan OOP (single)", "out of pocket single": "Current Health Plan OOP (single)",
    "current health plan oop family": "Current Health Plan OOP (family)", "health plan oop family": "Current Health Plan OOP (family)",
    "oop family": "Current Health Plan OOP (family)", "out of pocket family": "Current Health Plan OOP (family)",
    "current health plan deductible single": "Current Health Plan Deductible (single)", "health plan deductible single": "Current Health Plan Deductible (single)",
    "deductible single": "Current Health Plan Deductible (single)",
    "current health plan deductible family": "Current Health Plan Deductible (family)", "health plan deductible family": "Current Health Plan Deductible (family)",
    "deductible family": "Current Health Plan Deductible (family)",
    "current health plan er cost": "Current Health Plan ER Cost", "health plan er cost": "Current Health Plan ER Cost",
    "medical er cost": "Current Health Plan ER Cost", "er cost": "Current Health Plan ER Cost",
    "employer cost": "Current Health Plan ER Cost", "employer contribution": "Current Health Plan ER Cost",
    "current health plan ee cost": "Current Health Plan EE Cost", "health plan ee cost": "Current Health Plan EE Cost",
    "medical ee cost": "Current Health Plan EE Cost", "ee cost": "Current Health Plan EE Cost",
    "employee cost": "Current Health Plan EE Cost", "employee contribution": "Current Health Plan EE Cost",
    # Compensation
    "annual salary": "Annual Salary", "salary": "Annual Salary", "yearly salary": "Annual Salary", "base salary": "Annual Salary",
    "hourly rate": "Hourly Rate", "hourly": "Hourly Rate", "rate": "Hourly Rate", "hourly wage": "Hourly Rate",
    "hours per week": "Hours Per Week", "hours week": "Hours Per Week", "weekly hours": "Hours Per Week", "hrs per week": "Hours Per Week",
    # Notes
    "notes": "Notes", "note": "Notes", "comments": "Notes", "comment": "Notes", "remarks": "Notes",
    **{col.lower(): f["name"] for col, f in zip(
        ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O",
         "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD"], FIELDS
    )},
}

# =============================================================================
# 3. VALUE NORMALISATION
# =============================================================================

_RELATIONSHIP_MAP: dict[str, str] = {}

def _add_rel(canonical: str, *aliases: str) -> None:
    for a in aliases:
        _RELATIONSHIP_MAP[_norm(a)] = canonical

_add_rel("Employee", "employee", "ee", "emp", "subscriber", "self", "primary", "insured", "member", "worker")
_add_rel("Spouse", "spouse", "sp", "husband", "wife", "partner", "domestic partner", "life partner",
         "registered domestic partner", "common law spouse", "common law partner", "DomesticPartner", "domesticpartner", "spse", "sps")
_add_rel("Child", "child", "ch", "dependent", "dep", "daughter", "son", "stepchild", "step child",
         "adopted child", "foster child", "stepson", "stepdaughter", "child of domestic partner", "child child")

_ELECTION_ENROLL = {"enroll", "e", "yes", "y", "enrolled", "participating", "active", "covered", "elect", "elected"}
_ELECTION_WAIVE = {"waive", "w", "no", "n", "waived", "decline", "declined", "opt out", "optout", "not participating", "waiving"}

_TIER_EMP_ONLY = {"employee only", "employee", "ee only", "employee only coverage", "single", "individual", "self only"}
_TIER_EMP_SPOUSE = {"employee spouse", "employee and spouse", "employee + spouse", "ee spouse", "ee + spouse", "ee sp", "employee plus spouse"}
_TIER_EMP_CHILDREN = {"employee children", "employee child", "employee + children", "employee + child", "ee children", "ee child", "ee + children", "ee + child", "employee plus children", "parent child", "parent children"}
_TIER_FAMILY = {"family", "employee family", "employee + family", "ee family", "ee + family", "employee spouse children", "employee + spouse + children", "employee spouse child", "ee spouse children", "ee + spouse + children"}
_TIER_WAIVE = {"waive", "waived", "decline", "declined", "no coverage", "none", "not enrolled", "no election"}


def normalize_relationship(v: object) -> tuple[str, Optional[str]]:
    original = str(v or "").strip()
    key = _norm(original.replace("|", " "))
    if not key:
        return "", None
    if key in _RELATIONSHIP_MAP:
        canonical = _RELATIONSHIP_MAP[key]
        note = f"Relationship normalised: '{original}' → '{canonical}'" if canonical.lower() != original.lower() else None
        return canonical, note
    if FUZZY_AVAILABLE:
        match = rf_process.extractOne(key, _RELATIONSHIP_MAP.keys(), scorer=fuzz.ratio)
        if match:
            best, score, _ = match
            if score >= 82:
                canonical = _RELATIONSHIP_MAP[best]
                return canonical, f"Relationship fuzzy-matched: '{original}' → '{canonical}' (score {score})"
    return original, None


def normalize_election(v: object) -> tuple[str, Optional[str]]:
    original = str(v or "").strip()
    key = _norm(original)
    if not key:
        return "", None
    if key in _ELECTION_ENROLL:
        canonical = "Enroll"
    elif key in _ELECTION_WAIVE:
        canonical = "Waive"
    else:
        return original, None
    note = f"Health Election normalised: '{original}' → '{canonical}'" if original != canonical else None
    return canonical, note


def normalize_tier(v: object) -> tuple[str, Optional[str]]:
    original = str(v or "").strip()
    key = _norm(original.replace("&", " and ").replace("/", " ").replace("-", " "))
    key = key.replace(" plus ", " ").replace(" and ", " ")
    key = re.sub(r"\s+", " ", key).strip()
    if not key:
        return "", None

    def _hit(options: set[str]) -> bool:
        return key in {_norm(o).replace(" and ", " ").replace(" plus ", " ") for o in options}

    # Common carrier/export abbreviations.
    compact = _compact(original)
    abbreviation_map = {
        "eo": "Employee Only",
        "ee": "Employee Only",
        "e": "Employee Only",
        "es": "Employee + Spouse",
        "ec": "Employee + Children",
        "ech": "Employee + Children",
        "eech": "Employee + Children",
        "ef": "Family",
        "fa": "Family",
        "fam": "Family",
    }
    if compact in abbreviation_map:
        canonical = abbreviation_map[compact]
        note = f"Tier normalised: '{original}' -> '{canonical}'" if original != canonical else None
        return canonical, note

    # Keyword fallback catches common employer exports like "EE + Child(ren)".
    if _hit(_TIER_WAIVE) or any(x in compact for x in ["waive", "decline", "nocoverage"]):
        return "Waive", f"Tier normalised: '{original}' → 'Waive'" if original != "Waive" else None
    if _hit(_TIER_FAMILY) or "family" in compact or ("spouse" in compact and ("child" in compact or "children" in compact)):
        return "Family", f"Tier normalised: '{original}' → 'Family'" if original != "Family" else None
    if _hit(_TIER_EMP_SPOUSE) or ("spouse" in compact and "child" not in compact and "children" not in compact):
        return "Employee + Spouse", f"Tier normalised: '{original}' → 'Employee + Spouse'" if original != "Employee + Spouse" else None
    if _hit(_TIER_EMP_CHILDREN) or "child" in compact or "children" in compact:
        return "Employee + Children", f"Tier normalised: '{original}' → 'Employee + Children'" if original != "Employee + Children" else None
    if _hit(_TIER_EMP_ONLY) or compact in {"ee", "eo", "employeeonly", "selfonly", "individual"}:
        return "Employee Only", f"Tier normalised: '{original}' → 'Employee Only'" if original != "Employee Only" else None
    return original, None


def dependent_election_from_tier(employee_election: str, tier: str, relationship: str) -> str:
    if employee_election == "Waive" or tier == "Waive":
        return "Waive"
    if relationship == "Employee":
        return employee_election or "Enroll"
    if tier == "Employee Only":
        return "Waive"
    if tier == "Employee + Spouse":
        return "Enroll" if relationship == "Spouse" else "Waive"
    if tier == "Employee + Children":
        return "Enroll" if relationship == "Child" else "Waive"
    if tier == "Family":
        return "Enroll"
    return employee_election or "Enroll"

# =============================================================================
# 4. ZIP CODE LOOKUP
# =============================================================================

@st.cache_data(show_spinner=False)
def build_zip_cache() -> dict[str, dict[str, str]]:
    if not ZIPCODES_AVAILABLE:
        return {}
    cache: dict[str, dict[str, str]] = {}
    for entry in _zc.list_all():
        z = str(entry.get("zip_code", "")).strip()
        state = str(entry.get("state", "")).strip().upper()
        zip_type = str(entry.get("zip_code_type", "")).strip().upper()
        active = bool(entry.get("active", True))

        if not z or z in KNOWN_INVALID_OR_UNSUPPORTED_ZIPS:
            continue
        if state not in SUPPORTED_ZIP_STATES:
            continue
        if not active:
            continue
        if zip_type and zip_type != "STANDARD":
            continue

        if z not in cache:
            county_raw = str(entry.get("county") or "")
            county = county_raw.replace(" County", "").strip()
            cache[z] = {"abbr": state, "county": county}
    return cache


def _pad_zip(z: object) -> str:
    """Normalize U.S. ZIP values to a 5-digit ZIP code.

    Converts clear U.S. numeric ZIP formats only:
    - 525718303 -> 52571
    - 52571-8303 -> 52571
    - 52571 8303 -> 52571
    - 601 -> 00601
    - 601.0 -> 00601

    Alphanumeric postal codes, such as Canadian postal codes, are returned
    unchanged so validation can flag them instead of stripping digits.
    """
    v = str(z or "").strip().upper()
    if not v:
        return ""

    # Excel sometimes stores numeric ZIPs as 601.0 or 525718303.0.
    if re.fullmatch(r"\d+\.0", v):
        v = v[:-2]

    # Do not strip digits from non-U.S. alphanumeric postal codes.
    # Example: T8A2A4 should remain T8A2A4 and be flagged.
    if re.search(r"[A-Z]", v):
        return v

    # Accept only numeric ZIP, ZIP+4, or numeric ZIP+4 with spaces.
    compact = re.sub(r"[\s-]", "", v)
    if re.fullmatch(r"\d{9}", compact):
        return compact[:5]
    if re.fullmatch(r"\d{5}", compact):
        return compact
    if re.fullmatch(r"\d{1,4}", compact):
        return compact.zfill(5)

    return v


def lookup_zip(zip_raw: object, cache: dict[str, dict[str, str]]) -> Optional[dict[str, str]]:
    return cache.get(_pad_zip(zip_raw)[:5])

# =============================================================================
# 5. FILE INGESTION
# =============================================================================

# Encodings to try, in order. utf-8-sig handles BOM-prefixed UTF-8 (common from
# Excel "CSV UTF-8" exports). cp1252 / latin-1 handle the classic Windows export
# case where curly quotes, em-dashes, etc. show up as bytes like 0x92, which are
# not valid UTF-8 and otherwise throw UnicodeDecodeError.
CANDIDATE_ENCODINGS = ["utf-8-sig", "utf-8", "cp1252", "latin-1"]


def _read_csv_text(text: str) -> pd.DataFrame:
    """Read comma, tab, pipe, or semicolon delimited census text safely."""
    lines = [line for line in text.splitlines() if line.strip()]
    header = lines[0] if lines else ""
    delimiter_counts = {
        ",": header.count(","),
        "\t": header.count("\t"),
        "|": header.count("|"),
        ";": header.count(";"),
    }
    delimiter = max(delimiter_counts, key=delimiter_counts.get)
    if delimiter_counts[delimiter] == 0:
        delimiter = ","

    try:
        df = pd.read_csv(io.StringIO(text), dtype=str, keep_default_na=False, sep=delimiter)
    except Exception:
        df = pd.read_csv(io.StringIO(text), dtype=str, keep_default_na=False, sep=None, engine="python")

    # If a tab-delimited file was accidentally parsed as one column, retry as TSV.
    if df.shape[1] <= 1 and "\t" in header:
        df = pd.read_csv(io.StringIO(text), dtype=str, keep_default_na=False, sep="\t")

    df = df.fillna("").astype(str).reset_index(drop=True)
    df = df.reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _read_csv_with_fallback_encoding(file_obj) -> tuple[pd.DataFrame, str]:
    """Try a sequence of encodings until one parses cleanly.

    Accepts a file-like object (e.g. an uploaded file or BytesIO) and returns
    the parsed DataFrame along with the encoding that worked.
    """
    raw_bytes = file_obj.read()
    if isinstance(raw_bytes, str):
        # Already decoded (e.g. text pasted into a text area) — just parse it.
        df = _read_csv_text(raw_bytes)
        return df, "text"

    last_exc: Optional[Exception] = None
    for enc in CANDIDATE_ENCODINGS:
        try:
            text = raw_bytes.decode(enc)
            df = _read_csv_text(text)
            return df, enc
        except (UnicodeDecodeError, UnicodeError) as exc:
            last_exc = exc
            continue
        except Exception as exc:
            # A decode succeeded but pandas couldn't parse it — try next encoding
            # only if it's likely a decode artifact; otherwise re-raise immediately
            # on the last attempt.
            last_exc = exc
            continue

    # Last resort: decode with errors="replace" using latin-1 so we never crash,
    # though any replaced characters will show up as the U+FFFD marker.
    try:
        text = raw_bytes.decode("latin-1", errors="replace")
        df = _read_csv_text(text)
        return df, "latin-1 (with replacement)"
    except Exception:
        raise last_exc if last_exc else RuntimeError("Unable to decode file")


def read_uploaded_file(uploaded) -> pd.DataFrame:
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        df, _enc_used = _read_csv_with_fallback_encoding(uploaded)
    elif name.endswith((".xlsx", ".xls")):
        # Excel files are binary and encoding-agnostic at the pandas level;
        # openpyxl/xlrd handle internal text encoding themselves.
        df = pd.read_excel(uploaded, dtype=str, keep_default_na=False)
        df = df.fillna("").astype(str).reset_index(drop=True)
    else:
        raise ValueError(f"Unsupported file type: {uploaded.name}")
    df.columns = [str(c).strip() for c in df.columns]
    return df

# =============================================================================
# 6. COLUMN MAPPING FOR STANDARD VERTICAL FILES
# =============================================================================

FUZZY_THRESHOLD = 82


def map_columns(df: pd.DataFrame) -> tuple[dict[int, int], list[str], list[str]]:
    raw_headers = [str(c) for c in df.columns]
    headers_norm = [_norm(h) for h in raw_headers]
    mapping: dict[int, int] = {}
    used_df_cols: set[int] = set()
    mapping_notes: list[str] = []

    for fi, field_def in enumerate(FIELDS):
        for ci, hn in enumerate(headers_norm):
            if ci in used_df_cols:
                continue
            canonical = HEADER_ALIASES.get(hn)
            if canonical == field_def["name"]:
                mapping[fi] = ci
                used_df_cols.add(ci)
                break

    if FUZZY_AVAILABLE:
        for fi, field_def in enumerate(FIELDS):
            if fi in mapping:
                continue
            candidates = [(ci, headers_norm[ci]) for ci in range(len(raw_headers)) if ci not in used_df_cols]
            if not candidates:
                break
            best_ci, best_score = None, 0
            for ci, hn in candidates:
                score = max(
                    fuzz.ratio(field_def["name"].lower(), hn),
                    fuzz.token_sort_ratio(field_def["name"].lower(), hn),
                )
                if score > best_score:
                    best_score, best_ci = score, ci
            if best_ci is not None and best_score >= FUZZY_THRESHOLD:
                mapping[fi] = best_ci
                used_df_cols.add(best_ci)
                mapping_notes.append(f"Fuzzy mapped '{raw_headers[best_ci]}' → '{field_def['name']}' (score {best_score})")

    unrecognised = [raw_headers[ci] for ci in range(len(raw_headers)) if ci not in used_df_cols]
    return mapping, unrecognised, mapping_notes

# =============================================================================
# 7. HORIZONTAL CENSUS DETECTION + NORMALISATION
# =============================================================================

@dataclass
class HorizontalOptions:
    auto_placeholder_names: bool = True
    infer_dependents_enrolled_when_tier_missing: bool = True
    inherit_worksite_zip_to_dependents: bool = True
    exclude_interns_and_contractors: bool = True

@dataclass
class HorizontalColumnMap:
    employee_first: Optional[str] = None
    employee_last: Optional[str] = None
    employee_id: Optional[str] = None
    employee_dob: Optional[str] = None
    employee_gender: Optional[str] = None
    employee_email: Optional[str] = None
    address1: Optional[str] = None
    address2: Optional[str] = None
    city: Optional[str] = None
    state: Optional[str] = None
    zip_code: Optional[str] = None
    county: Optional[str] = None
    worksite_zip: Optional[str] = None
    worksite_county: Optional[str] = None
    ichra_class: Optional[str] = None
    worker_category: Optional[str] = None
    position_status: Optional[str] = None
    health_election: Optional[str] = None
    health_waive_reason: Optional[str] = None
    plan_vendor: Optional[str] = None
    plan_name: Optional[str] = None
    plan_tier: Optional[str] = None
    oop_single: Optional[str] = None
    oop_family: Optional[str] = None
    ded_single: Optional[str] = None
    ded_family: Optional[str] = None
    er_cost: Optional[str] = None
    ee_cost: Optional[str] = None
    annual_salary: Optional[str] = None
    hourly_rate: Optional[str] = None
    hours_per_week: Optional[str] = None
    notes: Optional[str] = None

@dataclass
class DepSlot:
    slot: str
    number: int = 9999
    relationship_col: Optional[str] = None
    dob_col: Optional[str] = None
    first_col: Optional[str] = None
    last_col: Optional[str] = None
    default_relationship: str = ""


def _is_ancillary_header(header: str) -> bool:
    h = _norm(header)
    # Do not treat HSA in a medical plan name as ancillary. Exclude only clear ancillary product headers.
    return any(term in h.split() or term in h for term in ANCILLARY_TERMS)


def _looks_health_header(header: str) -> bool:
    h = _norm(header)
    if _is_ancillary_header(h):
        return False
    return any(term in h for term in HEALTH_TERMS) or "medical" in h or "med " in f"{h} "


def _pick_col(
    df: pd.DataFrame,
    aliases: list[str],
    *,
    health_only: bool = False,
    allow_contains: bool = True,
    allow_fuzzy: bool = True,
) -> Optional[str]:
    headers = list(df.columns)
    norm_to_header: dict[str, str] = {_norm(h): h for h in headers}
    for a in aliases:
        n = _norm(a)
        if n in norm_to_header:
            h = norm_to_header[n]
            if not health_only or _looks_health_header(h):
                return h
    # Contains pass, useful for employer exports with long labels.
    # Can be disabled for fields where broad aliases like "Employee #" normalize
    # to "employee" and would otherwise match unrelated columns such as
    # "Employee Cost".
    alias_norm = [_norm(a) for a in aliases]
    if allow_contains:
        for h in headers:
            hn = _norm(h)
            if health_only and not _looks_health_header(h):
                continue
            if any(a and len(a) >= 4 and (hn == a or a in hn or hn in a) for a in alias_norm):
                return h
    if allow_fuzzy and FUZZY_AVAILABLE:
        for a in alias_norm:
            match = rf_process.extractOne(a, [_norm(h) for h in headers], scorer=fuzz.token_sort_ratio)
            if match:
                best_norm, score, _ = match
                if score >= 95:
                    h = norm_to_header.get(best_norm)
                    if h and (not health_only or _looks_health_header(h)):
                        return h
    return None


def build_horizontal_column_map(df: pd.DataFrame) -> HorizontalColumnMap:
    cm = HorizontalColumnMap()
    cm.employee_first = _pick_col(df, ["First Name", "Employee First Name", "EE First Name", "First", "Given Name"])
    cm.employee_last = _pick_col(df, ["Last Name", "Employee Last Name", "EE Last Name", "Last", "Surname", "Family Name"])
    cm.employee_id = _pick_col(df, ["Employee ID", "Employee Number", "Employee #", "Worker ID", "Person Number"], allow_contains=False, allow_fuzzy=False)
    cm.employee_dob = _pick_col(df, ["BIRTH DATE", "DOB", "Employee DOB", "Employee Date of Birth", "Date of Birth", "Birthdate"])
    cm.employee_gender = _pick_col(df, ["Gender", "Sex"])
    cm.employee_email = _pick_col(df, ["Email", "Email Address", "Work Email", "Personal Email"])
    cm.address1 = _pick_col(df, ["PRIMARY ADDRESS LINE 1", "Address Line 1", "Street Address", "Home Address Line 1"])
    cm.address2 = _pick_col(df, ["PRIMARY ADDRESS LINE 2", "Address Line 2", "Home Address Line 2"])
    cm.city = _pick_col(df, ["PRIMARY ADDRESS - CITY", "Primary Address City", "City", "Home City"])
    cm.state = _pick_col(df, ["PRIMARY ADDRESS - STATE / TERRITORY", "Primary Address State", "State", "Home State"])
    cm.zip_code = _pick_col(df, ["PRIMARY ADDRESS - ZIP CODE", "Primary Address Zip Code", "Zip Code", "Zip", "Home Zip"])
    cm.county = _pick_col(df, ["County", "Home County", "Primary Address County"])
    cm.worksite_zip = _pick_col(df, ["WORKSITE ZIP CODE", "Primary Worksite Zip Code", "Worksite Zip", "Work Zip", "Work Location Zip"])
    cm.worksite_county = _pick_col(df, ["Primary Worksite County", "Worksite County", "Work County", "Work Location County"])
    cm.ichra_class = _pick_col(df, ["ICHRA Class", "Benefit Class", "Employee Class", "Class"])
    cm.worker_category = _pick_col(df, ["WORKER CATEGORY", "Worker Category", "Employee Type", "Employment Type"])
    cm.position_status = _pick_col(df, ["POSITION STATUS", "Employment Status", "Status"])
    cm.health_election = _pick_col(df, ["Health Election", "Medical Election", "Medical Coverage Election", "Coverage Election"], health_only=True)
    cm.health_waive_reason = _pick_col(df, ["MEDICAL WAIVE REASON", "Health Waive Reason", "Waive Reason"], health_only=True)
    cm.plan_vendor = _pick_col(df, ["Current Health Plan Vendor", "Health Carrier", "Medical Carrier", "Carrier", "Medical Vendor"], health_only=True)
    cm.plan_name = (
        _pick_col(df, ["Product Name", "Current Plan Name", "Current Product Name", "Medical Product Name", "Health Product Name"], allow_fuzzy=False)
        or _pick_col(df, ["MEDICAL PLAN NAME", "Medical Plan Name", "Health Plan Name", "Current Health Plan", "Plan Name"], health_only=True)
    )
    cm.plan_tier = _pick_col(df, ["Current Health Plan Tier", "Medical Plan Tier", "Medical Coverage Tier", "Coverage Tier", "Coverage Level", "Benefit Tier", "Election Tier"], health_only=True)
    cm.oop_single = _pick_col(df, ["Current Health Plan OOP (single)", "Medical OOP Single", "OOP Single"], health_only=True)
    cm.oop_family = _pick_col(df, ["Current Health Plan OOP (family)", "Medical OOP Family", "OOP Family"], health_only=True)
    cm.ded_single = _pick_col(df, ["Current Health Plan Deductible (single)", "Medical Deductible Single", "Deductible Single"], health_only=True)
    cm.ded_family = _pick_col(df, ["Current Health Plan Deductible (family)", "Medical Deductible Family", "Deductible Family"], health_only=True)
    cm.er_cost = _pick_col(df, ["Current Health Plan ER Cost", "Medical ER Cost", "Employer Cost", "Pending Employer Cost", "Employer Contribution"], allow_fuzzy=False)
    cm.ee_cost = _pick_col(df, ["Current Health Plan EE Cost", "Medical EE Cost", "Employee Cost", "Pending Employee Cost", "Employee Contribution"], allow_fuzzy=False)
    cm.annual_salary = _pick_col(df, ["ANNUAL SALARY", "Annual Salary", "Salary", "Base Salary"])
    cm.hourly_rate = _pick_col(df, ["Hourly Rate", "Hourly Wage"])
    cm.hours_per_week = _pick_col(df, ["Hours Per Week", "Weekly Hours"])
    cm.notes = _pick_col(df, ["Notes", "Comments", "Remarks"])
    return cm


def find_dependent_slots(df: pd.DataFrame) -> tuple[list[DepSlot], bool]:
    """Return dependent slots and whether this is row-wise dependent data."""
    headers = list(df.columns)
    norm_headers = {_norm(h): h for h in headers}

    # Row-wise export: one Dependent Relationship and one Dependent DOB column, repeated employee rows.
    row_rel = _pick_col(df, ["Dependent Relationship", "Dependent Relation", "Dependent Type", "Dep Relationship", "Dep Type"])
    row_dob = _pick_col(df, ["Dependent DOB", "Dependent Date of Birth", "Dependent Birth Date", "Dep DOB", "Dep Birth Date"])
    row_num = _pick_col(df, ["Dependent Number", "Dependent #", "Dep Number", "Dep #"])
    row_first = _pick_col(df, ["Dependent First Name", "Dep First Name", "Dependent First"])
    row_last = _pick_col(df, ["Dependent Last Name", "Dep Last Name", "Dependent Last"])
    if row_rel or row_dob:
        return [DepSlot("rowwise", relationship_col=row_rel, dob_col=row_dob, first_col=row_first, last_col=row_last, number=1)], True

    slots: dict[str, DepSlot] = {}

    def get_slot(key: str, number: int, default_relationship: str = "") -> DepSlot:
        if key not in slots:
            slots[key] = DepSlot(slot=key, number=number, default_relationship=default_relationship)
        elif default_relationship and not slots[key].default_relationship:
            slots[key].default_relationship = default_relationship
        return slots[key]

    for h in headers:
        hn = _norm(h)
        hc = _compact(h)
        if _is_ancillary_header(hn):
            continue

        # Spouse DOB / Spouse First / Spouse Last
        if "spouse" in hn or "domestic partner" in hn:
            m = re.search(r"(\d+)", hn)
            number = int(m.group(1)) if m else 1
            slot = get_slot(f"spouse_{number}", 100 + number, "Spouse")
            if any(x in hn for x in ["dob", "date of birth", "birth date", "birthdate"]):
                slot.dob_col = h
            elif "first" in hn and "name" in hn:
                slot.first_col = h
            elif "last" in hn and "name" in hn:
                slot.last_col = h
            elif "relationship" in hn or "relation" in hn or "type" in hn:
                slot.relationship_col = h
            continue

        # Child 1 DOB / Child 2 First / Child 3 Last
        if "child" in hn or "dependent" in hn or re.search(r"\bdep\b", hn):
            m = re.search(r"(\d+)", hn)
            number = int(m.group(1)) if m else 999
            default_rel = "Child" if "child" in hn else ""
            slot_key = f"dep_{number}"
            slot = get_slot(slot_key, 200 + number, default_rel)
            if any(x in hn for x in ["dob", "date of birth", "birth date", "birthdate"]):
                slot.dob_col = h
            elif "first" in hn and "name" in hn:
                slot.first_col = h
            elif "last" in hn and "name" in hn:
                slot.last_col = h
            elif "relationship" in hn or "relation" in hn or "type" in hn:
                slot.relationship_col = h
            continue

    complete_or_partial = [s for s in slots.values() if s.dob_col or s.relationship_col or s.first_col or s.last_col]
    complete_or_partial.sort(key=lambda s: (0 if s.default_relationship == "Spouse" else 1, s.number, s.slot))
    return complete_or_partial, False


def detect_horizontal_census(df: pd.DataFrame) -> tuple[bool, str]:
    headers_norm = [_norm(c) for c in df.columns]
    has_canonical_relationship = any(h == "relationship" for h in headers_norm)
    has_canonical_dob = any(h in {"dob", "date of birth", "birth date"} for h in headers_norm)
    has_dep = any("dependent" in h or re.search(r"\bdep\b", h) or "spouse" in h or "child" in h for h in headers_norm)
    has_dep_health = any(("dependent" in h or "spouse" in h or "child" in h) and ("dob" in h or "birth" in h or "relationship" in h) for h in headers_norm)

    if has_dep_health:
        return True, "Dependent columns detected"
    if has_dep and not (has_canonical_relationship and has_canonical_dob):
        return True, "Horizontal dependent-style headers detected"
    return False, "Standard vertical format detected"


def _cell(row: pd.Series, col: Optional[str]) -> str:
    if not col or col not in row.index:
        return ""
    return str(row[col] or "").strip()


def _first_nonblank(rows: pd.DataFrame, col: Optional[str]) -> str:
    if not col or col not in rows.columns:
        return ""
    for v in rows[col].tolist():
        if str(v).strip():
            return str(v).strip()
    return ""


def _derive_ichra_class(row: pd.Series, cm: HorizontalColumnMap) -> str:
    direct = _cell(row, cm.ichra_class)
    if direct:
        return direct
    worker_category = _cell(row, cm.worker_category)
    state = _cell(row, cm.state).upper()
    if not worker_category:
        return "NY Full Time Benefits Eligible" if state in {"NY", "VT"} else "Full Time Benefits Eligible"
    wc = _norm(worker_category)
    if "intern" in wc or "contractor" in wc:
        return ""
    if "part" in wc:
        return "NY Part Time Benefits Eligible" if state in {"NY", "VT"} else "Part Time Benefits Eligible"
    return "NY Full Time Benefits Eligible" if state in {"NY", "VT"} else "Full Time Benefits Eligible"


def _should_exclude_employee(row: pd.Series, cm: HorizontalColumnMap, options: HorizontalOptions) -> bool:
    if not options.exclude_interns_and_contractors:
        return False
    wc = _norm(_cell(row, cm.worker_category))
    return "intern" in wc or "contractor" in wc


def _employee_group_key(row: pd.Series, cm: HorizontalColumnMap, row_index: int) -> tuple[Any, ...]:
    emp_id = _cell(row, cm.employee_id)
    if emp_id:
        return ("employee_id", emp_id)
    first = _cell(row, cm.employee_first)
    last = _cell(row, cm.employee_last)
    dob = _cell(row, cm.employee_dob)
    if first or last:
        return ("name_dob", first, last, dob, _cell(row, cm.zip_code))
    # For anonymous horizontal files, repeat rows are grouped by stable employee-level fields.
    stable_values = [
        _cell(row, cm.employee_dob), _cell(row, cm.address1), _cell(row, cm.city), _cell(row, cm.state),
        _cell(row, cm.zip_code), _cell(row, cm.worksite_zip), _cell(row, cm.worker_category),
        _cell(row, cm.position_status), _cell(row, cm.annual_salary), _cell(row, cm.plan_name),
        _cell(row, cm.plan_tier), _cell(row, cm.health_waive_reason),
    ]
    if any(stable_values):
        return ("anonymous", *stable_values)
    return ("row", row_index)


def _sort_dependents(deps: list[dict[str, str]]) -> list[dict[str, str]]:
    def key(d: dict[str, str]) -> tuple[int, int]:
        rel = d.get("Relationship", "")
        rel_order = 0 if rel == "Spouse" else 1 if rel == "Child" else 2
        raw_num = str(d.get("_dep_number", "9999"))
        m = re.search(r"\d+", raw_num)
        num = int(m.group(0)) if m else 9999
        return rel_order, num
    return sorted(deps, key=key)


def _infer_tier_from_dependents(deps: list[dict[str, str]]) -> str:
    has_spouse = any(d.get("Relationship") == "Spouse" for d in deps)
    has_child = any(d.get("Relationship") == "Child" for d in deps)
    if has_spouse and has_child:
        return "Family"
    if has_spouse:
        return "Employee + Spouse"
    if has_child:
        return "Employee + Children"
    return "Employee Only"


def convert_horizontal_to_vertical(df: pd.DataFrame, options: HorizontalOptions) -> dict[str, Any]:
    df = df.fillna("").astype(str).reset_index(drop=True)
    cm = build_horizontal_column_map(df)
    dep_slots, rowwise = find_dependent_slots(df)
    notes: list[dict[str, str]] = []
    output_rows: list[dict[str, str]] = []
    placeholder_counter = 1

    if not dep_slots:
        notes.append({"Kind": "Warning", "Issue": "Horizontal mode was selected, but no dependent DOB/relationship columns were detected."})

    groups: list[tuple[tuple[Any, ...], list[int]]] = []
    seen: dict[tuple[Any, ...], int] = {}
    for idx, row in df.iterrows():
        if all(str(v).strip() == "" for v in row):
            continue
        if _should_exclude_employee(row, cm, options):
            notes.append({"Kind": "Warning", "Issue": f"Input row {idx + 2} skipped because Worker Category appears to be intern/contractor."})
            continue
        key = _employee_group_key(row, cm, int(idx)) if rowwise else ("single_row", int(idx))
        if key not in seen:
            seen[key] = len(groups)
            groups.append((key, []))
        groups[seen[key]][1].append(int(idx))

    def base_row() -> dict[str, str]:
        return {c: "" for c in OUTPUT_COLUMNS}

    def apply_placeholder(row_out: dict[str, str], relationship: str) -> None:
        nonlocal placeholder_counter
        if row_out["First Name"] or row_out["Last Name"]:
            return
        if options.auto_placeholder_names:
            row_out["First Name"] = relationship or "Member"
            row_out["Last Name"] = str(placeholder_counter)
            placeholder_counter += 1

    for _, idxs in groups:
        rows = df.loc[idxs]
        first_row = rows.iloc[0]

        dep_records: list[dict[str, str]] = []
        if rowwise:
            slot = dep_slots[0]
            for idx in idxs:
                src = df.loc[idx]
                rel_raw = _cell(src, slot.relationship_col)
                dob = _cell(src, slot.dob_col)
                first = _cell(src, slot.first_col)
                last = _cell(src, slot.last_col)
                dep_num = _cell(src, _pick_col(df, ["Dependent Number", "Dependent #", "Dep Number", "Dep #"])) or str(len(dep_records) + 1)
                if not any([rel_raw, dob, first, last]):
                    continue

                rel_parts = [x.strip() for x in re.split(r"\s*[|;/]\s*", rel_raw) if x.strip()] or [rel_raw]
                dob_parts = [x.strip() for x in re.split(r"\s*[|;]\s*", dob) if x.strip()] or [dob]
                first_parts = [x.strip() for x in re.split(r"\s*[|;]\s*", first) if x.strip()] or [first]
                last_parts = [x.strip() for x in re.split(r"\s*[|;]\s*", last) if x.strip()] or [last]
                part_count = max(len(rel_parts), len(dob_parts), len(first_parts), len(last_parts), 1)
                if part_count > 1:
                    notes.append({"Kind": "Auto-fix", "Issue": f"Input row {idx + 2}: split one dependent cell into {part_count} dependent rows."})

                for part_i in range(part_count):
                    rel_piece = rel_parts[part_i] if part_i < len(rel_parts) else (rel_parts[-1] if rel_parts else "")
                    dob_piece = dob_parts[part_i] if part_i < len(dob_parts) else ""
                    first_piece = first_parts[part_i] if part_i < len(first_parts) else ""
                    last_piece = last_parts[part_i] if part_i < len(last_parts) else ""
                    rel, note = normalize_relationship(rel_piece)
                    if note:
                        notes.append({"Kind": "Auto-fix", "Issue": f"Input row {idx + 2}: {note}"})
                    if rel == "Employee":
                        # Row-wise horizontal files often include a self/member row in the
                        # dependent columns. The converter already creates the employee row,
                        # so keeping this record would duplicate the employee and break the
                        # intended Employee, Spouse, Child ordering.
                        continue
                    dep_records.append({"Relationship": rel, "DOB": dob_piece, "First Name": first_piece, "Last Name": last_piece, "_dep_number": f"{dep_num}.{part_i + 1}" if part_count > 1 else dep_num, "_source_row": str(idx + 2)})
        else:
            for slot in dep_slots:
                explicit_rel_raw = _cell(first_row, slot.relationship_col)
                dob = _cell(first_row, slot.dob_col)
                first = _cell(first_row, slot.first_col)
                last = _cell(first_row, slot.last_col)
                if not any([explicit_rel_raw, dob, first, last]):
                    continue
                rel_raw = explicit_rel_raw or slot.default_relationship
                rel, note = normalize_relationship(rel_raw)
                if note:
                    notes.append({"Kind": "Auto-fix", "Issue": f"Input row {idxs[0] + 2}: {note}"})
                dep_records.append({"Relationship": rel, "DOB": dob, "First Name": first, "Last Name": last, "_dep_number": str(slot.number), "_source_row": str(idxs[0] + 2)})

        dep_records = _sort_dependents(dep_records)

        # Employee-level tier/election decision.
        raw_election = _first_nonblank(rows, cm.health_election)
        raw_waive_reason = _first_nonblank(rows, cm.health_waive_reason)
        raw_plan_name = _first_nonblank(rows, cm.plan_name)
        raw_tier = _first_nonblank(rows, cm.plan_tier)
        tier, tier_note = normalize_tier(raw_tier)
        if tier_note:
            notes.append({"Kind": "Auto-fix", "Issue": tier_note})

        employee_election, election_note = normalize_election(raw_election)
        if election_note:
            notes.append({"Kind": "Auto-fix", "Issue": election_note})
        if not employee_election:
            if tier == "Waive" or raw_waive_reason:
                employee_election = "Waive"
            elif raw_plan_name or (tier and tier != "Waive"):
                employee_election = "Enroll"
            else:
                employee_election = "Waive"
                notes.append({"Kind": "Warning", "Issue": f"Input row {idxs[0] + 2}: no health election, product name, or coverage tier found; defaulted employee to Waive."})

        if not tier:
            if employee_election == "Waive":
                tier = "Waive"
            elif raw_plan_name and options.infer_dependents_enrolled_when_tier_missing:
                tier = _infer_tier_from_dependents(dep_records)
                notes.append({"Kind": "Warning", "Issue": f"Input row {idxs[0] + 2}: medical plan exists but no health tier was found; inferred '{tier}' from listed dependents."})
            elif raw_plan_name:
                tier = "Employee Only"
                notes.append({"Kind": "Warning", "Issue": f"Input row {idxs[0] + 2}: medical plan exists but no health tier was found; treated as Employee Only."})
            else:
                tier = "Waive"

        employee = base_row()
        employee["First Name"] = _first_nonblank(rows, cm.employee_first)
        employee["Last Name"] = _first_nonblank(rows, cm.employee_last)
        apply_placeholder(employee, "Employee")
        employee["Employee ID"] = _first_nonblank(rows, cm.employee_id)
        employee["Relationship"] = "Employee"
        employee["DOB"] = _first_nonblank(rows, cm.employee_dob)
        employee["Gender"] = _first_nonblank(rows, cm.employee_gender)
        employee["Email"] = _first_nonblank(rows, cm.employee_email)
        employee["Address Line 1"] = _first_nonblank(rows, cm.address1)
        employee["Address Line 2"] = _first_nonblank(rows, cm.address2)
        employee["City"] = _first_nonblank(rows, cm.city)
        employee["State"] = _first_nonblank(rows, cm.state)
        employee["Zip Code"] = _first_nonblank(rows, cm.zip_code)
        employee["County"] = _first_nonblank(rows, cm.county)
        employee["Primary Worksite Zip Code"] = _first_nonblank(rows, cm.worksite_zip)
        employee["Primary Worksite County"] = _first_nonblank(rows, cm.worksite_county)
        employee["ICHRA Class"] = _derive_ichra_class(first_row, cm)
        employee["Health Election"] = employee_election
        employee["Current Health Plan Vendor"] = _first_nonblank(rows, cm.plan_vendor)
        employee["Current Health Plan"] = raw_plan_name
        employee["Current Health Plan Tier"] = tier if tier != "Waive" else ""
        employee["Current Health Plan OOP (single)"] = _first_nonblank(rows, cm.oop_single)
        employee["Current Health Plan OOP (family)"] = _first_nonblank(rows, cm.oop_family)
        employee["Current Health Plan Deductible (single)"] = _first_nonblank(rows, cm.ded_single)
        employee["Current Health Plan Deductible (family)"] = _first_nonblank(rows, cm.ded_family)
        employee["Current Health Plan ER Cost"] = _first_nonblank(rows, cm.er_cost)
        employee["Current Health Plan EE Cost"] = _first_nonblank(rows, cm.ee_cost)
        employee["Annual Salary"] = _first_nonblank(rows, cm.annual_salary)
        employee["Hourly Rate"] = _first_nonblank(rows, cm.hourly_rate)
        employee["Hours Per Week"] = _first_nonblank(rows, cm.hours_per_week)
        employee["Notes"] = _first_nonblank(rows, cm.notes)
        output_rows.append(employee)

        for dep in dep_records:
            dep_row = base_row()
            dep_row["First Name"] = dep.get("First Name", "")
            dep_row["Last Name"] = dep.get("Last Name", "")
            dep_row["Relationship"] = dep.get("Relationship", "")
            apply_placeholder(dep_row, dep_row["Relationship"] or "Dependent")
            dep_row["DOB"] = dep.get("DOB", "")
            dep_row["Zip Code"] = employee["Zip Code"]
            dep_row["County"] = employee["County"]
            if options.inherit_worksite_zip_to_dependents:
                dep_row["Primary Worksite Zip Code"] = employee["Primary Worksite Zip Code"]
                dep_row["Primary Worksite County"] = employee["Primary Worksite County"]
            dep_row["ICHRA Class"] = employee["ICHRA Class"]
            dep_row["Health Election"] = dependent_election_from_tier(employee_election, tier, dep_row["Relationship"])
            output_rows.append(dep_row)

            if not dep_row["Relationship"]:
                notes.append({"Kind": "Warning", "Issue": f"Input row {dep.get('_source_row', '?')}: dependent has no relationship."})
            if not dep_row["DOB"]:
                notes.append({"Kind": "Warning", "Issue": f"Input row {dep.get('_source_row', '?')}: dependent has no DOB."})

    converted = pd.DataFrame(output_rows, columns=OUTPUT_COLUMNS)
    return {
        "converted_df": converted,
        "horizontal_notes": notes,
        "source_rows": len(df),
        "output_rows": len(converted),
        "employee_groups": len(groups),
        "rowwise_horizontal": rowwise,
    }

# =============================================================================
# 8. CELL-LEVEL VALIDATION
# =============================================================================

def validate_date(v: str) -> dict[str, Any]:
    original = str(v or "").strip()
    if not original:
        return {"ok": False, "msg": "DOB is required", "fixed_val": "", "fix_note": None}

    # Excel serial dates, conservative handling.
    if re.fullmatch(r"\d{5}", original):
        try:
            dt = pd.to_datetime(float(original), unit="D", origin="1899-12-30")
            fixed = dt.strftime("%m/%d/%Y")
            return {"ok": True, "msg": None, "fixed_val": fixed, "fix_note": f"Excel date serial converted: '{original}' → '{fixed}'"}
        except Exception:
            pass

    parsed = pd.to_datetime(original, errors="coerce")
    if pd.notna(parsed):
        fixed = parsed.strftime("%m/%d/%Y")
        year = int(parsed.year)
        if not (1900 <= year <= 2100):
            return {"ok": False, "msg": f"Year out of range ({year}) in date '{original}'", "fixed_val": fixed, "fix_note": None}
        note = f"Date reformatted: '{original}' → '{fixed}'" if fixed != original else None
        return {"ok": True, "msg": None, "fixed_val": fixed, "fix_note": note}

    return {"ok": False, "msg": f"Invalid date format '{original}' — expected mm/dd/yyyy", "fixed_val": original, "fix_note": None}


def validate_cell(field_def: dict[str, Any], raw: str) -> dict[str, Any]:
    v = str(raw or "").strip()
    if field_def["required"] and v == "":
        return {"ok": False, "msg": f"{field_def['name']} is required", "fixed_val": "", "fix_note": None}
    if v == "":
        return {"ok": True, "msg": None, "fixed_val": "", "fix_note": None}

    ftype = field_def["type"]
    if ftype == "date":
        return validate_date(v)
    if ftype == "relationship":
        canonical, note = normalize_relationship(v)
        if canonical not in {"Employee", "Spouse", "Child"}:
            return {"ok": False, "msg": f"Invalid Relationship '{v}' — expected Employee, Spouse, or Child", "fixed_val": canonical, "fix_note": note}
        return {"ok": True, "msg": None, "fixed_val": canonical, "fix_note": note}
    if ftype == "election":
        canonical, note = normalize_election(v)
        if canonical not in {"Enroll", "Waive"}:
            return {"ok": False, "msg": f"Invalid Health Election '{v}' — expected Enroll or Waive", "fixed_val": v, "fix_note": None}
        return {"ok": True, "msg": None, "fixed_val": canonical, "fix_note": note}
    if ftype == "tier":
        canonical, note = normalize_tier(v)
        if canonical in {"Employee Only", "Employee + Spouse", "Employee + Children", "Family", ""}:
            return {"ok": True, "msg": None, "fixed_val": canonical, "fix_note": note}
        return {"ok": False, "msg": f"Invalid value '{v}' for {field_def['name']} — expected Employee Only, Employee + Spouse, Employee + Children, or Family", "fixed_val": v, "fix_note": None}
    if ftype == "zipcode":
        us_states = {
            "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA",
            "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
            "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT",
            "VA", "WA", "WV", "WI", "WY", "DC", "PR", "VI", "GU", "AS", "MP",
        }
        if v.upper() in us_states:
            return {"ok": False, "msg": f"{field_def['name']}: looks like a state abbreviation ('{v}'), not a zip code", "fixed_val": v, "fix_note": None}
        fixed = _pad_zip(v)
        note = f"Zip standardised: '{v}' → '{fixed}'" if fixed != v else None
        if not re.fullmatch(r"\d{5}", fixed):
            if re.search(r"[A-Za-z]", v):
                msg = f"{field_def['name']} appears to be a non-U.S. postal code ('{v}'); expected a U.S. 5-digit ZIP code"
            else:
                msg = f"{field_def['name']} must be a 5-digit numeric zip, got '{v}'"
            return {"ok": False, "msg": msg, "fixed_val": fixed, "fix_note": note}
        if fixed in KNOWN_INVALID_OR_UNSUPPORTED_ZIPS:
            msg = f"{field_def['name']} '{fixed}' is not accepted by the CSA quoting tool"
            return {"ok": False, "msg": msg, "fixed_val": fixed, "fix_note": note}
        return {"ok": True, "msg": None, "fixed_val": fixed, "fix_note": note}
    if ftype == "numeric":
        clean = v.replace(",", "").replace("$", "").strip()
        if clean.endswith("%"):
            clean = clean[:-1].strip()
        if not re.fullmatch(r"-?\d+(\.\d+)?", clean):
            return {"ok": False, "msg": f"{field_def['name']} must be numeric, got '{v}'", "fixed_val": v, "fix_note": None}
        note = f"Stripped formatting: '{v}' → '{clean}'" if clean != v else None
        return {"ok": True, "msg": None, "fixed_val": clean, "fix_note": note}
    return {"ok": True, "msg": None, "fixed_val": v, "fix_note": None}

# =============================================================================
# 9. CORE VALIDATION + REFORMAT ENGINE
# =============================================================================

def run_validation(df: pd.DataFrame, zip_cache: dict[str, dict[str, str]], *, source_label: str = "input") -> dict[str, Any]:
    df = df.fillna("").astype(str).reset_index(drop=True)
    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    fixes: list[dict[str, Any]] = []
    clears: list[dict[str, Any]] = []

    col_map, unrecognised_cols, mapping_notes = map_columns(df)
    for uc in unrecognised_cols:
        warnings.append({"Row": "—", "Col": "—", "Field": uc, "Value": "", "Kind": "Warning", "Issue": f"Unrecognised column '{uc}' — not mapped to any canonical field"})

    out_rows: list[dict[str, str]] = []
    valid_positions: list[int] = []
    error_positions: list[int] = []

    for df_ri, row in df.iterrows():
        if all(str(v).strip() == "" for v in row):
            continue
        display_row = int(df_ri) + 2
        new_row = {col: "" for col in OUTPUT_COLUMNS}
        row_has_error = False
        zip_fields_with_errors: set[str] = set()

        for fi, field_def in enumerate(FIELDS):
            fname = field_def["name"]
            ci = col_map.get(fi)
            raw = str(row.iloc[ci]).strip() if ci is not None else ""

            if fname in CLEARED_FIELDS:
                if raw:
                    clears.append({"Row": display_row, "Col": field_def["col"], "Field": fname, "Value": raw, "Kind": "Cleared", "Issue": f"Field '{fname}' automatically cleared"})
                new_row[fname] = ""
                continue

            result = validate_cell(field_def, raw)
            if not result["ok"]:
                errors.append({"Row": display_row, "Col": field_def["col"], "Field": fname, "Value": raw, "Kind": "Error", "Issue": result["msg"]})
                row_has_error = True
                if field_def.get("type") == "zipcode":
                    zip_fields_with_errors.add(fname)
            if result["fix_note"]:
                fixes.append({"Row": display_row, "Col": field_def["col"], "Field": fname, "Value": raw, "Kind": "Auto-fix", "Issue": result["fix_note"]})
            new_row[fname] = result["fixed_val"] if result["ok"] else raw

        if zip_cache:
            for zip_fname, county_fname, col_letter in [
                ("Zip Code", "County", "L"),
                ("Primary Worksite Zip Code", "Primary Worksite County", "N"),
            ]:
                zv = new_row.get(zip_fname, "").strip()
                if not zv or zip_fname in zip_fields_with_errors:
                    continue

                # Only lookup values that already passed U.S. ZIP format validation.
                # Alphanumeric/non-U.S. postal codes and malformed ZIP values are
                # already captured above as validation errors, so avoid adding a
                # second lookup-related issue for the same bad value.
                if not re.fullmatch(r"\d{5}", zv):
                    continue

                info = lookup_zip(zv, zip_cache)
                if info:
                    if not new_row.get(county_fname, "").strip() and info.get("county"):
                        new_row[county_fname] = info["county"]
                        fixes.append({"Row": display_row, "Col": "—", "Field": county_fname, "Value": "", "Kind": "Auto-fix", "Issue": f"County auto-filled from zip '{zv}': '{info['county']}'"})
                else:
                    errors.append({"Row": display_row, "Col": col_letter, "Field": zip_fname, "Value": zv, "Kind": "Error", "Issue": f"{zip_fname} '{zv}' is not a recognized U.S. ZIP code"})
                    row_has_error = True

        pos = len(out_rows)
        if row_has_error:
            error_positions.append(pos)
        else:
            valid_positions.append(pos)
        out_rows.append(new_row)

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
            warnings.append({"Row": display_row, "Col": "Y/Z", "Field": "Cost Comparison", "Value": "", "Kind": "Warning", "Issue": "Both ER Cost and EE Cost should be present for Cost Comparison"})

    reformatted_df = pd.DataFrame(out_rows, columns=OUTPUT_COLUMNS).reset_index(drop=True)
    return {
        "errors": errors,
        "warnings": warnings,
        "fixes": fixes,
        "clears": clears,
        "reformatted_df": reformatted_df,
        "valid_positions": valid_positions,
        "error_positions": error_positions,
        "total_rows": len(out_rows),
        "mapping_notes": mapping_notes,
        "unrecognised_cols": unrecognised_cols,
        "source_label": source_label,
    }

# =============================================================================
# 10. DOWNLOAD HELPERS + SAMPLE DATA
# =============================================================================

def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def issues_to_csv_bytes(errors, warnings, fixes, clears, horizontal_notes=None) -> bytes:
    cols = ["Row", "Col", "Field", "Value", "Kind", "Issue"]
    combined = list(errors) + list(warnings) + list(fixes) + list(clears)
    if horizontal_notes:
        for n in horizontal_notes:
            combined.append({"Row": "—", "Col": "—", "Field": "Horizontal conversion", "Value": "", "Kind": n.get("Kind", "Note"), "Issue": n.get("Issue", "")})
    combined.sort(key=lambda x: (str(x.get("Row", "0")).zfill(6), x.get("Kind", ""), x.get("Field", "")))
    return pd.DataFrame(combined, columns=cols).to_csv(index=False).encode("utf-8")


SAMPLE_VERTICAL_CSV = """First Name,Last Name,Relationship,DOB,Zip Code,Primary Worksite Zip Code,ICHRA Class,Health Election,Current Health Plan Tier,Annual Salary
John,Smith,Employee,01/15/1980,53201,53201,Group A,Enroll,Employee Only,75000
Jane,Smith,Spouse,03/22/1982,53201,53201,Group A,Waive,,
Tim,Smith,Child,06/10/2010,53201,53201,Group A,Waive,,
"""

SAMPLE_HORIZONTAL_CSV = """BIRTH DATE,PRIMARY ADDRESS LINE 1,PRIMARY ADDRESS - CITY,PRIMARY ADDRESS - STATE / TERRITORY,PRIMARY ADDRESS - ZIP CODE,WORKSITE ZIP CODE,WORKER CATEGORY,ANNUAL SALARY,MEDICAL PLAN NAME,MEDICAL COVERAGE TIER,Dependent Number,Dependent Relationship,Dependent DOB
11/26/1986,10502 Barrichello St,Bakersfield,CA,93314,93314,F - Full Time,89139.96,HDHP HSA PRIME PLUS,Family,1,Child,3/4/2011
11/26/1986,10502 Barrichello St,Bakersfield,CA,93314,93314,F - Full Time,89139.96,HDHP HSA PRIME PLUS,Family,2,Child,11/2/2008
11/26/1986,10502 Barrichello St,Bakersfield,CA,93314,93314,F - Full Time,89139.96,HDHP HSA PRIME PLUS,Family,3,Spouse,6/30/1978
1/26/1989,10503 Barrichello St,Bakersfield,CA,93314,93314,F - Full Time,133240.64,HDHP HSA PRIME PLUS,Employee + Children,1,Child,2/8/2014
1/26/1989,10503 Barrichello St,Bakersfield,CA,93314,93314,F - Full Time,133240.64,HDHP HSA PRIME PLUS,Employee + Children,2,Child,2/16/2008
"""

# =============================================================================
# 11. STREAMLIT UI
# =============================================================================

def _style_issues(df_issues: list[dict[str, Any]]) -> "pd.io.formats.style.Styler":
    cols = ["Row", "Col", "Field", "Value", "Kind", "Issue"]
    df = pd.DataFrame(df_issues, columns=cols)
    if df.empty:
        df = pd.DataFrame(columns=cols)
    df = df.sort_values(["Row", "Kind"], key=lambda s: s.map(lambda x: str(x).zfill(6) if str(x).isdigit() else str(x))).reset_index(drop=True)

    def _color(val):
        if val == "Error": return "color:#993C1D;font-weight:600"
        if val == "Warning": return "color:#854F0B;font-weight:600"
        if val == "Auto-fix": return "color:#3B6D11;font-weight:600"
        if val == "Cleared": return "color:#555599;font-weight:600"
        return ""
    styler = df.style
    return styler.map(_color, subset=["Kind"]) if hasattr(styler, "map") else styler.applymap(_color, subset=["Kind"])


def main() -> None:
    st.set_page_config(page_title="Census Validator", page_icon="📋", layout="wide")
    st.markdown("""
    <style>
        .block-container { padding-top: 1.4rem; max-width: 1240px; }
        [data-testid="metric-container"] { background:#f7f7f5; border-radius:8px; padding:10px 14px; }
        .banner-ok { background:#EAF3DE; border:1px solid #97C459; border-radius:8px; padding:12px 16px; color:#3B6D11; font-size:14px; margin-bottom:1rem; }
        .info-box { background:#EEF2FB; border:1px solid #ADC0EF; border-radius:8px; padding:12px 16px; font-size:13px; color:#1a3a7a; margin-bottom:1rem; }
    </style>
    """, unsafe_allow_html=True)

    st.title("Census Validator & Reformatter")
    st.caption("Vertical CSA census validation · Horizontal dependent census conversion · Health-only plan/tier logic · System-ready CSV export")

    with st.sidebar:
        st.header("⚙️ Options")
        input_mode = st.radio(
            "Input format",
            ["Auto-detect", "Standard vertical census", "Horizontal dependent census"],
            index=0,
            help="Horizontal mode converts dependents into separate rows under each employee.",
        )
        use_zip_lookup = st.toggle("Auto-fill County from Zip", value=ZIPCODES_AVAILABLE, disabled=not ZIPCODES_AVAILABLE)
        if not ZIPCODES_AVAILABLE:
            st.caption("Install `zipcodes` to enable county auto-fill.")

        st.divider()
        st.markdown("**Horizontal conversion**")
        auto_placeholders = st.toggle("Create placeholder names when missing", value=True)
        infer_missing_tier = st.toggle("Infer dependent enrollment when medical tier is missing", value=True,
                                       help="If the file has a medical plan but no tier, listed dependents are treated as enrolled according to the dependent mix.")
        inherit_worksite = st.toggle("Copy worksite zip to dependent rows", value=True,
                                     help="Recommended because Primary Worksite Zip Code is required by the CSA import template.")
        exclude_interns = st.toggle("Skip interns/contractors in horizontal mode", value=True)

        st.divider()
        st.markdown("**Output column order**")
        for i, fdef in enumerate(FIELDS, 1):
            req = " ✳" if fdef["required"] else ""
            cleared = " 🚫" if fdef["name"] in CLEARED_FIELDS else ""
            st.caption(f"{i}. {fdef['name']}{req}{cleared}")
        st.caption("✳ = Required   🚫 = Always cleared")

    col_up, col_paste = st.columns([1, 1], gap="large")
    with col_up:
        uploaded = st.file_uploader("Upload census file (.csv / .xlsx / .xls)", type=["csv", "xlsx", "xls"])
    with col_paste:
        pasted = st.text_area("— or paste CSV data —", value=st.session_state.get("paste_text", ""), height=170, placeholder="Paste CSV including header row…")

    b1, b2, b3, _ = st.columns([1, 1, 1, 5])
    with b1:
        run_btn = st.button("✅ Validate", type="primary", use_container_width=True)
    with b2:
        if st.button("📄 Load vertical sample", use_container_width=True):
            st.session_state["paste_text"] = SAMPLE_VERTICAL_CSV
            st.rerun()
    with b3:
        if st.button("↔️ Load horizontal sample", use_container_width=True):
            st.session_state["paste_text"] = SAMPLE_HORIZONTAL_CSV
            st.rerun()

    if run_btn:
        df_input = None
        try:
            if uploaded:
                df_input = read_uploaded_file(uploaded)
            elif pasted.strip():
                df_input = _read_csv_text(pasted)
            else:
                st.warning("Please upload a file or paste CSV data first.")
        except Exception as exc:
            st.error(f"Could not read input: {exc}")

        if df_input is not None:
            with st.spinner("Validating and reformatting…"):
                zc = build_zip_cache() if (use_zip_lookup and ZIPCODES_AVAILABLE) else {}
                horizontal_detected, detect_reason = detect_horizontal_census(df_input)
                force_horizontal = input_mode == "Horizontal dependent census"
                force_vertical = input_mode == "Standard vertical census"
                use_horizontal = force_horizontal or (horizontal_detected and not force_vertical)

                horizontal_notes = []
                conversion_summary = None
                validation_input = df_input
                source_label = "standard vertical census"

                if use_horizontal:
                    opts = HorizontalOptions(
                        auto_placeholder_names=auto_placeholders,
                        infer_dependents_enrolled_when_tier_missing=infer_missing_tier,
                        inherit_worksite_zip_to_dependents=inherit_worksite,
                        exclude_interns_and_contractors=exclude_interns,
                    )
                    conv = convert_horizontal_to_vertical(df_input, opts)
                    validation_input = conv["converted_df"]
                    horizontal_notes = conv["horizontal_notes"]
                    conversion_summary = conv
                    source_label = "horizontal dependent census"

                results = run_validation(validation_input, zc, source_label=source_label)
                ref_df = results["reformatted_df"]
                results["valid_df"] = ref_df.iloc[results["valid_positions"]].reset_index(drop=True)
                results["error_df"] = ref_df.iloc[results["error_positions"]].reset_index(drop=True)
                results["horizontal_used"] = use_horizontal
                results["detect_reason"] = detect_reason
                results["horizontal_notes"] = horizontal_notes
                results["conversion_summary"] = conversion_summary
                st.session_state["results"] = results

    if "results" not in st.session_state:
        return

    res = st.session_state["results"]
    errors = res["errors"]
    warnings = res["warnings"]
    fixes = res["fixes"]
    clears = res["clears"]
    horizontal_notes = res.get("horizontal_notes", [])
    ref_df = res["reformatted_df"]
    valid_df = res["valid_df"]
    error_df = res["error_df"]
    total = res["total_rows"]
    err_rows = len(res["error_positions"])
    valid_rows = total - err_rows

    st.divider()
    if res.get("horizontal_used"):
        conv = res.get("conversion_summary") or {}
        st.markdown(
            f'<div class="info-box">↔️ Horizontal conversion applied: '
            f'{conv.get("source_rows", "?")} source rows → {conv.get("employee_groups", "?")} employee groups → '
            f'{conv.get("output_rows", total)} output rows. Dependents were ordered Employee, Spouse, Child.</div>',
            unsafe_allow_html=True,
        )
    else:
        st.caption(f"Input handled as standard vertical census. Detection note: {res.get('detect_reason', '')}")

    if res.get("mapping_notes") or res.get("unrecognised_cols"):
        with st.expander("🗺️ Column mapping notes", expanded=bool(res.get("unrecognised_cols"))):
            for note in res.get("mapping_notes", []):
                st.caption(f"• {note}")
            for uc in res.get("unrecognised_cols", []):
                st.caption(f"• Unrecognised: `{uc}`")

    m1, m2, m3, m4, m5, m6, m7, m8 = st.columns(8)
    m1.metric("Total rows", total)
    m2.metric("✅ Valid rows", valid_rows)
    m3.metric("❌ Error rows", err_rows)
    m4.metric("Errors", len(errors))
    m5.metric("Warnings", len(warnings) + len([n for n in horizontal_notes if n.get("Kind") == "Warning"]))
    m6.metric("Auto-fixes", len(fixes) + len([n for n in horizontal_notes if n.get("Kind") == "Auto-fix"]))
    m7.metric("Cleared fields", len(clears))
    if res.get("horizontal_used"):
        emp_count = int((ref_df["Relationship"] == "Employee").sum()) if "Relationship" in ref_df else 0
        m8.metric("Employees", emp_count)
    else:
        m8.metric("Mode", "Vertical")

    st.divider()
    st.subheader("📥 Download")
    st.caption("Canonical 30-column order · Dates normalised · Health Election normalised · Dependents expanded when horizontal mode is used")
    dc1, dc2, dc3, dc4 = st.columns(4)
    with dc1:
        st.markdown("##### Full reformatted census")
        st.caption(f"All {total} rows · ready for import")
        st.download_button("⬇ Download", data=df_to_csv_bytes(ref_df), file_name="census_reformatted.csv", mime="text/csv", use_container_width=True, key="dl_full")
    with dc2:
        st.markdown("##### Valid rows only")
        st.caption(f"{valid_rows} rows · no validation errors")
        st.download_button("⬇ Download", data=df_to_csv_bytes(valid_df), file_name="census_valid_rows.csv", mime="text/csv", use_container_width=True, key="dl_valid")
    with dc3:
        st.markdown("##### Error rows only")
        st.caption(f"{err_rows} rows · needs review")
        st.download_button("⬇ Download", data=df_to_csv_bytes(error_df), file_name="census_error_rows.csv", mime="text/csv", use_container_width=True, key="dl_errors")
    with dc4:
        st.markdown("##### Validation report")
        st.caption("Errors · warnings · fixes · clears · horizontal notes")
        st.download_button("⬇ Download", data=issues_to_csv_bytes(errors, warnings, fixes, clears, horizontal_notes), file_name="census_validation_report.csv", mime="text/csv", use_container_width=True, key="dl_report")

    st.divider()
    horizontal_as_issues = [{"Row": "—", "Col": "—", "Field": "Horizontal conversion", "Value": "", "Kind": n.get("Kind", "Note"), "Issue": n.get("Issue", "")} for n in horizontal_notes]
    all_issues = errors + warnings + fixes + clears + horizontal_as_issues
    if not all_issues:
        st.markdown(f'<div class="banner-ok">✅ All {total} rows passed validation with no issues.</div>', unsafe_allow_html=True)
    else:
        tab_all, tab_err, tab_warn, tab_fix, tab_clr = st.tabs([
            f"All ({len(all_issues)})", f"❌ Errors ({len(errors)})", f"⚠️ Warnings ({len(warnings) + len([n for n in horizontal_notes if n.get('Kind') == 'Warning'])})",
            f"🔧 Auto-fixes ({len(fixes) + len([n for n in horizontal_notes if n.get('Kind') == 'Auto-fix'])})", f"🚫 Cleared ({len(clears)})",
        ])
        with tab_all:
            st.dataframe(_style_issues(all_issues), use_container_width=True, hide_index=True)
        with tab_err:
            st.dataframe(_style_issues(errors), use_container_width=True, hide_index=True) if errors else st.success("No errors found.")
        with tab_warn:
            warn_list = warnings + [x for x in horizontal_as_issues if x["Kind"] == "Warning"]
            st.dataframe(_style_issues(warn_list), use_container_width=True, hide_index=True) if warn_list else st.success("No warnings found.")
        with tab_fix:
            fix_list = fixes + [x for x in horizontal_as_issues if x["Kind"] == "Auto-fix"]
            st.dataframe(_style_issues(fix_list), use_container_width=True, hide_index=True) if fix_list else st.info("No auto-fixes applied.")
        with tab_clr:
            st.dataframe(_style_issues(clears), use_container_width=True, hide_index=True) if clears else st.info("No fields were cleared.")

    st.divider()
    with st.expander("🔍 Preview reformatted output", expanded=False):
        st.dataframe(ref_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
