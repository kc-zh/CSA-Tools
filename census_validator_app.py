"""
CSA Census Validator, Reformatter, and Horizontal Census Converter
==================================================================

Single-file Streamlit app for testing and deployment.

Install:
    pip install streamlit pandas openpyxl rapidfuzz zipcodes

Run:
    streamlit run streamlit_census_app_single.py

Notes:
    - This file intentionally contains both the processing engine and a thin
      Streamlit test interface so it can be dropped into a GitHub repository as
      one app file.
    - The processing functions are pure enough to be reused outside Streamlit.
    - Streamlit is only used in the app wrapper at the bottom of this file.
"""

from __future__ import annotations

import io
import logging
import re
import traceback
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Iterable, Optional

import pandas as pd
import streamlit as st

try:
    from rapidfuzz import fuzz
    RAPIDFUZZ_AVAILABLE = True
except Exception:
    fuzz = None
    RAPIDFUZZ_AVAILABLE = False

try:
    import zipcodes as zipcodes_pkg
    ZIPCODES_AVAILABLE = True
except Exception:
    zipcodes_pkg = None
    ZIPCODES_AVAILABLE = False

APP_VERSION = "2026.06.15.single-file.1"
LOGGER = logging.getLogger("csa_census_app")
if not LOGGER.handlers:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")


class InputMode(str, Enum):
    AUTO = "Auto"
    VERTICAL = "Vertical"
    HORIZONTAL = "Horizontal"


@dataclass(frozen=True)
class FieldDef:
    col: str
    name: str
    required: bool
    kind: str
    values: tuple[str, ...] = ()


FIELDS: tuple[FieldDef, ...] = (
    FieldDef("A", "First Name", True, "text"),
    FieldDef("B", "Last Name", True, "text"),
    FieldDef("C", "Employee ID", False, "text"),
    FieldDef("D", "Relationship", True, "relationship"),
    FieldDef("E", "DOB", True, "date"),
    FieldDef("F", "Gender", False, "text"),
    FieldDef("G", "Email", False, "text"),
    FieldDef("H", "Address Line 1", False, "text"),
    FieldDef("I", "Address Line 2", False, "text"),
    FieldDef("J", "City", False, "text"),
    FieldDef("K", "State", False, "text"),
    FieldDef("L", "Zip Code", True, "zipcode"),
    FieldDef("M", "County", False, "text"),
    FieldDef("N", "Primary Worksite Zip Code", True, "zipcode"),
    FieldDef("O", "Primary Worksite County", False, "text"),
    FieldDef("P", "ICHRA Class", False, "text"),
    FieldDef("Q", "Health Election", True, "election"),
    FieldDef("R", "Current Health Plan Vendor", False, "text"),
    FieldDef("S", "Current Health Plan", False, "text"),
    FieldDef("T", "Current Health Plan Tier", False, "tier", ("Employee Only", "Employee + Spouse", "Employee + Children", "Family", "")),
    FieldDef("U", "Current Health Plan OOP (single)", False, "numeric"),
    FieldDef("V", "Current Health Plan OOP (family)", False, "numeric"),
    FieldDef("W", "Current Health Plan Deductible (single)", False, "numeric"),
    FieldDef("X", "Current Health Plan Deductible (family)", False, "numeric"),
    FieldDef("Y", "Current Health Plan ER Cost", False, "numeric"),
    FieldDef("Z", "Current Health Plan EE Cost", False, "numeric"),
    FieldDef("AA", "Annual Salary", False, "numeric"),
    FieldDef("AB", "Hourly Rate", False, "numeric"),
    FieldDef("AC", "Hours Per Week", False, "numeric"),
    FieldDef("AD", "Notes", False, "text"),
)

OUTPUT_COLUMNS = [f.name for f in FIELDS]
FIELD_BY_NAME = {f.name: f for f in FIELDS}
FIELD_INDEX = {f.name: idx for idx, f in enumerate(FIELDS)}

CLEARED_FIELDS = {
    "Employee ID",
    "Gender",
    "Email",
    "Address Line 1",
    "Address Line 2",
    "City",
    "State",
}

ANCILLARY_WORDS = {
    "dental", "dent", "vision", "vis", "life", "std", "ltd", "disability",
    "accident", "critical", "hospital", "supplemental", "voluntary", "ancillary",
}

US_STATE_CODES = {
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID",
    "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS",
    "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK",
    "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV",
    "WI", "WY", "DC", "PR", "VI", "GU", "AS", "MP",
}


@dataclass
class EngineOptions:
    input_mode: InputMode = InputMode.AUTO
    assumption_level: int = 2
    header_match_threshold: int = 88
    relationship_match_threshold: int = 86
    horizontal_detection_sensitivity: int = 3
    auto_fill_county_from_zip: bool = True
    strict_health_only: bool = True
    infer_dependents_enrolled_when_tier_missing: bool = True
    auto_placeholder_names: bool = True
    inherit_worksite_zip_to_dependents: bool = True
    exclude_interns_contractors_inactive: bool = True
    max_file_rows: int = 250_000
    max_file_columns: int = 500


@dataclass
class Issue:
    row: Any
    col: str
    field: str
    value: str
    kind: str
    issue: str

    def as_dict(self) -> dict[str, Any]:
        return {
            "Row": self.row,
            "Col": self.col,
            "Field": self.field,
            "Value": self.value,
            "Kind": self.kind,
            "Issue": self.issue,
        }


@dataclass
class ProcessResult:
    reformatted_df: pd.DataFrame
    errors: list[Issue] = field(default_factory=list)
    warnings: list[Issue] = field(default_factory=list)
    fixes: list[Issue] = field(default_factory=list)
    clears: list[Issue] = field(default_factory=list)
    notes: list[Issue] = field(default_factory=list)
    mapping_notes: list[str] = field(default_factory=list)
    unrecognized_columns: list[str] = field(default_factory=list)
    error_positions: list[int] = field(default_factory=list)
    valid_positions: list[int] = field(default_factory=list)
    source_rows: int = 0
    source_columns: int = 0
    horizontal_used: bool = False
    detect_reason: str = ""
    exception_text: str = ""

    @property
    def total_rows(self) -> int:
        return len(self.reformatted_df)

    @property
    def has_errors(self) -> bool:
        return bool(self.errors)

    @property
    def valid_df(self) -> pd.DataFrame:
        if self.reformatted_df.empty:
            return pd.DataFrame(columns=OUTPUT_COLUMNS)
        return self.reformatted_df.iloc[self.valid_positions].reset_index(drop=True)

    @property
    def error_df(self) -> pd.DataFrame:
        if self.reformatted_df.empty or not self.error_positions:
            return pd.DataFrame(columns=OUTPUT_COLUMNS)
        return self.reformatted_df.iloc[self.error_positions].reset_index(drop=True)

    def issue_frame(self) -> pd.DataFrame:
        rows = [i.as_dict() for i in self.errors + self.warnings + self.fixes + self.clears + self.notes]
        return pd.DataFrame(rows, columns=["Row", "Col", "Field", "Value", "Kind", "Issue"])

    def summary(self) -> dict[str, Any]:
        return {
            "source_rows": self.source_rows,
            "source_columns": self.source_columns,
            "output_rows": self.total_rows,
            "errors": len(self.errors),
            "warnings": len(self.warnings),
            "fixes": len(self.fixes),
            "clears": len(self.clears),
            "notes": len(self.notes),
            "horizontal_used": self.horizontal_used,
            "detect_reason": self.detect_reason,
            "exception": bool(self.exception_text),
        }


# =============================================================================
# Normalization helpers
# =============================================================================


def norm_key(value: Any) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\ufeff", " ").strip().lower()
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def cell(value: Any) -> str:
    if value is None:
        return ""
    text = str(value)
    if text.lower() in {"nan", "none", "nat"}:
        return ""
    return text.strip()


def row_value(row: pd.Series, column: Optional[str]) -> str:
    if not column or column not in row.index:
        return ""
    return cell(row[column])


def first_nonblank(*values: Any) -> str:
    for value in values:
        text = cell(value)
        if text:
            return text
    return ""


def split_multi(value: Any) -> list[str]:
    text = cell(value)
    if not text:
        return []
    return [part.strip() for part in re.split(r"\s*[|;]\s*", text) if part.strip()]


def contains_any(text: Any, words: Iterable[str]) -> bool:
    key = norm_key(text)
    return any(w in key.split() or w in key for w in words)


def is_ancillary_header(header: Any) -> bool:
    return contains_any(header, ANCILLARY_WORDS)


# =============================================================================
# Header aliases and mapping
# =============================================================================


def make_aliases() -> dict[str, str]:
    aliases = {
        "first name": "First Name", "firstname": "First Name", "first": "First Name", "given name": "First Name",
        "employee first name": "First Name", "ee first name": "First Name", "member first name": "First Name",
        "last name": "Last Name", "lastname": "Last Name", "last": "Last Name", "surname": "Last Name",
        "family name": "Last Name", "employee last name": "Last Name", "ee last name": "Last Name", "member last name": "Last Name",
        "employee id": "Employee ID", "emp id": "Employee ID", "empid": "Employee ID", "ee id": "Employee ID",
        "employee number": "Employee ID", "employee no": "Employee ID", "employee #": "Employee ID", "id": "Employee ID",
        "relationship": "Relationship", "relation": "Relationship", "rel": "Relationship", "member type": "Relationship",
        "dependent type": "Relationship", "coverage member type": "Relationship",
        "dob": "DOB", "date of birth": "DOB", "birth date": "DOB", "birthdate": "DOB", "employee dob": "DOB",
        "gender": "Gender", "sex": "Gender",
        "email": "Email", "email address": "Email", "e mail": "Email", "work email": "Email", "personal email": "Email",
        "address": "Address Line 1", "address 1": "Address Line 1", "address line 1": "Address Line 1", "street address": "Address Line 1",
        "address 2": "Address Line 2", "address line 2": "Address Line 2", "apt": "Address Line 2", "suite": "Address Line 2",
        "city": "City", "town": "City",
        "state": "State", "st": "State", "state code": "State",
        "zip": "Zip Code", "zipcode": "Zip Code", "zip code": "Zip Code", "postal code": "Zip Code", "home zip": "Zip Code",
        "home zip code": "Zip Code", "residential zip": "Zip Code", "employee zip": "Zip Code",
        "county": "County", "home county": "County",
        "primary worksite zip code": "Primary Worksite Zip Code", "primary worksite zip": "Primary Worksite Zip Code",
        "worksite zip": "Primary Worksite Zip Code", "worksite zip code": "Primary Worksite Zip Code", "work zip": "Primary Worksite Zip Code",
        "site zip": "Primary Worksite Zip Code", "primary work zip": "Primary Worksite Zip Code",
        "primary worksite county": "Primary Worksite County", "worksite county": "Primary Worksite County", "work county": "Primary Worksite County",
        "ichra class": "ICHRA Class", "class": "ICHRA Class", "benefit class": "ICHRA Class", "employee class": "ICHRA Class",
        "health election": "Health Election", "medical election": "Health Election", "benefit election": "Health Election",
        "coverage election": "Health Election", "election": "Health Election", "enrollment election": "Health Election",
        "current health plan vendor": "Current Health Plan Vendor", "current health vendor": "Current Health Plan Vendor",
        "health vendor": "Current Health Plan Vendor", "medical vendor": "Current Health Plan Vendor", "carrier": "Current Health Plan Vendor",
        "health carrier": "Current Health Plan Vendor", "medical carrier": "Current Health Plan Vendor", "insurance carrier": "Current Health Plan Vendor",
        "current health plan": "Current Health Plan", "health plan": "Current Health Plan", "medical plan": "Current Health Plan",
        "plan name": "Current Health Plan", "current medical plan": "Current Health Plan", "medical plan name": "Current Health Plan",
        "current health plan tier": "Current Health Plan Tier", "health plan tier": "Current Health Plan Tier",
        "medical tier": "Current Health Plan Tier", "medical plan tier": "Current Health Plan Tier", "coverage tier": "Current Health Plan Tier",
        "plan tier": "Current Health Plan Tier", "tier": "Current Health Plan Tier",
        "current health plan oop single": "Current Health Plan OOP (single)", "oop single": "Current Health Plan OOP (single)",
        "out of pocket single": "Current Health Plan OOP (single)", "single oop": "Current Health Plan OOP (single)",
        "current health plan oop family": "Current Health Plan OOP (family)", "oop family": "Current Health Plan OOP (family)",
        "out of pocket family": "Current Health Plan OOP (family)", "family oop": "Current Health Plan OOP (family)",
        "current health plan deductible single": "Current Health Plan Deductible (single)", "deductible single": "Current Health Plan Deductible (single)",
        "single deductible": "Current Health Plan Deductible (single)",
        "current health plan deductible family": "Current Health Plan Deductible (family)", "deductible family": "Current Health Plan Deductible (family)",
        "family deductible": "Current Health Plan Deductible (family)",
        "current health plan er cost": "Current Health Plan ER Cost", "health plan er cost": "Current Health Plan ER Cost",
        "er cost": "Current Health Plan ER Cost", "employer cost": "Current Health Plan ER Cost", "employer contribution": "Current Health Plan ER Cost",
        "current health plan ee cost": "Current Health Plan EE Cost", "health plan ee cost": "Current Health Plan EE Cost",
        "ee cost": "Current Health Plan EE Cost", "employee cost": "Current Health Plan EE Cost", "employee contribution": "Current Health Plan EE Cost",
        "annual salary": "Annual Salary", "salary": "Annual Salary", "yearly salary": "Annual Salary", "base salary": "Annual Salary",
        "hourly rate": "Hourly Rate", "hourly": "Hourly Rate", "hourly wage": "Hourly Rate", "wage": "Hourly Rate",
        "hours per week": "Hours Per Week", "hours week": "Hours Per Week", "hours/week": "Hours Per Week", "weekly hours": "Hours Per Week",
        "notes": "Notes", "note": "Notes", "comments": "Notes", "remarks": "Notes",
    }
    for field_def in FIELDS:
        aliases[norm_key(field_def.name)] = field_def.name
        aliases[norm_key(field_def.col)] = field_def.name
    return {norm_key(k): v for k, v in aliases.items()}


HEADER_ALIASES = make_aliases()


def similarity(a: str, b: str) -> int:
    a_key = norm_key(a)
    b_key = norm_key(b)
    if not a_key or not b_key:
        return 0
    if RAPIDFUZZ_AVAILABLE:
        return int(max(fuzz.ratio(a_key, b_key), fuzz.token_sort_ratio(a_key, b_key)))
    a_set = set(a_key.split())
    b_set = set(b_key.split())
    if not a_set or not b_set:
        return 0
    return int(100 * len(a_set & b_set) / len(a_set | b_set))


def map_columns(df: pd.DataFrame, options: EngineOptions) -> tuple[dict[str, str], list[str], list[str]]:
    raw_headers = [str(c) for c in df.columns]
    normalized = {str(c): norm_key(c) for c in df.columns}
    mapping: dict[str, str] = {}
    used: set[str] = set()
    notes: list[str] = []

    for col in raw_headers:
        if options.strict_health_only and is_ancillary_header(col):
            continue
        alias = HEADER_ALIASES.get(normalized[col])
        if alias and alias not in mapping:
            mapping[alias] = col
            used.add(col)

    for field_def in FIELDS:
        if field_def.name in mapping:
            continue
        best_col = ""
        best_score = 0
        for col in raw_headers:
            if col in used:
                continue
            if options.strict_health_only and is_ancillary_header(col):
                continue
            score = similarity(field_def.name, col)
            if score > best_score:
                best_col = col
                best_score = score
        if best_col and best_score >= options.header_match_threshold:
            mapping[field_def.name] = best_col
            used.add(best_col)
            notes.append(f"Fuzzy mapped '{best_col}' to '{field_def.name}' with score {best_score}.")

    unrecognized = [c for c in raw_headers if c not in used]
    return mapping, unrecognized, notes


# =============================================================================
# Value normalization and validation
# =============================================================================


RELATIONSHIP_ALIASES = {
    "employee": "Employee", "ee": "Employee", "emp": "Employee", "self": "Employee", "subscriber": "Employee",
    "primary": "Employee", "insured": "Employee", "member": "Employee",
    "spouse": "Spouse", "sp": "Spouse", "sps": "Spouse", "spse": "Spouse", "husband": "Spouse", "wife": "Spouse",
    "partner": "Spouse", "domestic partner": "Spouse", "registered domestic partner": "Spouse", "life partner": "Spouse",
    "child": "Child", "ch": "Child", "dep": "Child", "dependent": "Child", "son": "Child", "daughter": "Child",
    "stepchild": "Child", "step child": "Child", "stepson": "Child", "stepdaughter": "Child", "foster child": "Child",
}

ELECTION_ENROLL = {"enroll", "e", "yes", "y", "enrolled", "active", "participating", "covered", "elect", "elected"}
ELECTION_WAIVE = {"waive", "w", "no", "n", "waived", "decline", "declined", "not enrolled", "opt out", "opt-out", "waiving"}

TIER_ALIASES = {
    "employee only": "Employee Only", "ee only": "Employee Only", "eo": "Employee Only", "single": "Employee Only",
    "individual": "Employee Only", "employee": "Employee Only", "ee": "Employee Only", "employee-only": "Employee Only",
    "employee spouse": "Employee + Spouse", "employee plus spouse": "Employee + Spouse", "employee sp": "Employee + Spouse",
    "ee spouse": "Employee + Spouse", "ee plus spouse": "Employee + Spouse", "ee sp": "Employee + Spouse", "es": "Employee + Spouse",
    "employee child": "Employee + Children", "employee children": "Employee + Children", "employee plus child": "Employee + Children",
    "employee plus children": "Employee + Children", "ee child": "Employee + Children", "ee children": "Employee + Children",
    "ee plus child": "Employee + Children", "ee plus children": "Employee + Children", "ec": "Employee + Children", "ech": "Employee + Children",
    "employee children only": "Employee + Children", "employee child ren": "Employee + Children",
    "family": "Family", "fam": "Family", "employee family": "Family", "ee family": "Family", "ef": "Family",
    "employee spouse children": "Family", "employee plus spouse children": "Family", "ee spouse children": "Family",
}


def normalize_relationship(value: Any, threshold: int = 86) -> tuple[str, Optional[str]]:
    original = cell(value)
    key = norm_key(original)
    if not key:
        return "", None
    if key in RELATIONSHIP_ALIASES:
        normalized = RELATIONSHIP_ALIASES[key]
        note = None if normalized.lower() == key else f"Relationship normalized from '{original}' to '{normalized}'."
        return normalized, note
    best_alias = ""
    best_score = 0
    for alias in RELATIONSHIP_ALIASES:
        score = similarity(key, alias)
        if score > best_score:
            best_alias = alias
            best_score = score
    if best_alias and best_score >= threshold:
        normalized = RELATIONSHIP_ALIASES[best_alias]
        return normalized, f"Relationship fuzzy matched from '{original}' to '{normalized}' with score {best_score}."
    return original, None


def normalize_election(value: Any) -> tuple[str, Optional[str]]:
    original = cell(value)
    key = norm_key(original)
    if not key:
        return "", None
    if key in ELECTION_ENROLL:
        return "Enroll", None if original == "Enroll" else f"Health Election normalized from '{original}' to 'Enroll'."
    if key in ELECTION_WAIVE:
        return "Waive", None if original == "Waive" else f"Health Election normalized from '{original}' to 'Waive'."
    return original, None


def normalize_tier(value: Any) -> tuple[str, Optional[str]]:
    original = cell(value)
    key = norm_key(original).replace(" + ", " ")
    key = key.replace("plus", " plus ")
    key = re.sub(r"\s+", " ", key).strip()
    if not key:
        return "", None
    compact = key.replace(" ", "")
    candidates = {key, compact}
    for candidate in candidates:
        if candidate in TIER_ALIASES:
            normalized = TIER_ALIASES[candidate]
            return normalized, None if original == normalized else f"Current Health Plan Tier normalized from '{original}' to '{normalized}'."
    if ("spouse" in key or "sp" in key) and ("child" in key or "children" in key or "family" in key):
        return "Family", f"Current Health Plan Tier normalized from '{original}' to 'Family'."
    if "spouse" in key or re.search(r"\bsp\b", key):
        return "Employee + Spouse", f"Current Health Plan Tier normalized from '{original}' to 'Employee + Spouse'."
    if "child" in key or "children" in key or re.search(r"\bch\b", key):
        return "Employee + Children", f"Current Health Plan Tier normalized from '{original}' to 'Employee + Children'."
    if "family" in key or key == "fam":
        return "Family", f"Current Health Plan Tier normalized from '{original}' to 'Family'."
    return original, None


def normalize_zip(value: Any) -> tuple[str, Optional[str], Optional[str]]:
    original = cell(value)
    if not original:
        return "", None, None
    if original.upper() in US_STATE_CODES:
        return original, None, f"Zip field looks like a state abbreviation, not a zip code: '{original}'."
    digits = re.sub(r"\D", "", original)
    if len(digits) >= 9:
        fixed = f"{digits[:5]}-{digits[5:9]}"
    elif 0 < len(digits) < 5:
        fixed = digits.zfill(5)
    elif len(digits) == 5:
        fixed = digits
    else:
        return original, None, f"Zip Code must be 5 digits or ZIP+4, got '{original}'."
    note = None if fixed == original else f"Zip normalized from '{original}' to '{fixed}'."
    return fixed, note, None


def normalize_numeric(value: Any) -> tuple[str, Optional[str], Optional[str]]:
    original = cell(value)
    if not original:
        return "", None, None
    cleaned = original.replace("$", "").replace(",", "").replace("%", "").strip()
    cleaned = re.sub(r"\s+", "", cleaned)
    if not re.match(r"^-?\d+(\.\d+)?$", cleaned):
        return original, None, f"Numeric field must be numeric, got '{original}'."
    note = None if cleaned == original else f"Numeric formatting stripped from '{original}' to '{cleaned}'."
    return cleaned, note, None


def normalize_date(value: Any) -> tuple[str, Optional[str], Optional[str]]:
    original = cell(value)
    if not original:
        return "", None, None
    if re.match(r"^\d+(\.0)?$", original):
        try:
            serial = float(original)
            if 1 <= serial <= 60000:
                dt = pd.to_datetime(serial, unit="D", origin="1899-12-30", errors="coerce")
                if not pd.isna(dt):
                    fixed = dt.strftime("%m/%d/%Y")
                    return fixed, f"Excel serial date converted from '{original}' to '{fixed}'.", None
        except Exception:
            pass
    parsed = pd.to_datetime(original, errors="coerce")
    if pd.isna(parsed):
        return original, None, f"Invalid date format '{original}'. Expected mm/dd/yyyy."
    fixed = parsed.strftime("%m/%d/%Y")
    note = None if fixed == original else f"Date normalized from '{original}' to '{fixed}'."
    return fixed, note, None


def validate_field(field_def: FieldDef, value: Any, options: EngineOptions) -> tuple[str, list[str], list[str]]:
    raw = cell(value)
    notes: list[str] = []
    errors: list[str] = []

    if field_def.required and not raw:
        return "", notes, [f"{field_def.name} is required."]
    if not raw:
        return "", notes, errors

    if field_def.kind == "relationship":
        fixed, note = normalize_relationship(raw, options.relationship_match_threshold)
        if note:
            notes.append(note)
        if fixed not in {"Employee", "Spouse", "Child"}:
            errors.append(f"Invalid Relationship '{raw}'. Expected Employee, Spouse, or Child.")
        return fixed, notes, errors

    if field_def.kind == "election":
        fixed, note = normalize_election(raw)
        if note:
            notes.append(note)
        if fixed not in {"Enroll", "Waive"}:
            errors.append(f"Invalid Health Election '{raw}'. Expected Enroll or Waive.")
        return fixed, notes, errors

    if field_def.kind == "tier":
        fixed, note = normalize_tier(raw)
        if note:
            notes.append(note)
        if fixed not in field_def.values:
            errors.append(f"Invalid Current Health Plan Tier '{raw}'. Expected Employee Only, Employee + Spouse, Employee + Children, or Family.")
        return fixed, notes, errors

    if field_def.kind == "zipcode":
        fixed, note, err = normalize_zip(raw)
        if note:
            notes.append(note)
        if err:
            errors.append(err)
        return fixed, notes, errors

    if field_def.kind == "date":
        fixed, note, err = normalize_date(raw)
        if note:
            notes.append(note)
        if err:
            errors.append(err)
        return fixed, notes, errors

    if field_def.kind == "numeric":
        fixed, note, err = normalize_numeric(raw)
        if note:
            notes.append(note)
        if err:
            errors.append(err)
        return fixed, notes, errors

    return raw, notes, errors


# =============================================================================
# County lookup
# =============================================================================


@st.cache_data(show_spinner=False)
def build_zip_cache() -> dict[str, dict[str, str]]:
    if not ZIPCODES_AVAILABLE:
        return {}
    cache: dict[str, dict[str, str]] = {}
    for item in zipcodes_pkg.list_all():
        z = item.get("zip_code", "")
        county = cell(item.get("county", "")).replace(" County", "")
        if z and z not in cache:
            cache[z] = {"county": county, "state": item.get("state", "")}
    return cache


def lookup_county(zip_value: str, zip_cache: dict[str, dict[str, str]]) -> str:
    if not zip_value or not zip_cache:
        return ""
    five = re.sub(r"\D", "", zip_value)[:5]
    return zip_cache.get(five, {}).get("county", "")


# =============================================================================
# Horizontal census conversion
# =============================================================================


def find_column(df: pd.DataFrame, patterns: Iterable[str], *, exclude_ancillary: bool = True) -> str:
    for col in df.columns:
        if exclude_ancillary and is_ancillary_header(col):
            continue
        key = norm_key(col)
        if all(token in key for token in patterns):
            return str(col)
    return ""


def find_any_column(df: pd.DataFrame, pattern_sets: Iterable[Iterable[str]], *, exclude_ancillary: bool = True) -> str:
    for patterns in pattern_sets:
        found = find_column(df, patterns, exclude_ancillary=exclude_ancillary)
        if found:
            return found
    return ""


def find_dependent_columns(df: pd.DataFrame) -> dict[str, list[dict[str, str]]]:
    spouse_sets: list[dict[str, str]] = []
    child_sets: list[dict[str, str]] = []
    columns = [str(c) for c in df.columns]

    spouse_first = find_any_column(df, [["spouse", "first"], ["sp", "first"]])
    spouse_last = find_any_column(df, [["spouse", "last"], ["sp", "last"]])
    spouse_dob = find_any_column(df, [["spouse", "dob"], ["spouse", "birth"], ["sp", "dob"]])
    if spouse_dob:
        spouse_sets.append({"relationship": "Spouse", "first": spouse_first, "last": spouse_last, "dob": spouse_dob})

    child_indexes: set[str] = set()
    for col in columns:
        key = norm_key(col)
        if "child" not in key and not re.search(r"\bch\s*\d+\b", key):
            continue
        match = re.search(r"(?:child|ch)\s*(\d+)", key)
        child_indexes.add(match.group(1) if match else "1")

    for idx in sorted(child_indexes, key=lambda x: int(x) if x.isdigit() else 999):
        first = find_any_column(df, [["child", idx, "first"], ["ch", idx, "first"], ["child", "first"]])
        last = find_any_column(df, [["child", idx, "last"], ["ch", idx, "last"], ["child", "last"]])
        dob = find_any_column(df, [["child", idx, "dob"], ["child", idx, "birth"], ["ch", idx, "dob"], ["child", "dob"]])
        if dob and not any(existing["dob"] == dob for existing in child_sets):
            child_sets.append({"relationship": "Child", "first": first, "last": last, "dob": dob})

    return {"spouse": spouse_sets, "child": child_sets}


def detect_horizontal(df: pd.DataFrame, mapping: dict[str, str], options: EngineOptions) -> tuple[bool, str]:
    if options.input_mode == InputMode.VERTICAL:
        return False, "Input mode forced to vertical."
    if options.input_mode == InputMode.HORIZONTAL:
        return True, "Input mode forced to horizontal."

    headers = " ".join(norm_key(c) for c in df.columns)
    dep_header_score = 0
    for term in ["spouse", "child", "dependent", "dep dob", "member dob", "relationship"]:
        if term in headers:
            dep_header_score += 1

    rel_col = mapping.get("Relationship", "")
    vertical_relationships = 0
    if rel_col and rel_col in df.columns:
        values = df[rel_col].astype(str).map(norm_key)
        vertical_relationships = int(values.isin({"employee", "spouse", "child", "dependent", "dep", "ee", "self"}).sum())

    if dep_header_score >= max(2, options.horizontal_detection_sensitivity):
        return True, f"Dependent-oriented headers detected with score {dep_header_score}."
    if not rel_col and ("spouse" in headers or "child" in headers or "dependent" in headers):
        return True, "No standard relationship column was found and dependent headers are present."
    if rel_col and "employee dob" in headers and ("dependent dob" in headers or "spouse dob" in headers or "child dob" in headers):
        return True, "Row-wise horizontal layout detected."
    if vertical_relationships == 0 and ("medical" in headers or "health" in headers) and ("tier" in headers or "coverage" in headers):
        return True, "No vertical relationships found and health tier fields are present."
    return False, "Standard vertical layout detected."


def empty_output_row() -> dict[str, str]:
    return {col: "" for col in OUTPUT_COLUMNS}


def employee_group_key(row: pd.Series, mapping: dict[str, str], row_number: int) -> str:
    emp_id = row_value(row, mapping.get("Employee ID"))
    if emp_id:
        return f"id:{emp_id}"
    first = row_value(row, mapping.get("First Name"))
    last = row_value(row, mapping.get("Last Name"))
    dob = row_value(row, mapping.get("DOB"))
    zip_code = row_value(row, mapping.get("Zip Code"))
    if first or last or dob:
        return f"person:{norm_key(first)}:{norm_key(last)}:{norm_key(dob)}:{norm_key(zip_code)}"
    return f"row:{row_number}"


def should_skip_horizontal_row(row: pd.Series, options: EngineOptions) -> bool:
    if not options.exclude_interns_contractors_inactive:
        return False
    combined = " ".join(norm_key(v) for v in row.values)
    skip_terms = {"intern", "contractor", "terminated", "inactive", "seasonal ineligible", "not eligible"}
    return any(term in combined for term in skip_terms)


def infer_employee_election(row: pd.Series, mapping: dict[str, str]) -> str:
    raw = row_value(row, mapping.get("Health Election"))
    election, _ = normalize_election(raw)
    if election in {"Enroll", "Waive"}:
        return election
    tier, _ = normalize_tier(row_value(row, mapping.get("Current Health Plan Tier")))
    plan = row_value(row, mapping.get("Current Health Plan"))
    vendor = row_value(row, mapping.get("Current Health Plan Vendor"))
    if tier or plan or vendor:
        return "Enroll"
    return "Waive"


def dependent_election(relationship: str, employee_election: str, tier: str, options: EngineOptions, has_medical_plan: bool) -> str:
    if employee_election == "Waive":
        return "Waive"
    normalized_tier, _ = normalize_tier(tier)
    if normalized_tier == "Employee Only":
        return "Waive"
    if normalized_tier == "Employee + Spouse":
        return "Enroll" if relationship == "Spouse" else "Waive"
    if normalized_tier == "Employee + Children":
        return "Enroll" if relationship == "Child" else "Waive"
    if normalized_tier == "Family":
        return "Enroll"
    if options.infer_dependents_enrolled_when_tier_missing and has_medical_plan and employee_election == "Enroll":
        return "Enroll"
    return "Waive"


def build_employee_row(row: pd.Series, mapping: dict[str, str], sequence: int, options: EngineOptions) -> dict[str, str]:
    out = empty_output_row()
    for field_def in FIELDS:
        if field_def.name in CLEARED_FIELDS:
            continue
        raw = row_value(row, mapping.get(field_def.name))
        out[field_def.name] = raw

    if not out["First Name"] and options.auto_placeholder_names:
        out["First Name"] = f"Employee{sequence}"
    if not out["Last Name"] and options.auto_placeholder_names:
        out["Last Name"] = "Placeholder"
    out["Relationship"] = "Employee"
    out["Health Election"] = infer_employee_election(row, mapping)
    return out


def build_dependent_row(
    employee_row: dict[str, str],
    relationship: str,
    dob: str,
    first_name: str,
    last_name: str,
    sequence: int,
    dep_sequence: int,
    employee_election: str,
    tier: str,
    options: EngineOptions,
) -> dict[str, str]:
    out = empty_output_row()
    out["First Name"] = first_name or (f"{relationship}{dep_sequence}" if options.auto_placeholder_names else "")
    out["Last Name"] = last_name or employee_row.get("Last Name", "") or ("Placeholder" if options.auto_placeholder_names else "")
    out["Relationship"] = relationship
    out["DOB"] = dob
    out["Zip Code"] = employee_row.get("Zip Code", "")
    out["County"] = employee_row.get("County", "")
    if options.inherit_worksite_zip_to_dependents:
        out["Primary Worksite Zip Code"] = employee_row.get("Primary Worksite Zip Code", "") or employee_row.get("Zip Code", "")
        out["Primary Worksite County"] = employee_row.get("Primary Worksite County", "") or employee_row.get("County", "")
    out["ICHRA Class"] = employee_row.get("ICHRA Class", "")
    has_medical_plan = bool(employee_row.get("Current Health Plan") or employee_row.get("Current Health Plan Vendor") or employee_row.get("Current Health Plan Tier"))
    out["Health Election"] = dependent_election(relationship, employee_election, tier, options, has_medical_plan)
    return out


def extract_rowwise_dependent(row: pd.Series, mapping: dict[str, str], options: EngineOptions) -> Optional[dict[str, str]]:
    rel_raw = row_value(row, mapping.get("Relationship"))
    relationship, _ = normalize_relationship(rel_raw, options.relationship_match_threshold)

    dep_dob_col = find_any_column(pd.DataFrame(columns=row.index), [["dependent", "dob"], ["dependent", "birth"], ["member", "dob"], ["spouse", "dob"], ["child", "dob"]])
    dep_first_col = find_any_column(pd.DataFrame(columns=row.index), [["dependent", "first"], ["member", "first"], ["spouse", "first"], ["child", "first"]])
    dep_last_col = find_any_column(pd.DataFrame(columns=row.index), [["dependent", "last"], ["member", "last"], ["spouse", "last"], ["child", "last"]])

    dep_dob = row_value(row, dep_dob_col)
    if not dep_dob and relationship in {"Spouse", "Child"}:
        dep_dob = row_value(row, mapping.get("DOB"))
    if not dep_dob:
        return None

    if relationship not in {"Spouse", "Child"}:
        header_text = " ".join(norm_key(c) for c in row.index)
        if "spouse" in header_text:
            relationship = "Spouse"
        elif "child" in header_text or "dependent" in header_text:
            relationship = "Child"
        else:
            relationship = "Child"

    return {
        "relationship": relationship,
        "dob": dep_dob,
        "first": row_value(row, dep_first_col),
        "last": row_value(row, dep_last_col),
    }


def extract_wide_dependents(row: pd.Series, df: pd.DataFrame) -> list[dict[str, str]]:
    deps: list[dict[str, str]] = []
    dep_cols = find_dependent_columns(df)
    for item in dep_cols["spouse"] + dep_cols["child"]:
        dob = row_value(row, item.get("dob"))
        if not dob:
            continue
        for dob_part in split_multi(dob) or [dob]:
            deps.append({
                "relationship": item["relationship"],
                "dob": dob_part,
                "first": row_value(row, item.get("first")),
                "last": row_value(row, item.get("last")),
            })
    return deps


def convert_horizontal_to_vertical(df: pd.DataFrame, mapping: dict[str, str], options: EngineOptions) -> tuple[pd.DataFrame, list[Issue]]:
    out_rows: list[dict[str, str]] = []
    notes: list[Issue] = []
    groups: dict[str, dict[str, Any]] = {}

    for idx, row in df.iterrows():
        display_row = int(idx) + 2
        if should_skip_horizontal_row(row, options):
            notes.append(Issue(display_row, "-", "Horizontal Conversion", "", "Warning", "Skipped row because it appears ineligible, inactive, terminated, intern, or contractor."))
            continue

        key = employee_group_key(row, mapping, display_row)
        if key not in groups:
            groups[key] = {"first_row": row, "row_numbers": [], "dependents": []}
        groups[key]["row_numbers"].append(display_row)

        dep = extract_rowwise_dependent(row, mapping, options)
        if dep:
            groups[key]["dependents"].append(dep)
        for wide_dep in extract_wide_dependents(row, df):
            groups[key]["dependents"].append(wide_dep)

    for seq, (key, group) in enumerate(groups.items(), start=1):
        base_row = group["first_row"]
        employee = build_employee_row(base_row, mapping, seq, options)
        employee_election = employee["Health Election"]
        tier = employee.get("Current Health Plan Tier", "")
        out_rows.append(employee)

        seen_deps: set[tuple[str, str, str, str]] = set()
        cleaned_deps: list[dict[str, str]] = []
        for dep in group["dependents"]:
            dep_key = (dep.get("relationship", ""), norm_key(dep.get("dob", "")), norm_key(dep.get("first", "")), norm_key(dep.get("last", "")))
            if dep_key in seen_deps or not dep.get("dob"):
                continue
            seen_deps.add(dep_key)
            cleaned_deps.append(dep)

        cleaned_deps.sort(key=lambda d: (0 if d.get("relationship") == "Spouse" else 1, normalize_date(d.get("dob", ""))[0], d.get("first", "")))
        for dep_seq, dep in enumerate(cleaned_deps, start=1):
            out_rows.append(build_dependent_row(
                employee,
                dep.get("relationship", "Child"),
                dep.get("dob", ""),
                dep.get("first", ""),
                dep.get("last", ""),
                seq,
                dep_seq,
                employee_election,
                tier,
                options,
            ))

    notes.append(Issue("-", "-", "Horizontal Conversion", "", "Auto-fix", f"Converted {len(groups)} employee groups into {len(out_rows)} vertical output rows."))
    return pd.DataFrame(out_rows, columns=OUTPUT_COLUMNS), notes


# =============================================================================
# Core processing
# =============================================================================


def validate_and_reformat(df: pd.DataFrame, options: EngineOptions, zip_cache: Optional[dict[str, dict[str, str]]] = None, already_canonical: bool = False) -> ProcessResult:
    zip_cache = zip_cache or {}
    errors: list[Issue] = []
    warnings: list[Issue] = []
    fixes: list[Issue] = []
    clears: list[Issue] = []
    out_rows: list[dict[str, str]] = []
    valid_positions: list[int] = []
    error_positions: list[int] = []

    mapping, unrecognized, mapping_notes = ({name: name for name in OUTPUT_COLUMNS}, [], []) if already_canonical else map_columns(df, options)
    if not already_canonical:
        for col in unrecognized:
            warnings.append(Issue("-", "-", str(col), "", "Warning", f"Unrecognized column '{col}' was not mapped to a CSA field."))

    for source_idx, row in df.iterrows():
        if all(cell(v) == "" for v in row.values):
            continue
        display_row = int(source_idx) + 2
        output_pos = len(out_rows)
        row_has_error = False
        new_row = empty_output_row()

        for field_def in FIELDS:
            raw = row_value(row, mapping.get(field_def.name))
            if field_def.name in CLEARED_FIELDS:
                if raw:
                    clears.append(Issue(display_row, field_def.col, field_def.name, raw, "Cleared", f"Field '{field_def.name}' was cleared from output."))
                new_row[field_def.name] = ""
                continue

            fixed, note_list, error_list = validate_field(field_def, raw, options)
            new_row[field_def.name] = fixed
            for note in note_list:
                fixes.append(Issue(display_row, field_def.col, field_def.name, raw, "Auto-fix", note))
            for err in error_list:
                errors.append(Issue(display_row, field_def.col, field_def.name, raw, "Error", err))
                row_has_error = True

        if options.auto_fill_county_from_zip and zip_cache:
            for zip_field, county_field in [("Zip Code", "County"), ("Primary Worksite Zip Code", "Primary Worksite County")]:
                if new_row.get(zip_field) and not new_row.get(county_field):
                    county = lookup_county(new_row[zip_field], zip_cache)
                    if county:
                        new_row[county_field] = county
                        fixes.append(Issue(display_row, "-", county_field, "", "Auto-fix", f"County auto-filled from {zip_field} '{new_row[zip_field]}': '{county}'."))
                    else:
                        warnings.append(Issue(display_row, "-", zip_field, new_row[zip_field], "Warning", f"Zip code '{new_row[zip_field]}' was not found in the lookup table."))

        er = new_row.get("Current Health Plan ER Cost", "")
        ee = new_row.get("Current Health Plan EE Cost", "")
        if bool(er) ^ bool(ee):
            warnings.append(Issue(display_row, "Y/Z", "Cost Comparison", "", "Warning", "Both ER Cost and EE Cost should be present for cost comparison."))

        out_rows.append(new_row)
        if row_has_error:
            error_positions.append(output_pos)
        else:
            valid_positions.append(output_pos)

    return ProcessResult(
        reformatted_df=pd.DataFrame(out_rows, columns=OUTPUT_COLUMNS),
        errors=errors,
        warnings=warnings,
        fixes=fixes,
        clears=clears,
        mapping_notes=mapping_notes,
        unrecognized_columns=unrecognized,
        error_positions=error_positions,
        valid_positions=valid_positions,
        source_rows=len(df),
        source_columns=len(df.columns),
    )


def process_dataframe(df: pd.DataFrame, options: Optional[EngineOptions] = None, zip_cache: Optional[dict[str, dict[str, str]]] = None) -> ProcessResult:
    options = options or EngineOptions()
    zip_cache = zip_cache or {}
    try:
        df = df.fillna("").astype(str)
        if len(df) > options.max_file_rows:
            raise ValueError(f"Input has {len(df)} rows, exceeding max_file_rows={options.max_file_rows}.")
        if len(df.columns) > options.max_file_columns:
            raise ValueError(f"Input has {len(df.columns)} columns, exceeding max_file_columns={options.max_file_columns}.")

        mapping, unrecognized, mapping_notes = map_columns(df, options)
        horizontal, reason = detect_horizontal(df, mapping, options)
        if horizontal:
            converted, notes = convert_horizontal_to_vertical(df, mapping, options)
            result = validate_and_reformat(converted, options, zip_cache, already_canonical=True)
            result.notes.extend(notes)
            result.horizontal_used = True
            result.detect_reason = reason
            result.mapping_notes = mapping_notes
            result.unrecognized_columns = unrecognized
            result.source_rows = len(df)
            result.source_columns = len(df.columns)
            return result

        result = validate_and_reformat(df, options, zip_cache, already_canonical=False)
        result.horizontal_used = False
        result.detect_reason = reason
        return result
    except Exception:
        exception_text = traceback.format_exc()
        LOGGER.exception("Census processing failed.")
        return ProcessResult(
            reformatted_df=pd.DataFrame(columns=OUTPUT_COLUMNS),
            errors=[Issue("-", "-", "Engine", "", "Error", "Processing failed. Review exception text in diagnostics.")],
            source_rows=len(df) if isinstance(df, pd.DataFrame) else 0,
            source_columns=len(df.columns) if isinstance(df, pd.DataFrame) else 0,
            exception_text=exception_text,
        )


def read_census_file(uploaded_file: Any) -> pd.DataFrame:
    name = getattr(uploaded_file, "name", "uploaded.csv").lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file, dtype=str, keep_default_na=False).fillna("").astype(str)
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file, dtype=str, keep_default_na=False).fillna("").astype(str)
    raise ValueError("Unsupported file type. Upload a CSV, XLSX, or XLS file.")


def read_csv_text(text: str) -> pd.DataFrame:
    return pd.read_csv(io.StringIO(text), dtype=str, keep_default_na=False).fillna("").astype(str)


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


# =============================================================================
# Built-in sample censuses
# =============================================================================


SAMPLE_CENSUSES: dict[str, str] = {
    "01_standard_vertical_clean": """First Name,Last Name,Relationship,DOB,Zip Code,Primary Worksite Zip Code,Health Election,Current Health Plan Tier\nJohn,Smith,Employee,01/01/1980,53202,53202,Enroll,Family\nJane,Smith,Spouse,02/02/1982,53202,53202,Enroll,\nTim,Smith,Child,03/03/2012,53202,53202,Enroll,\n""",
    "02_vertical_informal_values": """first,last,rel,birthdate,zip,work zip,election,tier\nSam,Jones,ee,1980-01-15,601,601,e,ee only\nAlex,Jones,sp,1981-02-20,00601,00601,w,\n""",
    "03_wide_employee_only_with_dependents": """Employee First,Employee Last,Employee DOB,Home Zip,Worksite Zip,Medical Plan,Medical Tier,Spouse DOB,Child 1 DOB,Child 2 DOB\nChris,Lee,01/01/1975,60601,60601,Gold PPO,Employee Only,02/02/1977,03/03/2010,04/04/2012\n""",
    "04_wide_family": """Employee First Name,Employee Last Name,Employee DOB,Zip Code,Primary Worksite Zip Code,Current Medical Plan,Coverage Tier,Spouse First,Spouse Last,Spouse DOB,Child 1 First,Child 1 Last,Child 1 DOB\nPat,Miller,05/05/1985,77002,77002,Silver HMO,Fam,Riley,Miller,06/06/1986,Jordan,Miller,07/07/2015\n""",
    "05_wide_employee_spouse": """First Name,Last Name,DOB,Zip,Work Zip,Health Plan,Tier,Spouse DOB,Child 1 DOB\nAvery,Brown,01/01/1990,30301,30301,Medical PPO,EE+SP,02/02/1991,03/03/2020\n""",
    "06_wide_employee_children": """First Name,Last Name,DOB,Zip,Work Zip,Medical Plan,Medical Tier,Spouse DOB,Child 1 DOB,Child 2 DOB\nMorgan,Davis,01/01/1988,10001,10001,Medical EPO,EE+CH,02/02/1989,03/03/2011,04/04/2014\n""",
    "07_rowwise_horizontal": """Employee First,Employee Last,Employee DOB,Home Zip,Worksite Zip,Current Health Plan,Current Health Plan Tier,Dependent Relationship,Dependent DOB\nTaylor,Wilson,01/01/1980,53202,53202,Gold PPO,Family,Spouse,02/02/1981\nTaylor,Wilson,01/01/1980,53202,53202,Gold PPO,Family,Child,03/03/2011\n""",
    "08_missing_names_placeholders": """Employee DOB,Zip Code,Primary Worksite Zip Code,Medical Plan,Medical Tier,Spouse DOB,Child 1 DOB\n01/01/1970,75001,75001,Silver Plan,Family,02/02/1971,03/03/2010\n""",
    "09_ancillary_noise_ignored": """First Name,Last Name,DOB,Zip,Work Zip,Medical Plan,Medical Tier,Dental Tier,Vision Election,Life Amount,Spouse DOB\nJamie,Garcia,01/01/1980,33101,33101,Medical Plan,Fam,Employee Only,Waive,50000,02/02/1982\n""",
    "10_bad_zip_and_missing_dob": """First Name,Last Name,Relationship,DOB,Zip Code,Primary Worksite Zip Code,Health Election\nBad,Data,Employee,,WI,99999,Enroll\n""",
    "11_excel_serial_dates": """First Name,Last Name,Relationship,DOB,Zip Code,Primary Worksite Zip Code,Health Election\nSerial,Date,Employee,29221,53202,53202,E\n""",
    "12_zip_plus_four_and_currency": """First Name,Last Name,Relationship,DOB,Zip Code,Primary Worksite Zip Code,Health Election,Current Health Plan ER Cost,Current Health Plan EE Cost\nMoney,Case,Employee,01/01/1980,53202-1234,53202,$Enroll,$1,234.50,$250.00\n""".replace("$Enroll", "Enroll"),
    "13_ineligible_horizontal_skipped": """First Name,Last Name,DOB,Zip,Work Zip,Medical Plan,Tier,Status,Spouse DOB\nSkip,Intern,01/01/2000,53202,53202,Gold,Family,Intern,02/02/2000\nKeep,Employee,01/01/1980,53202,53202,Gold,Family,Active,02/02/1980\n""",
    "14_multiple_children_packed": """First Name,Last Name,DOB,Zip,Work Zip,Medical Plan,Tier,Child 1 DOB\nPack,Children,01/01/1980,60601,60601,Gold,Family,03/03/2010|04/04/2012;05/05/2014\n""",
    "15_no_tier_infer_dependents": """First Name,Last Name,DOB,Zip,Work Zip,Medical Plan,Spouse DOB,Child 1 DOB\nNo,Tier,01/01/1980,78701,78701,Medical PPO,02/02/1981,03/03/2012\n""",
    "16_waived_employee_dependents": """First Name,Last Name,DOB,Zip,Work Zip,Health Election,Medical Tier,Spouse DOB,Child 1 DOB\nWendy,Waive,01/01/1980,94105,94105,Waive,Family,02/02/1981,03/03/2010\n""",
    "17_duplicate_dependent_rows": """Employee First,Employee Last,Employee DOB,Home Zip,Worksite Zip,Current Health Plan,Current Health Plan Tier,Dependent Relationship,Dependent DOB\nDrew,Repeat,01/01/1980,53202,53202,Gold PPO,Family,Child,03/03/2011\nDrew,Repeat,01/01/1980,53202,53202,Gold PPO,Family,Child,03/03/2011\n""",
}


def run_self_tests(options: Optional[EngineOptions] = None) -> pd.DataFrame:
    options = options or EngineOptions(auto_fill_county_from_zip=False)
    rows: list[dict[str, Any]] = []
    for name, csv_text in SAMPLE_CENSUSES.items():
        df = read_csv_text(csv_text)
        result = process_dataframe(df, options=options, zip_cache={})
        rows.append({
            "Sample": name,
            "Source Rows": result.source_rows,
            "Output Rows": result.total_rows,
            "Errors": len(result.errors),
            "Warnings": len(result.warnings),
            "Horizontal Used": result.horizontal_used,
            "Detection Reason": result.detect_reason,
            "Status": "Review" if result.errors else "Pass",
        })
    return pd.DataFrame(rows)


# =============================================================================
# Streamlit app wrapper
# =============================================================================


def build_options_from_sidebar() -> EngineOptions:
    st.sidebar.header("Engine options")
    mode = st.sidebar.selectbox("Input mode", [m.value for m in InputMode], index=0)
    assumption_level = st.sidebar.slider("Assumption level", 0, 4, 2)
    auto_fill_county = st.sidebar.checkbox("Auto-fill county from zip", True)
    strict_health_only = st.sidebar.checkbox("Health-only horizontal conversion", True)
    infer_dependents = st.sidebar.checkbox("Infer dependent enrollment when tier is blank", True)
    placeholders = st.sidebar.checkbox("Create placeholder names when missing", True)
    inherit_worksite = st.sidebar.checkbox("Inherit worksite zip to dependents", True)
    exclude_ineligible = st.sidebar.checkbox("Exclude interns, contractors, inactive, and terminated rows", True)

    with st.sidebar.expander("Advanced controls", expanded=False):
        header_threshold = st.slider("Header match threshold", 70, 100, 88)
        relationship_threshold = st.slider("Relationship match threshold", 70, 100, 86)
        horizontal_sensitivity = st.slider("Horizontal detection sensitivity", 0, 5, 3)
        max_rows = st.number_input("Maximum rows", min_value=1, max_value=1_000_000, value=250_000, step=10_000)
        max_columns = st.number_input("Maximum columns", min_value=1, max_value=2_000, value=500, step=25)

    return EngineOptions(
        input_mode=InputMode(mode),
        assumption_level=assumption_level,
        header_match_threshold=header_threshold,
        relationship_match_threshold=relationship_threshold,
        horizontal_detection_sensitivity=horizontal_sensitivity,
        auto_fill_county_from_zip=auto_fill_county,
        strict_health_only=strict_health_only,
        infer_dependents_enrolled_when_tier_missing=infer_dependents,
        auto_placeholder_names=placeholders,
        inherit_worksite_zip_to_dependents=inherit_worksite,
        exclude_interns_contractors_inactive=exclude_ineligible,
        max_file_rows=int(max_rows),
        max_file_columns=int(max_columns),
    )


def render_metrics(result: ProcessResult) -> None:
    cols = st.columns(8)
    cols[0].metric("Source rows", result.source_rows)
    cols[1].metric("Source columns", result.source_columns)
    cols[2].metric("Output rows", result.total_rows)
    cols[3].metric("Valid rows", len(result.valid_positions))
    cols[4].metric("Error rows", len(result.error_positions))
    cols[5].metric("Errors", len(result.errors))
    cols[6].metric("Warnings", len(result.warnings))
    cols[7].metric("Fixes", len(result.fixes))
    status = "Review required" if result.has_errors else "Processed successfully"
    st.write(f"Status: {status}")
    st.write(f"Mode: {'Horizontal' if result.horizontal_used else 'Vertical'}")
    st.caption(result.detect_reason)


def render_downloads(result: ProcessResult) -> None:
    st.subheader("Downloads")
    c1, c2, c3, c4 = st.columns(4)
    c1.download_button("Full reformatted census", dataframe_to_csv_bytes(result.reformatted_df), "census_reformatted.csv", "text/csv", use_container_width=True)
    c2.download_button("Valid rows only", dataframe_to_csv_bytes(result.valid_df), "census_valid_rows.csv", "text/csv", use_container_width=True)
    c3.download_button("Error rows only", dataframe_to_csv_bytes(result.error_df), "census_error_rows.csv", "text/csv", use_container_width=True)
    c4.download_button("Issue report", dataframe_to_csv_bytes(result.issue_frame()), "census_issue_report.csv", "text/csv", use_container_width=True)


def render_result_tabs(result: ProcessResult) -> None:
    issues = result.issue_frame()
    tabs = st.tabs(["Reformatted output", "Issues", "Valid rows", "Error rows", "Diagnostics"])
    with tabs[0]:
        st.dataframe(result.reformatted_df, use_container_width=True, hide_index=True)
    with tabs[1]:
        if issues.empty:
            st.success("No issues reported.")
        else:
            kinds = sorted(issues["Kind"].dropna().astype(str).unique().tolist())
            selected = st.multiselect("Issue kinds", kinds, default=kinds)
            view = issues[issues["Kind"].astype(str).isin(selected)] if selected else issues
            st.dataframe(view, use_container_width=True, hide_index=True)
    with tabs[2]:
        st.dataframe(result.valid_df, use_container_width=True, hide_index=True)
    with tabs[3]:
        st.dataframe(result.error_df, use_container_width=True, hide_index=True)
    with tabs[4]:
        st.json(result.summary())
        if result.mapping_notes:
            st.write("Mapping notes")
            st.write(result.mapping_notes)
        if result.unrecognized_columns:
            st.write("Unrecognized columns")
            st.write(result.unrecognized_columns)
        if result.exception_text:
            st.code(result.exception_text)


def load_zip_cache_safely(options: EngineOptions) -> dict[str, dict[str, str]]:
    if not options.auto_fill_county_from_zip:
        return {}
    try:
        return build_zip_cache()
    except Exception:
        LOGGER.exception("Could not build zip cache.")
        st.warning("County auto-fill is unavailable. Processing will continue without zip lookup.")
        return {}


def main() -> None:
    st.set_page_config(page_title="CSA Census Validator", layout="wide")
    st.title("CSA Census Validator and Reformatter")
    st.caption(f"Version {APP_VERSION}. Single-file test app with deployable processing engine.")

    st.markdown(
        """
        <style>
            .block-container { padding-top: 1.5rem; max-width: 1280px; }
            [data-testid="metric-container"] { border: 1px solid #e5e7eb; border-radius: 8px; padding: 8px 12px; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    options = build_options_from_sidebar()

    left, right = st.columns([2, 1], gap="large")
    with left:
        uploaded = st.file_uploader("Upload census file", type=["csv", "xlsx", "xls"])
        pasted = st.text_area("Or paste CSV data", height=180, placeholder="Paste CSV text including a header row.")
    with right:
        sample_name = st.selectbox("Built-in sample census", [""] + sorted(SAMPLE_CENSUSES.keys()))
        sample_text = SAMPLE_CENSUSES.get(sample_name, "")
        if sample_text:
            st.text_area("Sample preview", sample_text, height=220, disabled=True)

    process_clicked = st.button("Process census", type="primary")
    if process_clicked:
        with st.spinner("Processing census"):
            try:
                if uploaded is not None:
                    df_input = read_census_file(uploaded)
                elif pasted.strip():
                    df_input = read_csv_text(pasted)
                elif sample_text.strip():
                    df_input = read_csv_text(sample_text)
                else:
                    st.warning("Upload a file, paste CSV data, or select a sample census.")
                    df_input = None
                if df_input is not None:
                    result = process_dataframe(df_input, options=options, zip_cache=load_zip_cache_safely(options))
                    st.session_state["last_result"] = result
            except Exception as exc:
                st.error(f"Could not read or process input: {exc}")
                with st.expander("Traceback", expanded=False):
                    st.code(traceback.format_exc())

    if "last_result" in st.session_state:
        result = st.session_state["last_result"]
        st.divider()
        render_metrics(result)
        st.divider()
        render_downloads(result)
        st.divider()
        render_result_tabs(result)

    st.divider()
    st.subheader("Built-in engine self-tests")
    if st.button("Run all sample tests"):
        st.session_state["self_tests"] = run_self_tests(options)
    if "self_tests" in st.session_state:
        tests = st.session_state["self_tests"]
        st.dataframe(tests, use_container_width=True, hide_index=True)
        st.download_button("Download self-test results", dataframe_to_csv_bytes(tests), "census_self_test_results.csv", "text/csv")


if __name__ == "__main__":
    main()
